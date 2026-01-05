import streamlit as st
import pandas as pd
import google.generativeai as genai
from docx import Document
from docx.shared import Inches
import io
import xlsxwriter

# 1. SETUP & SECRETS
st.set_page_config(page_title="Energy Data Pro", layout="wide")

# Safe API Key retrieval
api_key = st.secrets.get("GEMINI_API_KEY")
if api_key and api_key != "Your_Gemini_API_Key_Here":
    try:
        genai.configure(api_key=api_key)
    except Exception as e:
        st.error(f"Error configuring Gemini API: {e}")
else:
    st.warning("⚠️ Gemini API Key not found or default value in `.streamlit/secrets.toml`. AI reporting will be disabled.")


st.title("⚡ Solar Site Load Analyzer")
st.info("Upload your logger data (Seconds, Minutes, or Hourly) to generate a formula-based analysis.")

# 2. FILE UPLOADER (Main Page)
uploaded_file = st.file_uploader("Choose Logger File", type=['csv', 'xlsx'])

if uploaded_file:
    try:
        # Load and clean
        if uploaded_file.name.lower().endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
        
        # Auto-detection of columns
        df.columns = [str(c).lower().strip() for c in df.columns]
        time_col = next((c for c in df.columns if 'time' in c or 'date' in c or 'timestamp' in c), None)
        watt_col = next((c for c in df.columns if 'watt' in c or 'w' == c or 'power' in c), None) # Added 'power' just in case

        if time_col and watt_col:
            # Standardize Time with Robust Parsing
            # Hybrid Approach: Check if it looks like Numbers (Unix) or Strings (Dates)
            
            # 1. Try to convert to numeric, but keep original if it fails completely
            df_temp_numeric = pd.to_numeric(df[time_col], errors='coerce')
            
            # Check if we have valid numbers (at least 50% valid to be safe, or just non-empty)
            valid_numeric_count = df_temp_numeric.count()
            total_count = len(df)
            
            if valid_numeric_count > 0 and (valid_numeric_count / total_count > 0.5):
                # CASE A: It is numeric (Unix Timestamps)
                # Use the numeric version
                df[time_col] = df_temp_numeric.dropna()
                df = df.dropna(subset=[time_col]) # Drop invalid rows
                
                try:
                    first_val = df[time_col].iloc[0]
                    # Unix Timestamp Heuristics
                    if first_val > 1e12: # Milliseconds (13 digits)
                         df[time_col] = pd.to_datetime(df[time_col], unit='ms')
                    elif first_val > 5e9: # Deciseconds (11 digits, ~1.7e10) -> Convert to Seconds
                         df[time_col] = pd.to_datetime(df[time_col] / 10, unit='s')
                    else: # Seconds (10 digits)
                         df[time_col] = pd.to_datetime(df[time_col], unit='s')
                except Exception as e:
                    st.warning(f"Numeric timestamp error: {e}. Fallback to standard.")
                    df[time_col] = pd.to_datetime(df[time_col])
            else:
                # CASE B: It is likely String Dates (e.g. "2024-01-01 12:00")
                # Use standard string parsing on the ORIGINAL column
                # Attempt to parse 'dayfirst' for international formats like DD/MM/YYYY
                try:
                     df[time_col] = pd.to_datetime(df[time_col], dayfirst=True)
                except:
                     df[time_col] = pd.to_datetime(df[time_col]) # Fallback
            
            # Clean invalid dates
            df = df.dropna(subset=[time_col])

            df = df.sort_values(time_col)
            
            # Ensure Watt/Power column is Numeric (handle strings/errors)
            df[watt_col] = pd.to_numeric(df[watt_col], errors='coerce')
            df = df.dropna(subset=[watt_col]) # Drop rows where Power is invalid

            
            # Robust Interval Detection
            if len(df) > 1:
                # Calculate median diff in minutes
                interval_min = df[time_col].diff().median().total_seconds() / 60
                # Fallback if diff is 0 or NaN for some reason
                if pd.isna(interval_min) or interval_min <= 0:
                     interval_min = 1.0 
            else:
                st.warning("Single row detected. Defaulting to 1-minute interval.")
                interval_min = 1.0
            
            # 3. DYNAMIC ANALYSIS LOGIC based on FREQUENCY
            # Logic: 
            # - If Interval < 1.0 minute (High Frequency, e.g. Pirano), use AVERAGE.
            # - If Interval >= 1.0 minute (Standard/Low Frequency, e.g. Denmark), use SUM (User Reference Style).
            
            if interval_min < 1.0:
                analysis_method = "Time-Weighted Average"
                # Use MEAN for high frequency
            if interval_min < 1.0:
                analysis_method = "Time-Weighted Average"
                # Use MEAN for high frequency
                hourly_df = df.set_index(time_col).resample('H')[watt_col].mean().reset_index()
                hourly_df['Power (kW)'] = hourly_df[watt_col] / 1000
                col_name_w = "Avg Power (W)"
                col_name_kw = "Avg Power (kW)"
                st.info(f"ℹ️ **Smart Analysis**: Detected High Frequency data ({interval_min:.2f} min). Using **{analysis_method}**.")
            else:
                analysis_method = "Hourly Summation"
                # Use SUM for standard/low frequency using min_count=1 to detect Gaps vs Zeros
                hourly_df = df.set_index(time_col).resample('H')[watt_col].sum(min_count=1).reset_index()
                hourly_df['Power (kW)'] = hourly_df[watt_col] / 1000
                col_name_w = "Power (Sum W)"
                col_name_kw = "Power (Sum kW)"
                st.success(f"✅ **Smart Analysis**: Detected Standard data ({interval_min:.2f} min). Using **{analysis_method}**.")

            # Filter out empty hours (NaNs) created by resampling gaps
            before_len = len(hourly_df)
            hourly_df = hourly_df.dropna(subset=[watt_col])
            after_len = len(hourly_df)
            
            if before_len - after_len > 100:
                st.warning(f"⚠️ Note: Removed {before_len - after_len} empty hours (gaps in data) to optimize report.")
            
            if len(hourly_df) > 8760: # More than 1 year of hours
                st.warning(f"⚠️ Large Data Detected: {len(hourly_df)} hours. Excel generation may take a moment.")


            # Common Columns
            hourly_df['Day'] = hourly_df[time_col].dt.day
            hourly_df['Hour'] = hourly_df[time_col].dt.hour
            hourly_df['Power (W)'] = hourly_df[watt_col] # Raw aggregated value (Sum or Mean)
            
            # Display Basic Stats
            peak_val_kw = hourly_df['Power (kW)'].max()
            st.metric("Detected Interval", f"{interval_min:.2f} min")
            st.metric(f"Peak Hourly Load ({analysis_method})", f"{peak_val_kw:.2f} kW")

            # 4. GENERATE EXCEL WITH FORMULAS (Reference Style)
            output_excel = io.BytesIO()
            workbook = xlsxwriter.Workbook(output_excel)
            
            # Formats
            bold = workbook.add_format({'bold': True})
            header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D3D3D3', 'border': 1})
            num_fmt = workbook.add_format({'num_format': '#,##0.00'})
            # IMPORTANT: Date format for Formula referencing
            date_fmt = workbook.add_format({'num_format': 'yyyy-mm-dd hh:mm:ss'})
            
            # --- Sheet 1: Raw Data ---
            worksheet = workbook.add_worksheet("Raw Data")
            start_row = 0
            headers = ["Localtime", "Raw Watt", "Power (kW)", "Interval Energy (kWh)"]
            for col, header in enumerate(headers):
                worksheet.write(start_row, col, header, header_fmt)
            
            # Progress Bar for Large Files
            progress_text = "Generating Excel Formulas..."
            my_bar = st.progress(0, text=progress_text)
            total_rows = len(df)
            
            # Write Raw Data
            for i, (ts_val, watt_val) in enumerate(zip(df[time_col], df[watt_col])):
                excel_row = start_row + 1 + i
                row_str = str(excel_row + 1)
                
                # Write datetime object (essential for formulas)
                worksheet.write_datetime(excel_row, 0, ts_val, date_fmt)
                worksheet.write_number(excel_row, 1, watt_val, num_fmt)
                worksheet.write_formula(excel_row, 2, f"=B{row_str}/1000", num_fmt)
                worksheet.write_formula(excel_row, 3, f"=C{row_str}*({interval_min}/60)", num_fmt)
                
                # Update progress every 5%
                step_size = max(1, total_rows // 20)
                if i % step_size == 0:
                     # Scale Raw Data phase to 0-80% of total bar
                     progress_percent = min(80, int((i / total_rows) * 80))
                     my_bar.progress(progress_percent, text=f"{progress_text} ({progress_percent}%)")
            
            worksheet.set_column(0, 0, 22)
            worksheet.set_column(1, 3, 15)

            # --- Sheet 2: Analyzed (with Formulas referencing Sheet 1) ---
            my_bar.progress(85, text="Generating Analysis Sheet...")
            
            ws_analyzed = workbook.add_worksheet("Analyzed")
            analyzed_headers = ["Date & Time", "Day", "Hour", col_name_w, col_name_kw]
            
            for col, header in enumerate(analyzed_headers):
                ws_analyzed.write(0, col, header, header_fmt)
            
            # We still rely on hourly_df for structure/indexes, but values will be formulas
            num_rows = len(hourly_df)
            
            for i, (ts, day, hour) in enumerate(zip(hourly_df[time_col], hourly_df['Day'], hourly_df['Hour'])):
                row_idx = i + 1
                row_str = str(row_idx + 1)
                
                # Write Time bucket (as datetime)
                ws_analyzed.write_datetime(row_idx, 0, ts, date_fmt)
                ws_analyzed.write_number(row_idx, 1, day)
                ws_analyzed.write_number(row_idx, 2, hour)
                
                # FORMULA LOGIC based on Analysis Method
                range_time = "'Raw Data'!$A:$A"
                range_watt = "'Raw Data'!$B:$B"
                
                start_crit = f'">="&A{row_str}'  # >= Hourly Timestamp
                end_crit = f'"<"&A{row_str}+(1/24)' # < Hourly Timestamp + 1 Hour
                
                if analysis_method == "Hourly Summation":
                    formula = f'=SUMIFS({range_watt}, {range_time}, {start_crit}, {range_time}, {end_crit})'
                else:
                    formula = f'=AVERAGEIFS({range_watt}, {range_time}, {start_crit}, {range_time}, {end_crit})'
                
                ws_analyzed.write_formula(row_idx, 3, formula, num_fmt)
                ws_analyzed.write_formula(row_idx, 4, f"=D{row_str}/1000", num_fmt)
            
            ws_analyzed.set_column(0, 0, 22)
            ws_analyzed.set_column(1, 4, 18)

            # --- Add Native Excel Chart ---
            my_bar.progress(90, text="Adding Charts & Finalizing...")
            
            if num_rows > 0:
                try:
                    chart = workbook.add_chart({'type': 'line'})
                    chart.add_series({
                        'name':       f'Hourly Load ({analysis_method})',
                        'categories': ['Analyzed', 1, 0, num_rows, 0], # Timestamp column
                        'values':     ['Analyzed', 1, 4, num_rows, 4], # Power kW column
                        'line':       {'color': '#008000', 'width': 2.25},
                    })
                    chart.set_title({'name': 'Site Load Profile'})
                    chart.set_x_axis({'name': 'Time', 'date_axis': True})
                    chart.set_y_axis({'name': 'Power (kW)'})
                    ws_analyzed.insert_chart('G2', chart)
                except Exception as e:
                    st.warning(f"Could not add chart to Excel: {e}")

            try:
                workbook.close()
                my_bar.progress(100, text="Excel Generation Complete!")
            except Exception as e:
                st.error(f"Error saving Excel file: {e}")
            
            # 5. UI DISPLAY & DOWNLOADS
            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    label="📥 Download Analyzed Excel",
                    data=output_excel.getvalue(),
                    file_name="Analyzed_Energy_Data.xlsx",
                    mime="application/vnd.ms-excel"

                )
                
            # 6. AI REPORT TRIGGER
            if st.button("Generate Senior Engineer Report (AI)"):
                if not api_key or api_key == "Your_Gemini_API_Key_Here":
                     st.error("Please configure your Gemini API Key in .streamlit/secrets.toml to use this feature.")
                else:
                    with st.spinner("Generating technical report..."):
                        # --- CALCULATE METRICS FOR SUMMARY TABLE ---
                        peak_row = hourly_df.loc[hourly_df['Power (kW)'].idxmax()]
                        peak_val = peak_row['Power (kW)']
                        peak_time = peak_row[time_col]
                        
                        # Base Load (Global 5th Percentile for robustness)
                        base_load_kw = hourly_df['Power (kW)'].quantile(0.05)
                        
                        # Power Factor
                        pf_col = next((c for c in df.columns if 'pf' in c or 'factor' in c or 'cos' in c), None)
                        if pf_col:
                            avg_pf = df[pf_col].mean()
                            pf_str = f"{avg_pf:.2f}"
                        else:
                            pf_str = "N/A (Not Logged)"

                        # Prepare AI Prompt (Text Only)
                        prompt = f"""
                        Act as a Senior Solar Engineer. Write a technical site assessment report text.
                        
                        **Data Context**:
                        - Max Peak: {peak_val:.2f} kW
                        - Base Load: {base_load_kw:.2f} kW
                        - Analysis Type: {analysis_method}
                        
                        **Instructions**:
                        1. Write STRICTLY in plain text paragraphs. No Markdown.
                        2. **Executive Technical Summary**: Start with a dense summary of the load characteristics.
                        3. **Profile Analysis**: Describe the daily usage pattern (Day vs Night, Peak Timing).
                        4. **Recommendations**: Implication for storage (Base Load coverage) and Inverter sizing.
                        
                        (Do NOT generate a table.)
                        """
                        
                        try:
                            # Generate Content (Fallback Safe)
                            try:
                                model = genai.GenerativeModel('gemini-2.5-flash')
                                response = model.generate_content(prompt)
                            except:
                                st.warning("⚠️ Falling back to gemini-1.5-flash.")
                                model = genai.GenerativeModel('gemini-1.5-flash')
                                response = model.generate_content(prompt)
                            
                            # --- BUILD WORD DOC ---
                            doc = Document()
                            doc.add_heading('Technical Solar Site Assessment', 0)
                            
                            # 1. SUMMARY TABLE
                            table = doc.add_table(rows=1, cols=2)
                            table.style = 'Table Grid'
                            metrics = [
                                ("Peak Load", f"{peak_val:.2f} kW"),
                                ("Peak Timestamp", str(peak_time)),
                                ("Base Load", f"{base_load_kw:.2f} kW"),
                                ("Avg Power Factor", pf_str),
                                ("Analysis Method", analysis_method)
                            ]
                            for m, v in metrics:
                                row = table.add_row().cells
                                row[0].text = m
                                row[1].text = v
                            
                            doc.add_paragraph()
                            
                            # 2. REPORT TEXT
                            doc.add_heading('Technical Report', level=1)
                            doc.add_paragraph(response.text)
                            
                            # 3. ADVANCED CHARTS
                            doc.add_heading('Load Profile Analysis', level=1)
                            
                            import matplotlib.pyplot as plt
                            import seaborn as sns
                            
                            # Prepare Data for Plotting
                            hourly_df['DayType'] = hourly_df[time_col].dt.dayofweek.apply(lambda x: 'Weekend' if x >= 5 else 'Weekday')
                            
                            sns.set_theme(style="whitegrid")
                            
                            # Determine Subplots
                            # Show split ONLY if both Weekdays and Weekends are present in the data
                            unique_day_types = hourly_df['DayType'].unique()
                            show_split = 'Weekday' in unique_day_types and 'Weekend' in unique_day_types
                            
                            nrows = 3 if show_split else 2
                            fig, ax = plt.subplots(nrows, 1, figsize=(10, 5 * nrows))
                            
                            # Chart 1: Average Daily Profile (0-23h)
                            sns.lineplot(data=hourly_df, x='Hour', y='Power (kW)', estimator='mean', errorbar=None, color='#1f77b4', linewidth=3, ax=ax[0])
                            ax[0].axhline(y=base_load_kw, color='r', linestyle='--', label=f'Base Load ({base_load_kw:.2f} kW)')
                            ax[0].set_title("Average Daily Load Profile (0-23h)", fontweight='bold', fontsize=12)
                            ax[0].set_xticks(range(0, 24))
                            ax[0].legend()
                            ax[0].set_xlabel("Hour of Day")
                            
                            # Chart 2: Hourly Variability (Box Plot)
                            sns.boxplot(data=hourly_df, x='Hour', y='Power (kW)', palette="viridis", ax=ax[1])
                            ax[1].set_title("Hourly Variability (Box Plot)", fontweight='bold', fontsize=12)
                            ax[1].set_xlabel("Hour of Day")
                            
                            # Chart 3: Weekday vs Weekend (Conditional)
                            if show_split:
                                sns.lineplot(data=hourly_df, x='Hour', y='Power (kW)', hue='DayType', estimator='mean', errorbar=None, palette={'Weekday':'#2ca02c', 'Weekend':'#ff7f0e'}, linewidth=2.5, ax=ax[2])
                                ax[2].set_title("Weekday vs. Weekend Profile", fontweight='bold', fontsize=12)
                                ax[2].set_xticks(range(0, 24))
                                ax[2].set_xlabel("Hour of Day")
                            
                            plt.tight_layout()
                            
                            img_io = io.BytesIO()
                            plt.savefig(img_io, format='png', dpi=150)
                            img_io.seek(0)
                            plt.close()
                            
                            doc.add_picture(img_io, width=Inches(6.0))
                            
                            word_io = io.BytesIO()
                            doc.save(word_io)
                            
                            st.download_button(label="📥 Download Technical Report", data=word_io.getvalue(), file_name="Site_Assessment_Report.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

                        except Exception as e:
                            st.error(f"AI Generation Failed: {e}")

        else:
            st.error("Could not automatically identify 'Timestamp' and 'Watt' columns. Please ensure your file has identifiable headers.")
            st.write("Columns found:", df.columns.tolist())

    except Exception as e:
        st.error(f"Error processing file: {e}")
