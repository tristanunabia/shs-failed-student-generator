import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="SHS Class Record Processor", page_icon="📊", layout="wide")

# Navigation Sidebar
st.sidebar.title("📌 Navigation")
page = st.sidebar.radio("Go to:", ["⚠️ At-Risk Students (< 75) [Excel Only]", "📋 Full Consolidated Master List [CSV Only]"])

st.title("📊 Senior High School Class Record Processor")

# Helper function to check if name is valid or a placeholder/system text
def is_valid_student_name(name_str):
    if pd.isna(name_str):
        return False
    name_clean = str(name_str).strip()
    if name_clean == "":
        return False
    # Exclude if it explicitly starts with the word "Student" (case-insensitive) or matches system headers
    if name_clean.lower().startswith("student") or "student name" in name_clean.lower():
        return False
    return True

# ----------------- PAGE 1: AT RISK FILTER (EXCEL PIPELINE) -----------------
if page == "⚠️ At-Risk Students (< 75) [Excel Only]":
    st.subheader("📌 At-Risk Students Filter Tool (Upload `.xlsx` files)")
    uploaded_xlsx = st.file_uploader("Choose Class Record Excel files", type=["xlsx"], accept_multiple_files=True, key="xlsx_upload")
    
    if uploaded_xlsx:
        all_at_risk = []
        total_enrolled = 0
        
        for uploaded_file in uploaded_xlsx:
            try:
                xl = pd.ExcelFile(uploaded_file)
                sheet_names = xl.sheet_names
                settings_sheet = next((s for s in sheet_names if s.upper() == 'SETTINGS'), None)
                summary_sheet = next((s for s in sheet_names if s.upper() == 'SUMMARY'), None)
                
                if not settings_sheet or not summary_sheet:
                    continue
                
                # Parse Settings
                settings_df = pd.read_excel(uploaded_file, sheet_name=settings_sheet, header=None)
                section, subject_name = "Unknown Section", "Unknown Subject"
                for r_idx in range(len(settings_df)):
                    for c_idx in range(len(settings_df.columns)):
                        cell_val = str(settings_df.iloc[r_idx, c_idx]).strip()
                        if cell_val == "Section" and c_idx + 1 < len(settings_df.columns):
                            section = str(settings_df.iloc[r_idx, c_idx + 1]).strip()
                        if cell_val == "Subject" and c_idx + 1 < len(settings_df.columns):
                            subject_name = str(settings_df.iloc[r_idx, c_idx + 1]).strip()
                
                # Parse Summary
                summary_df = pd.read_excel(uploaded_file, sheet_name=summary_sheet, header=None)
                start_row, name_col, grade_col = 8, 2, 5
                for r_idx in range(min(15, len(summary_df))):
                    row_vals = [str(x).strip() for x in summary_df.iloc[r_idx].dropna()]
                    if "Student Name" in row_vals or "Stud No." in row_vals:
                        start_row = r_idx + 1
                        for c_i, c_v in enumerate(list(summary_df.iloc[r_idx])):
                            if str(c_v).strip() == "Student Name": name_col = c_i
                        break
                
                for r_idx in range(start_row, len(summary_df)):
                    student_name = summary_df.iloc[r_idx, name_col]
                    final_grade = summary_df.iloc[r_idx, grade_col]
                    
                    if is_valid_student_name(student_name):
                        total_enrolled += 1
                        numeric_grade = pd.to_numeric(final_grade, errors='coerce')
                        if pd.notna(numeric_grade) and numeric_grade < 75:
                            all_at_risk.append({
                                "Subject": subject_name,
                                "Section": section,
                                "Student Name": str(student_name).strip(),
                                "Final Grade": final_grade,
                                "Source File": uploaded_file.name
                            })
            except Exception as e:
                st.warning(f"Error parsing {uploaded_file.name}: {e}")
                
        if all_at_risk:
            at_risk_df = pd.DataFrame(all_at_risk)
            st.metric("Total Students Screened", total_enrolled)
            st.dataframe(at_risk_df, use_container_width=True)
            st.download_button("📥 Download At-Risk Students CSV", at_risk_df.to_csv(index=False).encode('utf-8'), "At_Risk_Students.csv", "text/csv")
        else:
            st.info("No students found with a final grade below 75.")

# ----------------- PAGE 2: FULL CONSOLIDATED EXPORT (CSV PIPELINE) -----------------
elif page == "📋 Full Consolidated Master List [CSV Only]":
    st.subheader("📌 Consolidated Master Record Generator (Upload `.csv` files)")
    st.markdown("Drop your exported CSV files here. You can upload pre-compiled summaries (like `Mr. Alifa.csv`) or raw worksheet pairs (like `SETTINGS.csv` and `Summary.csv`).")
    
    uploaded_csvs = st.file_uploader("Choose exported CSV files", type=["csv"], accept_multiple_files=True, key="csv_upload")
    
    if uploaded_csvs:
        all_consolidated_records = []
        
        # Temporary separation structures for raw sheets matching
        settings_map = {}
        summary_map = {}
        
        with st.spinner("Assembling master table components..."):
            for file in uploaded_csvs:
                try:
                    # Peek at headers to detect schema dynamically
                    file.seek(0)
                    peek_df = pd.read_csv(io.StringIO(file.getvalue().decode('utf-8', errors='ignore')), nrows=2)
                    file.seek(0)
                    
                    cols_lower = [str(c).lower().strip() for c in peek_df.columns]
                    
                    # ---- CASE A: File is an already structured summary list (e.g., Mr. Alifa.csv) ----
                    if "student name" in cols_lower and "final grade" in cols_lower:
                        full_df = pd.read_csv(file)
                        full_df.columns = [str(c).strip() for c in full_df.columns]
                        
                        col_mapping = {}
                        for c in full_df.columns:
                            c_low = c.lower()
                            if c_low == "subject": col_mapping[c] = "Subject"
                            elif c_low == "section": col_mapping[c] = "Section"
                            elif c_low == "student name": col_mapping[c] = "Student Name"
                            elif c_low == "final grade": col_mapping[c] = "Final Grade"
                        
                        filtered_summary = full_df[list(col_mapping.keys())].rename(columns=col_mapping)
                        
                        # Apply name exclusion filters to the structured lists
                        filtered_summary = filtered_summary[filtered_summary["Student Name"].apply(is_valid_student_name)]
                        
                        all_consolidated_records.extend(filtered_summary.to_dict(orient="records"))
                        
                    # ---- CASE B: File is a raw template worksheet segment export ----
                    else:
                        fname = file.name
                        root_match = re.split(r'\s-\s', fname, maxsplit=1)
                        root_name = root_match[0] if root_match else fname
                        
                        if "SETTINGS" in fname.upper():
                            settings_map[root_name] = file
                        elif "SUMMARY" in fname.upper():
                            summary_map[root_name] = file
                except Exception as e:
                    st.error(f"❌ Error diagnosing file format for '{file.name}': {e}")
            
            # Process paired raw sheets matching identical root naming patterns
            for root, summary_file in summary_map.items():
                settings_file = settings_map.get(root)
                
                if settings_file is None:
                    st.warning(f"⚠️ Skipped un-paired dataset '{root}': Missing its matching SETTINGS file.")
                    continue
                
                try:
                    # 1. Parse Raw CSV Settings
                    settings_file.seek(0)
                    settings_df = pd.read_csv(settings_file, header=None)
                    section, subject_name = "Unknown Section", "Unknown Subject"
                    
                    for r_idx in range(len(settings_df)):
                        for c_idx in range(len(settings_df.columns)):
                            cell_val = str(settings_df.iloc[r_idx, c_idx]).strip()
                            if cell_val == "Section" and c_idx + 1 < len(settings_df.columns):
                                section = str(settings_df.iloc[r_idx, c_idx + 1]).strip()
                            if cell_val == "Subject" and c_idx + 1 < len(settings_df.columns):
                                subject_name = str(settings_df.iloc[r_idx, c_idx + 1]).strip()
                    
                    # 2. Parse Raw CSV Summary
                    summary_file.seek(0)
                    summary_df = pd.read_csv(summary_file, header=None)
                    start_row, name_col, grade_col = 8, 2, 5
                    
                    for r_idx in range(min(15, len(summary_df))):
                        row_vals = [str(x).strip() for x in summary_df.iloc[r_idx].dropna()]
                        if "Student Name" in row_vals or "Stud No." in row_vals:
                            start_row = r_idx + 1
                            for c_i, c_v in enumerate(list(summary_df.iloc[r_idx])):
                                if str(c_v).strip() == "Student Name": name_col = c_i
                            break
                    
                    # 3. Extract Rows
                    for r_idx in range(start_row, len(summary_df)):
                        student_name = summary_df.iloc[r_idx, name_col]
                        final_grade = summary_df.iloc[r_idx, grade_col]
                        
                        if is_valid_student_name(student_name):
                            all_consolidated_records.append({
                                "Subject": subject_name,
                                "Section": section,
                                "Student Name": str(student_name).strip(),
                                "Final Grade": final_grade
                            })
                except Exception as e:
                    st.error(f"❌ Error compiling raw data pairs for '{root}': {str(e)}")
                    
        # Render Unified Tabular Display
        if all_consolidated_records:
            consolidated_df = pd.DataFrame(all_consolidated_records)
            
            # Enforce clean layout order (Strictly excluding Source File)
            consolidated_df = consolidated_df[["Subject", "Section", "Student Name", "Final Grade"]]
            
            st.success(f"📊 Successfully extracted and compiled {len(consolidated_df)} records.")
            st.metric("Total Rows Extracted", len(consolidated_df))
            
            # Present data to user in a clean tabular grid form
            st.dataframe(consolidated_df, use_container_width=True)
            
            # Download trigger
            csv_bytes = consolidated_df.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 Download Consolidated Master CSV",
                data=csv_bytes,
                file_name="Master_Consolidated_Class_Records.csv",
                mime="text/csv",
                key="download-master-csv"
            )
        else:
            st.info("No valid record entries could be parsed. Ensure your files are uploaded correctly and contain valid student names.")
else:
    st.info("👈 Please select a processing view page from the navigation sidebar menu.")