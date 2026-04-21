import streamlit as st
import base64
from google import genai
from google.genai import types
from PIL import Image
import io
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
import re
from typing import List, Dict
import tempfile
import os
import pdfplumber

# -------------------------
# Streamlit App
# -------------------------

def extract_bold_key_value(line: str):
    """
    Extracts **Key:** Value safely from a markdown line.
    Returns (key, value) or (None, None)
    """
    match = re.search(r"\*\*(.+?)\*\*:\s*(.+)", line)
    if match:
        return match.group(1).strip(), match.group(2).strip()
    return None, None

def parse_markdown_output(md_text: str):
    summary = {}
    standard = {}
    rows = []

    # -----------------------------
    # 1. Test Summary
    # -----------------------------
    summary_section = re.search(
        r"### 1\. Test Summary Information(.*?)---",
        md_text,
        re.S
    )

    if summary_section:
        for line in summary_section.group(1).splitlines():
            key, val = extract_bold_key_value(line)
            if key:
                summary[key] = val

    # -----------------------------
    # 2. Standard Block
    # -----------------------------
    std_section = re.search(
        r"### 2\. Verification with Standard Block(.*?)---",
        md_text,
        re.S
    )

    if std_section:
        for line in std_section.group(1).splitlines():
            key, val = extract_bold_key_value(line)
            if key:
                standard[key] = val

    # -----------------------------
    # 3. Hardness Table
    # -----------------------------
    table_match = re.search(
        r"\| Sr\. No\..*",
        md_text,
        re.S
    )

    if table_match:
        lines = table_match.group(0).splitlines()

        # Skip header + separator
        data_lines = [
            l for l in lines
            if l.startswith("|") and not "---" in l
        ][1:]

        for line in data_lines:
            cols = [c.strip() for c in line.strip("|").split("|")]

            if len(cols) < 7:
                continue

            # Check if test was conducted (has values)
            base_vals = [int(x.strip()) for x in cols[3].split(",") if x.strip().isdigit()]
            haz_vals = [int(x.strip()) for x in cols[4].split(",") if x.strip().isdigit()]
            weld_vals = [int(x.strip()) for x in cols[5].split(",") if x.strip().isdigit()]
            test_conducted = bool(base_vals or haz_vals or weld_vals)

            # CORRECT MAPPING for OBV sheet:
            # cols[1] = PIPE NO (starts with E or N like E26001361)
            # cols[2] = HEAT NO (plain digits like 2601253)

            pipe_no = cols[1]
            if pipe_no.startswith('E'):
                pipe_no = 'N' + pipe_no[1:]

            rows.append({
                "sample_id": pipe_no,   # PIPE NO (E26001361) - this is what gets validated
                "heat_no": cols[2],     # HEAT NO (2601253) - this is the key for lookup
                "base": base_vals,
                "haz": haz_vals,
                "weld": weld_vals,
                "remarks": cols[6] if len(cols) > 6 else "",
                "test_conducted": test_conducted
            })

    return summary, standard, rows

def populate_excel_from_markdown(md_text, template_path, output_path):
    summary, standard, samples = parse_markdown_output(md_text)

    wb = load_workbook(template_path)
    ws = wb.active

    # -----------------------------
    # Map summary fields
    # -----------------------------
    summary_map = {
        "Testing Laboratory": "B5",
        "Document Title": "B6",
        "Format No.": "B7",
        "Specification & Grade": "B8",
        "Test Method": "B9",
        "Pipe Size": "B12",
        "Atmospheric Conditions": "B13",
        "Date & Shift": "B14",
        "Requirements": "B15",
        "M/C No.": "B16",
    }

    for k, cell in summary_map.items():
        if k in summary:
            ws[cell] = summary[k]

    # -----------------------------
    # Standard Block
    # -----------------------------
    std_map = {
        "Standard Block ID No.": "E5",
        "Standard Block Value": "E6",
        "Reading 1": "E7",
        "Reading 2": "E8",
        "Reading 3": "E9",
        "Reading 4": "E10",
        "Reading 5": "E11",
        "Average (AVG)": "E12",
        "% Of Error": "E13",
        "Remark": "E14",
    }

    for k, cell in std_map.items():
        if k in standard:
            ws[cell] = standard[k]

    # -----------------------------
    # Hardness Table
    # -----------------------------
    START_ROW = 17

    COL_SRNO   = 1   # Column A
    COL_HEAT   = 2   # Column B
    COL_SAMPLE = 3  # Column C
    COL_BASE   = 4   # Column D (Points 1–6)
    COL_HAZ    = 10  # Column J (Points 7–24)
    COL_WELD   = 28  # Column AB (Points 25–33)
    COL_REMARK = 37  # Column AK

    for idx, s in enumerate(samples):
        r = START_ROW + idx
        # Sr. No.
        ws.cell(r, COL_SRNO, idx + 1)

        # IDs
        ws.cell(r, COL_SAMPLE, s["sample_id"])
        ws.cell(r, COL_HEAT, s["heat_no"])

        # Base (1–6)
        for i, v in enumerate(s["base"]):
            ws.cell(r, COL_BASE + i, v)

        # HAZ (7–24)
        for i, v in enumerate(s["haz"]):
            ws.cell(r, COL_HAZ + i, v)

        # Weld (25–33)
        for i, v in enumerate(s["weld"]):
            ws.cell(r, COL_WELD + i, v)

        # Remarks
        ws.cell(r, COL_REMARK, s["remarks"])

    wb.save(output_path)

def create_validation_excel(samples: List[Dict], master_map: Dict, output_path: str):
    """
    Create validation Excel file showing row-wise test results
    """
    results_data = []
    
    for sample in samples:
        heat_no = sample.get('heat_no', '').strip()
        pipe_used = sample.get('sample_id', '').strip()
        test_conducted = sample.get('test_conducted', False)
        
        if not test_conducted:
            results_data.append({
                "Heat No": heat_no,
                "Pipe No Used in Test": pipe_used,
                "Expected Pipe No (from Master)": "N/A",
                "Status": "❌ TEST NOT CONDUCTED",
                "Result": f"Test not conducted for Heat No {heat_no}"
            })
            continue
        
        if heat_no and pipe_used:
            if heat_no in master_map:
                expected_pipe = master_map[heat_no]
                if pipe_used == expected_pipe:
                    results_data.append({
                        "Heat No": heat_no,
                        "Pipe No Used in Test": pipe_used,
                        "Expected Pipe No (from Master)": expected_pipe,
                        "Status": "✅ MATCHED",
                        "Result": f"Test conducted using Pipe No: {pipe_used}"
                    })
                else:
                    results_data.append({
                        "Heat No": heat_no,
                        "Pipe No Used in Test": pipe_used,
                        "Expected Pipe No (from Master)": expected_pipe,
                        "Status": "⚠️ MISMATCH",
                        "Result": f"Pipe mismatch! Test used {pipe_used} but master shows {expected_pipe}"
                    })
            else:
                results_data.append({
                    "Heat No": heat_no,
                    "Pipe No Used in Test": pipe_used,
                    "Expected Pipe No (from Master)": "NOT FOUND",
                    "Status": "❌ NOT IN MASTER",
                    "Result": f"Heat No {heat_no} not found in master list"
                })
    
    # Create DataFrame
    df = pd.DataFrame(results_data)
    
    # Create Excel with formatting
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Validation Results', index=False)
        
        # Get the worksheet
        worksheet = writer.sheets['Validation Results']
        
        # Adjust column widths
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            worksheet.column_dimensions[column_letter].width = adjusted_width
        
        # Add color coding
        green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        yellow_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
        
        for row in worksheet.iter_rows(min_row=2, max_row=len(results_data) + 1):
            status_cell = row[3]  # Status column
            if status_cell.value == "✅ MATCHED":
                status_cell.fill = green_fill
            elif status_cell.value == "❌ NOT IN MASTER":
                status_cell.fill = red_fill
            elif status_cell.value == "⚠️ MISMATCH":
                status_cell.fill = yellow_fill

def validate_with_master_fast(samples: List[Dict], master_map: Dict) -> pd.DataFrame:
    """Fast validation using dictionary lookup"""
    results = []
    
    for sample in samples:
        if not sample.get("test_conducted", False):
            continue
            
        heat_no = sample.get('heat_no', '').strip()
        pipe_used = sample.get('sample_id', '').strip()
        
        if heat_no and pipe_used:
            if heat_no in master_map:
                expected_pipe = master_map[heat_no]
                if pipe_used == expected_pipe:
                    results.append({
                        "Heat No": heat_no,
                        "Pipe No Used": pipe_used,
                        "Status": "✅ MATCHED"
                    })
                else:
                    results.append({
                        "Heat No": heat_no,
                        "Pipe No Used": pipe_used,
                        "Expected Pipe": expected_pipe,
                        "Status": "⚠️ MISMATCH"
                    })
            else:
                results.append({
                    "Heat No": heat_no,
                    "Pipe No Used": pipe_used,
                    "Status": "❌ NOT IN MASTER"
                })
    
    return pd.DataFrame(results)

# -------------------------
# Initialize session state for caching
# -------------------------
if 'extracted_text' not in st.session_state:
    st.session_state.extracted_text = None
if 'samples' not in st.session_state:
    st.session_state.samples = None
if 'master_map' not in st.session_state:
    st.session_state.master_map = None
if 'master_list_grouped' not in st.session_state:
    st.session_state.master_list_grouped = None
if 'master_heat_to_pipes' not in st.session_state:
    st.session_state.master_heat_to_pipes = None
if 'excel_ready' not in st.session_state:
    st.session_state.excel_ready = False
if 'validation_excel_ready' not in st.session_state:
    st.session_state.validation_excel_ready = False

# -------------------------
# Streamlit UI
# -------------------------
st.set_page_config(page_title="Vickers Hardness Sheet Extractor & Validator", layout="wide")
st.title("Vickers Hardness Observation Sheet Extractor & Validator")

# Sidebar for validation option
st.sidebar.header("Validation Options")
enable_validation = st.sidebar.checkbox("Enable PDF Master Validation", value=False)

if enable_validation:
    st.sidebar.markdown("""
    **Upload Master File** containing:
    - Pipe No (starts with N)
    - Heat No (plain digits)
    
    This will validate tested samples against master list.
    """)
    
    # Choose file type
    master_file_type = st.sidebar.radio("Master file type:", ["PDF", "Excel"])
    
    if master_file_type == "PDF":
        master_file = st.sidebar.file_uploader("Upload Master PDF", type=["pdf"], key="master_pdf")
    else:
        master_file = st.sidebar.file_uploader("Upload Master Excel", type=["xlsx", "xls"], key="master_excel")
    
    if master_file and st.session_state.master_map is None:
        with st.sidebar.status("Loading master data..."):
            if master_file_type == "PDF":
                # PDF processing
                with pdfplumber.open(master_file) as pdf:
                    all_text = ""
                    for page in pdf.pages:
                        all_text += page.extract_text()
                
                master_map = {}
                heat_to_pipes = {}
                lines = all_text.split('\n')
                for line in lines:
                    pipe_match = re.search(r'([NE]\d{8,9})', line)
                    heat_matches = re.findall(r'\b(\d{7,10}(?:/\d{7,10})*)\b', line)
                    
                    if pipe_match and heat_matches:
                        pipe_no = pipe_match.group(1)
                        for heat_no_raw in heat_matches:
                            if '/' in heat_no_raw:
                                for h in heat_no_raw.split('/'):
                                    h = h.strip()
                                    master_map[h] = pipe_no
                                    if h not in heat_to_pipes:
                                        heat_to_pipes[h] = []
                                    if pipe_no not in heat_to_pipes[h]:
                                        heat_to_pipes[h].append(pipe_no)
                            else:
                                heat_no = heat_no_raw.strip()
                                master_map[heat_no] = pipe_no
                                if heat_no not in heat_to_pipes:
                                    heat_to_pipes[heat_no] = []
                                if pipe_no not in heat_to_pipes[heat_no]:
                                    heat_to_pipes[heat_no].append(pipe_no)
            else:
                # Excel processing
                try:
                    # Save uploaded file temporarily
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
                        tmp_file.write(master_file.getvalue())
                        tmp_path = tmp_file.name
                    
                    # Try to read directly first
                    try:
                        df_raw = pd.read_excel(tmp_path, header=None)
                    except:
                        # Fix corrupted XML if needed
                        import zipfile, shutil
                        with zipfile.ZipFile(tmp_path, 'r') as zip_ref:
                            zip_ref.extractall("temp_excel")
                        
                        workbook_path = "temp_excel/xl/workbook.xml"
                        with open(workbook_path, "r", encoding="utf-8") as f:
                            content = f.read()
                        
                        content = re.sub(r"<definedNames>.*?</definedNames>", "", content, flags=re.DOTALL)
                        
                        with open(workbook_path, "w", encoding="utf-8") as f:
                            f.write(content)
                        
                        fixed_path = tmp_path.replace('.xlsx', '_fixed.xlsx')
                        shutil.make_archive(fixed_path.replace('.xlsx', ''), 'zip', "temp_excel")
                        os.rename(fixed_path.replace('.xlsx', '') + '.zip', fixed_path)
                        
                        df_raw = pd.read_excel(fixed_path, header=None)
                        os.unlink(fixed_path)
                        shutil.rmtree("temp_excel", ignore_errors=True)
                    
                    # Find header row with "Pipe No."
                    header_row = df_raw[df_raw.eq("Pipe No.").any(axis=1)].index[0]
                    
                    # Clean dataframe
                    df = df_raw.iloc[header_row:]
                    df.columns = df.iloc[0]
                    df = df[1:]
                    
                    # Extract required columns
                    result_df = df[["Pipe No.", "Heat No."]].dropna().reset_index(drop=True)
                    
                    # Build master map
                    master_map = {}
                    heat_to_pipes = {}
                    
                    for _, row in result_df.iterrows():
                        pipe_no = str(row["Pipe No."]).strip()
                        heat_no_raw = str(row["Heat No."]).strip()
                        
                        if pipe_no and pipe_no != 'nan' and heat_no_raw and heat_no_raw != 'nan':
                            if '/' in heat_no_raw:
                                for h in heat_no_raw.split('/'):
                                    h = h.strip()
                                    master_map[h] = pipe_no
                                    if h not in heat_to_pipes:
                                        heat_to_pipes[h] = []
                                    if pipe_no not in heat_to_pipes[h]:
                                        heat_to_pipes[h].append(pipe_no)
                            else:
                                master_map[heat_no_raw] = pipe_no
                                if heat_no_raw not in heat_to_pipes:
                                    heat_to_pipes[heat_no_raw] = []
                                if pipe_no not in heat_to_pipes[heat_no_raw]:
                                    heat_to_pipes[heat_no_raw].append(pipe_no)
                    
                    os.unlink(tmp_path)
                    
                except Exception as e:
                    st.sidebar.error(f"Error loading Excel: {e}")
                    master_map = {}
                    heat_to_pipes = {}
            
            # Create grouped DataFrame for display
            grouped_data = []
            for heat_no, pipes in heat_to_pipes.items():
                pipes_sorted = sorted(pipes)
                pipes_display = "\n".join(pipes_sorted)
                grouped_data.append({
                    "Heat No": heat_no,
                    "Associated Pipe Nos": pipes_display,
                    "Number of Pipes": len(pipes)
                })
            
            grouped_data.sort(key=lambda x: x["Heat No"])
            
            st.session_state.master_map = master_map
            st.session_state.master_list_grouped = pd.DataFrame(grouped_data)
            st.session_state.master_heat_to_pipes = heat_to_pipes
            
            st.sidebar.success(f"✅ Loaded {len(master_map)} heat number mappings")
            st.sidebar.info(f"📊 {len(grouped_data)} unique Heat Numbers")

# Main content - Upload observation sheet
st.write("""
Upload an 'Observation Sheet (Mechanical - Vickers Hardness Test)' image and extract structured data.
""")

uploaded_file = st.file_uploader("Upload Observation Sheet Image", type=["png", "jpg", "jpeg", "tiff"])

if uploaded_file:
    image = Image.open(uploaded_file)
    st.image(image, caption="Uploaded Observation Sheet", use_container_width=True)

    # Convert PIL image to bytes
    buffered = io.BytesIO()
    image.save(buffered, format="JPEG")
    image_bytes = buffered.getvalue()

    # Prompt for Gemini 3 Flash Preview 
    prompt = """
Extract all information from the provided 'Observation Sheet (Mechanical - Vickers Hardness Test)' image into a structured Markdown format. 
Please organize the output into the following three sections:

1. Test Summary Information:
   Extract all metadata from the top section, including Customer, Spec & Grade, Test Method, Pipe Size, Atmospheric Conditions, Date, and Requirements.

2. Verification with Standard Block:
   Extract the block ID, standard value, the five individual readings, the average, the % error, and the remark.

3. Extracted Hardness Values Table:
   Create a table containing Sr. No., Sample ID No. (Pipe No,), and Heat No.
   For the 33 hardness measurement columns, group them into three sub-columns as labeled in the image: 
   Base (Points 1-6), HAZ (Points 7-24), and Weld (Points 25-33). 
   List the numbers for each row separated by commas within those groups. Include the 'Remarks' column at the end.

IMPORTANT: 
- Do NOT convert any numbers to decimals. Keep all numbers as integers.
- Do NOT add .0 to any numbers. For example, write "2601253" not "2601253.0".
- Heat No must be plain digits only, no decimal points.

Ensure all handwritten numbers are transcribed accurately and maintain a clean, professional layout.
"""

    # Initialize Gemini client
    client = genai.Client(api_key=st.secrets["GEMINI_API_KEY"])

    # Only process if not already cached
    if st.session_state.extracted_text is None:
        with st.spinner("Extracting data from image..."):
            try:
                user_content = types.Content(
                    role="user",
                    parts=[
                        types.Part.from_bytes(
                            mime_type="image/jpeg",
                            data=image_bytes
                        ),
                        types.Part.from_text(text=prompt)
                    ]
                )

                response = client.models.generate_content(
                    model="gemini-3-flash-preview",  
                    contents=user_content,
                )
                st.session_state.extracted_text = response.text
                
                # Parse and store samples
                _, _, samples = parse_markdown_output(st.session_state.extracted_text)
                st.session_state.samples = samples
                
                st.success("✅ Data extracted successfully!")
                
            except Exception as e:
                st.error(f"An error occurred: {e}")
                st.stop()
    
    # Display extracted markdown
    with st.expander("View Extracted Data"):
        st.code(st.session_state.extracted_text, language="markdown")
    
    # Show extracted samples
    if st.session_state.samples:
        st.subheader("📋 Extracted Test Samples")
        samples_df = pd.DataFrame([{
            "Pipe No (E/N prefixed)": s["sample_id"],
            "Heat No (plain digits)": s["heat_no"],
            "Test Conducted": "Yes" if s["test_conducted"] else "No",
            "Base Readings": len(s["base"]),
            "HAZ Readings": len(s["haz"]),
            "Weld Readings": len(s["weld"])
        } for s in st.session_state.samples])
        st.dataframe(samples_df, use_container_width=True)
    
    # Show Master List Preview (Grouped by Heat No) - ALL ENTRIES
    if enable_validation and st.session_state.master_list_grouped is not None:
        st.subheader("📚 Master List (Heat No → Associated Pipe Nos)")
        st.caption(f"Total: {len(st.session_state.master_list_grouped)} unique Heat Numbers")
        st.info("💡 Heat numbers are plain digits. Pipe numbers start with N or E.")
        
        # Show ALL entries with proper formatting
        st.dataframe(
            st.session_state.master_list_grouped,
            use_container_width=True,
            height=600,
            column_config={
                "Heat No": st.column_config.TextColumn("Heat No (plain digits)", width="small"),
                "Associated Pipe Nos": st.column_config.TextColumn("Associated Pipe Nos (N/E prefixed)", width="large"),
                "Number of Pipes": st.column_config.NumberColumn("Number of Pipes", width="small")
            }
        )

        st.subheader("🔍 View All Pipe Numbers by Heat Number")
        for _, row in st.session_state.master_list_grouped.iterrows():
            with st.expander(f"Heat No: {row['Heat No']} ({row['Number of Pipes']} pipes)"):
                st.code(row['Associated Pipe Nos'], language="text")
        
        # Show statistics
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total Unique Heat Nos", len(st.session_state.master_list_grouped))
        with col2:
            unique_pipes = len(set(st.session_state.master_map.values()))
            st.metric("Total Unique Pipe Nos", unique_pipes)
        with col3:
            total_relationships = st.session_state.master_list_grouped['Number of Pipes'].sum()
            st.metric("Total Heat-Pipe Relationships", total_relationships)
        
        # Option to download full master list (grouped)
        if st.button("📥 Download Full Master List (Grouped by Heat No)"):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                st.session_state.master_list_grouped.to_excel(writer, sheet_name='Master List Grouped', index=False)
            output.seek(0)
            st.download_button(
                "⬇️ Download Master List Excel",
                data=output,
                file_name="Master_List_Grouped_by_Heat.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    
    # Validation results (instant if master map exists)
    if enable_validation and st.session_state.master_map and st.session_state.samples:
        st.subheader("📊 Validation Results")
        
        # Fast validation
        results_df = validate_with_master_fast(st.session_state.samples, st.session_state.master_map)
        
        if not results_df.empty:
            st.dataframe(results_df, use_container_width=True)
            
            # Summary
            col1, col2, col3 = st.columns(3)
            with col1:
                matched = len(results_df[results_df['Status'] == '✅ MATCHED'])
                st.metric("✅ Matched", matched)
            with col2:
                mismatched = len(results_df[results_df['Status'] == '⚠️ MISMATCH'])
                st.metric("⚠️ Mismatched", mismatched)
            with col3:
                not_found = len(results_df[results_df['Status'] == '❌ NOT IN MASTER'])
                st.metric("❌ Not in Master", not_found)
            
            # Show missing heat numbers
            not_found_results = results_df[results_df['Status'] == '❌ NOT IN MASTER']
            if not not_found_results.empty:
                st.subheader("❌ Heat Numbers Not Found in Master List:")
                for _, row in not_found_results.iterrows():
                    st.warning(f"Heat No: {row['Heat No']} (Pipe No Used: {row['Pipe No Used']})")
            
            # Final output
            st.subheader("🎯 Test Conducted Using:")
            matched_results = results_df[results_df['Status'] == '✅ MATCHED']
            if not matched_results.empty:
                for _, row in matched_results.iterrows():
                    st.success(f"**Heat No:** {row['Heat No']} → **Pipe No:** {row['Pipe No Used']}")
            else:
                st.info("No matching tests found.")
    
    # Generate Excel files
    st.subheader("📥 Download Reports")
    
    col1, col2, col3 = st.columns([1, 1, 1])
    
    with col1:
        if st.button("📊 Generate Hardness Test Excel", type="primary", use_container_width=True):
            with st.spinner("Creating Hardness Test Excel file..."):
                TEMPLATE_PATH = "TEMPLATE MTC - Single Sheet.xlsx"
                OUTPUT_PATH = "Hardness_Test_Report.xlsx"
                
                try:
                    populate_excel_from_markdown(
                        st.session_state.extracted_text,
                        TEMPLATE_PATH,
                        OUTPUT_PATH
                    )
                    st.session_state.excel_ready = True
                    st.success("✅ Hardness Test Excel generated!")
                except Exception as e:
                    st.error(f"Error: {e}")
    
    with col2:
        if st.session_state.get('excel_ready', False):
            with open("Hardness_Test_Report.xlsx", "rb") as f:
                st.download_button(
                    "⬇️ Download Hardness Test Report",
                    data=f,
                    file_name="Vickers_Hardness_Test_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
    
    with col3:
        if enable_validation and st.session_state.master_map and st.session_state.samples:
            if st.button("✅ Generate Validation Report", type="secondary", use_container_width=True):
                with st.spinner("Creating Validation Excel file..."):
                    OUTPUT_PATH = "Validation_Report.xlsx"
                    
                    try:
                        create_validation_excel(
                            st.session_state.samples,
                            st.session_state.master_map,
                            OUTPUT_PATH
                        )
                        st.session_state.validation_excel_ready = True
                        st.success("✅ Validation Report generated!")
                    except Exception as e:
                        st.error(f"Error: {e}")
    
    # Validation download button
    if enable_validation and st.session_state.get('validation_excel_ready', False):
        with open("Validation_Report.xlsx", "rb") as f:
            st.download_button(
                "⬇️ Download Validation Report",
                data=f,
                file_name="Validation_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

# Clear cache button
st.sidebar.markdown("---")
if st.sidebar.button("🗑️ Clear Cache & Reset", use_container_width=True):
    st.session_state.extracted_text = None
    st.session_state.samples = None
    st.session_state.master_map = None
    st.session_state.master_list_grouped = None
    st.session_state.master_heat_to_pipes = None
    st.session_state.excel_ready = False
    st.session_state.validation_excel_ready = False
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.info("""
**Instructions:**
1. Upload observation sheet image
2. Enable validation (optional) and upload master PDF
3. View complete Master List (heat numbers with their associated pipe numbers)
4. Click 'Generate Hardness Test Excel'
5. Click 'Generate Validation Report' (if validation enabled)
6. Download both reports

**Note:** 
- Heat numbers are plain digits (e.g., 2601253, 25203481)
- Pipe numbers start with N or E (e.g., N26000425, E26001361)
""")