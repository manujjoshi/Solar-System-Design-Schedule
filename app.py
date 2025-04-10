import streamlit as st
import pandas as pd
import numpy as np
import io
import os
from datetime import datetime
import base64
from io import BytesIO
import json
from functools import lru_cache

# Cache PDF parsing functions with string conversion
@lru_cache(maxsize=32)
def extract_components_from_pdf(pdf_content):
    reader = PdfReader(BytesIO(pdf_content))
    components_data = {
        "inverters": [],
        "modules": [],
        "strings": []
    }
    
    for page in reader.pages:
        text = page.extract_text()
        
        # Extract inverter information
        inverter_matches = re.findall(r'Inverter\s*:\s*([^\n]+)', text)
        if inverter_matches:
            components_data["inverters"].extend(inverter_matches)
        
        # Extract module information
        module_matches = re.findall(r'Module\s*:\s*([^\n]+)', text)
        if module_matches:
            components_data["modules"].extend(module_matches)
        
        # Extract string information
        string_matches = re.findall(r'String\s*:\s*([^\n]+)', text)
        if string_matches:
            components_data["strings"].extend(string_matches)
    
    return json.dumps(components_data)  # Convert to string for caching

@lru_cache(maxsize=32)
def parse_helioscope_data(components_data_str):
    components_data = json.loads(components_data_str)  # Convert back to dict
    parsed_data = {
        "inverter_count": len(components_data["inverters"]),
        "inverter_model": components_data["inverters"][0] if components_data["inverters"] else "",
        "module_count": len(components_data["modules"]),
        "module_model": components_data["modules"][0] if components_data["modules"] else "",
        "string_count": len(components_data["strings"]),
        "string_config": components_data["strings"][0] if components_data["strings"] else ""
    }
    return parsed_data

def save_to_pdf_page1(data, filename="System_Summary.csv"):
    """Create CSV for System Summary"""
    # Convert data to DataFrame
    df = pd.DataFrame([data])
    df = df.T.reset_index()
    df.columns = ['Field', 'Value']
    
    # Save to CSV
    buffer = BytesIO()
    df.to_csv(buffer, index=False)
    buffer.seek(0)
    return buffer

def save_to_pdf_page2(dataframe, filename="Feed_Schedule.csv"):
    """Create CSV for Feed Schedule"""
    buffer = BytesIO()
    dataframe.to_csv(buffer, index=False)
    buffer.seek(0)
    return buffer

def save_to_pdf_page3(data, filename="Metco_Equipment.csv"):
    """Create CSV for Metco Equipment"""
    df = pd.DataFrame(data, columns=["EQUIPMENTS", "MANUFACTURER", "MODEL NUMBER", "FURNISHED BY", "INSTALLED BY"])
    buffer = BytesIO()
    df.to_csv(buffer, index=False)
    buffer.seek(0)
    return buffer

def save_to_pdf_page4(df, filename="inverter_schedule.csv"):
    """Create CSV for Inverter Schedule"""
    buffer = BytesIO()
    df.to_csv(buffer, index=False)
    buffer.seek(0)
    return buffer

def save_to_pdf_page5(df, filename="stringing_table.csv"):
    """Create CSV for String Table"""
    buffer = BytesIO()
    df.to_csv(buffer, index=False)
    buffer.seek(0)
    return buffer

def save_to_pdf_page6(panel_details, table_data, filename="panel_schedule.csv"):
    """Create CSV for Panel Schedule"""
    # Combine panel details and table data
    panel_df = pd.DataFrame([panel_details])
    table_df = pd.DataFrame(table_data)
    buffer = BytesIO()
    panel_df.to_csv(buffer, index=False)
    table_df.to_csv(buffer, mode='a', index=False)
    buffer.seek(0)
    return buffer

def merge_files(file_list, output_filename="Combined_Report.csv"):
    """Merge multiple CSV files into one"""
    with open(output_filename, 'w') as outfile:
        for i, file in enumerate(file_list):
            if file and os.path.exists(file):
                with open(file, 'r') as infile:
                    if i > 0:  # Skip header for all but first file
                        next(infile)
                    outfile.write(infile.read())
    return output_filename

# --- Page 1 Functions ---
def save_to_pdf_page1(data, filename="System_Summary.xlsx"):
    """Create Excel for System Summary"""
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        # Convert data to DataFrame
        df = pd.DataFrame([data])
        df = df.T.reset_index()
        df.columns = ['Field', 'Value']
        
        # Write to Excel
        df.to_excel(writer, sheet_name='System Summary', index=False)
        
        # Get the worksheet
        worksheet = writer.sheets['System Summary']
        
        # Set column widths
        worksheet.set_column('A:A', 40)  # Field column
        worksheet.set_column('B:B', 60)  # Value column
        
        # Add header formatting
        header_format = writer.book.add_format({
            'bold': True,
            'bg_color': '#D9D9D9',
            'border': 1
        })
        
        # Apply header formatting
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
    
    buffer.seek(0)
    return buffer

# --- Page 2 Functions ---
def generate_pdf_page2(dataframe, filename="Feed_Schedule.pdf"):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=8)
    pdf.add_page()
    pdf.set_font("Arial", "B", 12)
    pdf.cell(200, 10, "Feed Schedule - 600V MAX", ln=True, align="C")
    pdf.set_font("Arial", "B", 8)
    pdf.cell(40, 10, "TAG", border=1, align="C")
    pdf.cell(140, 10, "DESCRIPTION", border=1, align="C")
    pdf.ln()
    pdf.set_font("Arial", "", 8)
    for _, row in dataframe.iterrows():
        pdf.cell(40, 5, row["TAG"], border=1, align="C")
        pdf.cell(140, 5, row["DESCRIPTION"], border=1, align="L")
        pdf.ln()
    pdf.output(filename)
    return filename

# --- Page 3 Functions ---
def generate_equipment_pdf_page3(data, filename="Metco_Equipment.pdf"):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=12)
    pdf.add_page()
    pdf.set_font("Arial", style="B", size=14)
    pdf.cell(200, 10, "METCO Provided Equipments", ln=True, align='C')
    pdf.ln(10)
    pdf.set_font("Arial", size=6)
    pdf.set_font("Arial", style="B", size=10)
    columns = ["EQUIPMENTS", "MANUFACTURER", "MODEL NUMBER", "FURNISHED BY", "INSTALLED BY"]
    for col in columns:
        pdf.cell(38, 5, col, border=1, align='C')
    pdf.ln()
    pdf.set_font("Arial", style="", size=6)
    for row in data:
        for item in row:
            pdf.cell(38, 5, str(item), border=1, align='C')
        pdf.ln()
    pdf.output(filename)
    return filename

# --- Page 4 Functions ---
def create_default_values_page4(n, manufacturer_model, phase, volts, fla, kw, mocp, mocs, dcda, dcdnr, acda, acdnr, remarks):
    return [["INV", i + 1, "Inverter", manufacturer_model, phase, volts, fla, kw, 1, mocp,
             mocs, "20'", "0.08%", dcda, dcdnr, acda, acdnr, remarks]
            for i in range(n)]

def generate_pdf_page4(df, pdf_path="inverter_schedule.pdf"):
    pdf = FPDF(orientation='L', unit='mm', format=(800, 420))
    pdf.set_auto_page_break(auto=True, margin=10)
    pdf.add_page()
    pdf.set_font("Arial", size=10)
    col_widths = [20, 10, 30, 60, 15, 15, 15, 15, 15, 15, 40, 20, 20, 20, 20, 20, 20, 30]
    pdf.set_fill_color(200, 200, 200)
    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Arial", style="B", size=10)
    header_height = 25
    row_height = 10
    columns = df.columns.tolist()
    for col, width in zip(columns, col_widths):
        x_pos = pdf.get_x()
        y_pos = pdf.get_y()
        pdf.multi_cell(width, header_height / 5, col, border=1, align="C", fill=True)
        pdf.set_xy(x_pos + width, y_pos)
    pdf.ln(header_height)
    pdf.set_font("Arial", size=10)
    for _, row in df.iterrows():
        for item, width in zip(row, col_widths):
            pdf.cell(width, row_height, str(item), border=1, align="C", fill=False)
        pdf.ln()
    pdf.output(pdf_path)
    return pdf_path

# --- Page 5 Functions ---
def fill_table_page5(data, num_inverters, best_panels_per_string, remainder, no_of_col_to_be_filled, total_strings, no_of_mppt):
    count = 0
    step = 5
    index_order = []
    for i in range(step):
        for mppt in range(no_of_mppt):
            index = i + (mppt * step)
            if index < total_strings:
                index_order.append(index)
    for col_index in index_order:
        if count < no_of_col_to_be_filled:
            for inv in range(num_inverters):
                data[inv][col_index + 2] = best_panels_per_string
            count += 1
    while sum(sum(int(cell) for cell in row[2:] if isinstance(cell, int)) for row in data) < num_panels:
        for col_index in index_order:
            for inv in range(num_inverters):
                if isinstance(data[inv][col_index + 2], int):
                    data[inv][col_index + 2] += 1
                    if sum(sum(int(cell) for cell in row[2:] if isinstance(cell, int)) for row in data) == num_panels:
                        return data
    return data

def create_pdf_page5(df):
    """Create Excel for String Table"""
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='String Table', index=False)
        worksheet = writer.sheets['String Table']
        worksheet.set_column('A:Z', 15)
    buffer.seek(0)
    return buffer

# --- Page 6 Functions ---
def generate_pdf_page6(panel_details, table_data):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_font("Arial", style='B', size=16)
    pdf.cell(200, 10, "New Panel Schedule", ln=True, align='C')
    pdf.ln(5)
    pdf.set_font("Arial", size=10)
    pdf.set_fill_color(200, 220, 255)
    def row(label1, value1, label2, value2, label3, value3):
        pdf.cell(30, 8, label1, 1, 0, 'L', 1)
        pdf.cell(40, 8, str(value1), 1, 0, 'L')
        pdf.cell(30, 8, label2, 1, 0, 'L', 1)
        pdf.cell(40, 8, str(value2), 1, 0, 'L')
        pdf.cell(30, 8, label3, 1, 0, 'L', 1)
        pdf.cell(20, 8, str(value3), 1, 1, 'L')
    row("Panel", panel_details["Panel"], "Location", panel_details["Location"], "AMP MAIN BKR", panel_details["AMP MAIN BKR"])
    row("Mount", panel_details["Mount"], "Fed From", panel_details["Fed From"], "AMP MAIN BKR", panel_details["AMP MAIN BKR"])
    row("Volts", panel_details["Volts"], "PH", panel_details["PH"], "Wire", panel_details["Wire"])
    pdf.ln(10)
    pdf.set_font("Arial", style='B', size=11)
    pdf.cell(60, 8, "Circuit Description", 1)
    pdf.cell(30, 8, "Load VA", 1)
    pdf.cell(30, 8, "CKT BKR", 1)
    pdf.cell(30, 8, "Phase", 1)
    pdf.cell(30, 8, "Circuit #", 1)
    pdf.ln()
    pdf.set_font("Arial", size=10)
    for _, row_data in table_data.iterrows():
        pdf.cell(60, 8, str(row_data["Circuit Description"]), 1)
        pdf.cell(30, 8, str(row_data["Load VA"]), 1)
        pdf.cell(30, 8, str(row_data["CKT BKR"]), 1)
        pdf.cell(30, 8, str(row_data["Phase"]), 1)
        pdf.cell(30, 8, str(row_data["Circuit #"]), 1)
        pdf.ln()
    pdf_output = "panel_schedule.pdf"
    pdf.output(pdf_output)
    return pdf_output

# --- Common Functions ---
def merge_pdfs(pdf_list, output_filename="Combined_Report.xlsx"):
    """Merge multiple Excel files into one"""
    with pd.ExcelWriter(output_filename, engine='xlsxwriter') as writer:
        for pdf in pdf_list:
            if pdf and os.path.exists(pdf):
                df = pd.read_excel(pdf)
                sheet_name = os.path.splitext(os.path.basename(pdf))[0]
                df.to_excel(writer, sheet_name=sheet_name, index=False)
    return output_filename

# --- Streamlit Application ---
st.set_page_config(page_title="Solar System Design & Schedule", layout="wide")

# Custom CSS for styling
st.markdown("""
    <style>
    .main-section {
        background-color: #f0f2f6;
        padding: 20px;
        border-radius: 10px;
        margin-bottom: 20px;
        cursor: pointer;
        transition: transform 0.2s;
        border: 2px solid #1f77b4;
    }
    .main-section:hover {
        transform: scale(1.02);
    }
    .section-header {
        color: #1f77b4;
        font-size: 28px;
        font-weight: bold;
        margin-bottom: 15px;
        text-align: center;
    }
    .section-description {
        color: #555;
        font-size: 16px;
        text-align: center;
        margin-bottom: 20px;
    }
    .option-card {
        background-color: white;
        padding: 15px;
        border-radius: 5px;
        margin: 10px 0;
        border: 1px solid #dee2e6;
    }
    .big-title {
        font-size: 42px;
        font-weight: bold;
        color: #1f77b4;
        margin-bottom: 40px;
        text-align: center;
    }
    </style>
    """, unsafe_allow_html=True)

# Initialize session state with optimized structure
if 'current_section' not in st.session_state:
    st.session_state.current_section = None
if 'current_page' not in st.session_state:
    st.session_state.current_page = None
if 'helioscope_data' not in st.session_state:
    st.session_state.helioscope_data = None
if 'auto_populated' not in st.session_state:
    st.session_state.auto_populated = False
if 'pdf_cache' not in st.session_state:
    st.session_state.pdf_cache = {}

# Function to reset session state
def reset_session_state():
    st.session_state.current_section = None
    st.session_state.current_page = None
    st.session_state.pdf_cache = {}
    st.rerun()

# Home Page
if not st.session_state.current_section:
    st.markdown('<p class="big-title">Solar System Design & Schedule</p>', unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown('<div class="main-section">', unsafe_allow_html=True)
        st.markdown('<p class="section-header">Design Report</p>', unsafe_allow_html=True)
        st.markdown('<p class="section-description">Upload and process Helioscope reports to automatically populate system schedules.</p>', unsafe_allow_html=True)
        if st.button("Open Design Report", key="design_report_btn", use_container_width=True):
            st.session_state.current_section = "Design Report"
        st.markdown('</div>', unsafe_allow_html=True)
    
    with col2:
        st.markdown('<div class="main-section">', unsafe_allow_html=True)
        st.markdown('<p class="section-header">System Schedule</p>', unsafe_allow_html=True)
        st.markdown('<p class="section-description">Manually create and manage system schedules and generate reports.</p>', unsafe_allow_html=True)
        if st.button("Open System Schedule", key="system_schedule_btn", use_container_width=True):
            st.session_state.current_section = "System Schedule"
        st.markdown('</div>', unsafe_allow_html=True)

# Design Report Section
elif st.session_state.current_section == "Design Report":
    st.sidebar.markdown('<p class="section-header">Design Report</p>', unsafe_allow_html=True)
    if st.sidebar.button("← Back to Home"):
        reset_session_state()
    
    st.markdown('<p class="big-title">Helioscope Report Processing</p>', unsafe_allow_html=True)
    
    with st.container():
        st.markdown('<div class="workflow-step">', unsafe_allow_html=True)
        uploaded_file = st.file_uploader("Upload your Helioscope PDF Report", type=["pdf"])
        
        if uploaded_file is not None:
            try:
                with st.spinner("Processing Helioscope report..."):
                    # Cache the file content
                    file_content = uploaded_file.read()
                    components_data_str = extract_components_from_pdf(file_content)
                    parsed_data = parse_helioscope_data(components_data_str)
                    st.session_state.helioscope_data = parsed_data
                    
                    # Display extracted data
                    st.success("Report uploaded and processed successfully!")
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Inverters", f"{parsed_data['inverter_count']}x", parsed_data["inverter_model"])
                    with col2:
                        st.metric("Modules", f"{parsed_data['module_count']}x", parsed_data["module_model"])
                    with col3:
                        st.metric("Strings", f"{parsed_data['string_count']}x", parsed_data["string_config"])
                    
                    # Generate PDFs in parallel using cached data
                    with st.spinner("Generating system schedule PDFs..."):
                        # Cache PDF generation results
                        if "System_Summary.pdf" not in st.session_state.pdf_cache:
                            system_summary_data = {
                                "DC SYSTEM SIZE": f"{parsed_data['module_count'] * 0.575} kW",
                                "AC SYSTEM SIZE": f"{parsed_data['inverter_count'] * 50} kW",
                                "INVERTER (PRODUCT NAME)": parsed_data["inverter_model"],
                                "NO OF INVERTERS": str(parsed_data["inverter_count"]),
                                "SOLAR PV MODULE (PRODUCT NAME)": parsed_data["module_model"],
                                "NO OF SOLAR PV MODULES": str(parsed_data["module_count"]),
                                "RACKING (PRODUCT NAME)": "Standard Racking",
                                "NO OF RACKINGS": str(parsed_data["module_count"])
                            }
                            save_to_pdf_page1(system_summary_data)
                            st.session_state.pdf_cache["System_Summary.pdf"] = True
                            st.success("✅ System Summary PDF generated")
                        
                        # 2. Generate Feed Schedule PDF
                        tags = [20.4, 30.4, 40.4, 50.5, 60.4]  # Using first 5 tags for example
                        descriptions = [
                            "3#12,#12N,#12G-3/4\"C", "3#10,#10N,#10G-3/4\"C",
                            "3#8,#10N,#10G-1\"C", "3#6,#10N,#10G-1\"C",
                            "3#4,#10N,#10G-1 1/4\"C"
                        ]
                        feed_df = pd.DataFrame({"TAG": list(map(str, tags)), "DESCRIPTION": descriptions})
                        generate_pdf_page2(feed_df)
                        st.session_state.pdf_cache["Feed_Schedule.pdf"] = True
                        st.success("✅ Feed Schedule PDF generated")
                        
                        # 3. Generate Metco Equipment PDF
                        equipment_data = [
                            ["PV MODULES", parsed_data["module_model"], "Model X", "METCO", "Elec. Sub."],
                            ["INVERTERS", parsed_data["inverter_model"], "Model Y", "METCO", "Elec. Sub."],
                            ["RACKING", "Standard Racking", "Model Z", "METCO", "Elec. Sub."],
                            ["RAPID SHUTDOWN DEVICES", "RSD Device", "Model A", "METCO", "Elec. Sub."],
                            ["ENERGY METERING", "Standard Meter", "Model B", "METCO", "Elec. Sub."]
                        ]
                        generate_equipment_pdf_page3(equipment_data)
                        st.session_state.pdf_cache["Metco_Equipment.pdf"] = True
                        st.success("✅ Metco Equipment PDF generated")
                        
                        # 4. Generate Inverter Schedule PDF
                        inverter_data = create_default_values_page4(
                            parsed_data["inverter_count"],
                            parsed_data["inverter_model"],
                            3,  # phase
                            480,  # volts
                            "60.2",  # fla
                            50,  # kw
                            50,  # mocp
                            '3-#6, 2-#6 1 "C, EMT',  # mocs
                            "100A",  # dcda
                            "3R",  # dcdnr
                            "100A",  # acda
                            "3R",  # acdnr
                            "Standard Installation"  # remarks
                        )
                        inverter_df = pd.DataFrame(inverter_data, columns=[
                            "TAG", "#", "DESCRIPTION", "MANUFACTURER\nAND MODEL", "PHASE", "VOLTS", 
                            "FLA", "Kw", "FED TO", "MOCP", "MINIMUM CONDUIT\nAND CABLE SIZE", 
                            "APPROXIMATE\nLENGTH OF RUN", "VOLTAGE DROP\nINVERTER TO POI", 
                            "DC DISCONNECT\nAMPERAGE", "DC DISCONNECT\nNEMA RATING",
                            "AC DISCONNECT\nAMPERAGE", "AC DISCONNECT\nNEMA RATING", "REMARKS"
                        ])
                        generate_pdf_page4(inverter_df)
                        st.session_state.pdf_cache["inverter_schedule.pdf"] = True
                        st.success("✅ Inverter Schedule PDF generated")
                        
                        # 5. Generate String Table PDF
                        try:
                            # Get values from parsed data with validation
                            num_panels = int(parsed_data.get('module_count', 0))
                            num_inverters = int(parsed_data.get('inverter_count', 0))
                            num_strings = int(parsed_data.get('string_count', 0))
                            
                            # Validate values with detailed error messages
                            if num_panels <= 0:
                                raise ValueError("Invalid number of panels in the report. Please check the Helioscope report.")
                            if num_inverters <= 0:
                                raise ValueError("Invalid number of inverters in the report. Please check the Helioscope report.")
                            if num_strings <= 0:
                                raise ValueError("Invalid number of strings in the report. Please check the Helioscope report.")
                            
                            # Set up table parameters with validation
                            total_strings = max(10, num_strings)  # Ensure minimum of 10 strings
                            mppt_labels = [f"Mppt {((i // 5) + 1)}" for i in range(total_strings)]
                            string_labels = [f"String {i+1}" for i in range(total_strings)]
                            module_labels = ["MODULE COUNT"] * total_strings
                            
                            # Initialize table data with validation
                            if num_inverters > 0:
                                data = [["INV", i + 1] + ["-"] * total_strings for i in range(num_inverters)]
                            else:
                                raise ValueError("Cannot create table with zero inverters")
                            
                            # Calculate optimal panels per string with validation
                            if num_panels > 0:
                                best_panels_per_string = min(range(5, 18), key=lambda x: abs(num_panels - (x * (num_panels // x))))
                                remainder = num_panels - (best_panels_per_string * (num_panels // best_panels_per_string))
                            else:
                                raise ValueError("Cannot calculate panels per string with zero panels")
                            
                            # Calculate number of columns to be filled with validation
                            if num_inverters > 0 and num_strings > 0:
                                no_of_col_to_be_filled = min(7, max(1, int(num_strings / num_inverters)))
                            else:
                                raise ValueError("Cannot calculate columns to be filled with invalid inverter or string count")
                            
                            # Fill the table
                            data = fill_table_page5(
                                data,
                                num_inverters,
                                best_panels_per_string,
                                remainder,
                                no_of_col_to_be_filled,
                                total_strings,
                                2    # no_of_mppt
                            )
                            
                            # Create DataFrame
                            string_df = pd.DataFrame(data, columns=["TAG", "#"] + string_labels)
                            string_df.columns = pd.MultiIndex.from_tuples(
                                [("TAG", "", ""), ("#", "", "")] + list(zip(string_labels, mppt_labels, module_labels))
                            )
                            
                            # Generate PDF
                            pdf_buffer = create_pdf_page5(string_df)
                            with open("stringing_table.pdf", "wb") as f:
                                f.write(pdf_buffer.getvalue())
                            st.session_state.pdf_cache["stringing_table.pdf"] = True
                            st.success("✅ String Table PDF generated")
                            
                        except ValueError as ve:
                            st.error(f"Validation Error: {str(ve)}")
                            st.warning("Please check the Helioscope report data and try again.")
                        except Exception as e:
                            st.error(f"Error generating String Table: {str(e)}")
                            st.warning("Please check the input data and try again.")
                        
                        # 6. Generate Panel Schedule PDF
                        panel_details = {
                            "Panel": "PV-DP",
                            "Location": "Main Electric Room",
                            "AMP MAIN BKR": 350,
                            "AIC RATING": 14000,
                            "Mount": "SURFACE",
                            "Fed From": "Inverters",
                            "Volts": "480 / 277",
                            "PH": "3",
                            "Wire": "4 Wire"
                        }
                        
                        inverter_rows = []
                        for i in range(parsed_data["inverter_count"]):
                            for j in range(3):  # 3-phase
                                inverter_rows.append({
                                    "Circuit Description": f"INVERTER {i+1}" if j == 0 else "",
                                    "Load VA": "30000",
                                    "CKT BKR": "50" if j == 0 else "",
                                    "Phase": "3" if j == 0 else "",
                                    "Circuit #": len(inverter_rows) + 1
                                })
                        
                        panel_df = pd.DataFrame(inverter_rows)
                        generate_pdf_page6(panel_details, panel_df)
                        st.session_state.pdf_cache["panel_schedule.pdf"] = True
                        st.success("✅ Panel Schedule PDF generated")
                        
                        # Generate Combined PDF
                        st.info("Creating combined report...")
                        required_pdfs = [
                            "System_Summary.pdf",
                            "Feed_Schedule.pdf",
                            "Metco_Equipment.pdf",
                            "inverter_schedule.pdf",
                            "stringing_table.pdf",
                            "panel_schedule.pdf"
                        ]
                        
                        # Check if all required PDFs exist
                        missing_pdfs = [pdf for pdf in required_pdfs if not os.path.exists(pdf)]
                        if missing_pdfs:
                            st.error(f"Missing PDFs: {', '.join(missing_pdfs)}")
                            st.warning("Please ensure all PDFs are generated before creating the combined report.")
                        else:
                            try:
                                combined_pdf = merge_pdfs(required_pdfs)
                                if combined_pdf and os.path.exists(combined_pdf):
                                    st.success("✅ Complete system report generated successfully!")
                                    with open(combined_pdf, "rb") as file:
                                        st.download_button(
                                            "📥 Download Complete System Report",
                                            data=file.read(),
                                            file_name="Solar_System_Report.xlsx",
                                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                            key="download_complete"
                                        )
                            except Exception as e:
                                st.error(f"Error creating combined PDF: {str(e)}")
                                st.warning("Please try generating the report again.")
            except Exception as e:
                st.error(f"Error processing PDF: {str(e)}")
        st.markdown('</div>', unsafe_allow_html=True)

# System Schedule Section
elif st.session_state.current_section == "System Schedule":
    st.sidebar.markdown('<p class="section-header">System Schedule</p>', unsafe_allow_html=True)
    if st.sidebar.button("← Back to Home"):
        reset_session_state()
    
    # Cache the radio selection to improve performance
    if 'system_schedule_page' not in st.session_state:
        st.session_state.system_schedule_page = "System Summary"
    
    st.session_state.current_page = st.sidebar.radio(
        "Select Page",
        ["System Summary", "Feed Schedule", "Metco Equipment", "Inverter Schedule", 
         "String Table", "Panel Schedule", "Download Combined PDF"],
        key="system_schedule_radio"
    )
    
    # Display selected page content
    if st.session_state.current_page == "System Summary":
        st.title("System Summary Form")
        with st.form("system_summary_form"):
            # Use auto-populated data if available
            if st.session_state.get("auto_populated", False):
                dc_system_size = st.text_input("DC SYSTEM SIZE", st.session_state.system_summary_data["DC SYSTEM SIZE"])
                ac_system_size = st.text_input("AC SYSTEM SIZE", st.session_state.system_summary_data["AC SYSTEM SIZE"])
                inverters = st.text_area("INVERTER (PRODUCT NAME)", st.session_state.system_summary_data["INVERTER (PRODUCT NAME)"])
                no_of_inverters = st.text_input("NO OF INVERTERS", st.session_state.system_summary_data["NO OF INVERTERS"])
                solar_pv_module = st.text_input("SOLAR PV MODULE (PRODUCT NAME)", st.session_state.system_summary_data["SOLAR PV MODULE (PRODUCT NAME)"])
                no_of_solar_pv_modules = st.text_area("NO OF SOLAR PV MODULES", st.session_state.system_summary_data["NO OF SOLAR PV MODULES"])
            else:
                dc_system_size = st.text_input("DC SYSTEM SIZE", "181.13 kW")
                ac_system_size = st.text_input("AC SYSTEM SIZE", "150kW")
                inverters = st.text_area("INVERTER (PRODUCT NAME)", "CPS-SCA50KTL-DO/US-480")
                no_of_inverters = st.text_input("NO OF INVERTERS", "3")
                solar_pv_module = st.text_input("SOLAR PV MODULE (PRODUCT NAME)", "Sunsprint Engineering, SPISLE575-144TGG (575W)")
                no_of_solar_pv_modules = st.text_area("NO OF SOLAR PV MODULES", "315")
            
            racking = st.text_input("RACKING (PRODUCT NAME)", "")
            no_of_rackings = st.text_input("NO OF RACKINGS", "")
            submitted = st.form_submit_button("Generate PDF")

        if submitted:
            data = {
                "DC SYSTEM SIZE": dc_system_size,
                "AC SYSTEM SIZE": ac_system_size,
                "INVERTER (PRODUCT NAME)": inverters,
                "NO OF INVERTERS": no_of_inverters,
                "RACKING (PRODUCT NAME)": racking,
                "NO OF RACKINGS": no_of_rackings,
                "SOLAR PV MODULE (PRODUCT NAME)": solar_pv_module,
                "NO OF SOLAR PV MODULES": no_of_solar_pv_modules,
            }
            csv_buffer = save_to_pdf_page1(data)
            st.session_state["csv_data_page1"] = csv_buffer.getvalue()
            st.success("CSV file generated successfully!")
            st.download_button(
                "📄 Download System Summary",
                data=st.session_state["csv_data_page1"],
                file_name="System_Summary.csv",
                mime="text/csv"
            )
    
    elif st.session_state.current_page == "Feed Schedule":
        st.title("Feed Schedule - 600V MAX")
        tags = [
            20.4, 30.4, 40.4, 50.5, 60.4, 70.4, 80.4, 90.4, 100.4, 125.4, 150.4, 175.4, 200.4, 225.5,
            250.4, 300.4, 400.4, 500.4, 600.4, 700.4, 800.4, 900.4, 1000.4, 1600.4, 2000.4, 2500.4, 4000.4
        ]
        descriptions = [
            "3#12,#12N,#12G-3/4\"C", "3#10,#10N,#10G-3/4\"C", "3#8,#10N,#10G-1\"C", "3#6,#10N,#10G-1\"C",
            "3#4,#10N,#10G-1 1/4\"C", "3#4,#8N,#8G-1 1/4\"C", "3#3,#8N,#8G-1 1/4\"C", "3#2,#8N,#8G-1 1/2\"C",
            "3#2,#8N,#8G-1 1/2\"C", "3#1,#6N,#6G-1 1/2\"C", "3#1/0,#6N,#6G-2\"C", "3#2/0,#6N,#6G-2\"C",
            "3#3/0,#6N,#6G-2\"C", "3#4/0,#4N,#4G-2 1/2\"C", "3-250KCMIL, #4N,#4G-2 1/2\"C",
            "3-350KCMIL, #4N,#4G-3\"C", "3-500KCMIL, #3N,#3G-4\"C", "2 SETS OF 3-250KCMIL, #2N,#2G-2 1/2\"C",
            "3 SETS OF 3-350KCMIL, #1N,#1G-1\"C", "3 SETS OF 3-350KCMIL, #1/0N,#1/0G-3\"C",
            "3 SETS OF 3-350KCMIL, #2/0N,#2/0G-3\"C", "3 SETS OF 3-400KCMIL, #2/0N,#2/0G-3\"C",
            "3 SETS OF 3-350KCMIL, #3/0N,#3/0G-3\"C", "3 SETS OF 3-500KCMIL, #4/0N,#4/0G-4\"C",
            "5 SETS OF 3-500KCMIL, 1-250KCMILN,1-250KCMILG-4\"C",
            "6 SETS OF 3-500KCMIL, 1-350KCMILN,1-350KCMILG-4\"C",
            "7 SETS OF 3-500KCMIL, 1-400KCMILN,1-400KCMILG-4\"C"
        ]

        df = pd.DataFrame({"TAG": list(map(str, tags)), "DESCRIPTION": descriptions})
        st.dataframe(df)

        if st.button("Generate Feed Schedule PDF"):
            pdf_file = generate_pdf_page2(df)
            with open(pdf_file, "rb") as file:
                st.session_state["pdf_data_page2"] = file.read()
                st.success("PDF generated successfully!")
                st.download_button(
                    "📄 Download Feed Schedule PDF",
                    data=st.session_state["pdf_data_page2"],
                    file_name="Feed_Schedule.pdf",
                    mime="application/pdf"
                )
    
    elif st.session_state.current_page == "Metco Equipment":
        st.title("METCO Provided Equipment Form")
        
        equipment_list = [
            "PV MODULES", "RAPID SHUTDOWN DEVICES", "RACKING",
            "INVERTERS", "ENERGY METERING", "ENERGY ENCLOSURES", 
            "ENERGY WEATHER STATION", "WIRELESS BRIDGES"
        ]
        
        with st.form("equipment_form"):
            st.subheader("Enter Equipment Details")
            data = []
            
            for equipment in equipment_list:
                st.write(f"### {equipment}")
                col1, col2 = st.columns(2)
                with col1:
                    manufacturer = st.text_input(f"Manufacturer", key=f"manufacturer_{equipment}")
                    model_number = st.text_input(f"Model Number", key=f"model_{equipment}")
                with col2:
                    furnished_by = st.text_input(f"Furnished By", value="METCO", key=f"furnished_{equipment}")
                    installed_by = st.text_input(f"Installed By", value="Elec. Sub.", key=f"installed_{equipment}")
                
                data.append([equipment, manufacturer, model_number, furnished_by, installed_by])
            
            submitted = st.form_submit_button("Generate PDF")
        
        if submitted:
            pdf_file = generate_equipment_pdf_page3(data)
            with open(pdf_file, "rb") as file:
                st.session_state["pdf_page3"] = file.read()
                st.success("PDF generated successfully!")
                st.download_button(
                    "📄 Download Equipment PDF",
                    data=st.session_state["pdf_page3"],
                    file_name="Metco_Equipment.pdf",
                    mime="application/pdf"
                )

    elif st.session_state.current_page == "Inverter Schedule":
        st.title("Inverter Schedule Form")
        
        col1, col2 = st.columns(2)
        with col1:
            num_rows = st.number_input("Number of Inverters", min_value=1, value=1, step=1)
            manufacturer_model = st.text_input("Manufacturer and Model", "CPS- SCA50KTL-DO/US-480 - 50 kW")
            phase = st.number_input("Phase", min_value=1, max_value=1000, value=3, step=1)
            volts = st.number_input("Volts", min_value=100, max_value=1000, value=480, step=1)
            fla = st.text_input("FLA", "60.2")
            kw = st.number_input("Kilowatts (kW)", min_value=1, max_value=100, value=50, step=1)
        
        with col2:
            mocp = st.number_input("MOCP", min_value=1, max_value=100, value=50, step=1)
            mocs = st.text_input("MINIMUM CONDUIT AND CABLE SIZE", '3-#6, 2-#6 1 "C, EMT')
            dcda = st.text_input("DC DISCONNECT AMPERAGE", "N/A")
            dcdnr = st.text_input("DC DISCONNECT NEMA RATING", "N/A")
            acda = st.text_input("AC DISCONNECT AMPERAGE", "N/A")
            acdnr = st.text_input("AC DISCONNECT NEMA RATING", "N/A")
            remarks = st.text_input("REMARKS", "N/A")

        columns = [
            "TAG", "#", "DESCRIPTION", "MANUFACTURER\nAND MODEL", "PHASE", "VOLTS", "FLA", "Kw",
            "FED TO", "MOCP", "MINIMUM CONDUIT\nAND CABLE SIZE", "APPROXIMATE\nLENGTH OF RUN",
            "VOLTAGE DROP\nINVERTER TO POI", "DC DISCONNECT\nAMPERAGE", "DC DISCONNECT\nNEMA RATING",
            "AC DISCONNECT\nAMPERAGE", "AC DISCONNECT\nNEMA RATING", "REMARKS"
        ]

        df = pd.DataFrame(create_default_values_page4(num_rows, manufacturer_model, phase, volts, fla, kw, mocp, mocs, dcda, dcdnr, acda, acdnr, remarks), columns=columns)

        st.subheader("Additional Details")
        for i in range(num_rows):
            col1, col2, col3 = st.columns(3)
            with col1:
                df.at[i, "FED TO"] = st.selectbox(f"Distribution Panel (Row {i+1})", [str(x) for x in range(1, 11)], key=f"dist_{i}")
            with col2:
                df.at[i, "APPROXIMATE\nLENGTH OF RUN"] = st.text_input(f"Approx. Length of Run (Row {i+1})", "20'", key=f"length_{i}")
            with col3:
                df.at[i, "VOLTAGE DROP\nINVERTER TO POI"] = st.text_input(f"Voltage Drop (Row {i+1})", "0.08%", key=f"voltage_{i}")

        st.subheader("Preview")
        st.dataframe(df)

        if st.button("Generate PDF"):
            pdf_file = generate_pdf_page4(df)
            with open(pdf_file, "rb") as file:
                st.session_state["pdf_page4"] = file.read()
                st.success("PDF generated successfully!")
                st.download_button(
                    "📄 Download Inverter Schedule PDF",
                    data=st.session_state["pdf_page4"],
                    file_name="inverter_schedule.pdf",
                    mime="application/pdf"
                )

    elif st.session_state.current_page == "String Table":
        st.title("String Table Generator")
        
        with st.form("string_table_form"):
            col1, col2 = st.columns(2)
            with col1:
                num_inverters = st.number_input("Enter number of inverters:", min_value=1, value=1)
                num_panels = st.number_input("Enter number of panels:", min_value=1, value=1)
                panels_in_1_inverter = num_panels / num_inverters if num_inverters > 0 else 0
                st.write(f"Panels in one inverter: {panels_in_1_inverter}")
            
            with col2:
                no_of_mppt = st.number_input("No of MPPT (In 1 inverter):", min_value=1, value=1)
                no_of_string = st.number_input("Total No of String (In 1 MPPT):", min_value=1, value=1)
                total_strings = no_of_mppt * no_of_string
                st.write(f"Total number of Strings: {total_strings}")
            
            no_of_strings_used = st.number_input("Enter number of Strings used from report:", min_value=1, value=1)
            submit_button = st.form_submit_button("Generate Table")
        
        if submit_button:
            try:
                best_panels_per_string = min(range(5, 18), key=lambda x: abs(num_panels - (x * (num_panels // x))))
                remainder = num_panels - (best_panels_per_string * (num_panels // best_panels_per_string))
                no_of_col_to_be_filled = min(7, int(no_of_strings_used / num_inverters))
                
                total_strings = no_of_mppt * 5
                mppt_labels = [f"Mppt {((i // 5) + 1)}" for i in range(total_strings)]
                string_labels = [f"String {i+1}" for i in range(total_strings)]
                module_labels = ["MODULE COUNT"] * total_strings
                
                data = [["INV", i + 1] + ["-"] * total_strings for i in range(num_inverters)]
                data = fill_table_page5(
                    data,
                    num_inverters,
                    best_panels_per_string,
                    remainder,
                    no_of_col_to_be_filled,
                    total_strings,
                    no_of_mppt
                )
                
                df = pd.DataFrame(data, columns=["TAG", "#"] + string_labels)
                df.columns = pd.MultiIndex.from_tuples(
                    [("TAG", "", ""), ("#", "", "")] + list(zip(string_labels, mppt_labels, module_labels))
                )
                
                st.subheader("Generated String Table")
                st.dataframe(df)
                
                st.session_state.string_table_df = df
                st.session_state.string_table_data = data
                
                if st.button("Generate PDF"):
                    pdf_buffer = create_pdf_page5(df)
                    with open("stringing_table.pdf", "wb") as f:
                        f.write(pdf_buffer.getvalue())
                    st.success("PDF generated successfully!")
                    st.download_button(
                        "📄 Download String Table PDF",
                        data=pdf_buffer,
                        file_name="stringing_table.pdf",
                        mime="application/pdf"
                    )
                
                st.write(f"Optimal panels per string: {best_panels_per_string}")
                st.write(f"Optimal number of columns to be filled: {no_of_col_to_be_filled}")
                st.write(f"Total panels in table: {sum(sum(int(cell) for cell in row[2:] if isinstance(cell, int)) for row in data)}")
            except Exception as e:
                st.error(f"Error generating table: {str(e)}")

    elif st.session_state.current_page == "Panel Schedule":
        st.title("Panel Schedule")
        
        with st.form("panel_form"):
            st.subheader("Panel Information")
            col1, col2, col3 = st.columns(3)
            with col1:
                panel = st.text_input("Panel", "PV-DP")
                mount = st.text_input("Mount", "SURFACE")
                volts = st.text_input("Volts", "480 / 277")
            with col2:
                location = st.text_input("Location")
                fed_from = st.text_input("Fed From", "Inverters")
                ph = st.selectbox("PH", ["1", "3"], index=1)
            with col3:
                amp_main_bkr = st.number_input("AMP MAIN BKR", min_value=0, value=350)
                aic_rating = st.number_input("AIC RATING", min_value=0, value=14000)
                wire = st.text_input("Wire", "4 Wire")
            
            submitted = st.form_submit_button("Save Panel Details")
        
        st.subheader("Set Load VA and CKT BKR for Inverters")
        with st.form("global_values_form"):
            global_load_va = st.number_input("Global Load VA", min_value=0)
            global_ckt_bkr = st.number_input("Global CKT BKR", min_value=0)
            apply_values = st.form_submit_button("Apply Values to Inverter Rows")
        
        st.subheader("Panel Schedule Table")
        inverter_names = ["INVERTER 1", "INVERTER 2", "INVERTER 3"]
        inverter_rows = []
        phase_count = int(ph)
        
        for idx, name in enumerate(inverter_names):
            for i in range(phase_count):
                inverter_rows.append({
                    "Circuit Description": name if i == 0 else "",
                    "Load VA": str(global_load_va) if global_load_va else "",
                    "CKT BKR": str(global_ckt_bkr) if i == 0 and global_ckt_bkr else "",
                    "Phase": ph if i == 0 else "",
                    "Circuit #": len(inverter_rows) + 1
                })
        
        ae1_rows = [
            {
                "Circuit Description": "AE1",
                "Load VA": "2250",
                "CKT BKR": "20",
                "Phase": "2",
                "Circuit #": len(inverter_rows) + 1
            },
            {
                "Circuit Description": "",
                "Load VA": "2250",
                "CKT BKR": "",
                "Phase": "",
                "Circuit #": len(inverter_rows) + 2
            }
        ]
        
        final_rows = inverter_rows + ae1_rows
        df = pd.DataFrame(final_rows)
        edited_df = st.data_editor(df, num_rows="dynamic", use_container_width=True)
        st.session_state.panel_df = edited_df
        
        if st.button("Generate PDF"):
            panel_details = {
                "Panel": panel,
                "Location": location,
                "AMP MAIN BKR": amp_main_bkr,
                "AIC RATING": aic_rating,
                "Mount": mount,
                "Fed From": fed_from,
                "Volts": volts,
                "PH": ph,
                "Wire": wire
            }
            pdf_file = generate_pdf_page6(panel_details, edited_df)
            with open(pdf_file, "rb") as f:
                st.success("PDF generated successfully!")
                st.download_button(
                    "📄 Download Panel Schedule PDF",
                    data=f.read(),
                    file_name="panel_schedule.pdf",
                    mime="application/pdf"
                )

    elif st.session_state.current_page == "Download Combined PDF":
        st.title("Download Complete System Report")
        
        required_pdfs = [
            "System_Summary.pdf",
            "Feed_Schedule.pdf",
            "Metco_Equipment.pdf",
            "inverter_schedule.pdf",
            "stringing_table.pdf",
            "panel_schedule.pdf"
        ]
        
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("### Required Components")
            for pdf in required_pdfs:
                if os.path.exists(pdf):
                    st.success(f"✅ {pdf}")
                else:
                    st.error(f"❌ {pdf} (Missing)")
        
        with col2:
            st.markdown("### Actions")
            if all(os.path.exists(pdf) for pdf in required_pdfs):
                try:
                    combined_pdf = merge_pdfs(required_pdfs)
                    if combined_pdf and os.path.exists(combined_pdf):
                        with open(combined_pdf, "rb") as file:
                            st.download_button(
                                "📥 Download Complete System Report",
                                data=file.read(),
                                file_name="Solar_System_Report.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="download_complete"
                            )
                except Exception as e:
                    st.error(f"Error creating combined PDF: {str(e)}")
                    st.warning("Please try generating the report again.")
            else:
                st.warning("Please generate all required PDFs before downloading the complete report.")

# ... rest of the existing code ... 