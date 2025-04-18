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
from fpdf import FPDF
import re
import fitz  # PyMuPDF
from PyPDF2 import PdfMerger, PdfReader
from reportlab.lib.pagesizes import landscape, letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib import colors

def save_to_pdf_page1(data):
    """Generate System Summary PDF"""
    pdf = FPDF(orientation='L')
    pdf.add_page()
    
    # Set professional font and styling
    pdf.set_font("Arial", "B", 16)
    pdf.cell(0, 10, "System Summary", ln=True, align="C")
    pdf.ln(10)
    
    # Set column widths for better readability
    col_widths = [80, 200]  # Wider columns for landscape
    
    # Header row
    pdf.set_font("Arial", "B", 12)
    pdf.cell(col_widths[0], 10, "Field", 1, 0, 'C')
    pdf.cell(col_widths[1], 10, "Value", 1, 1, 'C')
    
    # Data rows
    pdf.set_font("Arial", "", 12)
    for key, value in data.items():
        pdf.cell(col_widths[0], 10, key, 1, 0, 'L')
        pdf.cell(col_widths[1], 10, str(value), 1, 1, 'L')
    
    output_path = "System_Summary.pdf"
    pdf.output(output_path)
    return output_path

def save_to_pdf_page2(dataframe):
    """Generate Feed Schedule PDF"""
    pdf = FPDF(orientation='L')
    pdf.add_page()
    
    # Set professional font and styling
    pdf.set_font("Arial", "B", 16)
    pdf.cell(0, 10, "Feed Schedule", ln=True, align="C")
    pdf.ln(10)
    
    # Set column widths for better readability
    col_widths = [80, 200]  # Wider columns for landscape
    
    # Header row
    pdf.set_font("Arial", "B", 12)
    pdf.cell(col_widths[0], 10, "TAG", 1, 0, 'C')
    pdf.cell(col_widths[1], 10, "DESCRIPTION", 1, 1, 'C')
    
    # Data rows
    pdf.set_font("Arial", "", 12)
    for _, row in dataframe.iterrows():
        pdf.cell(col_widths[0], 10, str(row['TAG']), 1, 0, 'L')
        pdf.cell(col_widths[1], 10, str(row['DESCRIPTION']), 1, 1, 'L')
    
    output_path = "Feed_Schedule.pdf"
    pdf.output(output_path)
    return output_path

def save_to_pdf_page3(data):
    """Generate Metco Equipment PDF"""
    pdf = FPDF(orientation='L')
    pdf.add_page()
    
    # Set professional font and styling
    pdf.set_font("Arial", "B", 16)
    pdf.cell(0, 10, "Metco Equipment", ln=True, align="C")
    pdf.ln(10)
    
    # Set column widths for better readability
    col_widths = [80, 200]  # Wider columns for landscape
    
    # Header row
    pdf.set_font("Arial", "B", 12)
    pdf.cell(col_widths[0], 10, "Field", 1, 0, 'C')
    pdf.cell(col_widths[1], 10, "Value", 1, 1, 'C')
    
    # Data rows
    pdf.set_font("Arial", "", 12)
    for field, value in data.items():
        pdf.cell(col_widths[0], 10, str(field), 1, 0, 'L')
        pdf.cell(col_widths[1], 10, str(value), 1, 1, 'L')
    
    output_path = "Metco_Equipment.pdf"
    pdf.output(output_path)
    return output_path

def generate_equipment_pdf_page3(data):
    """Generate Metco Equipment PDF with improved formatting"""
    # Create PDF in landscape mode with letter size
    doc = SimpleDocTemplate("Metco_Equipment.pdf", pagesize=landscape(letter))
    elements = []
    
    # Define column widths as percentages of the page width
    page_width = landscape(letter)[0]
    col_widths = [
        page_width * 0.25,  # Equipment
        page_width * 0.20,  # Manufacturer
        page_width * 0.20,  # Model Number
        page_width * 0.15,  # Furnished By
        page_width * 0.15   # Installed By
    ]
    
    # Create table data
    table_data = [["Equipment", "Manufacturer", "Model Number", "Furnished By", "Installed By"]]
    table_data.extend(data)
    
    # Create table with the specified column widths
    table = Table(table_data, colWidths=col_widths)
    
    # Add style to the table
    style = TableStyle([
        # Header styling
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1f77b4')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        
        # Content styling
        ('BACKGROUND', (0, 1), (-1, -1), colors.white),
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
        ('ALIGN', (0, 1), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 10),
        ('TOPPADDING', (0, 1), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 1), (-1, -1), 8),
        
        # Grid styling
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('LINEBEFORE', (0, 0), (-1, -1), 1, colors.black),
        ('LINEAFTER', (0, 0), (-1, -1), 1, colors.black),
        ('LINEBELOW', (0, 0), (-1, -1), 1, colors.black),
        ('LINEABOVE', (0, 0), (-1, -1), 1, colors.black),
    ])
    table.setStyle(style)
    
    # Add title
    title_style = TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 16),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 20),
    ])
    title = Table([[f"METCO Provided Equipment"]], colWidths=[sum(col_widths)])
    title.setStyle(title_style)
    
    # Add elements to the document
    elements.append(title)
    elements.append(table)
    
    # Build the document
    doc.build(elements)
    return "Metco_Equipment.pdf"

def generate_pdf_page4(df):
    """Generate Inverter Schedule PDF with improved formatting"""
    # Create PDF in landscape mode with double width
    doc = SimpleDocTemplate(
        "inverter_schedule.pdf",
        pagesize=(landscape(letter)[0] * 2, landscape(letter)[1]),  # Double width
        rightMargin=10,
        leftMargin=10,
        topMargin=10,
        bottomMargin=10
    )
    elements = []
    
    # Calculate column widths based on content
    page_width = (landscape(letter)[0] * 2) - 20  # Account for margins
    
    # Define minimum and maximum column widths
    min_col_width = page_width * 0.05  # 5% of page width
    max_col_width = page_width * 0.25  # 25% of page width
    
    # Calculate content-based widths
    col_widths = []
    for col in df.columns:
        # Get maximum content length in the column
        max_content_length = max(
            len(str(val)) for val in df[col].tolist() + [col]
        )
        # Calculate width based on content length
        content_width = min(
            max(min_col_width, max_content_length * 5),  # 5 points per character
            max_col_width
        )
        col_widths.append(content_width)
    
    # Normalize widths to fit page width
    total_width = sum(col_widths)
    if total_width > page_width:
        scale_factor = page_width / total_width
        col_widths = [w * scale_factor for w in col_widths]
    
    # Prepare table data with 2 rows for headers
    if isinstance(df.columns, pd.MultiIndex):
        # First row: MPPT labels
        mppt_row = ['TAG', '#'] + [f"MPPT {i//5 + 1}" for i in range(len(df.columns) - 2)]
        # Second row: String labels
        string_row = ['', ''] + [f"String {i%5 + 1}" for i in range(len(df.columns) - 2)]
        header_rows = [mppt_row, string_row]
    else:
        # Split column names into two rows if they contain spaces
        header_rows = []
        first_row = []
        second_row = []
        for col in df.columns:
            if isinstance(col, str) and ' ' in col:
                parts = col.split(' ', 1)
                first_row.append(parts[0])
                second_row.append(parts[1] if len(parts) > 1 else '')
            else:
                first_row.append(str(col))
                second_row.append('')
        header_rows = [first_row, second_row]
    
    # Convert all values to strings to avoid type errors
    table_data = header_rows + [[str(val) if val is not None else '' for val in row] for row in df.values.tolist()]
    
    # Create table
    table = Table(table_data, colWidths=col_widths, repeatRows=len(header_rows))
    
    # Add style to the table
    style = TableStyle([
        # Header styling for all header rows
        ('BACKGROUND', (0, 0), (-1, len(header_rows)-1), colors.HexColor('#1f77b4')),
        ('TEXTCOLOR', (0, 0), (-1, len(header_rows)-1), colors.white),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, len(header_rows)-1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, len(header_rows)-1), 10),
        ('BOTTOMPADDING', (0, 0), (-1, len(header_rows)-1), 8),
        ('TOPPADDING', (0, 0), (-1, len(header_rows)-1), 8),
        ('LEADING', (0, 0), (-1, len(header_rows)-1), 12),
        
        # Content styling
        ('BACKGROUND', (0, len(header_rows)), (-1, -1), colors.white),
        ('TEXTCOLOR', (0, len(header_rows)), (-1, -1), colors.black),
        ('FONTNAME', (0, len(header_rows)), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, len(header_rows)), (-1, -1), 9),
        ('TOPPADDING', (0, len(header_rows)), (-1, -1), 6),
        ('BOTTOMPADDING', (0, len(header_rows)), (-1, -1), 6),
        ('LEADING', (0, len(header_rows)), (-1, -1), 10),
        
        # Grid styling
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('LINEBEFORE', (0, 0), (-1, -1), 0.5, colors.black),
        ('LINEAFTER', (0, 0), (-1, -1), 0.5, colors.black),
        ('LINEBELOW', (0, 0), (-1, -1), 0.5, colors.black),
        ('LINEABOVE', (0, 0), (-1, -1), 0.5, colors.black),
        
        # Word wrapping
        ('WORDWRAP', (0, 0), (-1, -1), True),
        
        # Alternate row colors for better readability
        *[('BACKGROUND', (0, i), (-1, i), colors.HexColor('#f8f9fa')) 
          for i in range(len(header_rows), len(table_data), 2)]
    ])
    table.setStyle(style)
    
    # Add title
    title_style = TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 16),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 20),
        ('TOPPADDING', (0, 0), (-1, -1), 20),
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.HexColor('#1f77b4')),
    ])
    title = Table([["INVERTER SCHEDULE"]], colWidths=[sum(col_widths)])
    title.setStyle(title_style)
    
    # Add elements to the document
    elements.append(title)
    elements.append(table)
    
    # Build the document
    doc.build(elements)
    return "inverter_schedule.pdf"

def create_pdf_page5(df):
    """Generate String Table PDF with two-page view"""
    # Create PDF with double page size
    doc = SimpleDocTemplate(
        "stringing_table.pdf",
        pagesize=(landscape(letter)[0] * 2, landscape(letter)[1]),  # Double width
        rightMargin=10,
        leftMargin=10,
        topMargin=10,
        bottomMargin=10
    )
    elements = []
    
    # Calculate column widths
    page_width = (landscape(letter)[0] * 2) - 20  # Account for margins
    num_columns = len(df.columns)
    tag_width = page_width * 0.15  # Width for TAG column
    num_width = page_width * 0.12  # Width for # column
    remaining_width = page_width - (tag_width + num_width)
    string_width = remaining_width / (num_columns - 2)  # Equal width for remaining columns
    
    col_widths = [tag_width, num_width] + [string_width] * (num_columns - 2)
    
    # Prepare table data with 2 rows for headers
    if isinstance(df.columns, pd.MultiIndex):
        # First row: MPPT labels
        mppt_row = ['TAG', '#'] + [f"MPPT {i//5 + 1}" for i in range(num_columns - 2)]
        # Second row: String labels
        string_row = ['', ''] + [f"String {i%5 + 1}" for i in range(num_columns - 2)]
        header_rows = [mppt_row, string_row]
    else:
        # Split column names into two rows if they contain spaces
        header_rows = []
        first_row = []
        second_row = []
        for col in df.columns:
            if isinstance(col, str) and ' ' in col:
                parts = col.split(' ', 1)
                first_row.append(parts[0])
                second_row.append(parts[1] if len(parts) > 1 else '')
            else:
                first_row.append(str(col))
                second_row.append('')
        header_rows = [first_row, second_row]
    
    # Convert all values to strings to avoid type errors
    table_data = header_rows + [[str(val) if val is not None else '' for val in row] for row in df.values.tolist()]
    
    # Create table
    table = Table(table_data, colWidths=col_widths, repeatRows=len(header_rows))
    
    # Add style to the table
    style = TableStyle([
        # Header styling for all header rows
        ('BACKGROUND', (0, 0), (-1, len(header_rows)-1), colors.HexColor('#1f77b4')),
        ('TEXTCOLOR', (0, 0), (-1, len(header_rows)-1), colors.white),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, len(header_rows)-1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, len(header_rows)-1), 12),
        ('BOTTOMPADDING', (0, 0), (-1, len(header_rows)-1), 12),
        ('TOPPADDING', (0, 0), (-1, len(header_rows)-1), 12),
        ('LEADING', (0, 0), (-1, len(header_rows)-1), 14),
        
        # Content styling
        ('BACKGROUND', (0, len(header_rows)), (-1, -1), colors.white),
        ('TEXTCOLOR', (0, len(header_rows)), (-1, -1), colors.black),
        ('FONTNAME', (0, len(header_rows)), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, len(header_rows)), (-1, -1), 10),
        ('TOPPADDING', (0, len(header_rows)), (-1, -1), 8),
        ('BOTTOMPADDING', (0, len(header_rows)), (-1, -1), 8),
        ('LEADING', (0, len(header_rows)), (-1, -1), 12),
        
        # Grid styling
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('LINEBEFORE', (0, 0), (-1, -1), 0.5, colors.black),
        ('LINEAFTER', (0, 0), (-1, -1), 0.5, colors.black),
        ('LINEBELOW', (0, 0), (-1, -1), 0.5, colors.black),
        ('LINEABOVE', (0, 0), (-1, -1), 0.5, colors.black),
        
        # Span cells for TAG and # columns
        ('SPAN', (0, 0), (0, 1)),  # TAG
        ('SPAN', (1, 0), (1, 1)),  # #
        
        # Word wrapping
        ('WORDWRAP', (0, 0), (-1, -1), True),
        
        # Alternate row colors for better readability
        *[('BACKGROUND', (0, i), (-1, i), colors.HexColor('#f8f9fa')) 
          for i in range(len(header_rows), len(table_data), 2)]
    ])
    table.setStyle(style)
    
    # Add title
    title_style = TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 16),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 20),
        ('TOPPADDING', (0, 0), (-1, -1), 20),
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.HexColor('#1f77b4')),
    ])
    title = Table([["STRING TABLE"]], colWidths=[sum(col_widths)])
    title.setStyle(title_style)
    
    # Add elements to the document
    elements.append(title)
    elements.append(table)
    
    # Build the document
    doc.build(elements)
    return "stringing_table.pdf"

def generate_pdf_page6(panel_details, df):
    """Generate Panel Schedule PDF"""
    pdf = FPDF(orientation='L')
    pdf.add_page()
    
    # Set professional font and styling
    pdf.set_font("Arial", "B", 16)
    pdf.cell(0, 10, "Panel Schedule", ln=True, align="C")
    pdf.ln(10)
    
    # Add panel details
    pdf.set_font("Arial", "", 12)
    for key, value in panel_details.items():
        pdf.cell(60, 10, key, 1, 0, 'L')
        pdf.cell(0, 10, str(value), 1, 1, 'L')
    pdf.ln(10)
    
    # Set column widths for better readability
    col_widths = [60, 30, 30, 30, 30]
    headers = ["Circuit Description", "Load VA", "CKT BKR", "Phase", "Circuit #"]
    
    # Header row
    pdf.set_font("Arial", "B", 12)
    for i, header in enumerate(headers):
        pdf.cell(col_widths[i], 10, header, 1, 0, 'C')
    pdf.ln()
    
    # Data rows
    pdf.set_font("Arial", "", 12)
    for _, row in df.iterrows():
        for i, col in enumerate(headers):
            pdf.cell(col_widths[i], 10, str(row[col]), 1, 0, 'L')
        pdf.ln()
    
    output_path = "panel_schedule.pdf"
    pdf.output(output_path)
    return output_path

def merge_pdfs(pdf_files):
    """Merge multiple PDFs into a single report with consistent formatting"""
    merger = PdfMerger()
    
    # Add each PDF with proper page breaks
    for pdf in pdf_files:
        merger.append(pdf)
    
    # Create the final merged PDF
    output_path = "Solar_System_Report.pdf"
    merger.write(output_path)
    merger.close()
    
    return output_path

def create_default_values_page4(num_rows, manufacturer_model, phase, volts, fla, kw, mocp, mocs, dcda, dcdnr, acda, acdnr, remarks):
    """Create default values for inverter schedule"""
    data = []
    for i in range(num_rows):
        row = [
            f"INV{i+1}",  # TAG
            str(i+1),     # #
            f"Inverter {i+1}",  # DESCRIPTION
            manufacturer_model,  # MANUFACTURER AND MODEL
            str(phase),   # PHASE
            str(volts),   # VOLTS
            str(fla),     # FLA
            str(kw),      # Kw
            "",           # FED TO
            str(mocp),    # MOCP
            mocs,         # MINIMUM CONDUIT AND CABLE SIZE
            "",           # APPROXIMATE LENGTH OF RUN
            "",           # VOLTAGE DROP INVERTER TO POI
            dcda,         # DC DISCONNECT AMPERAGE
            dcdnr,        # DC DISCONNECT NEMA RATING
            acda,         # AC DISCONNECT AMPERAGE
            acdnr,        # AC DISCONNECT NEMA RATING
            remarks       # REMARKS
        ]
        data.append(row)
    return data

def create_default_values_page5(num_inverters, num_panels, no_of_mppt, no_of_string):
    """Create default values for string table"""
    data = []
    total_strings = no_of_mppt * no_of_string
    for i in range(num_inverters):
        row = ["INV", str(i+1)] + ["-"] * total_strings
        data.append(row)
    return data

def fill_table_page5(data, num_inverters, best_panels_per_string, remainder, no_of_col_to_be_filled, total_strings, no_of_mppt, num_panels):
    """Fill the string table with panel counts"""
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

# Set page config
st.set_page_config(
    page_title="Solar System Design & Schedule",
    page_icon="☀️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
    <style>
    .logo-container {
        display: flex;
        justify-content: center;
        margin-bottom: 20px;
    }
    .logo {
        max-width: 200px;
        height: auto;
    }
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
        font-size: 38px;
        font-weight: bold;
        color: #1f77b4;
        margin-bottom: 40px;
        text-align: center;
    }
    .stButton > button {
        width: 100%;
        background-color: #1f77b4;
        color: white;
        border: none;
        padding: 10px 20px;
        border-radius: 5px;
        font-size: 16px;
        transition: background-color 0.3s;
    }
    .stButton > button:hover {
        background-color: #1565c0;
    }
    </style>
    """, unsafe_allow_html=True)

# Initialize session state variables
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
if 'system_summary_data' not in st.session_state:
    st.session_state.system_summary_data = {
        "DC SYSTEM SIZE": "",
        "AC SYSTEM SIZE": "",
        "INVERTER (PRODUCT NAME)": "",
        "NO OF INVERTERS": "",
        "SOLAR PV MODULE (PRODUCT NAME)": "",
        "NO OF SOLAR PV MODULES": "",
        "RACKING (PRODUCT NAME)": "",
        "NO OF RACKINGS": ""
    }
if 'inverter_schedule_data' not in st.session_state:
    st.session_state.inverter_schedule_data = {
        "num_rows": 1,
        "manufacturer_model": "CPS- SCA50KTL-DO/US-480 - 50 kW",
        "kw": 50
    }
if 'string_table_data' not in st.session_state:
    st.session_state.string_table_data = {
        "num_inverters": 1,
        "num_panels": 1,
        "no_of_string": 1,
        "no_of_mppt": 1
    }

def extract_helioscope_data(pdf_content):
    """Extract data from Helioscope report using the extraction logic from app_3_report_extractor.py"""
    try:
        # Convert PDF content to text
        doc = fitz.open(stream=pdf_content, filetype="pdf")
        text = "\n".join([page.get_text() for page in doc])
        
        # Extract data using the enhanced extraction logic
        data = {}
        
        # Project Info
        project_name = re.search(r"Project Name\s+(.*)", text)
        address = re.search(r"Project\s+Address\s+(.*?)\s+USA", text, re.DOTALL)
        if project_name: data["PROJECT NAME"] = project_name.group(1).strip()
        if address: data["PROJECT ADDRESS"] = address.group(1).replace('\n', ', ').strip()
        
        # Production Info
        production = re.search(r"Annual\s+Production\s+([\d.]+)\s+MWh", text)
        ratio = re.search(r"Performance\s+Ratio\s+([\d.]+)%", text)
        if production: data["ANNUAL PRODUCTION"] = f"{production.group(1)} MWh"
        if ratio: data["PERFORMANCE RATIO"] = f"{ratio.group(1)}%"
        
        # Weather Dataset
        weather = re.search(r"Weather Dataset\s+(.+?)\s+Simulator Version", text, re.DOTALL)
        if weather: data["WEATHER DATASET"] = weather.group(1).replace("\n", " ").strip()
        
        # Extract components information
        components = []
        component_block = re.findall(
            r"(Inverters|Strings[^\n]*?|Module)\s+([A-Za-z0-9\-/,().\s]+?)\s+(\d+)\s+\(([\d.,]+)\s*(kW|ft)\)",
            text
        )
        
        for comp in component_block:
            components.append({
                "Component": comp[0].strip(),
                "Description": comp[1].strip(),
                "Count": comp[2].strip(),
                "Value": comp[3].strip(),
                "Unit": comp[4].strip()
            })
        
        # Extract system losses
        system_losses = []
        losses_pattern = r"([A-Za-z\s]+)\s+([\d.]+)%"
        losses_matches = re.findall(losses_pattern, text)
        for loss_type, loss_value in losses_matches:
            if "loss" in loss_type.lower() or "degradation" in loss_type.lower():
                system_losses.append({
                    "Loss Type": loss_type.strip(),
                    "Loss (%)": f"{loss_value}%"
                })
        data["SYSTEM LOSSES"] = system_losses
        
        # Store component details for auto-population
        st.session_state.component_details = {
            "inverter": next((c for c in components if "inverter" in c["Component"].lower()), None),
            "strings": next((c for c in components if "string" in c["Component"].lower()), None),
            "module": next((c for c in components if "module" in c["Component"].lower()), None)
        }
        
        # Store system summary data
        if st.session_state.component_details["module"]:
            st.session_state.system_summary_data = {
                "DC SYSTEM SIZE": f"{st.session_state.component_details['module']['Value']} kW",
                "AC SYSTEM SIZE": f"{st.session_state.component_details['inverter']['Value']} kW",
                "INVERTER (PRODUCT NAME)": st.session_state.component_details['inverter']['Description'],
                "NO OF INVERTERS": st.session_state.component_details['inverter']['Count'],
                "SOLAR PV MODULE (PRODUCT NAME)": st.session_state.component_details['module']['Description'],
                "NO OF SOLAR PV MODULES": st.session_state.component_details['module']['Count'],
                "RACKING (PRODUCT NAME)": "",
                "NO OF RACKINGS": ""
            }
        
        # Store inverter schedule data
        if st.session_state.component_details["inverter"]:
            st.session_state.inverter_schedule_data = {
                "num_rows": int(st.session_state.component_details['inverter']['Count']),
                "manufacturer_model": st.session_state.component_details['inverter']['Description'],
                "kw": float(st.session_state.component_details['inverter']['Value'])
            }
        
        # Store string table data
        if st.session_state.component_details["strings"] and st.session_state.component_details["inverter"]:
            st.session_state.string_table_data = {
                "num_inverters": int(st.session_state.component_details['inverter']['Count']),
                "num_panels": int(st.session_state.component_details['module']['Count']),
                "no_of_string": int(st.session_state.component_details['strings']['Count']) // int(st.session_state.component_details['inverter']['Count'])
            }
        
        return {
            **data,
            "COMPONENTS": components
        }
    except Exception as e:
        st.error(f"Error extracting data from PDF: {str(e)}")
        return None

def auto_populate_system_schedule():
    """Auto-populate system schedule with extracted data"""
    if not st.session_state.get('component_details'):
        return
    
    # Auto-populate System Summary
    if st.session_state.current_page == "System Summary":
        if st.session_state.component_details["module"]:
            st.session_state.system_summary_data = {
                "DC SYSTEM SIZE": f"{st.session_state.component_details['module']['Value']} kW",
                "AC SYSTEM SIZE": f"{st.session_state.component_details['inverter']['Value']} kW",
                "INVERTER (PRODUCT NAME)": st.session_state.component_details['inverter']['Description'],
                "NO OF INVERTERS": st.session_state.component_details['inverter']['Count'],
                "SOLAR PV MODULE (PRODUCT NAME)": st.session_state.component_details['module']['Description'],
                "NO OF SOLAR PV MODULES": st.session_state.component_details['module']['Count'],
                "RACKING (PRODUCT NAME)": "",
                "NO OF RACKINGS": ""
            }
    
    # Auto-populate Inverter Schedule
    elif st.session_state.current_page == "Inverter Schedule":
        if st.session_state.component_details["inverter"]:
            st.session_state.inverter_schedule_data = {
                "num_rows": int(st.session_state.component_details['inverter']['Count']),
                "manufacturer_model": st.session_state.component_details['inverter']['Description'],
                "kw": float(st.session_state.component_details['inverter']['Value'])
            }
    
    # Auto-populate String Table
    elif st.session_state.current_page == "String Table":
        if st.session_state.component_details["strings"] and st.session_state.component_details["inverter"]:
            st.session_state.string_table_data = {
                "num_inverters": int(st.session_state.component_details['inverter']['Count']),
                "num_panels": int(st.session_state.component_details['module']['Count']),
                "no_of_string": int(st.session_state.component_details['strings']['Count']) // int(st.session_state.component_details['inverter']['Count'])
            }

def show_design_report():
    st.title("Design Report")
    
    # Add navigation buttons at the top
    col1, col2 = st.columns([1, 1])
    with col1:
        if st.button("← Back to Home", key="design_report_back_home"):
            st.session_state.current_section = None
            st.rerun()
    with col2:
        if st.button("Go to System Schedule →", key="design_report_to_schedule"):
            st.session_state.current_section = "System Schedule"
            st.rerun()
    
    # File uploader for Helioscope PDF
    uploaded_file = st.file_uploader("Upload Helioscope PDF Report", type=["pdf"])
    
    if uploaded_file is not None:
        with st.spinner("Extracting data from PDF..."):
            # Read the PDF content
            pdf_content = uploaded_file.read()
            
            # Extract data from PDF
            extracted_data = extract_helioscope_data(pdf_content)
            
            if extracted_data:
                st.session_state.helioscope_data = extracted_data
                st.success("Data extracted successfully!")
                
                # Display extracted data
                st.subheader("Extracted Data")
                
                # Components Table
                if extracted_data.get('COMPONENTS'):
                    st.write("\n**Components Table**")
                    components_df = pd.DataFrame(extracted_data['COMPONENTS'])
                    
                    # Project Information
                    st.markdown("### 📍 Project Information")
                    st.write(f"**Project Name:** {extracted_data.get('PROJECT NAME', 'N/A')}")
                    st.write(f"**Project Address:** {extracted_data.get('PROJECT ADDRESS', 'N/A')}")
                    st.write(f"**Annual Production (MWh):** {extracted_data.get('ANNUAL PRODUCTION', 'N/A')}")
                    st.write(f"**Performance Ratio (%):** {extracted_data.get('PERFORMANCE RATIO', 'N/A')}")
                    st.write(f"**Weather Dataset:** {extracted_data.get('WEATHER DATASET', 'N/A')}")
                    
                    # Inverter Information
                    st.markdown("### 🔌 Inverter Information")
                    inverter_data = next((item for item in extracted_data['COMPONENTS'] if item['Component'] == 'Inverters'), None)
                    if inverter_data:
                        st.write(f"**Description:** {inverter_data['Description']}")
                        st.write(f"**Count:** {inverter_data['Count']}, **Capacity:** {inverter_data['Value']} kW")
                    
                    # Strings Information
                    st.markdown("### 🔗 Strings Information")
                    strings_data = next((item for item in extracted_data['COMPONENTS'] if item['Component'] == 'Strings'), None)
                    if strings_data:
                        st.write(f"**Description:** {strings_data['Description']}")
                        st.write(f"**Count:** {strings_data['Count']}, **Length:** {strings_data['Value']} {strings_data['Unit']}")
                    
                    # Module Information
                    st.markdown("### 📦 Module Information")
                    module_data = next((item for item in extracted_data['COMPONENTS'] if item['Component'] == 'Module'), None)
                    if module_data:
                        st.write(f"**Description:** {module_data['Description']}")
                        st.write(f"**Count:** {module_data['Count']}, **Capacity:** {module_data['Value']} kW")
                
                # System Losses
                if extracted_data.get('SYSTEM LOSSES'):
                    st.write("\n**System Losses**")
                    losses_df = pd.DataFrame(extracted_data['SYSTEM LOSSES'])
                    st.dataframe(losses_df)
                
                # Add button to auto-populate system schedule
                if st.button("Auto-populate System Schedule"):
                    st.session_state.auto_populated = True
                    st.session_state.current_section = "System Schedule"
                    st.rerun()
            else:
                st.error("Failed to extract data from PDF")

def show_home_page():
    # Centered logo and title
    st.markdown("""
        <div style='display: flex; flex-direction: column; align-items: center; justify-content: center; margin-bottom: 30px;'>
            <img src='https://media.licdn.com/dms/image/v2/C4D0BAQGVrL8T6AkjAg/company-logo_200_200/company-logo_200_200/0/1630563478826/metco_engineering_inc_logo?e=2147483647&v=beta&t=FfjIApcjfNOHWBq5YvwS24YukcS7o6IJi5kEC9WBR20' style='width: 150px; margin-bottom: 10px;'>
            <h1 style='color: #1f77b4; font-size: 36px; margin: 0; white-space: nowrap;'>Solar System Design & Schedule</h1>
        </div>
    """, unsafe_allow_html=True)
    
    # Create two columns for the main sections
    col1, col2 = st.columns(2)
    
    with col1:
        # Design Report Section
        st.markdown("""
            <div style='text-align: center;'>
                <h2 style='color: #1f77b4;'>Design Report</h2>
                <p style='color: #555; margin-bottom: 15px;'>Upload HelioScope PDF report and extract system data</p>
            </div>
        """, unsafe_allow_html=True)
        if st.button("Go to Design Report", key="design-report", use_container_width=True):
            st.session_state.current_section = "Design Report"
            st.rerun()
    
    with col2:
        # System Schedule Section
        st.markdown("""
            <div style='text-align: center;'>
                <h2 style='color: #1f77b4;'>System Schedule</h2>
                <p style='color: #555; margin-bottom: 15px;'>View and edit system schedule details</p>
            </div>
        """, unsafe_allow_html=True)
        if st.button("Go to System Schedule", key="system-schedule", use_container_width=True):
            st.session_state.current_section = "System Schedule"
            st.rerun()

def show_system_schedule():
    st.title("System Schedule")
    
    # Add navigation buttons at the top
    col1, col2 = st.columns([1, 1])
    with col1:
        if st.button("← Back to Home", key="system_schedule_back_home"):
            st.session_state.current_section = None
            st.rerun()
    with col2:
        if st.button("Go to Design Report →", key="system_schedule_to_design"):
            st.session_state.current_section = "Design Report"
            st.rerun()
    
    # Initialize session state variables if they don't exist
    if 'string_table_data' not in st.session_state:
        st.session_state.string_table_data = {
            "num_inverters": 1,
            "num_panels": 1,
            "no_of_string": 1,
            "no_of_mppt": 1
        }
    
    # Auto-populate data if available
    auto_populate_system_schedule()
    
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
            pdf_file = save_to_pdf_page1(data)
            with open(pdf_file, "rb") as file:
                st.session_state["pdf_data_page1"] = file.read()
                st.success("PDF generated successfully!")
        
        if st.session_state.get("pdf_data_page1"):
                st.download_button(
                    "📄 Download System Summary PDF",
                    data=st.session_state["pdf_data_page1"],
                    file_name="System_Summary.pdf",
                    mime="application/pdf"
                )
    
    elif st.session_state.current_page == "Feed Schedule":
        st.title("Feed Schedule - 600V MAX")
        tags = [
            20.4, 30.4, 40.4, 50.5, 60.4, 70.4, 80.4, 90.4, 100.4, 125.4, 150.4, 175.4, 200.4, 225.5,
            250.4, 300.4, 400.4, 500.4, 600.4, 700.4, 800.4, 900.4, 1000.4, 1600.4, 2000.4, 2500.4, 4000.4
        ]
        descriptions = [
            "3#12,#12N,#10G-3/4\"C", "3#10,#10N,#10G-3/4\"C", "3#8,#10N,#10G-1\"C", "3#6,#10N,#10G-1\"C",
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
            pdf_file = save_to_pdf_page2(df)
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
        
        with st.form("inverter_schedule_form"):
            col1, col2 = st.columns(2)
            with col1:
                num_rows = st.number_input("Number of Inverters", min_value=1, value=st.session_state.inverter_schedule_data.get("num_rows", 1), step=1)
                manufacturer_model = st.text_input("Manufacturer and Model", st.session_state.inverter_schedule_data.get("manufacturer_model", "CPS- SCA50KTL-DO/US-480 - 50 kW"))
                phase = st.number_input("Phase", min_value=1, max_value=1000, value=3, step=1)
                volts = st.number_input("Volts", min_value=100, max_value=1000, value=480, step=1)
                fla = st.text_input("FLA", "60.2")
                kw = st.number_input(
                    "Kilowatts (kW)",
                    min_value=1.0,
                    max_value=1000.0,
                    value=float(st.session_state.inverter_schedule_data.get("kw", 50)),
                    step=0.5
                )

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

            st.subheader("Preview")
            st.dataframe(df)

            submitted = st.form_submit_button("Generate PDF")
            
            if submitted:
                pdf_file = generate_pdf_page4(df)
                with open(pdf_file, "rb") as file:
                    st.session_state["pdf_page4"] = file.read()
                    st.success("PDF generated successfully!")
        
        if st.session_state.get("pdf_page4"):
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
                num_inverters = st.number_input(
                    "Enter number of inverters:", 
                    min_value=1, 
                    value=st.session_state.string_table_data.get("num_inverters", 1)
                )
                num_panels = st.number_input(
                    "Enter number of panels:", 
                    min_value=1, 
                    value=st.session_state.string_table_data.get("num_panels", 1)
                )
                panels_in_1_inverter = num_panels / num_inverters if num_inverters > 0 else 0
                st.write(f"Panels in one inverter: {panels_in_1_inverter}")
            
            with col2:
                no_of_mppt = st.number_input(
                    "No of MPPT (In 1 inverter):", 
                    min_value=1, 
                    value=st.session_state.string_table_data.get("no_of_mppt", 1)
                )
                no_of_string = st.number_input(
                    "Total No of String (In 1 MPPT):", 
                    min_value=1, 
                    value=st.session_state.string_table_data.get("no_of_string", 1)
                )
                total_strings = no_of_mppt * no_of_string
                st.write(f"Total number of Strings: {total_strings}")
            
            no_of_strings_used = st.number_input("Enter number of Strings used from report:", min_value=1, value=1)
            submit_button = st.form_submit_button("Generate Table")
        
        if submit_button:
            try:
                # Update session state with new values
                st.session_state.string_table_data = {
                    "num_inverters": num_inverters,
                    "num_panels": num_panels,
                    "no_of_mppt": no_of_mppt,
                    "no_of_string": no_of_string
                }
                
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
                    no_of_mppt,
                    num_panels
                )
                
                df = pd.DataFrame(data, columns=["TAG", "#"] + string_labels)
                df.columns = pd.MultiIndex.from_tuples(
                    [("TAG", "", ""), ("#", "", "")] + list(zip(string_labels, mppt_labels, module_labels))
                )
                
                st.subheader("Generated String Table")
                st.dataframe(df)
                
                st.session_state.string_table_df = df
        
                # Generate PDF when table is generated
                pdf_file = create_pdf_page5(st.session_state.string_table_df)
                with open(pdf_file, "rb") as file:
                    st.session_state["pdf_page5"] = file.read()
                    st.success("PDF generated successfully!")
            except Exception as e:
                st.error(f"Error generating string table: {str(e)}")
        
        if st.session_state.get("pdf_page5"):
            st.download_button(
                "📄 Download String Table PDF",
                data=st.session_state["pdf_page5"],
                file_name="stringing_table.pdf",
                mime="application/pdf"
            )

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
                                file_name="Solar_System_Report.pdf",
                                mime="application/pdf",
                                key="download_complete"
                            )
                except Exception as e:
                    st.error(f"Error creating combined PDF: {str(e)}")
                    st.warning("Please try generating the report again.")
            else:
                st.warning("Please generate all required PDFs before downloading the complete report.")

def main():
    # Home Page
    if not st.session_state.current_section:
        show_home_page()
    # Design Report Section
    elif st.session_state.current_section == "Design Report":
        show_design_report()
    # System Schedule Section
    elif st.session_state.current_section == "System Schedule":
        show_system_schedule()

if __name__ == "__main__":
    main() 