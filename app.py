import fitz   # PyMuPDF
import re
import pandas as pd
import streamlit as st
import io
from typing import Tuple, Optional, List

st.set_page_config(page_title="Drawing Dimension Extractor", page_icon="📐", layout="wide")
st.title("📐 Drawing Dimension Extractor to Excel")
st.markdown("Extract dimensional information from 2D technical drawings and export to Excel format.")

def parse_dimension(line: str) -> Tuple[str, Optional[float], str, str, str]:
    """
    Enhanced dimension parsing with better pattern recognition
    """
    # Clean the line
    line = line.strip()
    
    # Default values
    desc = "Dimension"
    nominal = None
    tol = "±0.10"  # default tolerance
    typ = ""
    inst = ""
    
    # Enhanced nominal value extraction (handles decimals, fractions, multiple numbers)
    nominal_patterns = [
        r"(\d+\.\d+)",      # Decimal numbers (priority)
        r"(\d+)",           # Whole numbers
        r"(\d+/\d+)",       # Fractions
        r"(\d+\s\d+/\d+)"   # Mixed numbers
    ]
    
    for pattern in nominal_patterns:
        m_val = re.search(pattern, line)
        if m_val:
            try:
                val_str = m_val.group(1)
                if '/' in val_str:
                    # Handle fractions
                    if ' ' in val_str:  # Mixed number
                        whole, frac = val_str.split(' ', 1)
                        num, den = frac.split('/')
                        nominal = float(whole) + float(num) / float(den)
                    else:  # Simple fraction
                        num, den = val_str.split('/')
                        nominal = float(num) / float(den)
                else:
                    nominal = float(val_str)
                break
            except (ValueError, ZeroDivisionError):
                continue
    
    # Enhanced tolerance extraction
    tol_patterns = [
        r"(±\s*\d+\.\d+)",          # ± decimal
        r"(±\s*\d+)",               # ± whole number
        r"([+]\d+\.\d+/-\d+\.\d+)", # +x.x/-x.x format
        r"([+]\d+/-\d+)",           # +x/-x format
        r"([+\-]\d+\.\d+)",         # + or - decimal
        r"([+\-]\d+)"               # + or - whole number
    ]
    
    for pattern in tol_patterns:
        m_tol = re.search(pattern, line)
        if m_tol:
            tol = m_tol.group(1).replace(' ', '')
            break
    
    # Enhanced classification with more patterns
    line_upper = line.upper()
    
    # Diameter detection
    if any(symbol in line for symbol in ["Ø", "DIA", "DIAM"]) or "DIAMETER" in line_upper:
        desc = "Diameter"
        inst = "DVC"
    
    # Radius detection
    elif line.startswith("R") or "RADIUS" in line_upper or " R " in line:
        desc = "Radius"
        inst = "VMS/IMM"
    
    # Thread detection
    elif re.search(r"M\d+", line) or "THREAD" in line_upper:
        desc = "Thread"
        inst = "Thread Gauge"
    
    # Chamfer detection
    elif ("°" in line and any(x in line_upper for x in ["X", "CHAM", "CHAMFER"])) or \
         re.search(r"\d+\s*[Xx]\s*\d+°", line):
        desc = "Chamfer"
        inst = "VMS/IMM"
    
    # Angle detection
    elif "°" in line or "ANGLE" in line_upper or "DEG" in line_upper:
        desc = "Angle"
        inst = "VMS/IMM"
    
    # Surface roughness
    elif any(symbol in line for symbol in ["Ra", "Rz", "Rt"]) or "SURFACE" in line_upper:
        desc = "Surface Roughness"
        inst = "Surface Tester"
    
    # Concentricity/Runout
    elif any(symbol in line for symbol in ["⌖", "↗"]) or "CONC" in line_upper or "RUNOUT" in line_upper:
        desc = "Concentricity/Runout"
        inst = "CMM"
    
    # Default to length
    else:
        desc = "Length"
        inst = "DVC"
    
    # Critical / Specification type detection
    if re.search(r"\bC\b", line_upper) or "CRITICAL" in line_upper:
        typ = "C"
    elif re.search(r"\bS\b", line_upper) or "SPEC" in line_upper:
        typ = "S"
    elif "KEY" in line_upper or "MAJOR" in line_upper:
        typ = "K"  # Key dimension
    
    return desc, nominal, tol, typ, inst

def extract_dimensions_from_pdf(pdf_file) -> List[List]:
    """
    Extract dimensions from PDF with improved text parsing
    """
    doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
    data = []
    
    for page_num, page in enumerate(doc, 1):
        # Try different text extraction methods
        text_methods = [
            ("text", page.get_text("text")),
            ("dict", page.get_text("dict")),
            ("blocks", page.get_text("blocks"))
        ]
        
        for method_name, text_content in text_methods:
            if method_name == "text":
                lines = text_content.splitlines()
            elif method_name == "dict":
                # Extract text from dictionary format
                lines = []
                for block in text_content.get("blocks", []):
                    if "lines" in block:
                        for line in block["lines"]:
                            for span in line.get("spans", []):
                                if span.get("text", "").strip():
                                    lines.append(span["text"].strip())
            else:  # blocks
                lines = [block[4] for block in text_content if len(block) > 4]
            
            for line in lines:
                line = line.strip()
                if not line:
                    continue
                
                # Look for balloon numbers (various formats)
                balloon_patterns = [
                    r"^(\d+)\s+(.+)",           # Standard: "1 dimension"
                    r"^\((\d+)\)\s*(.+)",       # Parentheses: "(1) dimension"
                    r"^(\d+)[\.\-\:]\s*(.+)",   # With separators: "1. dimension"
                    r"(\d+)\s*[^\d\w]*(.+)"     # Flexible: "1 - dimension"
                ]
                
                for pattern in balloon_patterns:
                    match = re.match(pattern, line)
                    if match:
                        try:
                            sr_no = int(match.group(1))
                            dim_text = match.group(2).strip()
                            
                            if dim_text and len(dim_text) > 1:  # Valid dimension text
                                desc, nominal, tol, typ, inst = parse_dimension(dim_text)
                                data.append([sr_no, desc, nominal, tol, typ, inst, f"Page {page_num}"])
                                break
                        except (ValueError, IndexError):
                            continue
            
            # If we found data with this method, use it
            if data:
                break
    
    doc.close()
    return data

# Sidebar for configuration
st.sidebar.header("Configuration")
show_preview = st.sidebar.checkbox("Show dimension preview", value=True)
default_tolerance = st.sidebar.text_input("Default tolerance", value="±0.10")
include_page_ref = st.sidebar.checkbox("Include page reference", value=True)

# File upload
uploaded_file = st.file_uploader(
    "Upload a 2D Drawing (PDF)", 
    type=["pdf"],
    help="Upload a PDF file containing technical drawings with dimensional callouts"
)

if uploaded_file:
    # Show file info
    file_details = {
        "Filename": uploaded_file.name,
        "File size": f"{uploaded_file.size / 1024:.2f} KB",
        "File type": uploaded_file.type
    }
    
    with st.expander("📋 File Information"):
        for key, value in file_details.items():
            st.write(f"**{key}:** {value}")
    
    # Process the PDF
    with st.spinner("Extracting dimensions from PDF..."):
        try:
            data = extract_dimensions_from_pdf(uploaded_file)
            
            if data:
                # Create DataFrame
                columns = ["Sr. No.", "Parameter", "Nominal Value", "Tolerance", "Type (C/S)", "Instrument"]
                if include_page_ref:
                    columns.append("Page")
                
                df = pd.DataFrame(data, columns=columns)
                df = df.sort_values("Sr. No.").reset_index(drop=True)
                
                # Display statistics
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("Total Dimensions", len(df))
                with col2:
                    st.metric("Critical Dimensions", len(df[df["Type (C/S)"] == "C"]))
                with col3:
                    st.metric("Unique Parameters", df["Parameter"].nunique())
                with col4:
                    st.metric("Pages Processed", df["Page"].nunique() if include_page_ref else "N/A")
                
                # Show parameter distribution
                if show_preview:
                    st.subheader("📊 Parameter Distribution")
                    param_counts = df["Parameter"].value_counts()
                    st.bar_chart(param_counts)
                
                # Display the dataframe
                st.subheader("📋 Extracted Dimensions")
                
                # Add filters
                col1, col2 = st.columns(2)
                with col1:
                    param_filter = st.multiselect(
                        "Filter by Parameter Type:",
                        options=df["Parameter"].unique(),
                        default=df["Parameter"].unique()
                    )
                with col2:
                    type_filter = st.multiselect(
                        "Filter by Type:",
                        options=df["Type (C/S)"].unique(),
                        default=df["Type (C/S)"].unique()
                    )
                
                # Apply filters
                filtered_df = df[
                    (df["Parameter"].isin(param_filter)) &
                    (df["Type (C/S)"].isin(type_filter))
                ]
                
                # Display filtered data
                st.dataframe(
                    filtered_df,
                    use_container_width=True,
                    hide_index=True
                )
                
                # Export to Excel with enhanced formatting
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    # Write main data
                    filtered_df.to_excel(writer, sheet_name='Dimensions', index=False)
                    
                    # Create summary sheet
                    summary_data = {
                        'Parameter Type': param_counts.index.tolist(),
                        'Count': param_counts.values.tolist()
                    }
                    summary_df = pd.DataFrame(summary_data)
                    summary_df.to_excel(writer, sheet_name='Summary', index=False)
                    
                    # Format the main sheet
                    workbook = writer.book
                    worksheet = writer.sheets['Dimensions']
                    
                    # Add formatting
                    header_format = workbook.add_format({
                        'bold': True,
                        'text_wrap': True,
                        'valign': 'top',
                        'fg_color': '#D7E4BC',
                        'border': 1
                    })
                    
                    # Apply header formatting
                    for col_num, value in enumerate(filtered_df.columns.values):
                        worksheet.write(0, col_num, value, header_format)
                    
                    # Auto-adjust column widths
                    for i, col in enumerate(filtered_df.columns):
                        max_length = max(
                            filtered_df[col].astype(str).map(len).max(),
                            len(col)
                        )
                        worksheet.set_column(i, i, min(max_length + 2, 20))
                
                output.seek(0)
                
                st.download_button(
                    label="📥 Download Excel File",
                    data=output.getvalue(),
                    file_name=f"extracted_dimensions_{uploaded_file.name.split('.')[0]}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                # Show sample entries for verification
                if show_preview and len(filtered_df) > 0:
                    st.subheader("🔍 Sample Entries")
                    st.dataframe(
                        filtered_df.head(5),
                        use_container_width=True,
                        hide_index=True
                    )
                
            else:
                st.warning("⚠️ No dimensional data found in the PDF. Please check if:")
                st.write("- The PDF contains technical drawings with numbered dimensions")
                st.write("- The dimensions follow standard callout formats")
                st.write("- The PDF is text-searchable (not a scanned image)")
                
        except Exception as e:
            st.error(f"❌ Error processing PDF: {str(e)}")
            st.write("Please try with a different PDF file or check the file format.")

else:
    st.info("👆 Please upload a PDF file to begin extraction")
    
    # Show example of expected format
    st.subheader("📖 Expected Format")
    st.write("The tool expects dimensions in the following formats:")
    
    example_data = [
        ["1", "Length", "25.4", "±0.1", "C", "DVC"],
        ["2", "Diameter", "Ø12.0", "±0.05", "S", "DVC"],
        ["3", "Radius", "R5.0", "±0.1", "", "VMS/IMM"],
        ["4", "Angle", "45°", "±1°", "", "VMS/IMM"],
        ["5", "Chamfer", "2 x 45°", "±0.1", "", "VMS/IMM"]
    ]
    
    example_df = pd.DataFrame(
        example_data, 
        columns=["Sr. No.", "Parameter", "Nominal Value", "Tolerance", "Type (C/S)", "Instrument"]
    )
    st.dataframe(example_df, hide_index=True)