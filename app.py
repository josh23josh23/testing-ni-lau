import io
import zipfile
import streamlit as st
import fitz  # PyMuPDF
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from io import BytesIO

# Define a list of predefined keywords
KEYWORDS = [
    "Activity Centre", "Amendment", "Amendments Report", "Annual Plan", "Annual Report",
    "Area Plan", "Assessments", "Broadacre", "Budget", "City Plan", "Code Amendment",
    "Concept Plan", "Corporate Business Plan", "Corporate Plan", "Council Action Plan",
    "Council Business Plan", "Council Plan", "Council Report", 
    "Development Investigation Area", "Development Plan", "Development Plan Amendment",
    "DPA", "Emerging community", "Employment land study", "Exhibition", "expansion",
    "Framework", "Framework plan", "Gateway Determination", "greenfield", "growth area", 
    "growth plan", "growth plans", "housing", "Housing Strategy",
    "Industrial land study", "infrastructure plan", "infrastructure planning", 
    "Inquiries", "Investigation area", "land use", "Land use strategy",
    "LDP", "Local Area Plan", "Local Development Area", "Local Development Plan",
    "Local Environmental Plan", "Local Planning Policy", "Local Planning Scheme",
    "Local Planning Strategy", "Local Strategic Planning Statement", "LPP", "LPS", "LSPS",
    "Major Amendment", "Major Update", "Master Plan", "Masterplan", "Neighbourhood Plan",
    "Operational Plan", "Planning Commission", "Planning Framework", "Planning Investigation Area",
    "Planning proposal", "Planning report", "Planning Scheme", "Planning Scheme Amendment",
    "Planning Strategy", "Precinct plan", "Priority Development Area",
    "Project Vision", "Rezoning", "settlement", "Strategy", "Structure Plan", "Structure Planning",
    "Study", "Territory plan", "Town Planning Scheme",
    "Township Plan", "TPS", "Urban Design Framework", "Urban growth", "Urban Release",
    "Urban renewal", "Variation", "Vision"
]

# Function to highlight text in PDF and track keyword occurrences
def highlight_text_in_pdf(uploaded_file, keywords, original_filename):
    pdf_file = io.BytesIO(uploaded_file.read())
    pdf_document = fitz.open(stream=pdf_file, filetype="pdf")

    # Dictionary to track keyword occurrences by page
    keyword_occurrences = {keyword: [] for keyword in keywords}
    keywords_found = False

    for page_num in range(len(pdf_document)):
        page = pdf_document.load_page(page_num)
        text = page.get_text("dict")

        for keyword in keywords:
            keyword_lower = keyword.lower()

            for block in text["blocks"]:
                if block["type"] == 0:  # Block is text
                    for line in block["lines"]:
                        for span in line["spans"]:
                            text_content = span["text"]
                            lower_text = text_content.lower()

                            start = 0
                            while True:
                                start = lower_text.find(keyword_lower, start)
                                if start == -1:
                                    break

                                # Track page number for each keyword occurrence
                                if page_num + 1 not in keyword_occurrences[keyword]:
                                    keyword_occurrences[keyword].append(page_num + 1)
                                    keywords_found = True  # At least one keyword found

                                # Highlight the keyword in the PDF
                                bbox = span["bbox"]
                                span_width = bbox[2] - bbox[0]
                                span_height = bbox[3] - bbox[1]
                                char_width = span_width / len(text_content)

                                keyword_bbox = fitz.Rect(
                                    bbox[0] + char_width * start,
                                    bbox[1],
                                    bbox[0] + char_width * (start + len(keyword_lower)),
                                    bbox[3]
                                )
                                
                                keyword_bbox = keyword_bbox.intersect(fitz.Rect(0, 0, page.rect.width, page.rect.height))
                                
                                if not keyword_bbox.is_empty:
                                    highlight = page.add_highlight_annot(keyword_bbox)
                                    highlight.set_colors(stroke=(1, 0.65, 0))  # Set color to orange
                                    highlight.update()

                                start += len(keyword_lower)

    # Save the highlighted PDF
    output_pdf = BytesIO()
    pdf_document.save(output_pdf)
    output_pdf.seek(0)

    if not keywords_found:
        return None, None, False  # No keywords found

    # Determine the maximum number of occurrences for any keyword
    max_occurrences = max(len(pages) for pages in keyword_occurrences.values() if pages)

    # Sort keywords alphabetically and prepare data for Excel
    sorted_keywords = sorted(
        (keyword for keyword in keyword_occurrences if keyword_occurrences[keyword]),
        key=lambda x: x.lower()
    )
    
    # Create a new Excel workbook and worksheet
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Keywords Report"

    # Write header with dynamic column names
    header = ["Keyword"] + [f"Occurrence {i+1}" for i in range(max_occurrences)]
    ws.append(header)

    # Write keyword occurrences
    for keyword in sorted_keywords:
        pages = keyword_occurrences[keyword]
        row = [keyword] + pages
        # Fill in empty columns if fewer occurrences are found
        while len(row) < len(header):
            row.append('')
        ws.append(row)

    # Auto-size columns based on content length
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter  # Get column name (e.g., 'A')
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)
        ws.column_dimensions[column].width = adjusted_width

    # Save workbook to BytesIO
    excel_output = BytesIO()
    wb.save(excel_output)
    excel_output.seek(0)

    new_filename = f"highlighted_{original_filename}"
    excel_filename = f"keywords_report_{original_filename.replace('.pdf', '.xlsx')}"

    # Create a zip file containing both files
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w') as zip_file:
        # Save PDF to zip
        zip_file.writestr(new_filename, output_pdf.getvalue())
        # Save Excel report to zip
        zip_file.writestr(excel_filename, excel_output.getvalue())
    
    zip_buffer.seek(0)
    zip_filename = f"highlighted_and_report_{original_filename.replace('.pdf', '.zip')}"

    return zip_buffer, zip_filename, keywords_found

# Main tool interface
def keyword_highlighter_page():
    st.title("PDF Keyword Highlighter")

    uploaded_file = st.file_uploader("Choose a PDF file", type="pdf")

    if uploaded_file:
        st.subheader("Select Keywords to Highlight")

        if "selected_keywords" not in st.session_state:
            st.session_state.selected_keywords = []

        # Add a "Select All" checkbox
        select_all = st.checkbox("Select All Keywords")
        
        col1, col2, col3 = st.columns(3)
        selected_keywords = st.session_state.selected_keywords

        # Automatically select or deselect all keywords if "Select All" is checked
        if select_all:
            selected_keywords = KEYWORDS.copy()
        else:
            selected_keywords = []

        with col1:
            for keyword in KEYWORDS[:len(KEYWORDS)//3]:
                if st.checkbox(keyword, value=(keyword in selected_keywords)):
                    if keyword not in selected_keywords:
                        selected_keywords.append(keyword)
                else:
                    if keyword in selected_keywords:
                        selected_keywords.remove(keyword)

        with col2:
            for keyword in KEYWORDS[len(KEYWORDS)//3:2*len(KEYWORDS)//3]:
                if st.checkbox(keyword, value=(keyword in selected_keywords)):
                    if keyword not in selected_keywords:
                        selected_keywords.append(keyword)
                else:
                    if keyword in selected_keywords:
                        selected_keywords.remove(keyword)

        with col3:
            for keyword in KEYWORDS[2*len(KEYWORDS)//3:]:
                if st.checkbox(keyword, value=(keyword in selected_keywords)):
                    if keyword not in selected_keywords:
                        selected_keywords.append(keyword)
                else:
                    if keyword in selected_keywords:
                        selected_keywords.remove(keyword)

        custom_keywords = st.text_area("Or add your own keywords (one per line):", "")
        if custom_keywords:
            custom_keywords_list = [kw.strip() for kw in custom_keywords.split('\n') if kw.strip()]
            selected_keywords.extend(custom_keywords_list)

        if st.button("Highlight Keywords"):
            if not selected_keywords:
                st.error("Please select or add at least one keyword.")
            else:
                zip_buffer, zip_filename, keywords_found = highlight_text_in_pdf(uploaded_file, selected_keywords, uploaded_file.name)

                if not keywords_found:
                    st.warning("No keywords found in the PDF.")
                else:
                    st.download_button(
                        label="Download Highlighted PDF and Report",
                        data=zip_buffer,
                        file_name=zip_filename,
                        mime="application/zip"
                    )

# Main function
def main():
    keyword_highlighter_page()

if __name__ == "__main__":
    main()
