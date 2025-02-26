import streamlit as st
import pdfplumber
import re
import pandas as pd
from io import BytesIO

def load_css():
    css = """
    <style>
    @font-face {
        font-family: 'EuclidFlex';
        src: url('https://raw.githubusercontent.com/TinaMuuto/pcon-converter/main/EuclidFlex-Regular.otf') format('opentype');
    }
    body {
        font-family: 'EuclidFlex', sans-serif;
    }
    .copy-button {
        background-color: #4CAF50;
        color: white;
        border: none;
        padding: 8px 12px;
        text-align: center;
        display: inline-block;
        font-size: 14px;
        margin: 4px 2px;
        cursor: pointer;
        border-radius: 4px;
    }
    </style>
    """
    st.markdown(css, unsafe_allow_html=True)

def extract_text_from_pdf(pdf_file):
    text = ""
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"
    return text

def format_product_name(text):
    if "/" in text:
        parts = text.split("/")
        formatted = parts[0].upper() + " / " + " / ".join(part.strip().capitalize() for part in parts[1:])
        return formatted
    return text.upper()

def parse_pcon_data(text):
    lines = text.split("\n")
    extracted_data = []
    current_item = None
    capture_product_name = 2  # Antal linjer til produktnavn

    for i, line in enumerate(lines):
        line = line.strip()

        # Ignorer linjer med prisinformation og metadata
        if any(ignore in line.lower() for ignore in ["pu/eur", "it/eur", "value added tax", "gross", "position net", "measurements", "stock version"]):
            continue

        # Matcher første linje med artikelnummer og antal
        quantity_match = re.match(r"(\d+)\s+([\w\-/]+)", line)
        if quantity_match:
            quantity = int(quantity_match.group(1))
            item_number = quantity_match.group(2)
            current_item = {"Quantity": quantity, "ItemNumber": item_number, "ProductName": "", "Details": ""}
            extracted_data.append(current_item)
            capture_product_name = 2  # De næste to linjer skal bruges som produktnavn
            continue

        # Matcher de næste to linjer som produktnavn
        if capture_product_name > 0 and current_item:
            if not current_item["ProductName"]:
                current_item["ProductName"] = format_product_name(line)
            else:
                current_item["ProductName"] += f" {line}"  # Hvis produktnavn er på flere linjer
            capture_product_name -= 1
            continue

        # Matcher materialer og farver
        if "material" in line.lower() or "color" in line.lower() or "remix" in line.lower():
            if current_item and current_item["Details"]:
                current_item["Details"] += f", {line.capitalize()}"
            elif current_item:
                current_item["Details"] = line.capitalize()
            continue

    formatted_data = []
    structured_data = []
    for item in extracted_data:
        formatted_entry = f"{item['Quantity']} x {item['ProductName']}"
        if item["Details"]:
            formatted_entry += f" / {item['Details']}"
        formatted_data.append(formatted_entry)
        structured_data.append([item["ItemNumber"], item["ProductName"], item["Quantity"]])

    return formatted_data, structured_data

def generate_excel(data, headers=False):
    df = pd.DataFrame(data, columns=["Item Number", "Product Name", "Quantity"] if headers else None)
    output = BytesIO()
    df.to_excel(output, index=False, header=headers)
    output.seek(0)
    return output

def main():
    st.title("Muuto pCon PDF Converter")
    load_css()
    
    st.write("""
    ### About this tool
    This tool allows you to upload a PDF file exported from pCon and automatically extract product data. 
    The extracted data is formatted into a structured list and two downloadable Excel files.
    """)
    
    uploaded_file = st.file_uploader("Upload pCon Export PDF", type=["pdf"])
    if uploaded_file is not None:
        pdf_text = extract_text_from_pdf(uploaded_file)
        formatted_product_list, structured_data = parse_pcon_data(pdf_text)
        item_list = [[row[0], row[2]] for row in structured_data]
        excel_file_1 = generate_excel(item_list, headers=False)
        excel_file_2 = generate_excel(structured_data, headers=True)

        st.subheader("Formatted Product List")
        product_list_text = "\n".join(formatted_product_list)
        
        st.text_area("Copy all", product_list_text, height=300)
        
        for item in formatted_product_list:
            st.write(item)

        st.subheader("Download Files")
        st.download_button(label="Download Item List", data=excel_file_1, file_name="item_numbers_and_quantities.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.download_button(label="Download Detailed Product List", data=excel_file_2, file_name="detailed_product_list.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    main()
