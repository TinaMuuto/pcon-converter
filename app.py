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

def format_text(description):
    if "/" in description:
        parts = description.split("/")
        formatted = parts[0].upper() + " / " + " / ".join(part.strip().capitalize() for part in parts[1:])
        return formatted
    return description.upper()

def parse_pcon_data(text):
    lines = text.split("\n")
    extracted_data = []
    current_item = None

    for line in lines:
        line = line.strip()

        # Ignorer linjer med overskrifter og prisinformation
        if any(ignore in line.lower() for ignore in ["pos article code", "description quantity", "eur", "position net", "value added tax", "gross"]):
            continue

        quantity_match = re.match(r"(\d+),?(\d*)\s+([\d\-]+)", line)
        if quantity_match:
            quantity = int(quantity_match.group(1))  # Henter kun heltal
            item_number = quantity_match.group(3)
            current_item = {"Quantity": quantity, "Description": "", "Details": "", "ItemNumber": item_number}
            extracted_data.append(current_item)
            continue

        if current_item is not None and "/" in line:
            current_item["Description"] = format_text(line)
            continue

        if current_item is not None and any(x in line.lower() for x in ["material", "color", "fabric"]):
            if current_item["Details"]:
                current_item["Details"] += f", {line.capitalize()}"
            else:
                current_item["Details"] = line.capitalize()

    formatted_data = []
    structured_data = []
    for item in extracted_data:
        if item["Description"]:
            formatted_entry = f"{item['Quantity']} x {item['Description']}"
            if item["Details"]:
                formatted_entry += f" / {item['Details']}"
            formatted_data.append(formatted_entry)
            
            product_name = item["Description"].split("/")[0].strip()
            structured_data.append([item["ItemNumber"], product_name, item["Quantity"]])

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
    
    st.write("**Example of pCon PDF format:**")
    st.markdown("[📄 Download Example PDF](https://github.com/TinaMuuto/pcon-converter/raw/main/pconexample.pdf)", unsafe_allow_html=True)
    
    uploaded_file = st.file_uploader("Upload pCon Export PDF", type=["pdf"])
    if uploaded_file is not None:
        pdf_text = extract_text_from_pdf(uploaded_file)
        formatted_product_list, structured_data = parse_pcon_data(pdf_text)
        item_list = [[row[0], row[2]] for row in structured_data]
        excel_file_1 = generate_excel(item_list, headers=False)
        excel_file_2 = generate_excel(structured_data, headers=True)

        st.subheader("Formatted Product List")
        product_list_text = "\n".join(formatted_product_list)
        
        st.button("📋 Copy to Clipboard", on_click=lambda: st.session_state.update(clipboard=product_list_text))
        
        for item in formatted_product_list:
            st.write(item)

        st.subheader("Download Files")
        st.download_button(label="Download Item List", data=excel_file_1, file_name="item_numbers_and_quantities.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.download_button(label="Download Detailed Product List", data=excel_file_2, file_name="detailed_product_list.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    main()
