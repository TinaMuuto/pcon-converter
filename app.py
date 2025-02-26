import streamlit as st
import pdfplumber
import re
import pandas as pd
from io import BytesIO

def extract_text_from_pdf(pdf_file):
    text = ""
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text += page.extract_text() + "\n"
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
        quantity_match = re.match(r"(\d+)\s+([\d\-]+)\s+([\d,]+)", line)
        if quantity_match:
            quantity = int(float(quantity_match.group(3).replace(",", ".")))
            current_item = {"Quantity": quantity, "Description": "", "Details": "", "ItemNumber": quantity_match.group(2)}
            extracted_data.append(current_item)
            continue

        if current_item is not None and "/" in line:
            current_item["Description"] = format_text(line.strip())
            continue

        if current_item is not None and ("Material" in line or "Color" in line or "Fabric" in line):
            if current_item["Details"]:
                current_item["Details"] += f", {line.strip().capitalize()}"
            else:
                current_item["Details"] = line.strip().capitalize()

    formatted_data = []
    structured_data = []
    for item in extracted_data:
        if item["Description"]:
            formatted_entry = f"{item['Quantity']} x {item['Description']}"
            if item["Details"]:
                formatted_entry += f" / {item['Details']}"
            formatted_data.append(formatted_entry)
            
            # Extract product name (text before first /)
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
    st.title("pCon PDF Converter")
    st.write("""
    ### About this tool
    This tool allows you to upload a pCon export PDF and automatically extract product data. 
    The extracted data is formatted into a structured list and two downloadable Excel files.
    
    **How it works:**
    1. Upload a pCon export PDF.
    2. The tool will process the file and extract relevant product details.
    3. You will see a formatted product list below.
    4. Download the output as either a **basic item list** (Item Number & Quantity) or a **detailed product list** (Item Number, Product Name, Quantity).
    
    **Example output:**
    - 3 x STACKED STORAGE SYSTEM / PLINTH - 131 X 35 H: 10 CM
    - 4 x STACKED STORAGE SYSTEM / LARGE / Material: Oak veneered MDF.
    - 2 x FIVE POUF / LARGE / Remix: 113
    
    **Example of pCon PDF format:**
    """)
    
    st.file_uploader("Example PDF", type=["pdf"], disabled=True)
    st.write("[Download Example PDF](sandbox:/mnt/data/pconexample.pdf)")
    
    uploaded_file = st.file_uploader("Upload pCon Export PDF", type=["pdf"])
    if uploaded_file is not None:
        pdf_text = extract_text_from_pdf(uploaded_file)
        formatted_product_list, structured_data = parse_pcon_data(pdf_text)
        item_list = [[row[0], row[2]] for row in structured_data]  # Extract item number and quantity
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
