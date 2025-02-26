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

def parse_pcon_data(text):
    lines = text.split("\n")
    extracted_data = []
    current_item = None

    for line in lines:
        quantity_match = re.match(r"(\d+)\s+([\d\-]+)\s+([\d,]+)", line)
        if quantity_match:
            quantity = int(float(quantity_match.group(3).replace(",", ".")))
            current_item = {"Quantity": quantity, "Description": "", "Details": ""}
            extracted_data.append(current_item)
            continue

        if current_item is not None and "/" in line:
            current_item["Description"] = line.strip()
            continue

        if current_item is not None and ("Material" in line or "Color" in line or "Fabric" in line):
            if current_item["Details"]:
                current_item["Details"] += f", {line.strip()}"
            else:
                current_item["Details"] = line.strip()

    formatted_data = []
    for item in extracted_data:
        if item["Description"]:
            formatted_entry = f"{item['Quantity']} x {item['Description']}"
            if item["Details"]:
                formatted_entry += f" / {item['Details']}"
            formatted_data.append(formatted_entry)

    return formatted_data

def extract_item_numbers_and_quantities(text):
    lines = text.split("\n")
    extracted_data = []
    for line in lines:
        match = re.match(r"(\d+)\s+([\d\-]+)\s+([\d,]+)", line)
        if match:
            quantity = int(float(match.group(3).replace(",", ".")))
            item_number = match.group(2)
            extracted_data.append([item_number, quantity])
    return extracted_data

def generate_excel(data):
    df = pd.DataFrame(data)
    output = BytesIO()
    df.to_excel(output, index=False, header=False)
    output.seek(0)
    return output

def main():
    st.title("pCon PDF Converter")
    uploaded_file = st.file_uploader("Upload pCon Export PDF", type=["pdf"])
    if uploaded_file is not None:
        pdf_text = extract_text_from_pdf(uploaded_file)
        formatted_product_list = parse_pcon_data(pdf_text)
        item_list = extract_item_numbers_and_quantities(pdf_text)
        excel_file = generate_excel(item_list)

        st.subheader("Formatted Product List")
        for item in formatted_product_list:
            st.write(item)

        st.subheader("Download Excel File")
        st.download_button(label="Download Item List", data=excel_file, file_name="item_numbers_and_quantities.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    main()
