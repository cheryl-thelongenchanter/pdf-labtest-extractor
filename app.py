import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import io
from collections import Counter
import re
import os

st.set_page_config(page_title="PDF Lab Test Extractor", layout="centered")
st.title("🧪 Lab Test PDF Extractor to Excel")

# Load static lookup tables
try:
    company_df = pd.read_csv("license_to_company_lookup.csv")
    product_df = pd.read_csv("product_lookup.csv")
except FileNotFoundError:
    st.error("Static CSV files not found. Please ensure 'license_to_company_lookup.csv' and 'product_lookup.csv' are in the same directory as this app.")
    st.stop()

company_df["License Number"] = company_df["License Number"].astype(str).str.strip()
license_to_company = dict(zip(company_df["License Number"].str.strip(), company_df["Company Name"].str.strip()))

product_map = dict(zip(product_df["Product Name"].str.lower().str.strip(), product_df["Standardized Name"].str.strip()))
desc_map = dict(zip(product_df["Standardized Name"].str.strip(), product_df["Description"].str.strip()))

uploaded_files = st.file_uploader("Upload one or more Transfer Manifest PDFs", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    all_rows = []

    for uploaded_file in uploaded_files:
        pdf_doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
        full_text = ""
        for page in pdf_doc:
            full_text += page.get_text()

        # Clean up footer text if it exists
        full_text = re.sub(r"By receiving this transfer in Metrc[\s\S]+?rejecting any items.*", "", full_text, flags=re.IGNORECASE)

        # Extract License Number
        license_match = re.search(r"Originating License Number\s+(\S+)", full_text)
        license_number = license_match.group(1).strip() if license_match else ""

        # Normalize and lookup Company Name from License Number
        customer = license_to_company.get(license_number.strip(), "")

        # Extract Manifest Number
        manifest_match = re.search(r"Manifest No\.\s+(\d{10})", full_text)
        manifest_number = manifest_match.group(1).strip() if manifest_match else ""

        # Split the text into packages
        package_blocks = re.split(r"\n\d+\. Package \| Accepted", full_text)
        all_services = []

        for block in package_blocks[1:]:
            service_match = re.search(r"Req'd Lab Test Batches\s*([^\n]*)", block)
            if service_match:
                raw_services = [s.strip() for s in service_match.group(1).split(",") if s.strip()]
                for service_raw in raw_services:
                    key = service_raw.lower().strip()
                    service_std = product_map.get(key, service_raw.strip())
                    all_services.append(service_std)

        service_counts = Counter(all_services)
        unique_services = sorted(service_counts.keys())

        for i, service in enumerate(unique_services):
            description = desc_map.get(service.strip(), "")
            row = ["", customer if i == 0 else "", "", "", "", "", "", "",
                   manifest_number if i == 0 else "", license_number if i == 0 else "", service, "", "", "", service_counts[service], description]
            all_rows.append(row)

    # Create DataFrame with correct column layout
    columns = list("ABCDEFGHIJKLM") + ["L", "M", "N"]
    df = pd.DataFrame(all_rows, columns=columns)
    df["L"] = ""  # Ensure column L remains blank

    # Export to Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="LabTestData")
    output.seek(0)

    st.success("✅ Data from all PDFs extracted and enriched successfully!")
    st.download_button(
        label="📥 Download Excel File",
        data=output,
        file_name="lab_test_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
