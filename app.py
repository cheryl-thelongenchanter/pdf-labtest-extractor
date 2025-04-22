import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import io
from collections import Counter
import re

st.set_page_config(page_title="PDF Lab Test Extractor", layout="centered")
st.title("🧪 Lab Test PDF Extractor to Excel")

uploaded_file = st.file_uploader("Upload a Transfer Manifest PDF", type=["pdf"])

if uploaded_file:
    # Read the PDF
    pdf_doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
    full_text = ""
    for page in pdf_doc:
        full_text += page.get_text()

    # Clean up footer text if it exists
    full_text = re.sub(r"By receiving this transfer in Metrc[\s\S]+?rejecting any items.*", "", full_text, flags=re.IGNORECASE)

    # Extract Customer Name
    customer_match = re.search(r"Originating\s+Entity\s+(.*?)\s+For MED", full_text)
    customer = customer_match.group(1).strip() if customer_match else ""

    # Extract License Number
    license_match = re.search(r"Originating License Number\s+(\S+)", full_text)
    license_number = license_match.group(1).strip() if license_match else ""

    # Extract Manifest Number
    manifest_match = re.search(r"Manifest No\.\s+(\d{10})", full_text)
    manifest_number = manifest_match.group(1).strip() if manifest_match else ""

    # Define valid services (updated product/service list)
    valid_services = [
        "Homogeneity",
        "Metals",
        "Microbial Contaminant",
        "Microbial Contaminant For Edible/Topical Products Only",
        "Microbial Contaminant For Remediated Concentrates Only",
        "Mycotoxin",
        "Pesticides",
        "Potency",
        "R & D Testing",
        "Residual Solvents",
        "Water Activity"
    ]

    # Split the text into packages
    package_blocks = re.split(r"\n\d+\. Package \| Accepted", full_text)
    all_services = []

    for block in package_blocks[1:]:
        service_match = re.search(r"Req'd Lab Test Batches\s*([^\n]*)", block)
        if service_match:
            raw_services = [s.strip() for s in service_match.group(1).split(",") if s.strip()]
            for service in raw_services:
                cleaned_service = service.lower().replace("  ", " ").strip()
                for valid in valid_services:
                    if valid.lower() in cleaned_service:
                        all_services.append(valid)
                        break

    # Count frequencies
    service_counts = Counter(all_services)
    unique_services = sorted(service_counts.keys())

    # Build rows with top row containing customer, manifest, and license
    rows = []
    for i, service in enumerate(unique_services):
        row = ["", customer if i == 0 else "", "", "", "", "", "", "",
               manifest_number if i == 0 else "", license_number if i == 0 else "", service, "", "", service_counts[service]]
        rows.append(row)

    # Create DataFrame with correct column layout
    columns = list("ABCDEFGHIJKLMN")  # A to N
    df = pd.DataFrame(rows, columns=columns)
    df["L"] = ""  # Overwrite column L explicitly with blanks

    # Export to Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="LabTestData")
    output.seek(0)

    st.success("✅ Data extracted successfully!")
    st.download_button(
        label="📥 Download Excel File",
        data=output,
        file_name="lab_test_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )