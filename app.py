import streamlit as st
from BulkPOItemExtractor import BulkPOItemExtractor
import os
import shutil
import pandas as pd
from datetime import datetime

# Set page configuration
st.set_page_config(page_title="PO Upload and Summary", layout="centered")

# Title and instructions
st.title("Purchase Order Upload and Summary")
st.write("Upload one or more PO PDF files to generate summaries. Results will be available in an Excel file.")

# File uploader
uploaded_files = st.file_uploader("Upload PO PDFs", accept_multiple_files=True, type="pdf")

if uploaded_files:
    with st.spinner("Processing your purchase orders..."):
        # Create temporary directory for uploads
        temp_dir = "temp_uploads"
        os.makedirs(temp_dir, exist_ok=True)
        
        # Save uploaded files to temp directory
        for uploaded_file in uploaded_files:
            with open(os.path.join(temp_dir, uploaded_file.name), "wb") as f:
                f.write(uploaded_file.getbuffer())
        
        # Initialize and run the extractor
        extractor = BulkPOItemExtractor(temp_dir)
        extractor.run_analysis()
        
        # Move the output file to temp directory
        output_file = os.path.join(temp_dir, "po_analysis.xlsx")
        
        # Provide download button
        with open(output_file, "rb") as f:
            st.download_button(
                label=f"Download Summary (po_analysis.xlsx) - Generated at {datetime.now().strftime('%H:%M:%S %d-%m-%Y')}",
                data=f,
                file_name="po_analysis.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        # Clean up temporary files
        shutil.rmtree(temp_dir)
        st.success("Processing complete! Download your summary below.")

# Add a footer with current date and time
st.write(f"Last updated: {datetime.now().strftime('%I:%M %p PKT, %d %B %Y')}")