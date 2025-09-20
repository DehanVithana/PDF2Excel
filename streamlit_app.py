import streamlit as st
import tabula
import pandas as pd
import io
import zipfile
import subprocess
import os

# Set up the page configuration
st.set_page_config(
    page_title="PDF to Excel Converter",
    page_icon="📄",
    layout="centered"
)

# --- Functions ---

def convert_pdf_to_excel(pdf_file, excel_buffer):
    """
    Reads tables from a PDF file and writes them to an Excel file buffer.
    Each detected table is written to a new sheet.
    """
    try:
        # Check if the 'java' command is available
        try:
            subprocess.run(["java", "-version"], check=True, capture_output=True)
        except (subprocess.CalledProcessError, FileNotFoundError):
            st.error("It looks like the 'java' command is not found. Please ensure Java is installed and configured correctly.")
            st.info("On Streamlit Cloud, you need to create a `packages.txt` file with `openjdk-11-jdk` inside it to install Java.")
            return

        # Read all tables from the PDF.
        # This will return a list of pandas DataFrames.
        tables = tabula.read_pdf(
            pdf_file,
            pages='all',
            multiple_tables=True,
            stream=True, # Use 'stream' mode for un-bordered tables
            guess=True,  # Automatically detect table areas
            lattice=False # Use 'lattice' for bordered tables. Can switch based on PDF.
        )
        
        # Check if any tables were detected
        if not tables:
            st.warning(f"No tables were detected in {pdf_file.name}. An empty Excel file will be created.")
            # Create an empty DataFrame with a message
            df_empty = pd.DataFrame({"Message": ["No tables found in PDF."]})
            tables.append(df_empty)
            
        # Create an Excel writer object to write to the in-memory buffer
        with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
            # Iterate through the list of tables and write each to a new sheet
            for i, df in enumerate(tables):
                # Clean up empty rows and columns
                df.dropna(how='all', axis=0, inplace=True)
                df.dropna(how='all', axis=1, inplace=True)
                
                # Check if the DataFrame is not empty after cleanup
                if not df.empty:
                    sheet_name = f"Table {i + 1}"
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                
    except Exception as e:
        st.error(f"An error occurred while processing {pdf_file.name}: {e}")
        # In case of an error, create an Excel file with an error message
        df_error = pd.DataFrame({"Error": [f"Failed to process PDF: {e}"]})
        with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
            df_error.to_excel(writer, sheet_name="Error", index=False)
            
    # Move the buffer's cursor to the beginning
    excel_buffer.seek(0)
    return excel_buffer

# --- UI Layout ---

st.title("📄 PDF to Excel Converter")
st.write("Upload one or more PDF files, and I will convert them into Excel spreadsheets.")
st.write("The app will automatically detect tables in your PDFs, even from unstructured documents, and export them. If you upload multiple files, they will be combined into a single ZIP file for download.")

# File uploader widget, allows multiple files
uploaded_files = st.file_uploader(
    "Choose PDF files...",
    type="pdf",
    accept_multiple_files=True
)

# Button to trigger the conversion
if st.button("Convert to Excel"):
    if uploaded_files:
        # Show a spinner while processing
        with st.spinner("Processing your PDFs..."):
            
            # Use an in-memory buffer to store the files
            excel_files = {}

            # Process each uploaded file
            for uploaded_file in uploaded_files:
                # Create an in-memory buffer for the output Excel file
                excel_buffer = io.BytesIO()
                
                # Perform the conversion
                convert_pdf_to_excel(uploaded_file, excel_buffer)
                
                # Store the result in a dictionary with the filename
                excel_filename = uploaded_file.name.replace('.pdf', '.xlsx')
                excel_files[excel_filename] = excel_buffer.getvalue()

        # Check if we have files to offer for download
        if excel_files:
            if len(excel_files) == 1:
                # If only one file, provide a single download button
                filename = list(excel_files.keys())[0]
                data = list(excel_files.values())[0]
                st.success("Conversion complete! Your file is ready to download.")
                st.download_button(
                    label=f"Download {filename}",
                    data=data,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                # If multiple files, create a zip file
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                    for filename, data in excel_files.items():
                        zip_file.writestr(filename, data)
                
                st.success("Conversion complete! All your files are ready in a ZIP archive.")
                # Provide a download button for the zip file
                st.download_button(
                    label="Download All as ZIP",
                    data=zip_buffer.getvalue(),
                    file_name="converted_pdfs.zip",
                    mime="application/zip"
                )
                
        else:
            st.error("Something went wrong. Please try uploading your files again.")

    else:
        st.warning("Please upload at least one PDF file to begin the conversion.")

# --- Information Section ---
st.markdown("---")
st.markdown(
    """
    ### About the App
    This application uses `tabula-py` and `pandas` to extract tabular data from your PDFs.
    It works best on documents with clear table structures but can also make a good guess on less structured data.
    
    You can deploy this application on Streamlit's Community Cloud by simply pushing this script to a GitHub repository.
    """
)
st.markdown("---")
