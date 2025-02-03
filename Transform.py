import streamlit as st
import pandas as pd
from io import BytesIO
import os
import zipfile
import concurrent.futures
import warnings

# Suppress Streamlit's warning related to threading
warnings.filterwarnings("ignore", message="missing ScriptRunContext!")

# Function to process and convert a single CSV file to XLSX
def process_file(uploaded_file, delimiter, custom_xlsx_names, adjust_column_width, zipf, progress_callback):
    try:
        # Read the CSV file into a DataFrame using the selected delimiter
        df = pd.read_csv(uploaded_file, delimiter=delimiter, low_memory=False)

        # Replace NaN values with blank (empty string)
        df = df.fillna("")

        # Convert all columns to string (reduce decimal places if not needed)
        def convert_to_string(x):
            # If the value is a float and is equivalent to an integer, remove the decimal point
            if isinstance(x, (int, float)):
                return str(int(x)) if x == int(x) else str(x)  # Convert to string without '.0' for whole numbers
            return str(x)

        # Apply conversion to each column using map()
        df = df.apply(lambda col: col.map(convert_to_string))

        # Get the original CSV file name (without extension) for naming the XLSX file
        base_filename = os.path.splitext(uploaded_file.name)[0]  # Remove extension from CSV
        
        # Allow the user to specify a custom name for each file, or use default CSV filename
        if custom_xlsx_names:
            xlsx_filename = base_filename
        else:
            xlsx_filename = base_filename

        xlsx_filename = f"{xlsx_filename}.xlsx"

        # Convert DataFrame to an XLSX file and save it to the ZIP file
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name="Sheet1")

            if adjust_column_width:
                # Adjust column widths to fit content only for smaller files
                worksheet = writer.sheets["Sheet1"]
                for col in worksheet.columns:
                    max_length = 0
                    column = col[0].column_letter  # Get column name (A, B, C, etc.)
                    for cell in col:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(cell.value)
                        except:
                            pass
                    adjusted_width = (max_length + 2)
                    worksheet.column_dimensions[column].width = adjusted_width

        output.seek(0)

        # Add the XLSX file to the ZIP archive
        zipf.writestr(xlsx_filename, output.read())

        # Call the progress callback function to update progress
        progress_callback()

    except Exception as e:
        return f"Error processing file {uploaded_file.name}: {e}"

# Streamlit app
def convert_multiple_csv_to_xlsx():
    st.title("Multiple CSV to XLSX Converter")

    # Ask for the delimiter option first
    delimiter = st.selectbox(
        "Select the delimiter for your CSV files",
        options=[",", ";", "\t", "|"],
        index=0,  # Default to comma
        help="Choose the appropriate delimiter used in your CSV file"
    )

    # Upload multiple CSV files
    uploaded_files = st.file_uploader("Choose CSV files", type=["csv"], accept_multiple_files=True)

    if uploaded_files:
        # Ask for the base name for the output ZIP file
        zip_name = st.text_input("Enter a name for the output ZIP file (without extension):", "converted_files")

        # Ask if the user wants to specify a custom name for each XLSX file
        custom_xlsx_names = st.checkbox("Specify custom names for the XLSX files inside the ZIP")

        # Ask if the user wants to adjust column width (only for smaller files)
        adjust_column_width = st.checkbox("Adjust column width to fit content")

        # Progress bar for tracking file conversion progress
        progress_bar = st.progress(0)

        # Start button to trigger conversion
        start_button = st.button("Start Conversion")

        # If Start Conversion button is clicked
        if start_button:
            # Use BytesIO to create an in-memory ZIP file
            zip_buffer = BytesIO()

            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
                # Use ThreadPoolExecutor to process multiple files concurrently
                total_files = len(uploaded_files)
                completed_files = 0

                # Update progress callback
                def update_progress():
                    nonlocal completed_files
                    completed_files += 1
                    progress = int((completed_files / total_files) * 100)
                    progress_bar.progress(progress)

                with concurrent.futures.ThreadPoolExecutor() as executor:
                    futures = []
                    for uploaded_file in uploaded_files:
                        futures.append(executor.submit(process_file, uploaded_file, delimiter, custom_xlsx_names, adjust_column_width, zipf, update_progress))

                    # Wait for all futures to complete and check for any errors
                    for future in concurrent.futures.as_completed(futures):
                        result = future.result()
                        if result:
                            st.error(result)

            # Allow the user to download the ZIP file containing all the converted XLSX files
            zip_buffer.seek(0)  # Move to the beginning of the BytesIO buffer
            zip_filename = f"{zip_name}.zip"  # Use custom name for ZIP file
            st.download_button(
                label="Download all XLSX files as a ZIP",
                data=zip_buffer.read(),  # Read the contents of the ZIP file as binary data
                file_name=zip_filename,
                mime="application/zip"
            )

# Run the app
if __name__ == "__main__":
    convert_multiple_csv_to_xlsx()
