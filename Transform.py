import streamlit as st
import pandas as pd
from io import BytesIO
import os
import zipfile
import re  # Used for valid sheet names
from Mainconversion import preview_csv  # Assuming this function is defined in your Mainconversion.py

# Function to create a valid sheet name from the file name
def create_valid_sheet_name(file_name):
    name = os.path.splitext(file_name)[0]  # Remove extension
    name = re.sub(r'[\\/*?:\[\]]', '', name)  # Remove invalid characters
    return name[:31]  # Limit to 31 characters

# Function to process and convert multiple CSV files into multiple XLSX files

def process_multiple_files_to_xlsx(uploaded_files, delimiter, custom_xlsx_names, adjust_column_width, zipf, progress_callback):
    try:
        total_files = len(uploaded_files)
        for idx, uploaded_file in enumerate(uploaded_files):
            # Read the CSV file into a DataFrame using the selected delimiter
            df = pd.read_csv(uploaded_file, delimiter=delimiter, low_memory=False)

            # Replace NaN values with blank (empty string)
            df = df.fillna("")

            # Convert all columns to string (reduce decimal places if not needed)
            def convert_to_string(x):
                if isinstance(x, (int, float)):
                    return str(int(x)) if x == int(x) else str(x)  # Convert to string without '.0' for whole numbers
                return str(x)

            # Apply conversion to each column using map()
            df = df.apply(lambda col: col.map(convert_to_string))

            # Get the original CSV file name (without extension) for naming the XLSX file
            base_filename = os.path.splitext(uploaded_file.name)[0]  # Remove extension from CSV
            
            # If the user wants custom XLSX sheet names, use the base filename
            sheet_name = create_valid_sheet_name(uploaded_file.name)  # Create valid sheet name from the CSV file name

            # Write DataFrame to Excel sheet
            xlsx_filename = f"{base_filename}.xlsx"
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name=sheet_name)

                # Adjust column width if necessary
                if adjust_column_width:
                    worksheet = writer.sheets[sheet_name]
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
            progress = int(((idx + 1) / total_files) * 100)  # Calculate progress as percentage
            progress_callback(progress)

    except Exception as e:
        return f"Error processing files: {e}"


# Function to process and convert multiple CSV files into one Excel file with multiple sheets
def process_multiple_files_to_single_excel(uploaded_files, delimiter, custom_xlsx_names, adjust_column_width, progress_callback):
    try:
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for idx, uploaded_file in enumerate(uploaded_files):
                # Read the CSV file into a DataFrame using the selected delimiter
                df = pd.read_csv(uploaded_file, delimiter=delimiter, low_memory=False)

                # Replace NaN values with blank (empty string)
                df = df.fillna("")

                # Convert all columns to string (reduce decimal places if not needed)
                def convert_to_string(x):
                    if isinstance(x, (int, float)):
                        return str(int(x)) if x == int(x) else str(x)  # Convert to string without '.0' for whole numbers
                    return str(x)

                # Apply conversion to each column using map()
                df = df.apply(lambda col: col.map(convert_to_string))

                # Get the original CSV file name (without extension) for naming the sheet
                sheet_name = create_valid_sheet_name(uploaded_file.name)  # Create valid sheet name
                
                # If the user wants custom XLSX sheet names, use the base filename
                if not custom_xlsx_names:
                    sheet_name = f"Sheet{idx + 1}"

                # Write DataFrame to Excel sheet
                df.to_excel(writer, index=False, sheet_name=sheet_name)

                # Adjust column width if necessary
                if adjust_column_width:
                    worksheet = writer.sheets[sheet_name]
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

                # Call progress callback to update the progress bar
                progress_callback((idx + 1) / len(uploaded_files))

        output.seek(0)
        return output

    except Exception as e:
        return f"Error processing files: {e}"

# Streamlit app
def main():
    st.title("CSV to Excel Converter")
    st.sidebar.title("Configurations")
    # Sidebar selection to choose between the two functionalities
    app_mode = st.sidebar.selectbox(
        "Choose Conversion Mode", 
        options=["Multiple CSV to Multiple XLSX", "Multiple CSV to Single XLSX"]
    )

    # Common configuration options
    delimiter = st.sidebar.selectbox(
        "Select the delimiter for your CSV files",
        options=[",", ";", "\t", "|"],
        index=0,  # Default to comma
        help="Choose the appropriate delimiter used in your CSV file"
    )

    # Upload multiple CSV files
    uploaded_files = st.sidebar.file_uploader("Choose CSV files", type=["csv"], accept_multiple_files=True)

    if uploaded_files:
        # Ask for the base name for the output file
        output_name = st.sidebar.text_input("Enter a name for the output file (without extension):", "output")

        # Ask if the user wants to specify custom names for each XLSX sheet/file
        custom_xlsx_names = st.sidebar.checkbox("Custom names for XLSX sheets for consolidation")

        # Ask if the user wants to adjust column width (only for smaller files)
        adjust_column_width = st.sidebar.checkbox("Adjust column width to fit content")

        # Start button to trigger conversion
        start_button = st.sidebar.button("Start Conversion")

        if start_button:
            if app_mode == "Multiple CSV to Multiple XLSX":
                # Process multiple CSV to XLSX files
                zip_buffer = BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
                    progress_placeholder = st.empty()  # Placeholder for the progress bar
                    total_files = len(uploaded_files)
            
                    # Define a progress callback that updates the progress bar
                    def update_progress(progress):
                        progress_placeholder.progress(progress)

                    process_multiple_files_to_xlsx(uploaded_files, delimiter, custom_xlsx_names, adjust_column_width, zipf, update_progress)
        
                # After all files are processed, allow the user to download the ZIP file
                zip_buffer.seek(0)
                zip_filename = f"{output_name}.zip"
                st.download_button(
                    label="Download ZIP file containing XLSX files",
                    data=zip_buffer.read(),
                    file_name=zip_filename,
                    mime="application/zip"
                )


            elif app_mode == "Multiple CSV to Single XLSX":
                # Process multiple CSV to single Excel file
                progress_placeholder = st.empty()
                excel_buffer = process_multiple_files_to_single_excel(uploaded_files, delimiter, custom_xlsx_names, adjust_column_width, lambda progress: progress_placeholder.progress(int(progress * 100)))

                if isinstance(excel_buffer, BytesIO):
                    # Allow the user to download the resulting Excel file
                    excel_filename = f"{output_name}.xlsx"
                    st.download_button(
                        label="Download the Excel file",
                        data=excel_buffer.read(),
                        file_name=excel_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

    for uploaded_file in uploaded_files:
        st.subheader(f"Preview of {uploaded_file.name}")
        preview_csv(uploaded_file, delimiter)

# Run the app
if __name__ == "__main__":
    main()
