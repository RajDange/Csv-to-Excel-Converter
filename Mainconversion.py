import pandas as pd
import streamlit as st

def preview_csv(uploaded_file, delimiter):
    """
    Function to preview the first few rows of a CSV file before processing.
    :param uploaded_file: The uploaded CSV file.
    :param delimiter: The delimiter used in the CSV file.
    :return: None
    """
    try:
        # Read the first few rows of the CSV file to preview
        df = pd.read_csv(uploaded_file, delimiter=delimiter, nrows=5)
        
        df = df.fillna("")
        
        # Remove commas from numbers to prevent thousand separator formatting
        df = df.apply(lambda col: col.apply(lambda x: str(x).replace(",", "") if isinstance(x, (int, float, str)) else x))

        # Show the preview of the DataFrame
        st.write("CSV File Preview:")
        st.dataframe(df)

    except Exception as e:
        st.error(f"Error loading the file: {e}")
