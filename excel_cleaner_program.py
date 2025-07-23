import openpyxl
from pathlib import Path
import pandas as pd
import streamlit as st
from io import BytesIO

st.title("Excel Cleanup App")

uploaded_file = st.file_uploader('Please upload the file that you need cleaned up')
if uploaded_file is not None:
    # Create a file uploader widget
    uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx"])

    # Check if a file has been uploaded
    if uploaded_file is not None:
        # Read the uploaded file as a BytesIO object
        file_contents = uploaded_file.getvalue()

        # Load the workbook from the BytesIO object using openpyxl
        workbook = openpyxl.load_workbook(BytesIO(file_contents))

    st.write("Successfully linked the workbook!")
    # Get the list of sheet names
    sheet_names_list = workbook.sheetnames

    sheet_name = st.selectbox("Found the following sheets in the uploaded workbook. Please choose which one contains the data to be extracted: ", sheet_names_list, index=None)
    # Select the worksheet by name
    worksheet = workbook[f'{sheet_name}']

    table_list = list(worksheet.tables.keys())
    print(table_list)

    extracted_table = []
    for t in table_list:
        lookup_table = worksheet.tables[f'{t}']
        lookup_range = lookup_table.ref

        # Access the data in the table range
        data = worksheet[lookup_range]
        rows_list = []

        # Loop through each row and get the values in the cells
        for row in data:
            # Get a list of all columns in each row
            cols = []
            for col in row:
                cols.append(col.value)
            rows_list.append(cols)

        # Create a pandas dataframe from the rows_list.
        # The first row is the column names
        df = pd.DataFrame(data=rows_list[1:], index=None, columns=rows_list[0])
        st.write(df)
        st.write(df.columns)
        # filtered_df = df.filter(items=[])