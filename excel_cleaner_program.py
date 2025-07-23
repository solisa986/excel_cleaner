import openpyxl
from pathlib import Path
import pandas as pd
import streamlit as st
import io
from io import BytesIO

# buffer to use for excel writer
buffer = io.BytesIO()

st.title("Excel Cleanup App")

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
    if sheet_name is not None:
        # Select the worksheet by name
        worksheet = workbook[f'{sheet_name}']

        table_list = list(worksheet.tables.keys())
        table = st.selectbox("Please choose which table contains the data that needs to be extracted: ", table_list, index=None)
        if table is not None:
            lookup_table = worksheet.tables[f'{table}']
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
            st.write(f"Successfully pulled in the data from table '{table}'. Showing the top 10 records below:")
            df = pd.DataFrame(data=rows_list[1:], index=None, columns=rows_list[0])
            st.write(df.head(10))
            column_list = list(df.columns)
            columns_selected = st.multiselect("Please choose which columns you would like to keep in your cleaned Excel file: ", column_list, index=None)
            if columns_selected is not None:
                st.write("Building a new excel file...")
                filtered_df = df[columns_selected]
                df_csv = filtered_df.to_csv(index=False).encode('utf-8')
                # download button 1 to download dataframe as csv
                download1 = st.download_button(
                    label="Download data as CSV",
                    data=df_csv,
                    file_name='large_df.csv',
                    mime='text/csv'
                )

                # download button 2 to download dataframe as xlsx
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    # Write each dataframe to a different worksheet.
                    filtered_df.to_excel(writer, sheet_name=f'{sheet_name}', index=False)

                    download2 = st.download_button(
                        label="Download data as Excel",
                        data=buffer,
                        file_name=f'{table}_cleaned.xlsx',
                        mime='application/vnd.ms-excel'
                    )
