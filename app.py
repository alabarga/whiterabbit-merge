import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from io import BytesIO
from datetime import datetime

def merge_excel_files(input_files, output_file):
    # Initialize DataFrames to store merged data
    merged_table_overview = pd.DataFrame()
    merged_fields_overview = pd.DataFrame()
    merged_sheets = {}

    # Define an empty row DataFrame with NaN values
    empty_row = pd.DataFrame([pd.NA], columns=["AddEmptyRow"]).drop(columns="AddEmptyRow")

    # Loop through each Excel file
    for file in input_files:
        # Load the workbook using openpyxl
        workbook = load_workbook(filename=BytesIO(file.getvalue()), data_only=True)
        # Convert workbook to a pandas ExcelFile object
        excel = pd.ExcelFile(BytesIO(file.getvalue()))

        # Merge "Table overview" sheet
        if "Table Overview" in excel.sheet_names:
            table_overview_df = pd.read_excel(excel, sheet_name="Table Overview")

            # Append the current DataFrame to the merged one, along with an empty row
            merged_table_overview = pd.concat(
                [merged_table_overview, table_overview_df], 
                ignore_index=True
            )
        
        # Merge "Fields overview" sheet
        if "Field Overview" in excel.sheet_names:
            fields_overview_df = pd.read_excel(excel, sheet_name="Field Overview")
            merged_fields_overview = pd.concat([merged_fields_overview, fields_overview_df, empty_row], ignore_index=True)

        # Handle other sheets (the ones not "Table overview" or "Fields overview")
        for sheet_name in excel.sheet_names:
            if sheet_name not in ["Table Overview", "Field Overview"]:
                if sheet_name not in merged_sheets:
                    # Initialize a DataFrame for this sheet if it doesn't exist in the dictionary
                    merged_sheets[sheet_name] = pd.DataFrame()

                # Read the data from the current sheet
                sheet_df = pd.read_excel(excel, sheet_name=sheet_name)
                # Concatenate the data to the corresponding DataFrame in merged_sheets
                merged_sheets[sheet_name] = pd.concat([merged_sheets[sheet_name], sheet_df], ignore_index=True)
    
    # Create a new Excel writer object using pandas
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Write merged "Table overview" DataFrame to Excel
        merged_table_overview.to_excel(writer, sheet_name='Table Overview', index=False)
        # Write merged "Fields overview" DataFrame to Excel
        merged_fields_overview.to_excel(writer, sheet_name='Field Overview', index=False)

        # Write all other merged sheets
        for sheet_name, sheet_df in merged_sheets.items():
            sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)

    return output_file

def main():
    st.title("WhiteRabbit Scan report Merger")

    # File uploader
    uploaded_files = st.file_uploader(
        "Upload Scan report files to merge",
        type=["xlsx"],
        accept_multiple_files=True,
        help="Drag and drop or browse to select multiple Excel files"
    )

    # Default filename
    default_filename = f"Scanreport_{datetime.now().strftime('%Y_%m_%d')}.xlsx"
    new_filename = st.text_input("Enter the new filename", value=default_filename)

    # Merge button
    if st.button("Merge Excel Files"):
        if uploaded_files and new_filename:
            # Merging the files
            try:
                output = BytesIO()
                merge_excel_files(uploaded_files, output)
                output.seek(0)
                st.success("Files merged successfully!")
                
                # Download link for merged file
                st.download_button(
                    label="Download Merged Scan report File",
                    data=output,
                    file_name=new_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"An error occurred while merging files: {e}")
        else:
            st.warning("Please upload Excel files and enter a filename.")

if __name__ == "__main__":
    main()
