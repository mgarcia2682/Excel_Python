import openpyxl as op

excel_file_path = input('Where is the excel file located ? ')

modified_file_path = input('Where would you like to save the file to ? ')

def apply_autofilter_to_all_sheets(excel_file_path): # this need to be the modified file path
    # Load the workbook
    workbook = op.load_workbook(excel_file_path)

    # Iterate through all sheets
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]

        # Apply AutoFilter to all columns
        sheet.auto_filter.ref = sheet.dimensions

    # Save the modified workbook
    #modified_file_path = 'C:\Python\modified_file.xlsx' # Output the xlsx to a file location
    workbook.save(modified_file_path)

    print(f"AutoFilter applied to all sheets. Modified file saved to {modified_file_path}")

# Replace 'path/to/your/excel/file.xlsx' with the path to your Excel file
apply_autofilter_to_all_sheets(excel_file_path)