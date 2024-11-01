import json, os, openpyxl

os.chdir(os.path.dirname(os.path.abspath(__file__)))


for excelFile in os.listdir('.'):
    # Skip non-xlsx files, load the workbook object.
    for file in os.listdir(os.getcwd()):
        if file.endswith('.xlsx'):
            print(file)
'''
    for sheetName in wb.get_sheet_names():
        # Loop through every sheet in the workbook.
        sheet = wb.get_sheet_by_name(sheetName)
        # TODO: Create the CSV filename from the Excel filename and sheet title.
        # TODO: Create the csv.writer object for this CSV file.
        # Loop through every row in the sheet.
        for rowNum in range(1, sheet.max_row + 1):
            rowData = [] # append each cell to this list
            # Loop through each cell in the row.
            for colNum in range(1, sheet.max_column + 1):
                # Append each cell's data to rowData.
            # Write the rowData list to the CSV file.
csvFile.close()'''