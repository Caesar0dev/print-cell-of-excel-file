import openpyxl

# Load the Excel file
workbook = openpyxl.load_workbook('E:/workspace/scraping/print-cell-of-excel-file/albury_R6andR7_20231216.xlsx', data_only=True)

# Optionally, you can interact with the file here

# Save the file
workbook.save('E:/workspace/scraping/print-cell-of-excel-file/albury_R6andR7_20231216(1).xlsx')

# Closing is automatic after saving
