import openpyxl

# Load the Excel file
excel_file = "home-hill_R7andR8_20231128.xlsx"  # Replace with the path to your Excel file
workbook = openpyxl.load_workbook(excel_file, data_only=True)
# Select the desired worksheet (assuming it's the first sheet)
worksheet = workbook.active

# # Get the formula of cell A16
# cell_formula = worksheet['A16'].value
# cell_formula = cell_formula.replace("=", "")
# print("cell formula", cell_formula)

# # Define the absolute cell reference
# absolute_reference = cell_formula

# # Get the cell value using the absolute reference
# cell_value = worksheet[absolute_reference].value

cell_value = worksheet['C16'].value

# Print the cell value
print("Value of cell A16:", cell_value)

# Close the Excel file
workbook.close()
