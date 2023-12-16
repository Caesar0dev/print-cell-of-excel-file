import openpyxl
import shutil
import time
import os

# Example usage
file_path = './albury_R6andR7_20231216.xlsx'

workbook = openpyxl.load_workbook(file_path, data_only=True)
worksheet = workbook.active

D1_cell_value = worksheet['D1'].value
D2_cell_value = worksheet['D2'].value
A16_cell_value = worksheet['A16'].value
B16_cell_value = worksheet['B16'].value
C16_cell_value = worksheet['C16'].value
D16_cell_value = worksheet['D16'].value

# Print the cell value
print("Value of cell D1:", D1_cell_value)
print("Value of cell D2:", D2_cell_value)
print("Value of cell A16:", A16_cell_value)
print("Value of cell B16:", B16_cell_value)
print("Value of cell C16:", C16_cell_value)
print("Value of cell D16:", D16_cell_value)

# Close the Excel file
workbook.close()
