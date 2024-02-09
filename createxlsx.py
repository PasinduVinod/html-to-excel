import json
from openpyxl import Workbook

# Read JSON data from file
with open('remapped_table_data.json', 'r') as json_file:
    data = json.load(json_file)

# Create a new workbook and select the active worksheet
wb = Workbook()
ws = wb.active

# Write header row
header = ['Order', 'Order Line', 'Quantity']
ws.append(header)

# Write data rows
for row_data in data:
    row = [row_data.get('Order', ''), row_data.get('Order Line', ''), row_data.get('Quantity', '')]
    ws.append(row)

# Save workbook to Excel file
wb.save('exl.xlsx')

print("Excel file 'output.xlsx' has been successfully created.")
