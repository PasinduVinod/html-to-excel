from bs4 import BeautifulSoup
import json

# Define file paths
html_file_path = 'Messaggio_2023-12-08.html'
json_file_path = 'table_data.json'

# Initialize list to store table data
table_data = []

# Read HTML content from file and write JSON data
with open(html_file_path, 'r') as html_file, open(json_file_path,
                                                  'w') as json_file:
  # Parse HTML content
  soup = BeautifulSoup(html_file, 'html.parser')

  # Find the table by name
  table = soup.find('table', {'name': 'tt'})

  # Extract table rows
  rows = table.find_all('tr')

  # Iterate over rows
  for row in rows:
    # Extract table cells from current row
    cells = row.find_all(['th', 'td'])
    row_data = {}

    # Iterate over cells
    for index, cell in enumerate(cells):
      # Only process the 2nd, 3rd, and 10th columns
      if index + 1 in [2, 3, 10, 17]:
        # Check if cell contains an input tag
        input_tag = cell.find('input')
        if input_tag:
          # Extract value from input tag
          cell_value = input_tag.get('value')
        else:
          # Use cell text as value
          cell_value = cell.get_text(strip=True)
        # Use cell value as key and value
        row_data[f'Column_{index+1}'] = cell_value

    # Append row data to table data list
    table_data.append(row_data)

  # Write table data to JSON file
  json.dump(table_data, json_file, indent=4)

print(f"Table data has been successfully written to '{json_file_path}'.")
import dataClear
import createxlsx
# import inputxl
