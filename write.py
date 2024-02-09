import pandas as pd

# Read JSON data into DataFrame
df = pd.read_json('table_data.json')

# Write DataFrame to Excel file
df.to_excel('output.xlsx', index=False)

print("Excel file 'output.xlsx' has been successfully created.")
