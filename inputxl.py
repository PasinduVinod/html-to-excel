import pandas as pd
import json

# Read the Excel file into a DataFrame
df = pd.read_excel('input.xlsx')

# Format the 'Quantity' column to have 3 decimal places if needed
df['Quantity'] = df['Quantity'].apply(lambda x: f'{x:.3f}' if x < 100 else str(int(x)))

# Convert the DataFrame to a list of dictionaries
data = df.to_dict(orient='records')

# Write the data to a JSON file
with open('output.json', 'w') as json_file:
    json.dump(data, json_file, indent=4)

print("JSON file 'output.json' has been successfully created.")
