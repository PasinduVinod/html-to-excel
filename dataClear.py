import json

# Read JSON data from file
with open('table_data.json', 'r') as json_file:
    data = json.load(json_file)

# Remove empty objects
cleaned_data = [item for item in data if item]

# Remap keys and filter the data
remapped_data = []
for item in cleaned_data:
    remapped_item = {}
    if "Column_2" in item:
        remapped_item["Order"] = item["Column_2"]
    if "Column_3" in item:
        remapped_item["Order Line"] = item["Column_3"]
    if "Column_10" in item:
        remapped_item["Quantity"] = item["Column_10"]
    remapped_data.append(remapped_item)

# Remove the first three objects
remapped_data = remapped_data[3:]

# Print or write the remapped data to a new JSON file
print(remapped_data)

# Write the remapped data to a new JSON file
with open('remapped_table_data.json', 'w') as remapped_json_file:
    json.dump(remapped_data, remapped_json_file, indent=4)
