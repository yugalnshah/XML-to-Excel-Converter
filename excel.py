import xml.etree.ElementTree as ET
from openpyxl import Workbook

# Function to recursively flatten nested XML elements with a depth limit
def flatten_xml(element, parent_key='', result=None, depth=0, max_depth=5):
    if result is None:
        result = {}
    
    # Prevent infinite recursion by limiting depth
    if depth > max_depth:
        return result

    # Extract attributes if present
    for key, value in element.attrib.items():
        new_key = f"{parent_key}.{key}" if parent_key else key
        result[new_key] = value

    # Extract text content if present
    if element.text and element.text.strip():
        result[parent_key] = element.text.strip()

    # Iterate through children of the current element
    for child in element:
        # Construct key using the parent's key and the child's tag
        new_key = f"{parent_key}.{child.tag}" if parent_key else child.tag

        # Recurse into the child
        flatten_xml(child, new_key, result, depth + 1, max_depth)
    
    return result

# Parse the XML file with ISO-8859-1 encoding to handle special characters
tree = ET.parse('/Users/yugal/Desktop/XML-Excel/XML Files/IA_Indvl_Feeds20.xml') # change this according to which XML file you want to convert
root = tree.getroot()

# Create a new Excel workbook and select the active worksheet
wb = Workbook()
ws = wb.active

print("Processing individual records...")

# Flatten each 'Indvl' element and collect headers
flattened_data = [flatten_xml(indvl) for indvl in root.findall('.//Indvl')]
print(f"Flattened {len(flattened_data)} records.")

# Extract headers from all keys in the flattened data
headers = sorted(set(key for item in flattened_data for key in item.keys()))
print(f"Headers extracted: {headers}")

# Write headers to the worksheet
ws.append(headers)
print("Headers written to the Excel file.")

# Write each flattened record as a row
for item in flattened_data:
    # Ensure data is properly encoded and sanitized
    row = [item.get(header, '').replace('\x00', '').strip() for header in headers]  # Removing null characters
    ws.append(row)

print("All rows written to the Excel file.")

# Save the workbook to a file
wb.save('/Users/yugal/Desktop/XML-Excel/XLSX Files/IA_Indvl_Feeds20.xlsx') # change this according the name of the new file you want to create
print("Conversion from nested XML to XLSX is complete! Data has been written to 'IA_Indvl_Feeds20.xlsx'.")