XML to Excel Converter

This Python script converts complex nested XML files into an Excel .xlsx format. It handles both attributes and nested elements within the XML, ensuring the data is flattened into a tabular format suitable for viewing and analysis in Excel.

Features

Recursive Flattening: The script recursively processes nested elements, extracting both text content and attributes, and flattens them into a structured Excel format.
Handles Special Characters: The script ensures that special characters and encoding issues are handled correctly (ISO-8859-1).
Customizable Depth Limit: Users can specify the maximum depth for recursion to control how deeply nested structures are processed.

How It Works
The script parses an XML file, extracts the data from each <Indvl> element and its nested child elements, and writes the data to an Excel file.

Prerequisites
Make sure you have the following Python packages installed:

xml.etree.ElementTree (part of the Python standard library)
openpyxl (for writing Excel files)
To install openpyxl, run:
pip install openpyxl

How to Use
Place Your XML File: Save your XML file in the XML Files directory.

Run the Script: Run the script by executing the following command in the terminal:
python3 /path/to/your/script.py

Make sure to update the file paths in the script if needed.

Generated Excel File: The Excel file will be saved in the XLSX Files directory with the name you specified in the script.

ustomization
Change XML File: Update the ET.parse() method with the correct path to your XML file.
Change Output File: Modify the wb.save() method to set the desired name and path for the generated Excel file.
Depth Limit: The max_depth parameter in the flatten_xml() function controls how deep the script will recurse into nested XML elements. You can modify this if necessary.
Troubleshooting
If Excel finds an issue with the file, ensure that the data is properly sanitized by removing any special characters, null values, or incorrect encoding.
If the script fails to find the XML file, double-check that the file path is correctly specified.

Final Notes
Feel free to modify this template based on the specifics of your project. This README.md provides a basic explanation of the functionality, how to use it, and common steps for customization. Let me know if you need further changes!