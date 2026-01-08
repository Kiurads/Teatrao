import openpyxl
from openpyxl.utils import get_column_letter
from datetime import datetime

# Initialize workbook and sheet
wb = openpyxl.Workbook()
sheet = wb.active
sheet.title = "Folha1"

# List of items to populate
items_to_extract = [
    "F1", "F5", "F7", "F3", "F15", "F13", "F9", "F11", "F17", "D24", "E24", 
    "D25", "E25", "D26", "E26", "D27", "E27", "D28", "E28", "D29", "E29",
    "D30", "E30", "D31", "E31", "D32", "E32", "D33", "E33", "D34", "E34",
    "D35", "E35", "D36", "E36", "D37", "E37", "D38", "E38", "D39", "E39",
    "D40", "E40", "D41", "E41", "D42", "E42", "D43", "E43", "D44", "E44",
    "D45", "E45", "D46", "E46", "D47", "E47", "D48", "E48", "D49", "F51",
    "H51", "J51", "D51", "E51", "B54"
]

# Populate the cells with appropriate values
for cell in items_to_extract:
    column, row = "", ""
    for char in cell:
        if char.isdigit():
            row += char
        else:
            column += char

    if cell == "F5":
        # If the cell is F5, populate with a date
        sheet[cell] = datetime.now().strftime("%Y-%m-%d")
    else:
        # Otherwise, populate with the cell's coordinate
        sheet[cell] = cell

# Save the workbook
wb.save("populated_cells.xlsx")
print("Excel file 'populated_cells.xlsx' created successfully.")
