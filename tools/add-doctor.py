import openpyxl

# Load the existing workbook
workbook_path = './excel-sheets/DUMMYDOC.xlsx' #Replace this with the correct file name located in the excel-sheets folder
workbook = openpyxl.load_workbook(workbook_path)

# Access the specific sheet
sheet_name = 'dummyPageName'
if sheet_name in workbook.sheetnames:
    sheet = workbook[sheet_name]
else:
    print(f"Sheet '{sheet_name}' does not exist in the workbook.")
    exit()

# Loop through rows 2 to 274 in column A and prepend "Dr." to the existing names
for row in range(2, 275):  # Note: range end is exclusive, so use 275 to include row 274
    cell = sheet[f'A{row}']
    if cell.value:
        cell.value = f"Dr. {cell.value}"

# Save the workbook
workbook.save(workbook_path)
print("Entries have been updated with 'Dr.' prefix.")
