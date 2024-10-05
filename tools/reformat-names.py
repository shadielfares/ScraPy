import openpyxl

def reformat_names(file_path, sheet_name):
    # Load the workbook and select the sheet
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook[sheet_name]

    # Iterate over the cells from A2 to A176
    for row in range(2, 177):  # Row numbers are 2 to 176
        cell = sheet[f'A{row}']
        name = cell.value
        if name:
            # Split the name by comma and trim any whitespace
            first_name, last_name = [part.strip() for part in name.split(',', 1)]
            # Reformat the name
            reformatted_name = f'{first_name} {last_name}'
            # Update the cell value
            cell.value = reformatted_name

    # Save the changes to the workbook
    workbook.save(file_path)

# Example usage
file_path = './FEEDERDOC.xlsx'  # Path to your Excel file
sheet_name = 'DUMMYNAME'  # Name of the sheet
reformat_names(file_path, sheet_name)
