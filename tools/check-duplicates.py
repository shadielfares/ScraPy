import pandas as pd

# Load the Excel workbook
file_path = 'FEEDERDOC.xlsx' #Replace this with the correct file name located in the excel-sheets folder
sheets = ['health-sci', 'fac-hum', 'fac-eng', 'fac-biz', 'fac-sci', 'fac-soc-sci'] #Example page names to feed from, CHANGE THESE ACCORDINGLY
df_dict = pd.read_excel(file_path, sheet_name=sheets)

# Combine all sheets into one DataFrame for checking duplicates
combined_df = pd.concat(df_dict.values(), keys=df_dict.keys()).reset_index(level=0).rename(columns={'level_0': 'Sheet'})

# Identify duplicates in names and emails
duplicate_names = combined_df[combined_df.duplicated(['Name'], keep=False)]
duplicate_emails = combined_df[combined_df.duplicated(['Email'], keep=False)]

def remove_duplicates(df, column_name):
    while True:
        print(f"\nDuplicates in {column_name}:")
        print(df[df.duplicated([column_name], keep=False)])
        choice = input(f"\nWould you like to remove duplicates in {column_name}? (yes/no): ").strip().lower()
        if choice == 'yes':
            # Drop only the second occurrence of each duplicate
            duplicates = df[df.duplicated([column_name], keep='first')]
            df = df.drop(duplicates.index)
            print(f"\nDuplicates in {column_name} removed.")
            break
        elif choice == 'no':
            print(f"\nNo changes made to {column_name}.")
            break
        else:
            print("Invalid input. Please type 'yes' or 'no'.")
    return df

# Check for duplicates in names and prompt user
if not duplicate_names.empty:
    combined_df = remove_duplicates(combined_df, 'Name')
else:
    print("\nNo duplicate names found.")

# Check for duplicates in emails and prompt user
if not duplicate_emails.empty:
    combined_df = remove_duplicates(combined_df, 'Email')
else:
    print("\nNo duplicate emails found.")

# Split the combined DataFrame back into individual sheets
updated_df_dict = {sheet: combined_df[combined_df['Sheet'] == sheet].drop(columns='Sheet') for sheet in sheets}

# Save the updated DataFrames to a new Excel file
output_file_path = 'UPDATEDDOC.xlsx' #Replace this with the correct file name located in the excel-sheets folder
with pd.ExcelWriter(output_file_path) as writer:
    for sheet_name, df in updated_df_dict.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"\nUpdated file saved as {output_file_path}")
