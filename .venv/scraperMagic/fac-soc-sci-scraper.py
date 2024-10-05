import requests
from bs4 import BeautifulSoup
import openpyxl
import time

# Load the existing workbook
workbook_path = './excel-sheets/DUMMYDOC.xlsx'
workbook = openpyxl.load_workbook(workbook_path)

# Create a new sheet for the scraped data
if 'dummyPageName' not in workbook.sheetnames:
    sheet = workbook.create_sheet('dummyPageName')
else:
    sheet = workbook['dummyPageName']

# Set up headers in the new sheet
sheet['A1'] = 'Name'
sheet['B1'] = 'Email'

# Function to scrape initial data from a specific page
def scrape_mcmaster_faculty(page_num):
    url = f'https://socialsciences.mcmaster.ca/contact-us/directory/?pg={page_num}'
    response = requests.get(url)
    soup = BeautifulSoup(response.text, 'html.parser')

    faculty_members = soup.find_all('div', class_='profile-item')
    data = []

    print(len(faculty_members))

    for member in faculty_members:
        name = member.find('div', class_='profile-item--name').get_text(strip=True)

        email_div = member.find('div', class_='profile-item-col profile-item-contact')
        email_tag = email_div.find('a', href=lambda href: href and 'mailto:' in href) if email_div else None
        email = email_tag['href'].replace('mailto:', '') if email_tag else 'N/A'
        data.append([name, email])
    return data

start_idx = 322

# Loop to scrape multiple pages
for page_num in range(32, 59):  # Adjust the range as needed to scrape more pages
    faculty_data = scrape_mcmaster_faculty(page_num)

    # Add data to the Excel sheet and scrape emails from profile pages
    for (name,email) in faculty_data:
        sheet[f'A{start_idx}'] = name
        sheet[f'B{start_idx}'] = email
        start_idx += 1

    # Delay between page requests
    time.sleep(5)

# Save the workbook and check to see if it changed

# Get the content of cell A1 before saving the workbook
pre_save_content = sheet['A2'].value

# Save the workbook
workbook.save(workbook_path)

# Load the workbook again to compare the content
workbook_post_save = openpyxl.load_workbook(workbook_path)
sheet_post_save = workbook_post_save['dummyPageName']

# Get the content of cell A1 after saving the workbook
post_save_content = sheet_post_save['A1'].value

# Compare the contents
if pre_save_content == post_save_content:
    print("The contents of cell A1 have not changed after saving the workbook.")
else:
    print("The contents of cell A1 have changed after saving the workbook.")