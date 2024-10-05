import requests
from bs4 import BeautifulSoup
import openpyxl

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

url = 'https://biology.mcmaster.ca/people/faculty/'

# Function to scrape initial data
def scrape_mcmaster_faculty(url):
    response = requests.get(url)
    soup = BeautifulSoup(response.text, 'html.parser')
    
    faculty_members = soup.find_all('div', class_='col-md-6 col-xl-3') #Changing the name of this class but same structure
    data = []
    print(len(faculty_members))

    for member in faculty_members:
        name = member.find('h3', class_='card-title no-line p-0 pb-2').get_text(strip=True)
        data.append(name)
    return data

# Function to scrape email from profile page
def scrape_email_from_profile(url):
    response = requests.get(url)
    soup = BeautifulSoup(response.text, 'html.parser')

    email_divs = soup.find_all('div', class_='list-group list-group-flush small mt-2')
    emails_array = []
    for emails in email_divs:
        email_tag = emails.find('a', href=lambda href: href and 'mailto:' in href) if emails else None
        email = email_tag['href'].replace('mailto:', '') if email_tag else 'N/A'
        emails_array.append(email)
    return emails_array

# Scrape the initial data
faculty_data = scrape_mcmaster_faculty(url)
faculty_emails = scrape_email_from_profile(url)

# Add data to the Excel sheet
for idx, (name, email) in enumerate(zip(faculty_data, faculty_emails), start=124):
    sheet[f'A{idx}'] = name
    sheet[f'B{idx}'] = email

# Save the workbook and check to see if it changed

# Get the content of cell A1 before saving the workbook
pre_save_content = sheet['A1'].value

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
