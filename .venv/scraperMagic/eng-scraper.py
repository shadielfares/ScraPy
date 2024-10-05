import requests
from bs4 import BeautifulSoup
import openpyxl
import time

# Load the existing workbook
workbook_path = './excel-sheets/DUMMYDOC.xlsx' #Replace this with the correct file name located in the excel-sheets folder
workbook = openpyxl.load_workbook(workbook_path)

# Create a new sheet for the scraped data
if 'dummyPageName' not in workbook.sheetnames:
    sheet = workbook.create_sheet('dummyPageName')
else:
    sheet = workbook['dummyPageName']

# Set up headers in the new sheet
sheet['A1'] = 'Name'
sheet['B1'] = 'Email'
#sheet['C1'] = 'Profile URL'

# Function to scrape initial data from a specific page
def scrape_mcmaster_faculty(page_num):
    url = f'https://www.eng.mcmaster.ca/faculty-staff/faculty-directory/?pg={page_num}'
    response = requests.get(url)
    soup = BeautifulSoup(response.text, 'html.parser')

    faculty_members = soup.find_all('li', class_='faculty-listing__faculty')
    data = []

    for member in faculty_members:
        name = member.find('h2', class_='faculty-card__name').get_text(strip=True)
        profile_link = member.find('a', class_='faculty-card__link')['href'] if member.find('a', class_='faculty-card__link') else None
        data.append((name, profile_link))
    return data

# Function to scrape email from profile page
def scrape_email_from_profile(profile_url):
    response = requests.get(profile_url)
    soup = BeautifulSoup(response.text, 'html.parser')
    email_div = soup.find('div', class_='single-faculty__contact__option-content')
    email_tag = email_div.find('p').find('a', href=lambda href: href and 'mailto:' in href) if email_div else None
    email = email_tag['href'].replace('mailto:', '') if email_tag else 'N/A'
    return email

# Initialize the starting row index
start_idx = 274

# Loop to scrape multiple pages
for page_num in range(17, 18):  # Adjust the range as needed to scrape more pages
    faculty_data = scrape_mcmaster_faculty(page_num)

    # Add data to the Excel sheet and scrape emails from profile pages
    for name, profile_url in faculty_data:
        email = scrape_email_from_profile(profile_url)
        sheet[f'A{start_idx}'] = name
        sheet[f'B{start_idx}'] = email
        #sheet[f'C{start_idx}'] = profile_url
        start_idx += 1

    # Delay between page requests
    time.sleep(10)

# Save the workbook
workbook.save(workbook_path)
print("Data has been scraped and added to the Excel file successfully.")
