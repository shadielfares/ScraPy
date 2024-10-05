import time
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service

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
#sheet['C1'] = 'Profile URL'

# Initialize the starting row index
start_idx = 2

chromedriver_path = './LOCATIONOFCHROMEDRIVER'  # Update this with the actual path

# Setup WebDriver
driver = webdriver.Chrome(service=Service(chromedriver_path))
# Function to scrape initial data from a specific page
def scrape_mcmaster_faculty():
    faculty_members = driver.find_elements(By.XPATH, '//tr[contains(@class, "odd") or contains(@class, "even")]')
    data = []
    for member in faculty_members:
        name = member.find_element(By.CLASS_NAME, 'sorting_1').find_element(By.TAG_NAME, 'a').text.strip()
        try:
            email_element = member.find_element(By.XPATH, './/td[@class="text-center"]/a[contains(@href, "mailto:")]')
            email = email_element.get_attribute('href').replace('mailto:', '')
        except:
            email = 'N/A'
        #profile_link = member.find_element(By.CLASS_NAME, 'sorting_1').find_element(By.TAG_NAME, 'a').get_attribute('href')
        data.append((name, email))
    return data

# Open the initial page
driver.get('https://www.degroote.mcmaster.ca/contact/directory/')

# Loop to scrape multiple pages
while True:
    # Scrape faculty data
    faculty_data = scrape_mcmaster_faculty()
    
    # Add data to the Excel sheet and scrape emails from profile pages
    for (name,email) in faculty_data:
        sheet[f'A{start_idx}'] = name
        sheet[f'B{start_idx}'] = email
        #sheet[f'C{start_idx}'] = url
        start_idx += 1

    # Find the next page button
    try:
        next_button = driver.find_element(By.XPATH, '//a[@class="paginate_button next"]')
        if 'disabled' in next_button.get_attribute('class'):
            break  # Exit the loop if the "next" button is disabled (no more pages)
        next_button.click()
        time.sleep(5)  # Adjust the sleep time if necessary to allow page load
    except:
        break  # Exit the loop if no more pages are found

# Save the workbook
workbook.save(workbook_path)
print("Data has been scraped and added to the Excel file successfully.")

# Close the WebDriver
driver.quit()
