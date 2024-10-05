import openpyxl
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.base import MIMEBase
from email import encoders
import time

from dotenv import load_dotenv
import os

# Load the environment variables from the .env file
load_dotenv()

def send_email(sender_email, sender_password, receiver_email, subject, body,  image_path):
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'html'))


    # Attach the image
    with open(image_path, 'rb') as img:
        mime = MIMEBase('image', 'png', filename="AppendedFileLocation") #Change application to image, and pdf to jpg or png
        mime.add_header('Content-Disposition', 'attachment', filename="Brochure")
        # The below code is for attaching an image
        mime.add_header('X-Attachment-Id', 'structure')
        mime.add_header('Content-ID', '<image1>')
        mime.set_payload(img.read())
        # mime.set_payload(pdf.read()) #Comment this line out to keep the avaliability to attach pdf
        encoders.encode_base64(mime)
        msg.attach(mime)
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, sender_password)
        text = msg.as_string()
        server.sendmail(sender_email, receiver_email, text)
        print(f"Email sent to {receiver_email} successfully!")
        time.sleep(10)  # Add a delay between emails
    except Exception as e:
        print(f"Error: {e}")
    finally:
        server.quit()

def send_emails_from_excel(file_path, sender_email, sender_password):
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook["dummyPageName"] #Adjust reading workbook sheet.
    #Change the range to the rows you want to send emails to
    for row in sheet.iter_rows(min_row=1, max_row=2, values_only=True):  # assuming data starts from row 1 and ends at row 2
        receiver_name, receiver_email = row
        subject = "This is a Test E-Mail"
        body = f""" 
        <html>
        <body>
        <p>Hi {receiver_name},</p>
        <p>This is the first email</p>
        </body>
        </html>
        """
        send_email(sender_email, sender_password, receiver_email, subject, body, 'AppendedFileLocation')

if __name__ == "__main__":
    file_path = "./excel-sheets/DUMMYDOC.xlsx"  # Replace with the path to your Excel file
    sender_email = "emailForScript"  # Replace with your Gmail address    
    sender_password = "senderPassword"  # Replace with your Gmail password / app password
    send_emails_from_excel(file_path, sender_email, sender_password)
