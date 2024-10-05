import openpyxl
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.base import MIMEBase
from email import encoders
import time

def send_email(sender_email, sender_password, receiver_email, subject, body,  pdf_path):
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'html'))


    # Attach the image
    with open(pdf_path, 'rb') as pdf:
        mime = MIMEBase('application', 'pdf', filename='2024-ShadiE.pdf') #Change application to image, and pdf to jpg or png
        mime.add_header('Content-Disposition', 'attachment', filename='My Resume')
        # The below code is for attaching an image
        mime.add_header('X-Attachment-Id', 'structure')
        mime.add_header('Content-ID', '<pdf1>') #Change this to pdf1
        # mime.set_payload(img.read()) #Comment this line out to keep the avaliability to attach img
        mime.set_payload(pdf.read()) #Comment this line out to keep the avaliability to attach pdf
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
    sheet = workbook["Sheet1"]
    #Change the range to the rows you want to send emails to
    for row in sheet.iter_rows(min_row=2, max_row=42, values_only=True):  # assuming data starts from row 2 and ends at row 42
        receiver_name, receiver_email = row
        subject = "Shadi - Development Roles in CAS"
        body = f""" 
        <html>
        <body>
       <p>Hi {receiver_name},</p>

        <p>I hope this message finds you well.</p>

        <p>My name is Shadi El-Fares, and I am reaching out to express my strong interest in opportunities within the Computing and Software Faculty at McMaster University, particularly in roles related to Data Science and Machine Learning. <br><br>
        
        While my formal education is in Software Engineering, I have a background in data analysis through hands-on projects utilizing Python, Pandas, and OpenPyXL. <br><br>
        
        Additionally, I’ve built and optimized data pipelines using tools like Selenium, BeautifulSoup4, and PyAutoGUI, which have proven invaluable in extracting and processing web data. <br><br>
        
        <em><b>This email</b> was sent through a Python email server (SMTP), part of the data pipeline I developed, allowing me to efficiently reach out to the entire Computing and Software Faculty at McMaster.</em></p>

        <p>I want to share that I am approved for the Work Study Program, and I am actively seeking a role for the upcoming academic year. While I have predominantly worked in Web Development and Design, I am eager to transition into roles that involve Data Science and Machine Learning, areas I am highly passionate about.</p>

        <p>I would greatly appreciate the opportunity to discuss how my technical expertise can support your needs or to explore any referrals for positions where I can contribute effectively.</p>

        <p>Thank you for considering my application. I look forward to the possibility of contributing to your team and the broader academic community at McMaster University.</p>

        <p>Warm regards,</p>


        <p>Name: Shadi El-Fares<br>
        Program: Engineering II, Software Engineering<br>
        Social: <a href="https://www.linkedin.com/in/shadielfares">LinkedIn</a><br>
        <a href="mailto:elfaress@mcmaster.ca">elfaress@mcmaster.ca</a><br></p>

        <p><img src="cid:image1" style="max-width: 100%; height: auto;"></p>
        </body>
        </html>
        """
        send_email(sender_email, sender_password, receiver_email, subject, body, './2024-ShadiE.pdf')

if __name__ == "__main__":
    file_path = "./profs_software_email_v.xlsx"  # Replace with the path to your Excel file
    sender_email = "shadi.elfares@gmail.com"  # Replace with your Gmail address
    sender_password = "glsi cbfw sthh ruru"  # Replace with your Gmail password
    send_emails_from_excel(file_path, sender_email, sender_password)
