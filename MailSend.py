import os
import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# Function to send email
def send_email(to_email, name, subject, body, from_email, password, attachment_path):
    try:
        msg = MIMEMultipart()
        msg['From'] = from_email
        msg['To'] = to_email
        msg['Subject'] = subject

        # Personalize the greeting
        personalized_body = body.replace("{name}", name)
        msg.attach(MIMEText(personalized_body, 'plain'))

        # Attach the file
        with open(attachment_path, "rb") as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f"attachment; filename= {os.path.basename(attachment_path)}")
            msg.attach(part)

        # Setup the server
        server = smtplib.SMTP('smtp.hostinger.io', 587)
        server.starttls()
        server.login(from_email, password)

        # Send the email
        server.send_message(msg)
        server.quit()

        print(f"Email sent to {name} at {to_email}")

    except Exception as e:
        print(f"Failed to send email to {name} at {to_email}. Error: {e}")

# Your email credentials
from_email = "kamini@gmail.com"
password = "Kamini24344"

# Read the Excel sheet
excel_file = r'C:\Users\Dell\OneDrive\Desktop\File1.xlsx'
df = pd.read_excel(excel_file)

# Check for required columns
required_columns = {'Name', 'Email'}
if not required_columns.issubset(df.columns):
    print(f"Excel file must contain the following columns: {required_columns}")
else:
    # Email content
    subject = "Write mail subject here.."
    body = """
    Mail Content
    """

    # Path to the attachment file (use absolute path)
    attachment_path = r'C:\Users\Dell\OneDrive\Desktop\FileAttach.pdf'

    # Send emails to each contact
    for index, row in df.iterrows():
        to_email = row['Email']
        name = row['Name']
        send_email(to_email, name, subject, body, from_email, password, attachment_path)

    print("Emails sent successfully!")


