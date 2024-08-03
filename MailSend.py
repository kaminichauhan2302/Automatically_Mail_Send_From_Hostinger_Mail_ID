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
from_email = "kamini.chauhan@akiyam.in"
password = "Kamini@24344"

# Read the Excel sheet
excel_file = r'C:\Users\Dell\OneDrive\Desktop\File1.xlsx'
df = pd.read_excel(excel_file)

# Check for required columns
required_columns = {'Name', 'Email'}
if not required_columns.issubset(df.columns):
    print(f"Excel file must contain the following columns: {required_columns}")
else:
    # Email content
    subject = "Request from Akiyam Solutions Pvt. Ltd."
    body = """
    Dear {name},
    I trust this message finds you well. My name is Kamini Chauhan, and I am writing to introduce you to an exceptional investment opportunity with Akiyam Solutions Pvt. Ltd., a pioneering gaming and technology services company based in India.
    Our company has made significant strides in the gaming industry, notably with our first FPS PC game, Assassin - The First List, which is available on both Steam and Epic Games. Building on this success, we are now poised to launch our next major project, Beyonders, a mobile game designed to engage and captivate a global audience. Beyond our gaming ventures, Akiyam Solutions Pvt. Ltd. offers a comprehensive suite of technology services. Our expertise spans several critical areas, including Web Development, SEO, Data Science, AR/VR, Machine Learning, and DevOps. Our proficiency in key technologies includes:

    Frontend and Cloud: React, Angular, AWS, Azure
    Backend: Java, Dotnet, NodeJS
    Mobile: React Native, Flutter, Xamarin, iOS, Android
    Platform: Power Platform, SharePoint, WordPress, Magento, Drupal

    We believe that our diversified service offerings, combined with our innovative gaming projects, position us as a leading player in the tech industry. We are confident that, with your investment, Akiyam Solutions Pvt. Ltd. can achieve unprecedented growth and success.
    We would welcome the opportunity to discuss this investment proposal in more detail and explore how we can achieve mutual success. We are also keen to introduce you to our company, products, business model, marketing strategy, and address any questions you may have. Please let us know a convenient time for a meeting or call.

    Thanks for your valuable time.

    Best Regards,
    Kamini Chauhan 
    Akiyam Solutions Pvt. Ltd.

    Phone: 0265 461 4248
    Email: kamini.chauhan@akiyam.in
    LinkedIn: Akiyam Solutions
    Instagram: Akiyam Solutions
    X (formerly Twitter): Akiyam Solutions
    Facebook: Akiyam Solutions
    Office Address = SF-16, Dwarkesh High view, Near Susen - Tarsali Ring Rd, NH No. 8, Tarsali, Vadodara, Gujarat 390009
    """

    # Path to the attachment file (use absolute path)
    attachment_path = r'C:\Users\Dell\OneDrive\Desktop\FileAttach.pdf'

    # Send emails to each contact
    for index, row in df.iterrows():
        to_email = row['Email']
        name = row['Name']
        send_email(to_email, name, subject, body, from_email, password, attachment_path)

    print("Emails sent successfully!")
