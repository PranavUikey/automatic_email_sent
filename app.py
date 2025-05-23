import gspread
from oauth2client.service_account import ServiceAccountCredentials
from __init__ import logger

import yaml
import os
from dotenv import load_dotenv
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart





# Connect to Google Sheets API
def send_job_email():

    

    try:
        
        load_dotenv()
        # Open the Google Sheet
        sheet = client.open("Job_Coll").sheet1
        data = sheet.get_all_records()
        last_row_index = len(data) + 1  # actual row number in sheet (including header)
        last_entry = data[-1]

        recipient_name = last_entry.get('Name')
        recipient_email = last_entry.get('Email-ID')

        # Get headers and check for 'Email Status'
        header = sheet.row_values(1)
        if "Email Status" not in header:
            header.append("Email Status")
            sheet.insert_row(header, 1)
        status_col_index = header.index("Email Status") + 1

        # Skip if email already sent
        status_cell_value = sheet.cell(last_row_index, status_col_index).value
        if status_cell_value == "✅":
            logger.info(f"✅ Email already sent to {recipient_email}, skipping.")
            return

        if not recipient_email:
            sheet.update_cell(last_row_index, status_col_index, "❌ Email Not Found")
            logger.warning("Email ID missing in the last entry.")
            return

        with open('job_colab_email_content.txt') as f:
            content = f.read()

        subject = "Collaboration Request: AI & ML Talent from AIAdventures"
        message = f"Dear {recipient_name},<br>{content}"


        # Load email configuration


        email_config = {
            'SMTP_USERNAME': os.getenv('JOB_SMTP_USERNAME'),
            'SMTP_PASSWORD': os.getenv('JOB_SMTP_PASSWORD'),
            # Add other fields as needed
        }

        # Create email
        msg = MIMEMultipart()
        msg["From"] = email_config['SMTP_USERNAME']
        msg["To"] = recipient_email
        msg["Subject"] = subject
        msg.attach(MIMEText(message, "html"))

        logger.info(f"Connecting to SMTP server: {email_config['SMTP_SERVER']} on port {email_config['SMTP_PORT']}")

        # Connect and send email
        with smtplib.SMTP_SSL(email_config['SMTP_SERVER'], email_config['SMTP_PORT']) as server:
            server.login(email_config['SMTP_USERNAME'], email_config['SMTP_PASSWORD'])
            logger.info("Login successful, sending email...")
            server.sendmail(email_config['SMTP_USERNAME'], recipient_email, msg.as_string())

        logger.info(f"✅ Email sent successfully to {recipient_email}")
        sheet.update_cell(last_row_index, status_col_index, "✅")

    except Exception as e:
        logger.error(f"❌ Error sending email: {e}")
        sheet.update_cell(last_row_index, status_col_index, f"❌ {str(e)}")






def send_course_email():
    try:
        # Open Google Sheet
        sheet = client.open("Untitled form submissions").sheet1
        data = sheet.get_all_records()
        last_entry = data[-1]
        last_row_index = len(data) + 1  # Sheet rows are 1-indexed (including header)

        # Check or create "Email Status" column
        headers = sheet.row_values(1)
        if 'Email Status' in headers:
            email_status_col = headers.index('Email Status') + 1
        else:
            sheet.update_cell(1, len(headers) + 1, 'Email Status')
            email_status_col = len(headers) + 1

        # Don't send if already marked as ✅
        existing_status = sheet.cell(last_row_index, email_status_col).value
        if existing_status == '✅':
            logger.info(f"⚠️ Email already sent to row {last_row_index}, skipping.")
            return

        # Extract student info
        name = last_entry.get('Name')
        email = last_entry.get('Email')
        interests_raw = last_entry.get('Intrested Technology', '')

        # Normalize input
        interests_cleaned = [i.strip() for i in interests_raw.strip().split(',') if i.strip()]
        interests = [tech.lower().replace(" ", "_") for tech in interests_cleaned]
        if not interests:
            logger.warning("No valid technologies found. Using default content.")
            interests = ['course_default']

        email_config = {
            'SMTP_USERNAME': os.getenv('COURSE_SMTP_USERNAME'),
            'SMTP_PASSWORD': os.getenv('COURSE_SMTP_PASSWORD'),
            # Add other fields as needed
        }
        home = 'https://www.aiadventures.in/'

        # Build course link HTML
        tech_list_html = ""
        for interest in interests:
            course_name = interest.replace("_", " ").title()
            course_id = interest.replace("_", "-")
            url_suffix = '-course-in-pune/' if course_id not in ['power-bi', 'gen-ai'] else '-course-pune/'
            course_url = f"{home}{course_id}{url_suffix}"
            tech_list_html += f'<li><a href="{course_url}" target="_blank">{course_name}</a></li>'

        # Email content
        with open('registration_mail_content.txt','r') as f:
          template = f.read()

        body = template.format(name=name, course_links=tech_list_html)

        # Compose and send email
        msg = MIMEMultipart()
        msg["From"] = email_config['SMTP_USERNAME']
        msg["To"] = email
        msg["Subject"] = "Detailed course content from AIAdventures"
        msg.attach(MIMEText(body, "html"))

        logger.info(f"📤 Sending course links to {email}")
        with smtplib.SMTP_SSL(email_config['SMTP_SERVER'], email_config['SMTP_PORT']) as server:
            server.login(email_config['SMTP_USERNAME'], email_config['SMTP_PASSWORD'])
            server.sendmail(email_config['SMTP_USERNAME'], email, msg.as_string())

        logger.info(f"✅ Course email sent successfully to {email}")
        sheet.update_cell(last_row_index, email_status_col, '✅')

    except Exception as e:
        logger.error(f"❌ Course email error: {e}")
        try:
            sheet.update_cell(last_row_index, email_status_col, '❌')
        except:
            logger.error("⚠️ Could not update Email Status in sheet after failure.")


if __name__ == "__main__":
    # Load configuration
    with open('config.yaml') as yaml_file:
        server_config = yaml.safe_load(yaml_file)


    SERVICE_ACCOUNT_FILE = "automation-450206-b4c4c3a8a4ec.json"

    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(SERVICE_ACCOUNT_FILE, scope)
    client = gspread.authorize(creds)

    # Send emails
    send_job_email()
    send_course_email()
    logger.info("All emails processed.")
        