import gspread
from oauth2client.service_account import ServiceAccountCredentials
from __init__ import logger

import yaml
import os
from dotenv import load_dotenv
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from datetime import datetime




def send_job_email(job_sender_name,job_sender_number):
    try:
        # Open the Google Sheet
        sheet = client.open("Job_Coll").sheet1
        data = sheet.get_all_records()
        last_row_index = len(data) + 1  # actual row number in sheet (including header)
        last_entry = data[-1]

        recipient_name = last_entry.get('Name')
        recipient_email = last_entry.get('Email-ID')

        # Get headers
        header = sheet.row_values(1)

        # Ensure 'Email Status' exists
        if "Email Status" not in header:
            header.append("Email Status")

        # Ensure 'Email Sent DateTime' exists after 'Email Status'
        if "Email Sent DateTime" not in header:
            status_index = header.index("Email Status")
            header.insert(status_index + 1, "Email Sent DateTime")

        # Update header row
        sheet.update('1:1', [header])

        # Recalculate column indexes
        status_col_index = header.index("Email Status") + 1
        datetime_col_index = header.index("Email Sent DateTime") + 1

        # Skip if email already sent
        status_cell_value = sheet.cell(last_row_index, status_col_index).value
        if status_cell_value == "✅":
            logger.info(f"✅ Email already sent to {recipient_email}, skipping.")
            return

        if not recipient_email:
            sheet.update_cell(last_row_index, status_col_index, "❌ Email Not Found")
            logger.warning("Email ID missing in the last entry.")
            return

        # Load email content
        with open('job_colab_email_content.txt') as f:
            template = f.read()


        body = template.format(recipient_name= recipient_name, job_sender_name = job_sender_name, job_sender_number = job_sender_number)

        # Email config
        email_config = server_config['JOB']

        # Create email
        msg = MIMEMultipart()
        msg["From"] = email_config['SMTP_USERNAME']
        msg["To"] = recipient_email
        msg["Subject"] = "Collaboration Request: AI & ML Talent from AIAdventures"
        msg.attach(MIMEText(body, "html"))

        logger.info(f"Connecting to SMTP server: {email_config['SMTP_SERVER']} on port {email_config['SMTP_PORT']}")

        # Send email
        logger.info(f"📤 Sending job colab request to {recipient_email}")
        with smtplib.SMTP_SSL(email_config['SMTP_SERVER'], email_config['SMTP_PORT']) as server:
            server.login(email_config['SMTP_USERNAME'], email_config['SMTP_PASSWORD'])
            logger.info("Login successful, sending email...")
            server.sendmail(email_config['SMTP_USERNAME'], recipient_email, msg.as_string())

        logger.info(f"✅ Email sent successfully to {recipient_email}")
        sheet.update_cell(last_row_index, status_col_index, "✅")
        sheet.update_cell(last_row_index, datetime_col_index, datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

    except Exception as e:
        logger.error(f"❌ Error sending email: {e}")
        try:
            # Safely update only if header exists and indexes are valid
            if 'status_col_index' in locals() and 'last_row_index' in locals():
                sheet.update_cell(last_row_index, status_col_index, f"❌ {str(e)}")
        except:
            pass  # Prevent double error logging




def send_course_email(course_sender_name, course_sender_number):
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

        # Extract student info
        name = last_entry.get('Name')
        email = last_entry.get('Email')
        interests_raw = last_entry.get('Intrested Technologies', '')

        # Normalize input and decide PDF
        li = interests_raw.split(',')
        for i in range(len(li)):
            li[i] = li[i].strip().lower().replace(" ", "_")

        interests = li if any(li) else ['course_default']

        # Decide which PDF to attach
        if 'gen_ai'  in li or 'machine_learning' in li or 'deep_learning' in li:
            selected_pdf = 'gen_ai_with_ml.pdf'
        else:
            selected_pdf = 'data_science_with_powerbi.pdf'

        logger.info(f"📎 PDF to attach: {selected_pdf}")

        # Email content using PDF name in template
        with open('registration_mail_content.txt', 'r') as f:
            template = f.read()

        body = template.format(
            name=name,
            course_pdf=selected_pdf,
            course_sender_name=course_sender_name,
            course_sender_number=course_sender_number
        )

        # Compose email
        email_config = server_config['COURSE']
        msg = MIMEMultipart()
        msg["From"] = email_config['SMTP_USERNAME']
        msg["To"] = email
        msg["Subject"] = "Detailed course content from AIAdventures"
        msg.attach(MIMEText(body, "html"))

        # Attach the selected PDF
        try:
            with open('course_content/'+selected_pdf, 'rb') as pdf_file:
                part = MIMEApplication(pdf_file.read(), _subtype='pdf')
                part.add_header('Content-Disposition', 'attachment', filename=selected_pdf)
                msg.attach(part)
        except Exception as e:
            logger.warning(f"⚠️ Could not attach PDF: {selected_pdf} — {e}")

        # Check if already sent
        existing_status = sheet.cell(last_row_index, email_status_col).value
        if existing_status == '✅':
            logger.info(f"✅ Email already sent to {email}, skipping.")
            return

        logger.info(f"📤 Sending course email to {email}")
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
    # Load environment variables
    load_dotenv()

    # Load configuration and resolve environment variables
    with open('config.yaml') as yaml_file:
        raw_config = yaml.safe_load(yaml_file)

    server_config = {
        account: {
            key: os.getenv(value.strip("${}"), value)  # Resolve env vars or use default
            for key, value in details.items()
        }
        for account, details in raw_config['EMAIL_ACCOUNTS'].items()
    }

    # Set up logging
    logger.info("Starting email sending process...")

    # Set up Google Sheets API
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    SERVICE_ACCOUNT_FILE = os.getenv("GOOGLE_APPLICATION_CREDENTIALS")
    if not SERVICE_ACCOUNT_FILE:
        logger.error("Google credentials file path is missing. Check your .env file.")
        exit(1)

    creds = ServiceAccountCredentials.from_json_keyfile_name(SERVICE_ACCOUNT_FILE, scope)
    client = gspread.authorize(creds)

    job_sender_name = os.getenv("JOB_SENDER_NAME")
    job_sender_number = os.getenv("JOB_SENDER_NUMBER")

    course_sender_name = os.getenv("COURSE_SENDER_NAME")
    course_sender_number = os.getenv("COURSE_SENDER_NUMBER")

    # Send emails
    try:
        send_job_email(job_sender_name,job_sender_number)
        send_course_email(course_sender_name, course_sender_number)

        logger.info("All emails processed.")
    except Exception as e:
        logger.error(f"Error during email processing: {e}")