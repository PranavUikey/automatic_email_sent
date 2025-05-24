# Automatic Email Sender

This project automates the process of sending emails for job collaborations and course details using data from Google Sheets. It uses the Google Sheets API to fetch data and sends emails via SMTP.

---

## Features
- Fetches data from Google Sheets.
- Sends job collaboration emails.
- Sends course detail emails with personalized links.
- Logs email statuses in Google Sheets.

---

## Prerequisites
1. **Python**: Ensure Python 3.9+ is installed.
2. **Google Cloud Service Account**:
   - Create a service account in the Google Cloud Console.
   - Enable the Google Sheets API and Google Drive API.
   - Download the service account JSON file and place it in the project directory.
3. **Hostinger Email Account**:
   - Ensure SMTP is enabled for your email account.
   - Use app-specific passwords if required.

---

## Installation
1. Clone the repository:
   ```bash
   git clone https://github.com/your-username/automatic_email_sent.git
   cd automatic_email_sent