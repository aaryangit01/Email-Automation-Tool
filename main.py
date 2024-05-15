import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os

#read email addresses from Excel spreadsheet
def read_emails_from_excel(excel_file, sheet_name, column_name):
    try:
        df = pd.read_excel(excel_file, sheet_name=sheet_name)
        return df[column_name].tolist()
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return []

#send email
def send_email(sender_email, sender_password, receiver_email, subject, body, attachment_path=None):
    try:
        message = MIMEMultipart()
        message['From'] = sender_email
        message['To'] = receiver_email
        message['Subject'] = subject
        message.attach(MIMEText(body, 'plain'))

        if attachment_path and os.path.exists(attachment_path):
            with open(attachment_path, 'rb') as attachment:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename= {attachment_path}')
            message.attach(part)

        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, receiver_email, message.as_string())
        server.quit()
        print(f"Email sent to {receiver_email}")
    except Exception as e:
        print(f"Error sending email to {receiver_email}: {e}")


if __name__ == "__main__":
    sender_email = 'yourgmail.com'
    sender_password = 'password'
    excel_file = 'excel.xlsx'
    sheet_name = 'Sheet1'
    column_name = 'Email'
    subject = 'Resume Submission'
    body = 'Dear recipient,\n\nPlease find attached my resume for your consideration.\n\nBest regards,\nYour Name'
    attachment_path = 'resume.pdf'

    email_addresses = read_emails_from_excel(excel_file, sheet_name, column_name)

    for email in email_addresses:
        send_email(sender_email, sender_password, email, subject, body, attachment_path)
