# EmailAutomationTool

EmailAutomationTool is a Python script designed to automate the process of sending personalized emails with attachments to multiple recipients. The email addresses are read from an Excel spreadsheet, making it easy to manage and update the list of recipients.

## Features

- Read email addresses from an Excel file
- Send personalized emails to multiple recipients
- Attach files to the emails
- Use secure SMTP connection for sending emails

## Requirements

- Python 3.x
- pandas
- openpyxl
- smtplib

## Installation

1. Clone the repository or download the script.
2. Install the required Python packages using pip:

    ```bash
    pip install pandas openpyxl
    ```

3. Update the script with your email credentials and file paths.

## Usage

1. Ensure you have an Excel file (`excel.xlsx`) with a sheet (`Sheet1`) containing a column (`Email`) with the email addresses of the recipients.
2. Update the following variables in the script with your details:

    ```python
    sender_email = 'yourgmail@gmail.com'
    sender_password = 'yourpassword'
    excel_file = 'excel.xlsx'
    sheet_name = 'Sheet1'
    column_name = 'Email'
    subject = 'Resume Submission'
    body = 'Dear recipient,\n\nPlease find attached my resume for your consideration.\n\nBest regards,\nYour Name'
    attachment_path = 'resume.pdf'
    ```

3. Run the script:

    ```bash
    python email_automation_tool.py
    ```

4. The script will read the email addresses from the specified Excel file and send the email with the attachment to each recipient.

## Functions

### `read_emails_from_excel(excel_file, sheet_name, column_name)`

Reads email addresses from the specified Excel file, sheet, and column.

- `excel_file`: Path to the Excel file.
- `sheet_name`: Name of the sheet containing the email addresses.
- `column_name`: Name of the column containing the email addresses.

Returns a list of email addresses.

### `send_email(sender_email, sender_password, receiver_email, subject, body, attachment_path=None)`

Sends an email with the specified subject, body, and attachment to the receiver.

- `sender_email`: Sender's email address.
- `sender_password`: Sender's email password.
- `receiver_email`: Receiver's email address.
- `subject`: Subject of the email.
- `body`: Body of the email.
- `attachment_path`: Path to the file to be attached (optional).

## Contributing

If you would like to contribute to this project, please open an issue or submit a pull request.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
