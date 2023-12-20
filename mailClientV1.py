import smtplib
import openpyxl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import credential

# SMTP server and port
smtp_server = credential.smtp_server
smtp_port = credential.smtp_port  # Port 25 for unauthenticated SMTP

# Email content
subject = 'Test Email with Attachment'
body = 'This is a test email with an attachment sent via SMTP on port 25 without authentication.'

# Excel file containing email addresses and attachment file names
excel_file = 'email_data.xlsx'

try:
    # Create an SMTP object without authentication
    server = smtplib.SMTP(smtp_server, smtp_port)

    # Load the Excel file
    workbook = openpyxl.load_workbook(excel_file)
    sheet = workbook.active

    for row in sheet.iter_rows(values_only=True):
        recipient_email = row[0]  # Email address from Column A
        attachment_file = row[1]  # Attachment file name from Column B

        # Create an email message
        message = MIMEMultipart()
        message['From'] = 'sender@example.com'
        message['To'] = recipient_email
        message['Subject'] = subject

        message.attach(MIMEText(body, 'plain'))

        with open(attachment_file, 'rb') as file:
            attachment = MIMEApplication(file.read(), _subtype="pdf")
            attachment.add_header('Content-Disposition', 'attachment', filename=attachment_file)
            message.attach(attachment)

        # Send the email
        server.sendmail('sender@example.com', recipient_email, message.as_string())

    # Quit the server
    server.quit()

    print("Emails sent successfully")

except Exception as e:
    print(f"An error occurred: {str(e)}")
