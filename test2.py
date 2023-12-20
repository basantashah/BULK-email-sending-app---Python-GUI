from exchangelib import Credentials, Message, DELEGATE, Account, Configuration
import credential

# Set your Exchange Server credentials
exchange_username = credential.email
exchange_password = credential.password
exchange_server = credential.exchange_server  # Replace with your Exchange Server URL

# Create Exchange Server credentials
credentials = Credentials(exchange_username, exchange_password)

config = Configuration(server=exchange_server)

# Connect to the Exchange Server account
account = Account(
    primary_smtp_address=exchange_username,
    credentials=credentials,
    autodiscover=False,
    access_type=DELEGATE,
    config=config
)

# Compose the email
email = Message(
    account=account,
    subject='Test Email',
    body='This is a test email sent via Microsoft Exchange Server.'
)

# Set the recipient email address
recipient_email = credential.receipt
email.to(recipient_email)

try:
    # Send the email
    email.send()
    print('Email sent successfully')
except Exception as e:
    print(f'An error occurred: {str(e)}')
