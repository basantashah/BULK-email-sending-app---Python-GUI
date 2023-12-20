# BULK-EMAIL-SENDING-APP

Kindly rename credential file as credential.py from credential1.py

## Overview

Helping lazy people send thousands of mails with unique recipients and attachment with customized body and subject.

'This python based mail server will help individual to sent multiple mails from relay server using customized Heading, Body and attachments'

## Table of Contents

- [Installation](#installation)
- [Usage](#usage)
- [Configuration](#configuration)
- [Dependencies](#dependencies)
- [File Structure](#file-structure)
- [Contributing](#contributing)
- [License](#license)

## Installation

For libraries and other, I have created a requirements.txt from where you can import all the packages required. In the end, I have exported .exe file for the use of laymans and daily operation people. Basically, for those who are not much into python programming and has basic to no idea on how to use it.


```bash
# Example installation commands
pip install -r requirements.txt


project-root/
|-- main.py
|-- email_body.txt
|-- credential.py
|-- gui_properties.py
|-- Ncell-logo.png
|-- requirements.txt
|-- README.md
|-- other_files/
```

## Usage

For now you can use 
python mailClientV2.py

Choose an Excel file containing recipient information and attachments.
Select the authentication option (with or without a password).
Enter the sender's email and, if applicable, the password.
Click the "Send Emails" button to initiate the email sending process.

I have made two version of this client, one with GUI and other without

### File Structure
project-root/
|-- main.py
|-- email_body.txt
|-- credential.py
|-- gui_properties.py
|-- Ncell-logo.png
|-- requirements.txt
|-- README.md
|-- other_files/


## Dependencies
smtplib
openpyxl
tkinter
Pillow (PIL)

## Contributing
Contributions are welcome! Feel free to submit bug reports, feature requests, or contribute code.

