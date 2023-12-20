import smtplib
import openpyxl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import tkinter as tk
from tkinter import filedialog
# from tkinter import simpledialog
from PIL import Image, ImageTk
# from tkinter import PhotoImage


import gui_properties
from gui_properties import window_title, window_width, window_height, background_color, button_color, copyright_text, label_font


# Import the SMTP server and port from credential.py
import credential

# Function to send emails
def send_emails():
    smtp_server = credential.smtp_server
    smtp_port = credential.smtp_port


    # Get the selected Excel file
    excel_file_path = file_path_var.get()

    try:
        # Create an SMTP object without authentication
        server = smtplib.SMTP(smtp_server, smtp_port)

        # Load the Excel file
        workbook = openpyxl.load_workbook(excel_file_path)
        sheet = workbook.active

        for row in sheet.iter_rows(values_only=True, min_row=2):
            recipient_email = row[0]  # Email address from Column A
            attachment_file = row[2]  # Attachment file name from Column B
            subject = row[3]
            body = row[4]
            cc_recipients = row[1].split(",")
            

            # Create an email message
            message = MIMEMultipart()
            message['From'] = sender_email_var.get()
            message['To'] = recipient_email
            message['Cc'] = ", ".join(cc_recipients) 
            message['Subject'] = subject
            
            # with open("signature.py", "r") as body_file:
            #     email_body = body_file.read()


            message.attach(MIMEText(body, 'plain'))
            # message.attach(MIMEText(email_body, 'html'))

            with open(attachment_file, 'rb') as file:
                attachment = MIMEApplication(file.read(), _subtype="pdf")
                attachment.add_header('Content-Disposition', 'attachment', filename=attachment_file)
                message.attach(attachment)

            # Send the email
            recipients = [recipient_email] + cc_recipients
            server.sendmail(sender_email_var.get(), recipients, message.as_string())

        # Quit the server
        server.quit()

        result_label.config(text="Emails sent successfully")

    except Exception as e:
        result_label.config(text=f"An error occurred: {str(e)}")


# Function to toggle between authentication options
def toggle_authentication():
    if authentication_option_var.get() == "Without Authentication":
        sender_email_entry.config(state='normal')
        sender_password_entry.config(state='disabled')
    else:
        sender_email_entry.config(state='normal')
        sender_password_entry.config(state='normal')

# Create the main application window

app = tk.Tk()
app.title(window_title)
app.geometry(f"{window_width}x{window_height}")
app.configure(bg=background_color)


# Choose Excel File Button
choose_file_button = tk.Button(app, text="Choose Excel File", command=lambda: file_path_var.set(filedialog.askopenfilename()), bg=button_color)
choose_file_button.pack()
choose_file_button.config(width=20,height=2)
choose_file_button.place(x=20, y=150)


# Authentication Option
authentication_option_var = tk.StringVar()
authentication_option_var.set("Without Authentication")
without_auth_radio = tk.Radiobutton(app, text="Without Authentication", variable=authentication_option_var, value="Without Authentication", command=toggle_authentication)
with_auth_radio = tk.Radiobutton(app, text="With Authentication", variable=authentication_option_var, value="With Authentication", command=toggle_authentication)
without_auth_radio.pack()
with_auth_radio.pack()
without_auth_radio.place(x=20, y=100)
with_auth_radio.place(x=20, y=120)

# Sender Email Label and Entry
sender_email_var = tk.StringVar()
sender_email_label = tk.Label(app, text="Sender Email:", bg=background_color,pady=2, padx=2, font=label_font, fg='white')
sender_email_label.place(x=20, y=120)
sender_email_entry = tk.Entry(app, textvariable=sender_email_var, state='normal', width=40)
sender_email_entry.place(x=130, y=120)  # Adjusted x-coordinate


def check_sender_email():
    sender_email = sender_email_var.get()
    if sender_email:
        send_button.config(state='normal')
    else:
        send_button.config(state='disabled')

sender_email_entry.bind("<FocusOut>", lambda e: check_sender_email())


# Sender Password Label and Entry
sender_password_var = tk.StringVar()
sender_password_label = tk.Label(app, text="Password:", bg=background_color, pady=2, padx=2, font=label_font, fg='white')
sender_password_label.place(x=20, y=160)
sender_password_entry = tk.Entry(app, textvariable=sender_password_var, show="*", state='disabled')
sender_password_entry.place(x=130, y=160) 


sender_email_label.pack()
sender_email_entry.pack()
sender_password_label.pack()
sender_password_entry.pack()

# Result Label
result_label = tk.Label(app, text="", bg=background_color)
result_label.place(x=20, y=240)

# Send Button
send_button = tk.Button(app, text="Send Emails", command=send_emails, state='disabled')
send_button.pack()
send_button.place(x=20, y=200)


# Copyright Label
copyright_label = tk.Label(app, text=copyright_text, bg=background_color)
copyright_label.pack()
copyright_label.place(x=400, y=300)

# App logo
png_image = Image.open("Ncell-logo.png")
photo_image = ImageTk.PhotoImage(png_image)
app.tk.call('wm', 'iconphoto', app._w, photo_image)



# Excel file path variable
file_path_var = tk.StringVar()
app.resizable(False,False)
app.mainloop()

