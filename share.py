'''
Final Project
Sydney Fowler and Matt Hileman
15-12-2019
Description: Emails a copy of the passed in file to a list of email addresses
'''

import os
import sys
import re
import smtplib
from getpass import getpass
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


# ================ REFERENCES ================
# Sending email attachments: https://www.tutorialspoint.com/send-mail-with-attachment-from-your-gmail-account-using-python
# Sending Excel attachments: https://stackoverflow.com/questions/25346001/add-excel-file-attachment-when-sending-python-email
# Rules for email_regex: https://help.returnpath.com/hc/en-us/articles/220560587-What-are-the-rules-for-email-address-syntax-


# ================== SETUP ===================

email_regex = re.compile(r'''(
([a-zA-Z0-9](([a-zA-Z0-9!#$%&'*+/=?^_`{|.-]){,62}[a-zA-Z0-9])?)     # Recipient name
(@)                                                                 # @ symbol
([a-zA-Z0-9](([a-zA-Z0-9.-]){,251}[a-zA-Z0-9])?)                    # Domain name
(\.)                                                                # . symbol
(com|org|net)                                                       # Top-level domain
)''', re.VERBOSE)


# ================ GET INPUT ================

# Get Excel file from user
while (True):  # Loop until you get a valid Excel file
    wb_path = input("Type path of your Excel file: ")
    if os.path.exists(wb_path):
        if wb_path[-5:] != ".xlsx":
            print("ERROR: Must be a .xlsx file.")
        else:
            break
    else:
        print("ERROR: Invalid file path.")

# Get email list from user
while (True):  # Loop until you get a valid text file
    email_list_file_path = input("Type path of the text file containing your email addresses: ")
    email_list_file_path = os.path.abspath(email_list_file_path)
    if os.path.exists(email_list_file_path):
        if email_list_file_path[-4:] == ".txt":
            try:
                email_list_file = open(email_list_file_path)
                email_list = email_list_file.readlines()
                email_list_file.close()
                break
            except Exception:
                print("ERROR: Unable to open file.")
        else:
            print("ERROR: Must be a .txt file.")
    else:
        print("ERROR: Invalid file path.")

# Get email and password of user
sender_email = input("Enter your email address (supports gmail, outlook, hotmail, and yahoo): ")
password = getpass("Enter your email password: ")


# ================ SEND EMAILS ================

# Determine smtp server
if sender_email.find("gmail") != -1:
    smtp = "smtp.gmail.com"
elif (sender_email.find("outlook") != -1) or (sender_email.find("hotmail") != -1):
    smtp = "smtp-mail.outlook.com"
elif sender_email.find("yahoo") != -1:
    smtp = "smtp.mail.yahoo.com"
else:
    print("ERROR: Invalid email address. Must be gmail, outlook, hotmail, or yahoo.")

# Connect to Server
session = smtplib.SMTP(smtp, 587)
session.ehlo()
session.starttls()

try:
    session.login(sender_email, password)
except Exception:
    print("ERROR: Unable to login to your email account")
    sys.exit()

invalid_emails = []
message_content = "Hello,\n\n" + sender_email + " has shared this file with you.\n\n"

for email in email_list:
    # Check the sending email is valid
    if not email_regex.search(email):
        invalid_emails.append(email)
        email_list.remove(email)
        continue

    # Setup the MIME
    message = MIMEMultipart()
    message['From'] = sender_email
    message['To'] = email
    message['Subject'] = "File Share From " + sender_email

    # The body and the attachments for the mail
    message.attach(MIMEText(message_content, 'plain'))
    attach_file = open(wb_path, 'rb')   # Open the file as binary mode
    payload = MIMEBase('application', 'vnd.ms-excel')
    payload.set_payload((attach_file).read())
    attach_file.close()
    encoders.encode_base64(payload)     # Encode the attachment

    # Add payload header with filename
    payload.add_header("Content-Disposition", "attachment", filename=os.path.basename(wb_path))
    message.attach(payload)

    # Create SMTP session for sending the email
    text = message.as_string()
    session.sendmail(sender_email, email, text)

session.quit()

# Print success message and list of invalid email addresses if there were any
print("Success!")
if len(invalid_emails) != 0:
    print("The following emails were deemed invalid: ")
    for email in invalid_emails:
        print(email)
