import imaplib
import email
from email.header import decode_header
import os

# Account credentials
username = "domus@asb-ibv.com"  # Replace with your IONOS email
password = "!asb-ibv.com!"  # Replace with your email password
imap_server = "imap.ionos.es"  # IONOS IMAP server

def clean(text):
    # Clean text for creating a folder
    return "".join(c if c.isalnum() else "_" for c in text)

# Connect to the IMAP server
def connect_imap(server, email_user, email_pass):
    try:
        mail = imaplib.IMAP4_SSL(server, port=993)  # Specify port explicitly
        mail.login(email_user, email_pass)  # Authenticate
        print("Login successful!")
        return mail
    except imaplib.IMAP4.error as e:
        print(f"Error connecting: {e}")
        return None

# Check the inbox and download attachments
def download_attachments(mail):
    # Select the inbox or "Posteingang"
    mail.select("inbox")  # Change to "Posteingang" if necessary
    status, messages = mail.search(None, 'ALL')  # Fetch all emails
    mail_ids = messages[0].split()

    for mail_id in mail_ids:
        status, msg_data = mail.fetch(mail_id, "(RFC822)")
        for response_part in msg_data:
            if isinstance(response_part, tuple):
                msg = email.message_from_bytes(response_part[1])
                subject, encoding = decode_header(msg["Subject"])[0]
                if isinstance(subject, bytes):
                    subject = subject.decode(encoding if encoding else "utf-8")

                print(f"Processing email: {subject}")

                # Check for attachments
                if msg.is_multipart():
                    for part in msg.walk():
                        content_disposition = str(part.get("Content-Disposition"))
                        if "attachment" in content_disposition:
                            filename = part.get_filename()
                            if filename:
                                folder_name = clean(subject)
                                if not os.path.isdir(folder_name):
                                    os.mkdir(folder_name)  # Create folder for this email
                                filepath = os.path.join(folder_name, filename)
                                open(filepath, "wb").write(part.get_payload(decode=True))
                                print(f"Downloaded: {filename} to {folder_name}")
                else:
                    print("No attachments found in this email.")

# Connect to IONOS IMAP
mail = connect_imap(imap_server, username, password)

if mail:
    download_attachments(mail)
    mail.logout()  # Ensure proper logout
