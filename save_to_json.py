import imaplib
import email
from email.header import decode_header
import time
import re
from pathlib import Path
import tkinter as tk
from tkinter import filedialog

# Function to select the folder where PDFs will be saved
def select_folder():
    root = tk.Tk()
    root.withdraw()  # Hide root window
    folder_selected = filedialog.askdirectory(title="Select folder to save PDFs")
    if not folder_selected:
        print("No folder selected. Exiting.")
        exit()
    return Path(folder_selected)

# Function to clean filenames by removing unsafe characters
def sanitize_filename(filename):
    return re.sub(r'[^0-9a-zA-Z\.]+', '', filename)

# Function to connect to the IMAP server
def connect_imap(server, email_user, email_pass):
    try:
        mail = imaplib.IMAP4_SSL(server, port=993)  # Specify port explicitly
        mail.login(email_user, email_pass)  # Authenticate
        print("Login successful!")
        return mail
    except imaplib.IMAP4.error as e:
        print(f"Error connecting: {e}")
        return None

# Function to check the inbox and download PDF attachments from unread emails
def check_inbox(mail, re_dir):
    try:
        mail.select("inbox")  # Select inbox
        status, messages = mail.search(None, '(UNSEEN)')  # Only unread emails
        mail_ids = messages[0].split()

        if mail_ids:
            for mail_id in mail_ids:
                status, msg_data = mail.fetch(mail_id, "(RFC822)")
                for response_part in msg_data:
                    if isinstance(response_part, tuple):
                        msg = email.message_from_bytes(response_part[1])
                        subject, encoding = decode_header(msg["Subject"])[0]
                        if isinstance(subject, bytes):
                            subject = subject.decode(encoding if encoding else "utf-8")

                        print(f"Processing email: {subject}")

                        # Get attachments if available
                        if msg.is_multipart():
                            for part in msg.walk():
                                content_disposition = str(part.get("Content-Disposition"))
                                if "attachment" in content_disposition:
                                    filename = part.get_filename()
                                    if filename and filename.lower().endswith('.pdf'):
                                        filename = sanitize_filename(filename)
                                        filepath = re_dir / filename
                                        with open(filepath, "wb") as f:
                                            f.write(part.get_payload(decode=True))
                                        print(f"PDF saved: {filename}")
                        else:
                            print(f"No attachments in email: {subject}")

        else:
            print("No new emails.")
    
    except Exception as e:
        print(f"Error checking inbox: {e}")

# IONOS IMAP server settings
IMAP_SERVER = "imap.ionos.es"  # IONOS IMAP server
EMAIL_USER = "domus@asb-ibv.com"  # Your IONOS email
EMAIL_PASS = "!asb-ibv.com!"  # Your email password

# Connect to IONOS IMAP
mail = connect_imap(IMAP_SERVER, EMAIL_USER, EMAIL_PASS)

if mail:
    # Select the folder where PDFs will be saved
    re_dir = select_folder()

    # Loop to check for new emails every 30 seconds
    try:
        while True:
            check_inbox(mail, re_dir)
            print("Waiting 30 seconds for the next check...")
            time.sleep(30)
    except KeyboardInterrupt:
        print("Exiting script.")
    finally:
        mail.logout()  # Ensure proper logout
