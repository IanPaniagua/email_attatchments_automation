import imaplib
import email
from email.header import decode_header
import time
import os
import re
import json
from pathlib import Path

# User configuration
EMAIL_USER = "domus@asb-ibv.com"  # Enter your email here
EMAIL_PASS = "!asb-ibv.com!"  # Enter your password here
IMAP_SERVER = "imap.ionos.es"

# Function to clean filenames by removing unsafe characters
def sanitize_filename(filename):
    return re.sub(r'[^0-9a-zA-Z\.]+', '', filename)

# Function to connect to the IMAP server
def connect_imap(server, email_user, email_pass):
    try:
        mail = imaplib.IMAP4_SSL(server, port=993)
        mail.login(email_user, email_pass)
        print("Login successful!")
        return mail
    except imaplib.IMAP4.error as e:
        print(f"Error connecting: {e}")
        return None

# Function to save email information to JSON
def save_email_info(email_data, json_file):
    try:
        if os.path.exists(json_file):
            with open(json_file, "r") as f:
                try:
                    data = json.load(f)
                except json.JSONDecodeError:
                    print("Error decoding JSON, starting fresh.")
                    data = []
        else:
            data = []

        data.append(email_data)

        with open(json_file, "w") as f:
            json.dump(data, f, indent=4)
    except Exception as e:
        print(f"Error saving email info: {e}")

# Function to download all attachments from unread emails
def check_inbox(mail, re_dir, json_file):
    try:
        mail.select("inbox")
        status, messages = mail.search(None, '(UNSEEN)')
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

                        date = msg.get("Date")
                        sender, encoding = decode_header(msg["From"])[0]
                        if isinstance(sender, bytes):
                            sender = sender.decode(encoding if encoding else "utf-8")

                        print(f"Processing email: {subject}")

                        email_data = {
                            "date": date,
                            "sender": sender,
                            "subject": subject
                        }
                        save_email_info(email_data, json_file)

                        if msg.is_multipart():
                            print("Email is multipart, checking attachments...")
                            for part in msg.walk():
                                content_disposition = str(part.get("Content-Disposition"))
                                if "attachment" in content_disposition:
                                    filename = part.get_filename()
                                    if filename:
                                        filename = sanitize_filename(filename)
                                        filepath = re_dir / filename
                                        print(f"Saving to: {filepath}")  # Print the full file path
                                        try:
                                            with open(filepath, "wb") as f:
                                                f.write(part.get_payload(decode=True))
                                            print(f"Attachment saved: {filename}")
                                        except Exception as e:
                                            print(f"Error saving attachment {filename}: {e}")
                        else:
                            print(f"No attachments in email: {subject}")
        else:
            print("No new emails.")
    
    except Exception as e:
        print(f"Error checking inbox: {e}")

# Main function to run the script
def main():
    re_dir = Path(r"C:\Users\MaxEDV\Desktop\re2_")  # Specify your save directory
    re_dir.mkdir(parents=True, exist_ok=True)  # Create the directory if it doesn't exist

    mail = connect_imap(IMAP_SERVER, EMAIL_USER, EMAIL_PASS)

    if mail:
        json_file = Path("data/email_info.json")
        while True:
            check_inbox(mail, re_dir, json_file)
            print("Waiting 30 seconds for the next check...")
            time.sleep(30)
    else:
        print("Failed to connect to the mail server.")

if __name__ == "__main__":
    main()
