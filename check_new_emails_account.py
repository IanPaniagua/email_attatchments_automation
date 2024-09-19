import win32com.client
import pythoncom  # Required for COM event handling
from pathlib import Path
import re

# Class to handle the new mail event
class NewMailHandler:
    def OnItemAdd(self, item):
        try:
            # Check if the item is a mail item (Class 43 is olMailItem)
            if item.Class == 43:
                attachments = item.Attachments

                # Process attachments
                for attachment in attachments:
                    # Only save PDFs
                    if attachment.FileName.lower().endswith('.pdf'):
                        # Create a safe filename
                        filename = re.sub(r'[^0-9a-zA-Z\.]+', '', attachment.FileName)

                        # Save the PDF to the folder
                        attachment.SaveAsFile(re_dir / filename)
                        print(f"PDF saved: {filename}")

        except Exception as e:
            print(f"Error processing new email: {e}")

# Ask if Outlook is in English or German
language = input("Is your Outlook in English or German? (Enter 'E' for English, 'G' for German): ").strip().lower()

# Connect to Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Get the list of all accounts in Outlook
accounts = outlook.Folders
account_names = [account.Name for account in accounts]

# Display available accounts and let the user select one
print("Available accounts:")
for idx, account_name in enumerate(account_names):
    print(f"{idx + 1}. {account_name}")

# Ask the user to select an account by number
selected_index = int(input("Enter the number of the account you want to use: ")) - 1
selected_account = accounts[selected_index]

# Determine the name of the inbox folder
if language == 'e':
    inbox_name = "Inbox"
elif language == 'g':
    inbox_name = "Posteingang"
else:
    print("Invalid input. Please enter 'E' for English or 'G' for German.")
    exit()

# Try to access the 'Inbox' or 'Posteingang' folder
try:
    inbox = selected_account.Folders[inbox_name]
    print(f"Connected to {selected_account.Name} - {inbox_name}")
except Exception as e:
    print(f"Error: Could not find the folder '{inbox_name}' for {selected_account.Name}.")
    print(f"Exception: {e}")
    exit()

# Set up the output folder for PDFs
re_dir = Path(r"C:\Users\MaxEDV\Desktop\re_")
re_dir.mkdir(parents=True, exist_ok=True)

# Set up the event handler for the Inbox
items = inbox.Items
event_handler = win32com.client.WithEvents(items, NewMailHandler)

# Keep the script running to listen for new emails
print("Monitoring for new emails...")

# Infinite loop to keep the script running
while True:
    # Process any waiting COM events
    pythoncom.PumpWaitingMessages()
