import win32com.client
import pythoncom
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

# Filter out folders that start with "Öffentliche Ordner" to remove duplicates
filtered_accounts = [account for account in accounts if not account.Name.startswith("Öffentliche Ordner")]

# Display available accounts and let the user select one
print("Available accounts:")
for idx, account in enumerate(filtered_accounts):
    print(f"{idx + 1}. {account.Name}")

# Ask the user to select an account by number
try:
    selected_index = int(input("Enter the number of the account you want to use: ")) - 1
    if selected_index < 0 or selected_index >= len(filtered_accounts):
        raise ValueError("Invalid account number.")
    selected_account = filtered_accounts[selected_index]
except (ValueError, IndexError) as e:
    print(f"Error: {e}")
    exit()

# Print the selected account for debugging
print(f"Selected account: {selected_account.Name}")

# Function to find folder by name, including subfolders
def find_folder(folders, name):
    """Find a folder by name, including subfolders."""
    for folder in folders:
        if folder.Name.lower() == name.lower():
            return folder
        # Recursively search in subfolders
        if folder.Folders.Count > 0:
            found_folder = find_folder(folder.Folders, name)
            if found_folder:
                return found_folder
    return None

# Ask the user to enter the name of the folder they want to use, and validate it
while True:
    folder_name = input("Enter the name of the folder you want to use: ").strip()
    selected_folder = find_folder(selected_account.Folders, folder_name)
    
    if selected_folder:
        print(f"Selected folder: {selected_folder.Name}")
        break
    else:
        print(f"Folder '{folder_name}' not found. Please enter a valid folder name.")

# Set up the output folder for PDFs
re_dir = Path(r"C:\Users\MaxEDV\Desktop\re_")
re_dir.mkdir(parents=True, exist_ok=True)

# Set up the event handler for the selected folder
items = selected_folder.Items
event_handler = win32com.client.WithEvents(items, NewMailHandler)

# Keep the script running to listen for new emails
print(f"Monitoring {selected_folder.Name} for new emails...")

# Infinite loop to keep the script running
while True:
    # Process any waiting COM events
    pythoncom.PumpWaitingMessages()
