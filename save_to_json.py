import win32com.client
import pythoncom
from pathlib import Path
import re
from email_utils import save_email_info  # Import the function to save email info
import tkinter as tk
from tkinter import filedialog, messagebox

# Function to open a folder selection dialog and return the selected path
def select_folder():
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    root.attributes('-topmost', True)  # Ensure the dialog is always on top
    folder_selected = filedialog.askdirectory(title="Select Folder for Saving PDFs")
    root.attributes('-topmost', False)  # Remove the topmost attribute
    if not folder_selected:
        print("No folder selected. Exiting.")
        exit()
    return Path(folder_selected)

# Function to display a list of accounts and let the user select one
def select_account(accounts):
    def on_select():
        selected_idx = listbox.curselection()
        if selected_idx:
            selected_account.set(accounts[selected_idx[0]].Name)
            root.destroy()
        else:
            messagebox.showerror("Error", "No account selected.")

    root = tk.Tk()
    root.title("Select Email Account")
    root.attributes('-topmost', True)  # Ensure the dialog is always on top

    tk.Label(root, text="Select an email account:").pack(pady=10)

    listbox = tk.Listbox(root, width=50, height=10)
    listbox.pack(pady=10)

    for account in accounts:
        listbox.insert(tk.END, account.Name)

    selected_account = tk.StringVar()
    tk.Button(root, text="Select", command=on_select).pack(pady=10)

    root.mainloop()

    return selected_account.get()

# Class to handle the new mail event
class NewMailHandler:
    def OnItemAdd(self, item):
        try:
            # Check if the item is a mail item (Class 43 is olMailItem)
            if item.Class == 43:
                attachments = item.Attachments

                # Save email info to JSON
                email_info = {
                    "date": item.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S"),
                    "sender": item.SenderEmailAddress,
                    "subject": item.Subject
                }
                save_email_info(email_info)

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

# Connect to Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Get the list of all accounts in Outlook
accounts = outlook.Folders

# Filter out folders that start with "Öffentliche Ordner" to remove duplicates
filtered_accounts = [account for account in accounts if not account.Name.startswith("Öffentliche Ordner")]

# Let the user select an account via UI
selected_account_name = select_account(filtered_accounts)

# Find the selected account from the filtered accounts
selected_account = next((account for account in filtered_accounts if account.Name == selected_account_name), None)

if selected_account is None:
    print(f"Account '{selected_account_name}' not found. Exiting.")
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

# Default folder names
inbox_folder_name = "Posteingang"  # For German
# inbox_folder_name = "Inbox"  # Uncomment for English

# Find the Inbox folder
selected_folder = find_folder(selected_account.Folders, inbox_folder_name)

if selected_folder:
    print(f"Selected folder: {selected_folder.Name}")
else:
    print(f"Folder '{inbox_folder_name}' not found. Exiting.")
    exit()

# Let the user select the output folder for PDFs
re_dir = select_folder()

# Set up the event handler for the selected folder
items = selected_folder.Items
event_handler = win32com.client.WithEvents(items, NewMailHandler)

# Keep the script running to listen for new emails
print(f"Monitoring {selected_folder.Name} for new emails...")

# Infinite loop to keep the script running
while True:
    # Process any waiting COM events
    pythoncom.PumpWaitingMessages()
