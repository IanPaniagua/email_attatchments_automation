import tkinter as tk
from tkinter import filedialog, ttk
import threading
from pathlib import Path
from email_downloader import connect_imap, check_inbox  # Import necessary functions

# Function to select the folder where PDFs will be saved
def select_folder():
    folder_selected = filedialog.askdirectory(title="Select folder to save PDFs")
    if not folder_selected:
        print("No folder selected. Exiting.")
        exit()
    return Path(folder_selected)

# Function to start the application
def start_app():
    root = tk.Tk()
    root.title("Email PDF Downloader")
    root.geometry("400x300")

    # Center the window on the screen
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width // 2) - (400 // 2)
    y = (screen_height // 2) - (300 // 2)
    root.geometry(f"400x300+{x}+{y}")

    tk.Label(root, text="Email:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
    email_entry = tk.Entry(root, width=30)
    email_entry.grid(row=0, column=1, padx=5, pady=5)

    tk.Label(root, text="Password:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
    password_entry = tk.Entry(root, show='*', width=30)
    password_entry.grid(row=1, column=1, padx=5, pady=5)

    tk.Label(root, text="Email Provider:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
    provider_combobox = ttk.Combobox(root, values=["IONOS", "Outlook", "Gmail", "Yahoo"])
    provider_combobox.grid(row=2, column=1, padx=5, pady=5)

    def submit():
        email_user = email_entry.get()
        email_pass = password_entry.get()
        provider = provider_combobox.get()

        provider_map = {
            "IONOS": "imap.ionos.es",
            "Outlook": "outlook.office365.com",
            "Gmail": "imap.gmail.com",
            "Yahoo": "imap.mail.yahoo.com"
        }
        server = provider_map.get(provider)

        if not server:
            print("Invalid provider selected.")
            return

        mail = connect_imap(server, email_user, email_pass)

        if mail:
            re_dir = select_folder()

            # Open a new window to display the success message
            success_window = tk.Toplevel(root)
            success_window.title("Connection Successful")
            success_window.geometry("300x100")
            success_label = tk.Label(success_window, text=f"Connection established successfully.\nEmail: {email_user}")
            success_label.pack(pady=20)

            # Start checking inbox in a new thread
            threading.Thread(target=check_inbox, args=(mail, re_dir, Path("data/email_info.json")), daemon=True).start()

            # Hide the main window
            root.withdraw()

    # Submit button
    submit_button = tk.Button(root, text="Start", command=submit)
    submit_button.grid(row=3, columnspan=2, pady=10)

    root.mainloop()

if __name__ == "__main__":
    start_app()
