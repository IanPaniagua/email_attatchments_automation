import imaplib

# IONOS IMAP server settings
IMAP_SERVER = "imap.ionos.es"  # Change this to your IONOS IMAP server address
EMAIL_USER = "domus@asb-ibv.com"  # Replace with your IONOS email
EMAIL_PASS = "!asb-ibv.com!"  # Replace with your email password

# Function to connect to the IMAP server
def connect_imap(server, email_user, email_pass):
    try:
        mail = imaplib.IMAP4_SSL(server)  # Create IMAP4 SSL connection
        mail.login(email_user, email_pass)  # Authenticate
        print("Login successful!")
        return mail
    except imaplib.IMAP4.error as e:
        print(f"Error connecting: {e}")

# Connect to IONOS IMAP
mail = connect_imap(IMAP_SERVER, EMAIL_USER, EMAIL_PASS)

# Logout if connected successfully
if mail:
    mail.logout()
