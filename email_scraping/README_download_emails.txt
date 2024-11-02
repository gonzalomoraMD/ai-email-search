## En este código:

Se utiliza imaplib para conectarse al servidor IMAP de tu correo corporativo.
Las credenciales de acceso (dirección de correo electrónico y contraseña) se obtienen desde variables de entorno (EMAIL_ADDRESS y EMAIL_PASSWORD).
Se define una función download_emails que se encarga de buscar y descargar los correos electrónicos desde una carpeta específica en el servidor IMAP.
Se utiliza email.message_from_bytes para procesar y guardar los correos electrónicos en formato .eml en carpetas locales.
Asegúrate de ajustar IMAP_SERVER, IMAP_PORT y los nombres de las carpetas según la configuración de tu servidor de correo corporativo.


import os
import imaplib
import email
import base64
from dotenv import load_dotenv
import time
import random

# Load environment variables
load_dotenv()

# Function to clean up the folder names and save emails
def save_email(email_message, folder, email_id):
    try:
        # Clean up the folder name
        folder = ''.join(c for c in folder if c.isalnum() or c in (' ', '_', '-'))
        folder = folder.strip()

        # Create the folder if it doesn't exist
        os.makedirs(folder, exist_ok=True)

        # Save the email
        filename = os.path.join(folder, f"{email_id}.eml")
        with open(filename, 'wb') as f:
            f.write(email_message)
        print(f"Saved email with ID: {email_id} to {filename}")
        return True
    except Exception as e:
        print(f"Error saving email with ID {email_id}: {str(e)}")
        return False

# IMAP server details (replace with your corporate email server details)
IMAP_SERVER = os.getenv("IMAP_SERVER")
IMAP_PORT = os.getenv("IMAP_PORT")
EMAIL_ADDRESS = os.getenv("EMAIL_ADDRESS")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")

# Connect to the IMAP server
mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
mail.login(EMAIL_ADDRESS, EMAIL_PASSWORD)

def download_emails(folder_name, local_folder):
    # Select the folder
    mail.select(folder_name)
    
    # Search for all emails in the selected folder
    result, data = mail.search(None, 'ALL')
    if result == 'OK':
        email_ids = data[0].split()
        print(f"Found {len(email_ids)} emails in {folder_name}")
        
        saved_count = 0
        for email_id in email_ids:
            result, msg_data = mail.fetch(email_id, '(RFC822)')
            if result == 'OK':
                try:
                    raw_email = msg_data[0][1]
                    email_message = email.message_from_bytes(raw_email)
                    if save_email(raw_email, local_folder, email_id.decode()):
                        saved_count += 1
                except Exception as e:
                    print(f"Error processing email {email_id.decode()}: {str(e)}")
            else:
                print(f"Failed to retrieve email {email_id}")
        
        print(f"Downloaded and saved {saved_count} out of {len(email_ids)} emails from {folder_name}")
    else:
        print(f"Failed to fetch emails from {folder_name}")

    mail.close()

print("Starting email download process...")

# Download emails from the inbox
print("\nDownloading inbox emails...")
download_emails("INBOX", "Emails/Inbox")

# Download emails from the sent folder (replace with your folder name)
print("\nDownloading sent emails...")
download_emails("Sent", "Emails/Sent")

mail.logout()
print("\nEmail download process complete.")
