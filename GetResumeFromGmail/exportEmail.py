import imaplib
import email
from email.header import decode_header
import os

# Gmail IMAP server
imap = imaplib.IMAP4_SSL("imap.gmail.com")

# Login
username = "jothivenkat.professional@gmail.com"
password = ""  # Use Gmail app password
imap.login(username, password)

# Select the inbox
imap.select("INBOX")

# Search for all emails (read + unread)
status, messages = imap.search(None, 'ALL')

# Specify the folder path where you want to save attachments using an f-string
base_path = r"E:\Venkat"
folder_name = "Extract_from_gmail"
attachment_dir = f"{base_path}\\{folder_name}"

# Create the folder if it doesn't exist
os.makedirs(attachment_dir, exist_ok=True)

attachment_found = False

# Process each email
for num in messages[0].split()[::-1]:  # Optional: reverse for latest first
    _, msg_data = imap.fetch(num, "(RFC822)")
    for response_part in msg_data:
        if isinstance(response_part, tuple):
            msg = email.message_from_bytes(response_part[1])

            # Decode subject
            subject, encoding = decode_header(msg["Subject"])[0]
            if isinstance(subject, bytes):
                subject = subject.decode(encoding or "utf-8", errors="ignore")
            print("Subject:", subject)
            print("From:", msg["From"])
            print("Date:", msg["Date"])

            # Check for attachments
            if msg.is_multipart():
                for part in msg.walk():
                    content_disposition = str(part.get("Content-Disposition") or "")
                    if "attachment" in content_disposition.lower():
                        filename = part.get_filename()
                        if filename:
                            decoded_filename, enc = decode_header(filename)[0]
                            if isinstance(decoded_filename, bytes):
                                filename = decoded_filename.decode(enc or "utf-8", errors="ignore")

                            # Normalize for comparison
                            filename_lower = filename.lower()
                            base_name, ext = os.path.splitext(filename_lower)

                            # Check if 'resume' is at the beginning or end of the filename (before extension)
                            if base_name.startswith("resume") or base_name.endswith("resume"):
                                filepath = os.path.join(attachment_dir, filename)
                                with open(filepath, "wb") as f:
                                    f.write(part.get_payload(decode=True))
                                print(f"Downloaded: {filepath}")
                                attachment_found = True
            else:
                print("Not a multipart message.")

# Final message
if not attachment_found:
    print("No attachment found.")

imap.logout()
