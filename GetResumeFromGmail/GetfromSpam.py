import imaplib
import email
from email.header import decode_header
import os

# Gmail IMAP server
imap = imaplib.IMAP4_SSL("imap.gmail.com")

# Login
username = "jothivenkat.professional@gmail.com"
password = "apgm mumi qgiq pysd"  # Use Gmail app password
imap.login(username, password)

# Specify the folder path where you want to save attachments
base_path = r"E:\Venkat"
folder_name = "Extract_from_gmail"
attachment_dir = os.path.join(base_path, folder_name)
os.makedirs(attachment_dir, exist_ok=True)

def process_mailbox(mailbox_name):
    status, _ = imap.select(mailbox_name)
    if status != "OK":
        print(f"Failed to select mailbox: {mailbox_name}")
        return

    status, messages = imap.search(None, 'ALL')
    if status != "OK":
        print(f"Failed to search in {mailbox_name}")
        return

    for num in messages[0].split()[::-1]:  # reverse to get latest first
        _, msg_data = imap.fetch(num, "(RFC822)")
        for response_part in msg_data:
            if isinstance(response_part, tuple):
                msg = email.message_from_bytes(response_part[1])

                # Decode subject
                subject, encoding = decode_header(msg["Subject"])[0]
                if isinstance(subject, bytes):
                    subject = subject.decode(encoding or "utf-8", errors="ignore")
                print(f"\nMailbox: {mailbox_name}")
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

                                filename_lower = filename.lower()
                                base_name, ext = os.path.splitext(filename_lower)

                                if base_name.startswith("resume") or base_name.endswith("resume"):
                                    filepath = os.path.join(attachment_dir, filename)
                                    with open(filepath, "wb") as f:
                                        f.write(part.get_payload(decode=True))
                                    print(f"Downloaded: {filepath}")
                else:
                    print("Not a multipart message.")

# Process INBOX and SPAM
process_mailbox("INBOX")
process_mailbox("[Gmail]/Spam")

imap.logout()
