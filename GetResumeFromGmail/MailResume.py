import imaplib
import email
import os

EMAIL = "jothivenkat.professional@gmail.com"
PASSWORD = "apgm mumi qgiq pysd"
IMAP_SERVER = "imap.gmail.com"
FOLDER_PATH = r"E:\Venkat\Extract_from_gmail"

def save_resume_attachments():
    # Connect to email server
    mail = imaplib.IMAP4_SSL(IMAP_SERVER)
    mail.login(EMAIL, PASSWORD)
    mail.select("inbox")

    # Search for emails with attachments (could add filters like "SUBJECT resume")
    result, data = mail.search(None, '(HASATTACHMENT)')
    email_ids = data[0].split()

    for eid in email_ids:
        result, msg_data = mail.fetch(eid, "(RFC822)")
        raw_email = msg_data[0][1]
        msg = email.message_from_bytes(raw_email)

        for part in msg.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get("Content-Disposition") is None:
                continue

            filename = part.get_filename()
            if filename and (filename.endswith(".pdf") or filename.endswith(".docx") or filename.endswith(".doc")):
                filepath = os.path.join(FOLDER_PATH, filename)
                with open(filepath, "wb") as f:
                    f.write(part.get_payload(decode=True))
                print(f"Saved: {filepath}")

    mail.logout()

save_resume_attachments()
