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

# numbers = [1, 2, 3, 4, 5]
#
# print(numbers[::1])   # [1, 2, 3, 4, 5] — normal order
# print(numbers[::2])   # [1, 3, 5]       — every 2nd element
# print(numbers[::-1])  # [5, 4, 3, 2, 1] — reversed

# 'messages' is a list with one bytes object
# Example: [b'1 2 3 4 5 6 7 8 9 10 ...']

for num in messages[0].split():  # Split bytes into list like [b'1', b'2', ..., b'35']
    print(num.decode())  # Convert bytes to string before printing: '1', '2', ..., '35'


for num in messages[0].split()[::1]:  # Access first element (bytes), split by space, take all elements
    print(num)  # Each 'num' is a bytes object like b'1', b'2', etc.

    # "RFC822" is a standard format for the full email content (headers + body).
    # _, result = some_function()
    # status, _ = ('OK', 42)  # Ignore the second value

    # Fetch the email with the unique ID 'num' from the IMAP server.
    # The "(RFC822)" argument tells the server to return the full raw email (headers + body) in RFC 822 format.
    # The fetch method returns a tuple: (status, data).
    # We use '_' to ignore the status (like 'OK'), and store the actual message data in 'msg_data'.
    _, msg_data = imap.fetch(num, "(RFC822)")
    _ ,msg_data = imap.fetch(num, "(RFC822)")
    print(msg_data)
    for response_part in msg_data:
        print()
        #  checks if the response_part is tuple
        if isinstance(response_part, tuple):
            # print(response_part,"response part")
            msg = email.message_from_bytes(response_part[1])
            print(msg,"msg")




