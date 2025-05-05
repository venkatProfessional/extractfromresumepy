from flask import Flask, render_template_string, request
import imaplib
import email
from email.header import decode_header
import os

app = Flask(__name__)

HTML_TEMPLATE = '''
<!doctype html>
<html>
<head><title>Resume Extractor</title></head>
<body>
    <h2>Resume Attachment Extractor</h2>
    <form method="post" action="/extract">
        <label>Email:</label><br>
        <input type="email" name="email" required><br><br>

        <label>App Password:</label><br>
        <input type="password" name="password" required><br><br>

        <button type="submit">Start Extraction</button>
    </form>
    {% if message %}
        <p><strong>{{ message }}</strong></p>
    {% endif %}
</body>
</html>
'''

@app.route("/", methods=["GET"])
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route("/extract", methods=["POST"])
def extract_resumes():
    email_user = request.form["email"]
    email_pass = request.form["password"]

    try:
        # Connect to Gmail
        imap = imaplib.IMAP4_SSL("imap.gmail.com")
        imap.login(email_user, email_pass)
        imap.select("INBOX")

        # Search all emails
        status, messages = imap.search(None, "ALL")
        message_ids = messages[0].split()[::-1]  # Reverse for latest first

        # Save folder
        base_path = r"E:\Venkat"
        folder_name = "Extract_from_gmail"
        attachment_dir = os.path.join(base_path, folder_name)
        os.makedirs(attachment_dir, exist_ok=True)

        attachment_found = False

        for num in message_ids:
            _, msg_data = imap.fetch(num, "(RFC822)")
            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_bytes(response_part[1])

                    subject, encoding = decode_header(msg["Subject"])[0]
                    if isinstance(subject, bytes):
                        subject = subject.decode(encoding or "utf-8", errors="ignore")
                    print("Subject:", subject)

                    if msg.is_multipart():
                        for part in msg.walk():
                            content_disposition = str(part.get("Content-Disposition") or "")
                            if "attachment" in content_disposition.lower():
                                filename = part.get_filename()
                                if filename:
                                    decoded_filename, enc = decode_header(filename)[0]
                                    if isinstance(decoded_filename, bytes):
                                        filename = decoded_filename.decode(enc or "utf-8", errors="ignore")

                                    base_name, ext = os.path.splitext(filename.lower())
                                    if base_name.startswith("resume") or base_name.endswith("resume"):
                                        filepath = os.path.join(attachment_dir, filename)
                                        with open(filepath, "wb") as f:
                                            f.write(part.get_payload(decode=True))
                                        print(f"Downloaded: {filepath}")
                                        attachment_found = True

        imap.logout()
        if attachment_found:
            return render_template_string(HTML_TEMPLATE, message="Resume attachments downloaded successfully.")
        else:
            return render_template_string(HTML_TEMPLATE, message="No resume attachments found.")

    except Exception as e:
        return render_template_string(HTML_TEMPLATE, message=f"Error: {str(e)}")

if __name__ == "__main__":
    app.run(debug=True)
