import imaplib
import email
from datetime import datetime, timedelta

username = "user@example.com"
password = "YOUR_PASSWORD"
imap_server = "imap.exmail.qq.com"
imap_port = 993

imap_conn = imaplib.IMAP4_SSL(imap_server, imap_port)
imap_conn.login(username, password)
imap_conn.select("INBOX")

since_date = (datetime.now() - timedelta(days=21)).strftime("%d-%b-%Y")
status, messages = imap_conn.search(None, f'SINCE {since_date}')
email_ids = messages[0].split()[-10:]

for email_id in email_ids:
    status, msg_data = imap_conn.fetch(email_id, "(RFC822)")
    if status == "OK":
        raw_email = msg_data[0][1]
        msg = email.message_from_bytes(raw_email)

        raw_from = msg.get("From", "")
        print(f"Email ID: {email_id}")
        print(f"Raw From: {repr(raw_from)}")

        decoded = email.header.decode_header(raw_from)
        print(f"Decoded: {decoded}")
        print()
