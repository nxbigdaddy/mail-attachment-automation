import imaplib
import smtplib
import email
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os
from datetime import datetime, timedelta


class TencentMailClient:
    def __init__(self, username, password):
        self.username = username
        self.password = YOUR_PASSWORD
        self.imap_server = "imap.exmail.qq.com"
        self.smtp_server = "smtp.exmail.qq.com"
        self.imap_port = 993
        self.smtp_port = 465
        self.imap_conn = None
        self.smtp_conn = None

    def connect_imap(self):
        try:
            self.imap_conn = imaplib.IMAP4_SSL(self.imap_server, self.imap_port)
            self.imap_conn.login(self.username, self.password)
            print(f"IMAP connected successfully: {self.username}")
            return True
        except Exception as e:
            print(f"IMAP connection failed: {e}")
            return False

    def connect_smtp(self):
        try:
            self.smtp_conn = smtplib.SMTP_SSL(self.smtp_server, self.smtp_port)
            self.smtp_conn.login(self.username, self.password)
            print(f"SMTP connected successfully: {self.username}")
            return True
        except Exception as e:
            print(f"SMTP connection failed: {e}")
            return False

    def connect_all(self):
        imap_ok = self.connect_imap()
        smtp_ok = self.connect_smtp()
        return imap_ok and smtp_ok

    def disconnect(self):
        if self.imap_conn:
            self.imap_conn.logout()
            print("IMAP disconnected")
        if self.smtp_conn:
            self.smtp_conn.quit()
            print("SMTP disconnected")

    def receive_emails(self, folder="INBOX", limit=10, days=None):
        if not self.imap_conn:
            if not self.connect_imap():
                return []

        try:
            self.imap_conn.select(folder)

            search_criteria = "ALL"
            if days:
                since_date = (datetime.now() - timedelta(days=days)).strftime("%d-%b-%Y")
                search_criteria = f'SINCE {since_date}'

            status, messages = self.imap_conn.search(None, search_criteria)
            email_ids = messages[0].split()

            if not email_ids:
                print("No emails found")
                return []

            emails = []
            total = len(email_ids)
            start = max(0, total - limit)

            for email_id in email_ids[start:total]:
                status, msg_data = self.imap_conn.fetch(email_id, "(RFC822)")
                if status == "OK":
                    raw_email = msg_data[0][1]
                    msg = email.message_from_bytes(raw_email)
                    email_info = self._parse_email(msg)
                    email_info['id'] = email_id
                    emails.append(email_info)

            return emails
        except Exception as e:
            print(f"Failed to receive emails: {e}")
            return []

    def search_emails_by_sender(self, sender_keyword, folder="INBOX", days=21):
        if not self.imap_conn:
            if not self.connect_imap():
                return []

        try:
            self.imap_conn.select(folder)
            since_date = (datetime.now() - timedelta(days=days)).strftime("%d-%b-%Y")
            search_criteria = f'SINCE {since_date} FROM "{sender_keyword}"'

            status, messages = self.imap_conn.search(None, search_criteria)
            email_ids = messages[0].split()

            if not email_ids:
                print(f"No emails found from {sender_keyword}")
                return []

            emails = []
            for email_id in email_ids:
                status, msg_data = self.imap_conn.fetch(email_id, "(RFC822)")
                if status == "OK":
                    raw_email = msg_data[0][1]
                    msg = email.message_from_bytes(raw_email)
                    email_info = self._parse_email(msg)
                    email_info['id'] = email_id
                    emails.append(email_info)

            return emails
        except Exception as e:
            print(f"Failed to search emails: {e}")
            return []

    def download_attachments(self, emails, save_dir="attachments"):
        if not os.path.exists(save_dir):
            os.makedirs(save_dir)

        downloaded_files = []
        for email_info in emails:
            if not email_info.get('attachments'):
                continue

            msg_num = email_info['id']
            self.imap_conn.select("INBOX")

            try:
                status, msg_data = self.imap_conn.fetch(msg_num, "(BODY.PEEK[])")
                if status != "OK":
                    continue

                raw_email = msg_data[0][1]
                msg = email.message_from_bytes(raw_email)

                for part in msg.walk():
                    disposition = str(part.get("Content-Disposition"))
                    if "attachment" in disposition:
                        filename = part.get_filename()
                        if filename:
                            decoded_filename = self._decode_filename(filename)
                            save_path = os.path.join(save_dir, f"{msg_num}_{decoded_filename}")

                            with open(save_path, "wb") as f:
                                f.write(part.get_payload(decode=True))

                            downloaded_files.append({
                                "email_subject": email_info['subject'],
                                "email_date": email_info['date'],
                                "filename": decoded_filename,
                                "path": save_path
                            })
                            print(f"Downloaded: {decoded_filename}")
            except Exception as e:
                print(f"Failed to download attachment from email {msg_num}: {e}")

        return downloaded_files

    def _decode_filename(self, filename):
        decoded = email.header.decode_header(filename)
        if decoded:
            part = decoded[0]
            if isinstance(part[0], bytes):
                charset = part[1] or "utf-8"
                try:
                    return part[0].decode(charset, errors="ignore")
                except:
                    return part[0].decode("utf-8", errors="ignore")
            return part[0]
        return filename

    def _parse_email(self, msg):
        subject = self._get_header(msg, "Subject")
        sender = self._get_header(msg, "From")
        recipient = self._get_header(msg, "To")
        date = self._get_header(msg, "Date")

        body = ""
        attachments = []

        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                disposition = str(part.get("Content-Disposition"))

                if content_type == "text/plain" and "attachment" not in disposition:
                    try:
                        body = part.get_payload(decode=True).decode("utf-8", errors="ignore")
                    except:
                        body = part.get_payload(decode=True).decode("gbk", errors="ignore")
                elif "attachment" in disposition:
                    filename = part.get_filename()
                    if filename:
                        attachments.append(filename)
        else:
            try:
                body = msg.get_payload(decode=True).decode("utf-8", errors="ignore")
            except:
                body = msg.get_payload(decode=True).decode("gbk", errors="ignore")

        return {
            "subject": subject,
            "sender": sender,
            "recipient": recipient,
            "date": date,
            "body": body[:500],
            "attachments": attachments
        }

    def _get_header(self, msg, header_name):
        header = msg.get(header_name, "")
        if header:
            decoded = email.header.decode_header(header)
            if decoded:
                result = ""
                for part, charset in decoded:
                    if isinstance(part, bytes):
                        try:
                            charset = charset or "utf-8"
                            result += part.decode(charset, errors="ignore")
                        except:
                            result += str(part)
                    else:
                        result += str(part)
                return result.strip()
        return ""

    def send_email(self, to_addr, subject, body, is_html=False, attachments=None):
        if not self.smtp_conn:
            if not self.connect_smtp():
                return False

        try:
            msg = MIMEMultipart()
            msg["From"] = self.username
            msg["To"] = to_addr
            msg["Subject"] = subject
            msg["Date"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            content_type = "html" if is_html else "plain"
            msg.attach(MIMEText(body, content_type, "utf-8"))

            if attachments:
                for file_path in attachments:
                    if os.path.exists(file_path):
                        with open(file_path, "rb") as f:
                            part = MIMEBase("application", "octet-stream")
                            part.set_payload(f.read())
                        encoders.encode_base64(part)
                        filename = os.path.basename(file_path)
                        part.add_header("Content-Disposition", f"attachment; filename={filename}")
                        msg.attach(part)

            self.smtp_conn.sendmail(self.username, to_addr, msg.as_string())
            print(f"Email sent successfully to {to_addr}")
            return True
        except Exception as e:
            print(f"Failed to send email: {e}")
            return False

    def list_folders(self):
        if not self.imap_conn:
            if not self.connect_imap():
                return []

        try:
            status, folders = self.imap_conn.list()
            return folders
        except Exception as e:
            print(f"Failed to list folders: {e}")
            return []


def main():
    username = "user@example.com"
    password = "YOUR_PASSWORD"

    client = TencentMailClient(username, password)

    print("=" * 50)
    print("Connecting to Tencent Enterprise Email...")
    print("=" * 50)

    if not client.connect_all():
        print("Connection failed, exiting...")
        return

    while True:
        print("\n" + "=" * 50)
        print("Email Client Menu")
        print("=" * 50)
        print("1. Receive recent 3 weeks emails")
        print("2. Send email")
        print("3. List folders")
        print("4. Exit")
        print("=" * 50)

        choice = input("Please select (1-4): ").strip()

        if choice == "1":
            emails = client.receive_emails(limit=100, days=21)
            if emails:
                print(f"\nFound {len(emails)} emails from last 3 weeks:\n")
                for i, em in enumerate(emails, 1):
                    print(f"--- Email {i} ---")
                    print(f"From: {em['sender']}")
                    print(f"Subject: {em['subject']}")
                    print(f"Date: {em['date']}")
                    print(f"Body: {em['body'][:150]}...")
                    if em['attachments']:
                        print(f"Attachments: {', '.join(em['attachments'])}")
                    print()
            else:
                print("No emails found in last 3 weeks")

        elif choice == "2":
            to_addr = input("Recipient email address: ").strip()
            subject = input("Subject: ").strip()
            print("Body (press Ctrl+Z on Windows / Ctrl+D to finish):")
            lines = []
            try:
                while True:
                    lines.append(input())
            except EOFError:
                pass
            body = "\n".join(lines)
            client.send_email(to_addr, subject, body)

        elif choice == "3":
            folders = client.list_folders()
            print("\nFolders:")
            for f in folders:
                print(f"  {f}")

        elif choice == "4":
            client.disconnect()
            print("Goodbye!")
            break

        else:
            print("Invalid option, please try again.")


if __name__ == "__main__":
    main()
