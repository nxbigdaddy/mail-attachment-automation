from tencent_mail import TencentMailClient
import os
from datetime import datetime

def main():
    username = "user@example.com"
    password = "YOUR_PASSWORD"

    client = TencentMailClient(username, password)

    print("=" * 60)
    print("Downloading attachments from Ding Hailong and Shui Bin")
    print("=" * 60)

    if not client.connect_all():
        print("Connection failed, exiting...")
        return

    save_dir = f"attachments_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    os.makedirs(save_dir, exist_ok=True)

    senders = ["示例员工", "水彬"]
    all_downloaded = []

    print("\nFetching all emails from last 3 weeks...")
    all_emails = client.receive_emails(limit=500, days=21)

    print(f"Total emails fetched: {len(all_emails)}")

    for sender in senders:
        print(f"\n{'=' * 60}")
        print(f"Filtering emails from: {sender}")
        print(f"{'=' * 60}")

        filtered_emails = [e for e in all_emails if sender in e.get('sender', '')]

        if not filtered_emails:
            print(f"No emails found from {sender}")
            continue

        print(f"Found {len(filtered_emails)} emails from {sender}")

        sender_dir = os.path.join(save_dir, sender)
        os.makedirs(sender_dir, exist_ok=True)

        downloaded = client.download_attachments(filtered_emails, save_dir=sender_dir)
        all_downloaded.extend(downloaded)

        print(f"\nDownloaded {len(downloaded)} attachments from {sender}")

    client.disconnect()

    print(f"\n{'=' * 60}")
    print("Download Summary")
    print(f"{'=' * 60}")
    print(f"Total attachments downloaded: {len(all_downloaded)}")
    print(f"Saved to directory: {save_dir}")

    if all_downloaded:
        print("\nDownloaded files:")
        for item in all_downloaded:
            print(f"  - {item['filename']}")
            print(f"    From: {item['email_subject']}")
            print(f"    Date: {item['email_date']}")
            print(f"    Path: {item['path']}")
            print()

if __name__ == "__main__":
    main()
