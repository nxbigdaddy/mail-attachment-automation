from tencent_mail import TencentMailClient

client = TencentMailClient('user@example.com', 'YOUR_PASSWORD')
if client.connect_all():
    emails = client.receive_emails(limit=100, days=21)
    if emails:
        print(f'\nFound {len(emails)} emails from last 3 weeks:\n')
        for i, em in enumerate(emails, 1):
            print(f'--- Email {i} ---')
            print(f'From: {em["sender"]}')
            print(f'Subject: {em["subject"]}')
            print(f'Date: {em["date"]}')
            print(f'Body: {em["body"][:200]}...')
            if em['attachments']:
                print(f'Attachments: {", ".join(em["attachments"])}')
            print()
    else:
        print('No emails found in last 3 weeks')
    client.disconnect()
