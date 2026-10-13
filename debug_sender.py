from tencent_mail import TencentMailClient

client = TencentMailClient('user@example.com', 'YOUR_PASSWORD')
if client.connect_all():
    emails = client.receive_emails(limit=100, days=21)

    print("Sample sender fields:")
    senders = set()
    for em in emails[:20]:
        sender = em.get('sender', '')
        senders.add(sender)
        print(f"  '{sender}'")

    print(f"\nAll unique senders:")
    for s in sorted(senders):
        print(f"  '{s}'")

    client.disconnect()
