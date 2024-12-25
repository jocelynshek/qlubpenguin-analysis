import mailbox
import csv
from email.header import decode_header
from datetime import datetime
from pytz import timezone, utc

# Define timezones
eastern = timezone('US/Eastern')

# Open MBOX file
mbox = mailbox.mbox('Qlubpenguin.mbox')

# Write to CSV
with open('emails_timeadjusted.csv', 'w', newline='', encoding='utf-8') as csvfile:
    writer = csv.writer(csvfile)
    writer.writerow(['Subject', 'From', 'Date', 'Time', 'Body'])

    for message in mbox:
        # Decode the subject line
        raw_subject = message['subject']
        subject = None  # Initialize subject as None for fallback

        if raw_subject:
            try:
                # Decode the subject
                decoded_parts = decode_header(raw_subject)
                subject_parts = []
                for part, encoding in decoded_parts:
                    if isinstance(part, bytes):  # Decode bytes
                        subject_parts.append(part.decode(encoding or 'utf-8', errors='ignore'))
                    else:  # Append strings directly
                        subject_parts.append(part)
                subject = ''.join(subject_parts)
            except Exception as e:
                print(f"Error decoding subject: {raw_subject} -> {e}")
                subject = raw_subject  # Fallback to raw subject

        # Remove unwanted prefix
        if subject and subject.startswith("=?utf-8?q?=5BQlubpenguin=5D_?="):
            subject = subject.replace("=?utf-8?q?=5BQlubpenguin=5D_?=", "").strip()

        # Get sender
        sender = message['from'] or "Unknown"

        # Parse and convert date
        raw_date = message['date']
        try:
            # Parse the raw date to a datetime object
            gmt_date = datetime.strptime(raw_date.split(' (')[0], "%a, %d %b %Y %H:%M:%S %z")
            # Convert to Eastern Time
            est_date = gmt_date.astimezone(eastern)
            date = est_date.strftime("%Y-%m-%d")  # Format as YYYY-MM-DD
            time = est_date.strftime("%H:%M:%S")  # Format as HH:MM:SS
        except Exception as e:
            print(f"Error parsing date: {raw_date} -> {e}")
            date = "Unknown"
            time = "Unknown"

        # Extract body content
        body = None
        try:
            if message.is_multipart():
                body = ''.join(
                    part.get_payload(decode=True).decode(errors='ignore')
                    for part in message.walk()
                    if part.get_content_type() == 'text/plain'
                )
            else:
                body = message.get_payload(decode=True).decode(errors='ignore')
        except Exception as e:
            print(f"Error decoding body: {e}")
            body = "Unable to decode body"

        # Write row
        writer.writerow([subject, sender, date, time, body])
