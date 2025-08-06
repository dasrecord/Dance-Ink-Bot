#!/usr/bin/env python3

import imaplib
import datetime
import email
import re
from config import email_username, email_password

def debug_emails():
    try:
        # Connect to the email server
        mail = imaplib.IMAP4_SSL('imap.gmail.com')
        mail.login(email_username, email_password)
        mail.select('inbox')

        # Search for today's emails
        today = datetime.date.today().strftime("%d-%b-%Y")
        result, data = mail.search(None, f'(SENTSINCE {today})')
        
        if data[0] is None:
            print("No emails found for today.")
            mail.logout()
            return
            
        email_ids = data[0].split()
        print(f"Found {len(email_ids)} emails from today.")

        for i, email_id in enumerate(email_ids):
            result, msg_data = mail.fetch(email_id, '(RFC822)')
            msg = email.message_from_bytes(msg_data[0][1])
            
            print(f"\n=== EMAIL {i+1} ===")
            print(f"Subject: {msg['Subject']}")
            print(f"From: {msg['From']}")
            print(f"Reply-To: {msg.get('Reply-To', 'Not set')}")
            
            # Check if it's an e-transfer
            if msg["Subject"] and "e-Transfer" in msg["Subject"]:
                print("*** THIS IS AN E-TRANSFER EMAIL ***")
                
                # Handle message body extraction for multipart messages
                if msg.is_multipart():
                    message_body = ""
                    for part in msg.walk():
                        if part.get_content_type() == "text/plain":
                            message_body = part.get_payload(decode=True).decode('utf-8')
                            break
                else:
                    message_body = msg.get_payload(decode=True).decode('utf-8')

                print("=== EMAIL BODY ===")
                print(message_body)
                print("==================")
                
                # Try to find reference number
                patterns = [
                    r"Reference Number: ([A-Za-z0-9]+)",
                    r"Reference #: ([A-Za-z0-9]+)",
                    r"Reference: ([A-Za-z0-9]+)",
                    r"Ref #: ([A-Za-z0-9]+)",
                    r"Ref: ([A-Za-z0-9]+)",
                    r"Reference Number:\s*([A-Za-z0-9]+)",
                    r"Reference #:\s*([A-Za-z0-9]+)",
                    r"Transaction ID: ([A-Za-z0-9]+)",
                    r"Transaction #: ([A-Za-z0-9]+)",
                    r"Confirmation #: ([A-Za-z0-9]+)",
                    r"Confirmation Number: ([A-Za-z0-9]+)",
                ]
                
                reference_found = False
                for pattern in patterns:
                    reference_match = re.search(pattern, message_body, re.IGNORECASE)
                    if reference_match:
                        reference_number = reference_match.group(1)
                        print(f"*** FOUND REFERENCE: {reference_number} using pattern '{pattern}' ***")
                        reference_found = True
                        break
                
                if not reference_found:
                    print("*** NO REFERENCE NUMBER FOUND ***")
                    print("Looking for any numbers in the text...")
                    # Find all numbers in the text
                    numbers = re.findall(r'\d+', message_body)
                    print(f"All numbers found: {numbers}")
            
            print("-" * 50)

        mail.logout()
        
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    debug_emails()
