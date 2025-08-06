#!/usr/bin/env python3

import imaplib
import datetime
import email
from config import email_username, email_password

def test_email_connection():
    print(f"Testing email connection...")
    print(f"Username: {email_username}")
    print(f"Password configured: {'Yes' if email_password else 'No'}")
    
    try:
        # Connect to the email server
        mail = imaplib.IMAP4_SSL('imap.gmail.com')
        print("Connected to Gmail IMAP server")
        
        # Login
        mail.login(email_username, email_password)
        print("Successfully logged in!")
        
        # Select inbox
        mail.select('inbox')
        print("Selected inbox")
        
        # Search for today's emails
        today = datetime.date.today().strftime("%d-%b-%Y")
        print(f"Searching for emails from: {today}")
        
        result, data = mail.search(None, f'(SENTSINCE {today})')
        
        if data[0] is None:
            print("No emails found for today.")
        else:
            email_ids = data[0].split()
            print(f"Found {len(email_ids)} emails from today")
            
            # Fetch and display email content
            for i, email_id in enumerate(email_ids[:3]):  # Limit to first 3 emails
                print(f"\n--- Email {i+1} ---")
                try:
                    result, msg_data = mail.fetch(email_id, '(RFC822)')
                    msg = email.message_from_bytes(msg_data[0][1])
                    
                    print(f"Subject: {msg['Subject']}")
                    print(f"From: {msg['From']}")
                    print(f"Date: {msg['Date']}")
                    print(f"Reply-To: {msg.get('Reply-To', 'Not set')}")
                    
                    # Extract message body
                    if msg.is_multipart():
                        print("Body:")
                        for part in msg.walk():
                            if part.get_content_type() == "text/plain":
                                body = part.get_payload(decode=True).decode('utf-8')
                                print(body[:500] + "..." if len(body) > 500 else body)
                                break
                    else:
                        body = msg.get_payload(decode=True).decode('utf-8')
                        print("Body:")
                        print(body[:500] + "..." if len(body) > 500 else body)
                    
                    print("-" * 50)
                except Exception as e:
                    print(f"Error processing email {i+1}: {e}")
        
        mail.logout()
        print("Connection test successful!")
        return True
        
    except imaplib.IMAP4.error as e:
        print(f"IMAP error: {e}")
        return False
    except Exception as e:
        print(f"Error: {e}")
        return False

if __name__ == "__main__":
    test_email_connection()
