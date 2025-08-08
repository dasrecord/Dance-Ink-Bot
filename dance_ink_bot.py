#!/usr/bin/env python3

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
import time
import imaplib
import datetime
import email
import re
from email.utils import parsedate_to_datetime
from config import studio_director_url, studio_director_username, studio_director_password, headless, safe_mode, buffer, email_username, email_password

# Add debugging for email credentials
print(f"Email username: {email_username}")
print(f"Email password configured: {'Yes' if email_password else 'No'}")
if email_password:
    print(f"Password length: {len(email_password)} characters")
    print(f"Password ends with: ...{email_password[-4:]}")
print(f"Safe mode: {safe_mode}")

# Initialize WebDriver as a global variable
driver = None
mail = None  # Add global mail variable

def login_to_studio_director():
    global driver
    
    # Configure Chrome options
    chrome_options = Options()
    if headless:
        chrome_options.add_argument("--headless")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    
    # Initialize the WebDriver
    driver = webdriver.Chrome(options=chrome_options)
    
    try:
        print("Logging in to Studio Director...")
        
        # Navigate to the login page
        driver.get(studio_director_url)
        print(f"Navigated to: {driver.current_url}")
        
        # Find username field and enter username
        print("Looking for username field...")
        username_field = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.NAME, "username"))
        )
        username_field.send_keys(studio_director_username)
        print("Username entered")
        
        # Find password field and enter password
        print("Looking for password field...")
        password_field = driver.find_element(By.NAME, "password")
        password_field.send_keys(studio_director_password)
        print("Password entered")
        
        # Submit the login form using the correct button selector
        print("Attempting to submit login...")
        login_button = driver.find_element(By.XPATH, "/html/body/div/div/div[1]/div/form/div[5]/a")
        print(f"Found login button with selector: /html/body/div/div/div[1]/div/form/div[5]/a")
        
        # Click the login button
        driver.execute_script("arguments[0].click();", login_button)
        print("Clicked login button")
        
        print("Login submission attempted, waiting for response...")
        time.sleep(buffer)  # Wait for page to load
        
        # Check if login was successful
        print(f"After login attempt - Current URL: {driver.current_url}")
        print(f"After login attempt - Page title: {driver.title}")
        
        # Look for elements that indicate successful login
        try:
            # Check if we're on the admin page (successful login)
            if "admin.sd" in driver.current_url:
                print("✅ Login successful - Reached admin page")
                return True
            else:
                # Check for search bar or other elements that appear after login
                search_element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.ID, "search"))
                )
                print("✅ Login successful - Found search bar")
                return True
            
        except Exception as e:
            # If we're on admin page but search element not found, still consider it successful
            if "admin.sd" in driver.current_url:
                print("✅ Login successful - On admin page (search element may have different ID)")
                return True
            else:
                print(f"❌ Login may have failed - Could not find expected elements: {e}")
                print(f"Current page source contains: {driver.page_source[:500]}...")
                return False
            
    except Exception as e:
        print(f"❌ Login failed with error: {e}")
        return False

def fetch_emails():
    global mail  # Make mail global so we can use it later
    try:
        # Connect to the email server
        mail = imaplib.IMAP4_SSL('imap.gmail.com')
        mail.login(email_username, email_password)
        mail.select('inbox')

        # Get emails from the last n days
        n = 1
        start_date = (datetime.date.today() - datetime.timedelta(days=n)).strftime("%d-%b-%Y")
        result, data = mail.search(None, f'(SENTSINCE {start_date})')
        
        email_ids = data[0].split() if data[0] else []
        print(f"Fetched {len(email_ids)} emails from today.")

        emails = []
        for email_id in email_ids:
            result, msg_data = mail.fetch(email_id, '(RFC822)')
            msg = email.message_from_bytes(msg_data[0][1])
            
            subject = msg["Subject"] or ""
            print(f"Found email with subject: '{subject}'")
            
            # Only process e-transfer emails
            if msg["Subject"] and "e-Transfer" in msg["Subject"]:
                print(f"✅ This is an e-transfer email: {subject}")
                emails.append((msg, email_id))  # Store both message and ID
            else:
                print(f"❌ Skipping non-e-transfer email: {subject}")

        # Don't logout here - we need the connection for later
        return emails
        
    except Exception as e:
        print(f"Error fetching emails: {e}")
        return []

def mark_email_processed(email_id, reference_number):
    """Mark email as read and apply '2025 Payments EFT's' label"""
    global mail
    try:
        # Mark email as read
        mail.store(email_id, '+FLAGS', '\\Seen')
        print(f"✅ Marked email {email_id} as read")
        
        # Create/apply the Gmail label
        label_name = "2025 Payments EFT's"
        
        # First, try to create the label (in case it doesn't exist)
        try:
            mail.create(label_name)
            print(f"Created Gmail label: {label_name}")
        except:
            # Label already exists, which is fine
            pass
        
        # Apply the label to the email
        try:
            mail.store(email_id, '+X-GM-LABELS', label_name)
            print(f"✅ Applied label '{label_name}' to email {email_id}")
        except Exception as label_error:
            print(f"⚠️ Could not apply label: {label_error}")
            # Try alternative Gmail labeling method
            try:
                mail.copy(email_id, label_name)
                print(f"✅ Applied label '{label_name}' using copy method")
            except Exception as copy_error:
                print(f"❌ Failed to apply label with any method: {copy_error}")
        
        print(f"Email processing completed for reference {reference_number}")
        
    except Exception as e:
        print(f"❌ Error marking email as processed: {e}")

def cleanup_email_connection():
    """Close the email connection"""
    global mail
    try:
        if mail:
            mail.logout()
            print("Email connection closed")
    except:
        pass

def process_emails():
    global driver
    
    emails = fetch_emails()
    print(f"Found {len(emails)} e-transfer emails to process")
    
    if len(emails) == 0:
        print("No e-transfer emails found to process")
        return
    
    # Keep track of processed reference numbers to avoid duplicates
    processed_references = set()
    
    for msg, email_id in emails:  # Unpack message and email ID
        try:
            # Extract payment details
            payment_date = parsedate_to_datetime(msg["Date"])
            print(f"Payment date: {msg['Date']}")
            
            # Parse month, day, year for form fields
            month_name = payment_date.strftime("%b")  # 3-letter month abbreviation
            day = payment_date.day
            year = payment_date.year
            print(f"Parsed month: {month_name}, day: {day}, year: {year}")
            
            # Extract sender information
            reply_to = msg.get("Reply-To", "")
            print(f"Original reply-to: {reply_to}")
            
            # Clean the email address (remove name part)
            if "<" in reply_to and ">" in reply_to:
                replyto_address = reply_to.split("<")[1].split(">")[0]
            else:
                replyto_address = reply_to
                
            print(f"Clean email for search: {replyto_address}")
            
            # Handle message body extraction for multipart messages
            if msg.is_multipart():
                message_body = ""
                for part in msg.walk():
                    if part.get_content_type() == "text/plain":
                        message_body = part.get_payload(decode=True).decode('utf-8')
                        break
            else:
                message_body = msg.get_payload(decode=True).decode('utf-8')

            print("=== EMAIL BODY DEBUG ===")
            print(message_body)
            print("========================")

            # Extract reference number - Updated pattern to handle alphanumeric references
            reference_match = re.search(r'Reference Number: ([A-Za-z0-9]+)', message_body)
            if reference_match:
                reference_number = reference_match.group(1)
                print(f"Found reference number using pattern 'Reference Number: ([A-Za-z0-9]+)': {reference_number}")
                
                # Check if we've already processed this reference number
                if reference_number in processed_references:
                    print(f"⚠️ Reference number {reference_number} already processed, skipping duplicate")
                    continue
                
                # Add to processed set
                processed_references.add(reference_number)
                print(f"✅ Added {reference_number} to processed references")
            else:
                print("No reference number found in email")
                continue

            # Extract amount
            amount_match = re.search(r'\$([0-9,]+\.?[0-9]*)', message_body)
            if amount_match:
                amount = amount_match.group(1)
            else:
                print("No amount found in email")
                continue

            # Extract sender name from the message body
            sender_match = re.search(r'Sent From: (.+)', message_body)
            if sender_match:
                sender_name = sender_match.group(1).strip()
            else:
                sender_name = "Unknown"

            print(f"Processing e-transfer: ${amount} from {sender_name} <{replyto_address}>")

            # Navigate to main page and search for the sender
            driver.get("https://app.thestudiodirector.com/danceink/admin.sd")
            time.sleep(buffer)

            # Search for the sender
            search_field = None
            # Try different selectors for the search field
            search_selectors = [
                (By.ID, "search"),
                (By.NAME, "search"),
                (By.XPATH, "//input[@type='text' and contains(@placeholder, 'search')]"),
                (By.XPATH, "//input[@type='text']"),
                (By.CSS_SELECTOR, "input[type='search']"),
                (By.CSS_SELECTOR, "input.search")
            ]
            
            for selector_type, selector_value in search_selectors:
                try:
                    search_field = driver.find_element(selector_type, selector_value)
                    print(f"Found search field using: {selector_type}='{selector_value}'")
                    break
                except:
                    continue
                    
            if not search_field:
                print("Could not find search field, skipping this email")
                continue
                
            search_field.clear()
            search_field.send_keys(replyto_address)
            
            # Click search button
            search_button = None
            search_button_selectors = [
                (By.XPATH, "//input[@value='Search']"),
                (By.XPATH, "//button[contains(text(), 'Search')]"),
                (By.XPATH, "//input[@type='submit']"),
                (By.XPATH, "//button[@type='submit']"),
                (By.CSS_SELECTOR, "input[type='submit']"),
                (By.CSS_SELECTOR, "button[type='submit']")
            ]
            
            for selector_type, selector_value in search_button_selectors:
                try:
                    search_button = driver.find_element(selector_type, selector_value)
                    print(f"Found search button using: {selector_type}='{selector_value}'")
                    break
                except:
                    continue
                    
            if search_button:
                search_button.click()
                print(f"Successfully searched for: {replyto_address}")
            else:
                print("Could not find search button, trying Enter key...")
                search_field.send_keys("\n")  # Try pressing Enter
                print(f"Tried Enter key for search: {replyto_address}")
            
            time.sleep(buffer)

            # Click on the first search result in searchResultItem div
            try:
                search_result_div = driver.find_element(By.CLASS_NAME, "searchResultItem")
                first_result_link = search_result_div.find_element(By.TAG_NAME, "a")
                first_result_link.click()
                print("Clicked first search result from searchResultItem div")
                time.sleep(buffer)
            except Exception as search_result_error:
                print(f"Could not find search result in searchResultItem div: {search_result_error}")
                # Fallback to original method
                try:
                    first_result = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH, "//table[@id='accountsTable']//tr[2]//a"))
                    )
                    first_result.click()
                    print("Clicked first search result using fallback method")
                    time.sleep(buffer)
                except:
                    print("Could not find search result with any method, skipping this email")
                    continue

            # Click the Ledger tab to view account details
            try:
                ledger_tab = driver.find_element(By.ID, "tab-ledger")
                ledger_tab.click()
                print("Clicked Ledger tab")
                time.sleep(buffer)
                
                # Read the ledger table to determine what category this payment should be
                payment_category = "Tuition"  # Default
                try:
                    ledger_table = driver.find_element(By.ID, 'Ledger')
                    ledger_tbody = ledger_table.find_element(By.TAG_NAME, 'tbody')
                    ledger_rows = ledger_tbody.find_elements(By.TAG_NAME, 'tr')
                    
                    if len(ledger_rows) > 0:  # Make sure there's at least one data row
                        # Get the last row (most recent entry) and second td element (description)
                        last_row = ledger_rows[-1]
                        cells = last_row.find_elements(By.TAG_NAME, 'td')
                        
                        if len(cells) >= 2:
                            payment_category_text = cells[1].text.strip()
                            print(f"Found existing payment category in ledger: {payment_category_text}")
                            
                            # Determine payment category based on ledger text
                            if "Private Lesson" in payment_category_text:
                                payment_category = "Private Lesson"
                                print("Will categorize this payment as Private Lesson")
                            elif "Auto Tuition" in payment_category_text:
                                payment_category = "Tuition"
                                print("Will categorize this payment as Tuition")
                            else:
                                print("Will categorize this payment as Tuition (default)")
                        else:
                            print("Ledger row doesn't have enough cells, using default Tuition")
                    else:
                        print("No data rows found in ledger table, using default Tuition")
                        
                except Exception as ledger_read_error:
                    print(f"Error reading ledger table: {ledger_read_error}")
                    print("Using default Tuition category")
                    
            except Exception as ledger_error:
                print(f"Could not find Ledger tab: {ledger_error}")
                payment_category = "Tuition"  # Default fallback
                print("Using default Tuition category")

            # Click the Add New Payment button
            try:
                add_payment_button = driver.find_element(By.ID, "addnewpayment")
                add_payment_button.click()
                print("Clicked Add New Payment button")
                time.sleep(buffer)
            except Exception as add_payment_error:
                print(f"Could not find Add New Payment button: {add_payment_error}")
                print("Skipping this email")
                continue

            # After clicking Add New Payment, click the "Cash, check, trade" link
            try:
                cash_check_trade_link = driver.find_element(By.XPATH, "//a[contains(text(), 'Cash, check, trade')]")
                cash_check_trade_link.click()
                print("Clicked 'Cash, check, trade' link")
                time.sleep(buffer)
            except Exception as cash_link_error:
                print(f"Could not find 'Cash, check, trade' link: {cash_link_error}")
                # Try alternative selectors
                try:
                    cash_link_alt = driver.find_element(By.XPATH, "//a[contains(text(), 'Cash')]")
                    cash_link_alt.click()
                    print("Clicked cash link (alternative)")
                    time.sleep(buffer)
                except:
                    print("Could not find any cash/check/trade link, skipping this email")
                    continue

            # Set payment amount (using name attribute as mentioned)
            amount_field = driver.find_element(By.NAME, "amount")
            amount_field.clear()
            amount_field.send_keys(amount)
            print("Successfully set payment amount")

            # Set reference in notes field (since there's no dedicated reference field)
            try:
                notes_field = driver.find_element(By.CSS_SELECTOR, '[name="notes"]')
                notes_field.clear()
                notes_field.send_keys(f"{reference_number}")
                print(f"Successfully set reference in notes field: {reference_number}")
            except Exception as e:
                print(f"Could not find notes field: {e}")
            
            # Set payment date using the correct field names (due_date__mon, etc.) - these are SELECT dropdowns
            try:
                # Set month (using due_date__mon) - this takes 3-letter month abbreviation
                month_select = Select(driver.find_element(By.NAME, "due_date__mon"))
                month_select.select_by_value(month_name)  # Use 3-letter month like "Aug"
                print(f"Successfully set month: {month_name}")
                
                # Set day (likely due_date__day) - this is a select dropdown
                day_select = Select(driver.find_element(By.NAME, "due_date__day"))
                day_select.select_by_value(str(day))
                print(f"Successfully set day: {day}")
                
                # Set year (likely due_date__year) - this is a select dropdown
                year_select = Select(driver.find_element(By.NAME, "due_date__year"))
                year_select.select_by_value(str(year))
                print(f"Successfully set year: {year}")
                
            except Exception as date_error:
                print(f"Error setting payment date: {date_error}")
                # Try alternative date field names using ID selectors
                try:
                    month_select = Select(driver.find_element(By.ID, "due_date__mon"))
                    month_select.select_by_value(month_name)  # Use 3-letter month
                    print(f"Successfully set month using ID: {month_name}")
                    
                    day_select = Select(driver.find_element(By.ID, "due_date__day"))
                    day_select.select_by_value(str(day))
                    print(f"Successfully set day using ID: {day}")
                    
                    year_select = Select(driver.find_element(By.ID, "due_date__year"))
                    year_select.select_by_value(str(year))
                    print(f"Successfully set year using ID: {year}")
                except Exception as id_date_error:
                    print(f"Could not set payment date with any method: {id_date_error}")

            # Try to set payment method if available
            try:
                # Look for method field - it might be a select or input
                method_selectors = [
                    '[name="method"]',
                    '[name="payment_method"]', 
                    '[id="method"]',
                    '[id="payment_method"]',  # Added this selector
                    'select[name*="method"]'
                ]
                
                method_field = None
                for selector in method_selectors:
                    try:
                        method_field = driver.find_element(By.CSS_SELECTOR, selector)
                        print(f"Found method field with selector: {selector}")
                        break
                    except:
                        continue
                
                if method_field:
                    if method_field.tag_name == 'select':
                        method_select = Select(method_field)
                        # Try different method values for e-transfer (EFT first)
                        method_options = ["EFT", "eTransfer", "Electronic", "Bank Transfer"]
                        for method_option in method_options:
                            try:
                                method_select.select_by_visible_text(method_option)
                                print(f"Selected payment method: {method_option}")
                                break
                            except:
                                continue
                    else:
                        method_field.clear()
                        method_field.send_keys("EFT")  # Use EFT instead of eTransfer
                        print("Set payment method to EFT")
                else:
                    print("Could not find payment method field")
                    
            except Exception as method_error:
                print(f"Error setting payment method: {method_error}")
            
            # Use the payment category we determined earlier from the account ledger
            print(f"Using predetermined payment category: {payment_category}")
            
            if payment_category == "Private Lesson":
                split_category = "Private Lesson"
                print("Will split payment as Private Lesson")
            else:
                split_category = "Tuition"
                print("Will split payment as Tuition")
            
            # BEFORE saving: Click Split Payment button to make paid_toward1 visible
            try:
                split_payment_button = driver.find_element(By.ID, 'splitpayment')
                split_payment_button.click()
                print("Clicked Split Payment button")
                time.sleep(buffer)
                
                # Set the split payment category in the SELECT dropdown
                try:
                    paid_toward_select = Select(driver.find_element(By.NAME, 'paid_toward1'))
                    
                    # Get available options for selection logic
                    all_options = paid_toward_select.options
                    available_options = [(opt.get_attribute('value'), opt.text) for opt in all_options]
                    
                    # Try different selection methods based on the split_category
                    selection_successful = False
                    
                    if split_category == "Private Lesson":
                        # Try to find Private Lesson option
                        for value, text in available_options:
                            if "Private" in text or "private" in text or "lesson" in text.lower():
                                try:
                                    paid_toward_select.select_by_value(value)
                                    print(f"Selected Private Lesson option: {value} ({text})")
                                    selection_successful = True
                                    break
                                except:
                                    continue
                    
                    if not selection_successful and split_category == "Tuition":
                        # Try to find Tuition option
                        for value, text in available_options:
                            if "Tuition" in text or "tuition" in text:
                                try:
                                    paid_toward_select.select_by_value(value)
                                    print(f"Selected Tuition option: {value} ({text})")
                                    selection_successful = True
                                    break
                                except:
                                    continue
                    
                    # If still not successful, try exact match
                    if not selection_successful:
                        try:
                            paid_toward_select.select_by_visible_text(split_category)
                            print(f"Selected by visible text: {split_category}")
                            selection_successful = True
                        except:
                            try:
                                paid_toward_select.select_by_value(split_category)
                                print(f"Selected by value: {split_category}")
                                selection_successful = True
                            except:
                                pass
                    
                    if not selection_successful:
                        print(f"Could not select any option for category: {split_category}")
                        
                except Exception as split_error:
                    print(f"Error setting split payment category: {split_error}")
                    
            except Exception as split_button_error:
                print(f"Error clicking Split Payment button: {split_button_error}")
            
            # NOW save the payment (with split category already set)
            print("Looking for save/submit button...")
            try:
                save_button = driver.find_element(By.ID, "savepayment")
                print("Found save button with ID: savepayment")
                if not safe_mode:
                    save_button.click()
                    print("Successfully clicked save button")
                    time.sleep(buffer)  # Wait for save to complete
                else:
                    print("SAFE MODE: Skipping save button click")
            except Exception as save_error:
                print(f"Could not find save button with ID 'savepayment': {save_error}")
                # Fallback to other selectors
                save_selectors = [
                    'input[type="submit"]',
                    'button[type="submit"]', 
                    'input[value*="Save"]',
                    'button[value*="Save"]',
                    'input[value*="Add"]',
                    'button[value*="Add"]',
                    '.save-btn',
                    '#save-payment',
                    '[name="save"]',
                    '[name="submit"]'
                ]
                
                save_button = None
                for selector in save_selectors:
                    try:
                        save_button = driver.find_element(By.CSS_SELECTOR, selector)
                        print(f"Found save button with selector: {selector}")
                        break
                    except:
                        continue
                
                if save_button:
                    try:
                        if not safe_mode:
                            save_button.click()
                            print("Successfully clicked save button (fallback)")
                            time.sleep(buffer)  # Wait for save to complete
                        else:
                            print("SAFE MODE: Skipping save button click (fallback)")
                    except Exception as e:
                        print(f"Error clicking save button: {e}")
                else:
                    print("Could not find save button - payment form filled but not submitted")
            
            print("Payment processing completed for this e-transfer")

            # After saving payment, look for "Review the account ledger" link and click it
            try:
                # Wait a moment for the page to update after saving
                time.sleep(buffer)
                
                # Look for the "Review the account ledger" link in the specific location:
                # <a> tag inside the last <p> tag inside the div with class="contentInfo"
                review_ledger_link = None
                
                try:
                    # Find the contentInfo div, then get the last p tag, then find the a tag inside it
                    content_info_div = driver.find_element(By.CLASS_NAME, "contentInfo")
                    p_tags = content_info_div.find_elements(By.TAG_NAME, "p")
                    
                    if p_tags:
                        last_p_tag = p_tags[-1]  # Get the last p tag
                        review_ledger_link = last_p_tag.find_element(By.TAG_NAME, "a")
                        print("Found 'Review the account ledger' link in last <p> tag of contentInfo div")
                    else:
                        print("No <p> tags found in contentInfo div")
                        
                except Exception as specific_error:
                    print(f"Could not find link in contentInfo div: {specific_error}")
                    
                    # Fallback to original selectors if the specific location fails
                    review_selectors = [
                        (By.XPATH, "//div[@class='contentInfo']//p[last()]//a"),
                        (By.XPATH, "//a[contains(text(), 'Review the account ledger')]"),
                        (By.XPATH, "//a[contains(text(), 'Review')]"),
                        (By.XPATH, "//a[contains(text(), 'ledger')]"),
                        (By.XPATH, "//a[contains(text(), 'account ledger')]")
                    ]
                    
                    for selector_type, selector_value in review_selectors:
                        try:
                            review_ledger_link = driver.find_element(selector_type, selector_value)
                            print(f"Found 'Review the account ledger' link using fallback: {selector_type}='{selector_value}'")
                            break
                        except:
                            continue
                
                if review_ledger_link:
                    review_ledger_link.click()
                    print("Clicked 'Review the account ledger' link")
                    time.sleep(buffer)  # Wait for the ledger page to load
                # Wait a moment for the page to update after saving
                time.sleep(buffer)
                
                # Look for the "Review the account ledger" link
                review_ledger_link = None
                review_selectors = [
                    (By.XPATH, "//a[contains(text(), 'Review the account ledger')]"),
                    (By.XPATH, "//a[contains(text(), 'Review')]"),
                    (By.XPATH, "//a[contains(text(), 'ledger')]"),
                    (By.XPATH, "//a[contains(text(), 'account ledger')]")
                ]
                
                for selector_type, selector_value in review_selectors:
                    try:
                        review_ledger_link = driver.find_element(selector_type, selector_value)
                        print(f"Found 'Review the account ledger' link using: {selector_type}='{selector_value}'")
                        break
                    except:
                        continue
                
                if review_ledger_link:
                    review_ledger_link.click()
                    print("Clicked 'Review the account ledger' link")
                    time.sleep(buffer)  # Wait for the ledger page to load
                    
                    # Now check the current balance on this page
                    try:
                        # Try multiple selectors for the current balance
                        current_balance = None
                        balance_selectors = [
                            (By.ID, "current-balance"),
                            (By.CSS_SELECTOR, "#current-balance"),
                            (By.XPATH, "//*[@id='current-balance']"),
                            (By.XPATH, "//span[@id='current-balance']"),
                            (By.XPATH, "//div[@id='current-balance']"),
                            (By.XPATH, "//*[contains(@id, 'balance')]"),
                            (By.XPATH, "//*[contains(text(), '$')]")
                        ]
                        
                        for selector_type, selector_value in balance_selectors:
                            try:
                                current_balance_element = WebDriverWait(driver, 3).until(
                                    EC.presence_of_element_located((selector_type, selector_value))
                                )
                                current_balance = current_balance_element.text.strip()
                                print(f"Found balance using {selector_type}='{selector_value}': {current_balance}")
                                if current_balance:  # If we got some text, break
                                    break
                            except Exception as selector_error:
                                print(f"Selector {selector_type}='{selector_value}' failed: {selector_error}")
                                continue
                        
                        if current_balance:
                            print(f"Current balance: {current_balance}")
                            
                            # Check if balance is zero (could be $0.00, 0.00, or similar)
                            if "0.00" in current_balance or current_balance == "0" or current_balance == "$0":
                                print("✅ Balance is correctly zeroed out - payment successful")
                                
                                # Mark email as processed since payment was successful
                                mark_email_processed(email_id, reference_number)
                                
                            else:
                                print(f"⚠️ Balance is NOT zero: {current_balance} - moving to next e-transfer")
                                continue  # Skip to next email
                        else:
                            print("Could not read balance from any selector - moving to next e-transfer")
                            continue  # Skip to next email
                            
                    except Exception as balance_error:
                        print(f"Error during balance check: {balance_error}")
                        print("Moving to next e-transfer")
                        continue  # Skip to next email
                        
                else:
                    print("Could not find 'Review the account ledger' link - moving to next e-transfer")
                    continue  # Skip to next email
                    
            except Exception as ledger_link_error:
                print(f"Error finding 'Review the account ledger' link: {ledger_link_error}")
                print("Moving to next e-transfer")
                continue  # Skip to next email

            print(f"Payment processing completed for e-transfer from {sender_name}")
                
        except Exception as e:
            print(f"Error processing e-transfer email: {e}")
            continue


if __name__ == "__main__":
    try:
        print("=== Dance Ink Bot Starting ===")
        
        # Step 1: Login to Studio Director
        login_to_studio_director()
        
        # Step 2: Process emails
        print("\n=== Processing Emails ===")
        process_emails()
        
        print("=== Dance Ink Bot Finished Successfully ===")
        
    except Exception as e:
        print(f"Fatal error in main execution: {e}")
        print("=== Dance Ink Bot Finished with Errors ===")
    finally:
        # Close the browser
        try:
            driver.quit()
            print("Browser closed.")
        except:
            pass
        
        # Close email connection
        cleanup_email_connection()
