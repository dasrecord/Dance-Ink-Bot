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

        # Get emails from the last n days - ONLY UNREAD emails
        n = 11  # Changed from 1 to 30 days to capture all October e-transfers
        start_date = (datetime.date.today() - datetime.timedelta(days=n)).strftime("%d-%b-%Y")
        
        # Search for UNREAD emails only from the specified date range
        search_criteria = f'(SENTSINCE {start_date} UNSEEN)'
        result, data = mail.search(None, search_criteria)
        
        email_ids = data[0].split() if data[0] else []
        print(f"Fetched {len(email_ids)} UNREAD emails from the last {n} days.")

        emails = []
        for email_id in email_ids:
            # Use BODY.PEEK instead of RFC822 to avoid marking email as read during fetch
            result, msg_data = mail.fetch(email_id, '(BODY.PEEK[])')
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

def parse_unpaid_charges(driver):
    """Parse unpaid charges from the Current Unpaid Charges section"""
    unpaid_charges = {}
    try:
        # Look for table with class "ReportTable" that contains unpaid charges
        print("Looking for ReportTable with unpaid charges...")
        
        # First, let's see what tables are available on the page
        all_tables = driver.find_elements(By.TAG_NAME, "table")
        print(f"Found {len(all_tables)} total tables on the page")
        
        for i, table in enumerate(all_tables):
            table_class = table.get_attribute("class") or ""
            print(f"Table {i}: class='{table_class}'")
        
        try:
            # Use the exact XPath structure you provided
            try:
                # Method 1: Use the exact XPath
                top_td = driver.find_element(By.XPATH, "//*[@id='mainContent']/div/div/div[2]/table/tbody/tr/td[2]")
                print("✅ Found exact td element using XPath")
                
                # Find the ReportTable within this td
                report_table = top_td.find_element(By.CLASS_NAME, "ReportTable")
                print("✅ Found ReportTable within the td element")
                
            except Exception as xpath_error:
                print(f"Exact XPath failed: {xpath_error}")
                
                # Method 2: Look for td with both class="Top" and the specific style
                try:
                    top_td = driver.find_element(By.XPATH, "//td[@class='Top' and contains(@style, 'padding-left:30px')]")
                    print("✅ Found td with class='Top' and padding-left:30px")
                    
                    report_table = top_td.find_element(By.CLASS_NAME, "ReportTable")
                    print("✅ Found ReportTable within the Top td")
                    
                except Exception as style_error:
                    print(f"Style-based search failed: {style_error}")
                    
                    # Method 3: Look for MultiTable div, then find ReportTable in same parent
                    try:
                        multitable_div = driver.find_element(By.XPATH, "//div[@class='MultiTable'][contains(text(), 'Current Unpaid Charges')]")
                        print("✅ Found MultiTable div with 'Current Unpaid Charges'")
                        
                        # Get the parent td, then find ReportTable within it
                        parent_td = multitable_div.find_element(By.XPATH, "./ancestor::td")
                        report_table = parent_td.find_element(By.CLASS_NAME, "ReportTable")
                        print("✅ Found ReportTable via MultiTable div parent")
                        
                    except Exception as multitable_error:
                        print(f"MultiTable method failed: {multitable_error}")
                        
                        # Method 4: Search all ReportTables for the one with actual charge data
                        try:
                            all_report_tables = driver.find_elements(By.CLASS_NAME, "ReportTable")
                            print(f"Method 4: Checking all {len(all_report_tables)} ReportTables for charge data...")
                            
                            report_table = None
                            for i, table in enumerate(all_report_tables):
                                table_html = table.get_attribute('outerHTML')
                                print(f"ReportTable #{i} HTML preview: {table_html[:200]}...")
                                
                                if 'Tuition' in table_html and 'Costume Deposit' in table_html:
                                    print(f"✅ Found ReportTable #{i} with charge data")
                                    report_table = table
                                    break
                                else:
                                    print(f"ReportTable #{i} doesn't contain charge data")
                            
                            if not report_table:
                                raise Exception("No ReportTable contains charge data")
                                
                        except Exception as method4_error:
                            print(f"Method 4 failed: {method4_error}")
                            raise Exception("All methods failed to find ReportTable")
            
            # If we found a report_table, parse it
            rows = report_table.find_elements(By.TAG_NAME, "tr")
            print(f"Found ReportTable with {len(rows)} rows")
            
            for i, row in enumerate(rows):
                cells = row.find_elements(By.TAG_NAME, "td")
                th_cells = row.find_elements(By.TAG_NAME, "th")
                
                if len(cells) >= 2:
                    category = cells[0].text.strip()
                    amount_text = cells[1].text.strip()
                    
                    print(f"Row {i}: Category='{category}', Amount='{amount_text}'")
                    
                    # Skip header rows and summary rows
                    skip_categories = [
                        "Category", 
                        "Total unpaid charges", 
                        "Current payments not applied to unpaid charges or current charges paid by future payments",
                        "Current Balance Due"
                    ]
                    
                    if category and category not in skip_categories:
                        # Extract numeric amount (remove any currency symbols)
                        amount_match = re.search(r'([\d,]+\.?\d*)', amount_text)
                        if amount_match:
                            amount = float(amount_match.group(1).replace(',', ''))
                            if amount > 0:
                                unpaid_charges[category] = amount
                                print(f"✅ Found unpaid charge: {category} = ${amount}")
                            else:
                                print(f"⚠️ Skipping zero amount: {category} = ${amount}")
                        else:
                            print(f"⚠️ Could not parse amount from: '{amount_text}'")
                    else:
                        print(f"⚠️ Skipping category: '{category}'")
                elif len(th_cells) >= 2:
                    # Header row
                    header1 = th_cells[0].text.strip()
                    header2 = th_cells[1].text.strip()
                    print(f"Header row {i}: '{header1}' | '{header2}'")
                else:
                    print(f"Row {i}: Insufficient cells (td: {len(cells)}, th: {len(th_cells)})")
                        
        except Exception as table_error:
            print(f"Error finding ReportTable with all methods: {table_error}")
            
            # Fallback: Try all ReportTables until we find one with data
            print("Trying fallback method - checking all ReportTables...")
            try:
                report_tables = driver.find_elements(By.CLASS_NAME, "ReportTable")
                print(f"Found {len(report_tables)} ReportTable elements total")
                
                for table_index, report_table in enumerate(report_tables):
                    print(f"Checking ReportTable #{table_index}...")
                    rows = report_table.find_elements(By.TAG_NAME, "tr")
                    
                    if len(rows) > 1:  # Must have more than just header
                        print(f"ReportTable #{table_index} has {len(rows)} rows - checking content...")
                        
                        # Check if this table contains charge data
                        table_text = report_table.text
                        if "Tuition" in table_text or "Costume Deposit" in table_text:
                            print(f"✅ ReportTable #{table_index} contains charge data!")
                            
                            for i, row in enumerate(rows):
                                cells = row.find_elements(By.TAG_NAME, "td")
                                if len(cells) >= 2:
                                    category = cells[0].text.strip()
                                    amount_text = cells[1].text.strip()
                                    
                                    print(f"Table #{table_index} Row {i}: Category='{category}', Amount='{amount_text}'")
                                    
                                    skip_categories = [
                                        "Category", 
                                        "Total unpaid charges", 
                                        "Current payments not applied to unpaid charges or current charges paid by future payments",
                                        "Current Balance Due"
                                    ]
                                    
                                    if category and category not in skip_categories:
                                        amount_match = re.search(r'([\d,]+\.?\d*)', amount_text)
                                        if amount_match:
                                            amount = float(amount_match.group(1).replace(',', ''))
                                            if amount > 0:
                                                unpaid_charges[category] = amount
                                                print(f"✅ Found unpaid charge: {category} = ${amount}")
                            break  # Found the right table, stop looking
                        else:
                            print(f"ReportTable #{table_index} doesn't contain charge data")
                    else:
                        print(f"ReportTable #{table_index} has only {len(rows)} rows - skipping")
                            
            except Exception as fallback_error:
                print(f"Fallback method failed: {fallback_error}")
            
            # Fallback: Look for any table containing unpaid charges patterns
            print("Trying fallback method with page source...")
            try:
                page_text = driver.page_source
                
                # Debug: Show a portion of page source
                if "Current Unpaid Charges" in page_text:
                    print("✅ Found 'Current Unpaid Charges' text in page source")
                else:
                    print("❌ 'Current Unpaid Charges' text not found in page source")
                
                if "Tuition" in page_text and "Costume Deposit" in page_text:
                    print("✅ Found charge category keywords in page source")
                else:
                    print("❌ Charge category keywords not found in page source")
                
                # Look for the specific patterns in the HTML structure
                # <td>Tuition</td><td class="AR">65.00</td>
                charge_patterns = [
                    (r'<td>Tuition</td><td[^>]*>([\d,]+\.?\d*)</td>', 'Tuition'),
                    (r'<td>Costume Deposit</td><td[^>]*>([\d,]+\.?\d*)</td>', 'Costume Deposit'),
                    (r'<td>Private Lesson</td><td[^>]*>([\d,]+\.?\d*)</td>', 'Private Lesson'),
                    (r'<td>Registration</td><td[^>]*>([\d,]+\.?\d*)</td>', 'Registration'),
                ]
                
                for pattern, category in charge_patterns:
                    matches = re.findall(pattern, page_text)
                    if matches:
                        amount = float(matches[0].replace(',', ''))
                        if amount > 0:
                            unpaid_charges[category] = amount
                            print(f"✅ Found unpaid charge (fallback): {category} = ${amount}")
                            
            except Exception as fallback_error:
                print(f"Fallback method failed: {fallback_error}")
        
        print(f"Total unpaid charges parsed: {unpaid_charges}")
        return unpaid_charges
        
    except Exception as e:
        print(f"Error parsing unpaid charges: {e}")
        return {}

def calculate_payment_allocation(payment_amount, unpaid_charges):
    """Calculate how to allocate payment across unpaid charges"""
    payment_amount = float(str(payment_amount).replace(',', ''))
    
    # Priority order for payment allocation
    priority_order = ["Tuition", "Costume Deposit", "Private Lesson", "Registration"]
    
    allocations = []
    remaining_payment = payment_amount
    total_unpaid = sum(unpaid_charges.values())
    
    print(f"Payment amount: ${payment_amount}")
    print(f"Total unpaid: ${total_unpaid}")
    print(f"Unpaid charges: {unpaid_charges}")
    
    # Scenario 1: Exact match - payment equals total unpaid
    if abs(payment_amount - total_unpaid) < 0.01:  # Allow for small rounding differences
        print("✅ Exact match - splitting payment across all unpaid charges")
        for category, amount in unpaid_charges.items():
            allocations.append((category, amount))
        return allocations
    
    # Scenario 2: Payment less than total unpaid - allocate by priority
    if payment_amount < total_unpaid:
        print("⚠️ Partial payment - allocating by priority order")
        for category in priority_order:
            if category in unpaid_charges and remaining_payment > 0:
                charge_amount = unpaid_charges[category]
                if remaining_payment >= charge_amount:
                    # Pay full amount for this category
                    allocations.append((category, charge_amount))
                    remaining_payment -= charge_amount
                    print(f"Allocated ${charge_amount} to {category} (full)")
                else:
                    # Pay partial amount for this category
                    allocations.append((category, remaining_payment))
                    print(f"Allocated ${remaining_payment} to {category} (partial)")
                    remaining_payment = 0
                    break
        return allocations
    
    # Scenario 3: Payment more than total unpaid - apply to highest priority category
    if payment_amount > total_unpaid:
        print("💰 Overpayment - applying full amount to highest priority category")
        # Find the highest priority category that exists
        for category in priority_order:
            if category in unpaid_charges:
                allocations.append((category, payment_amount))
                print(f"Allocated full payment ${payment_amount} to {category} (overpayment)")
                return allocations
        
        # Fallback to first available category
        if unpaid_charges:
            first_category = list(unpaid_charges.keys())[0]
            allocations.append((first_category, payment_amount))
            print(f"Allocated full payment ${payment_amount} to {first_category} (fallback)")
            return allocations
    
    # Fallback: Default to Tuition
    print("🔄 Fallback - defaulting to Tuition")
    allocations.append(("Tuition", payment_amount))
    return allocations

def cleanup_email_connection():
    """Close the email connection"""
    global mail
    try:
        if mail:
            mail.logout()
            print("Email connection closed")
    except:
        pass

def verify_family_email_match(target_email):
    """Verify if current family page has matching email in email or extra_emails fields within Overview tab"""
    try:
        # We should be on the Overview tab by default when clicking a family
        # Look for email fields within the overview tab content
        print("Looking for email fields in Overview tab...")
        
        primary_email = ""
        extra_emails = ""
        
        # Look for the email field inside a td element with id="email"
        try:
            email_field = driver.find_element(By.XPATH, "//td//input[@id='email']")
            primary_email = email_field.get_attribute("value")
            if primary_email is None:
                primary_email = email_field.text.strip()
            if primary_email is None:
                primary_email = ""
            primary_email = primary_email.strip().lower()
            print(f"Found primary email field: '{primary_email}'")
        except Exception as email_error:
            print(f"Could not find primary email field: {email_error}")
            # Try alternative selector
            try:
                email_field = driver.find_element(By.ID, "email")
                primary_email = email_field.get_attribute("value") or email_field.text or ""
                primary_email = primary_email.strip().lower()
                print(f"Found primary email field (alternative): '{primary_email}'")
            except:
                print("Could not find primary email field with any method")
        
        # Check if primary email matches
        if primary_email and primary_email == target_email.lower():
            print(f"✅ Primary email matches target: {target_email}")
            return True
        
        # Look for the extra_emails field inside a td element with id="extra_emails"
        try:
            extra_emails_field = driver.find_element(By.XPATH, "//td//input[@id='extra_emails']")
            extra_emails = extra_emails_field.get_attribute("value")
            if extra_emails is None:
                extra_emails = extra_emails_field.text.strip()
            if extra_emails is None:
                extra_emails = ""
            extra_emails = extra_emails.strip().lower()
            print(f"Found extra emails field: '{extra_emails}'")
        except Exception as extra_error:
            print(f"Could not find extra emails field: {extra_error}")
            # Try alternative selector
            try:
                extra_emails_field = driver.find_element(By.ID, "extra_emails")
                extra_emails = extra_emails_field.get_attribute("value") or extra_emails_field.text or ""
                extra_emails = extra_emails.strip().lower()
                print(f"Found extra emails field (alternative): '{extra_emails}'")
            except:
                print("Could not find extra emails field with any method")
        
        # Check if target email is in the extra emails (could be comma-separated)
        if extra_emails and target_email.lower() in extra_emails:
            print(f"✅ Target email found in extra emails: {target_email}")
            return True
        
        print(f"❌ Email mismatch - target: {target_email}, primary: '{primary_email}', extra: '{extra_emails}'")
        return False
        
    except Exception as e:
        print(f"Error verifying family email: {e}")
        return False

def find_correct_family_result(target_email):
    """Find and click the correct family result by verifying email fields"""
    try:
        # First try searchResultItem divs
        search_result_divs = driver.find_elements(By.CLASS_NAME, "searchResultItem")
        if search_result_divs:
            print(f"Found {len(search_result_divs)} search results to check")
            
            for i, result_div in enumerate(search_result_divs):
                try:
                    result_link = result_div.find_element(By.TAG_NAME, "a")
                    result_text = result_link.text.strip()
                    print(f"Checking result {i+1}: '{result_text}'")
                    
                    # Click the result
                    result_link.click()
                    time.sleep(buffer)
                    
                    # Check if this family has the correct email
                    if verify_family_email_match(target_email):
                        print(f"✅ Found correct family: '{result_text}'")
                        
                        # Since we found the correct family and we're on Overview tab,
                        # we can now directly access the Ledger tab (they're at same level)
                        print("Email verified, ready to proceed to Ledger tab")
                        
                        return True
                    else:
                        print(f"❌ Wrong family: '{result_text}', going back to search results")
                        # Go back to search results
                        driver.back()
                        time.sleep(buffer)
                        
                except Exception as result_error:
                    print(f"Error checking result {i+1}: {result_error}")
                    continue
        
        # Fallback to table method if searchResultItem didn't work
        print("Trying fallback table method...")
        try:
            accounts_table = driver.find_element(By.ID, "accountsTable")
            result_rows = accounts_table.find_elements(By.XPATH, ".//tr[position()>1]")  # Skip header
            
            print(f"Found {len(result_rows)} table results to check")
            
            for i, row in enumerate(result_rows):
                try:
                    result_link = row.find_element(By.TAG_NAME, "a")
                    result_text = result_link.text.strip()
                    print(f"Checking table result {i+1}: '{result_text}'")
                    
                    # Click the result
                    result_link.click()
                    time.sleep(buffer)
                    
                    # Check if this family has the correct email
                    if verify_family_email_match(target_email):
                        print(f"✅ Found correct family in table: '{result_text}'")
                        
                        # Since we found the correct family and we're on Overview tab,
                        # we can now directly access the Ledger tab (they're at same level)
                        print("Email verified, ready to proceed to Ledger tab")
                        
                        return True
                    else:
                        print(f"❌ Wrong family in table: '{result_text}', going back to search results")
                        # Go back to search results
                        driver.back()
                        time.sleep(buffer)
                        
                except Exception as table_result_error:
                    print(f"Error checking table result {i+1}: {table_result_error}")
                    continue
        
        except Exception as table_error:
            print(f"Error with table fallback method: {table_error}")
        
        print("❌ Could not find family with matching email")
        return False
        
    except Exception as e:
        print(f"Error in find_correct_family_result: {e}")
        return False

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
            month_number = payment_date.month  # Numeric month for date input
            day = payment_date.day
            year = payment_date.year
            print(f"Parsed month: {month_name} ({month_number}), day: {day}, year: {year}")
            
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

            # Extract message field from the e-transfer email
            message_match = re.search(r'Message: (.+)', message_body)
            if message_match:
                etransfer_message = message_match.group(1).strip()
                print(f"Found e-transfer message: '{etransfer_message}'")
            else:
                etransfer_message = ""
                print("No message found in e-transfer email")

            print(f"Processing e-transfer: ${amount} from {sender_name} <{replyto_address}>")
            if etransfer_message:
                print(f"E-transfer message: '{etransfer_message}'")

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

            # Try to find search results - check if email search was successful
            search_successful = False
            try:
                # Use the new function to find correct family by email verification
                if find_correct_family_result(replyto_address):
                    print("✅ Found and verified correct family from email search")
                    search_successful = True
                else:
                    print("❌ Could not find family with matching email address")
                    search_successful = False
            except Exception as search_result_error:
                print(f"Error during email search result verification: {search_result_error}")
                search_successful = False

            # If email search failed and we have a message, try searching with the message
            if not search_successful and etransfer_message:
                print(f"Email search failed, trying to search with e-transfer message: '{etransfer_message}'")
                
                # Navigate back to main page for new search
                driver.get("https://app.thestudiodirector.com/danceink/admin.sd")
                time.sleep(buffer)
                
                # Find search field again
                search_field = None
                for selector_type, selector_value in search_selectors:
                    try:
                        search_field = driver.find_element(selector_type, selector_value)
                        print(f"Found search field for message search using: {selector_type}='{selector_value}'")
                        break
                    except:
                        continue
                
                if search_field:
                    search_field.clear()
                    search_field.send_keys(etransfer_message)
                    
                    # Click search button for message search
                    search_button = None
                    for selector_type, selector_value in search_button_selectors:
                        try:
                            search_button = driver.find_element(selector_type, selector_value)
                            break
                        except:
                            continue
                    
                    if search_button:
                        search_button.click()
                        print(f"Successfully searched for message: {etransfer_message}")
                    else:
                        search_field.send_keys("\n")
                        print(f"Tried Enter key for message search: {etransfer_message}")
                    
                    time.sleep(buffer)
                    
                    # Try to find results from message search
                    try:
                        search_result_div = driver.find_element(By.CLASS_NAME, "searchResultItem")
                        first_result_link = search_result_div.find_element(By.TAG_NAME, "a")
                        first_result_link.click()
                        print("Clicked first search result from message search (searchResultItem div)")
                        search_successful = True
                        time.sleep(buffer)
                    except Exception as message_search_error:
                        print(f"Could not find search result with message search: {message_search_error}")
                        try:
                            first_result = WebDriverWait(driver, 5).until(
                                EC.element_to_be_clickable((By.XPATH, "//table[@id='accountsTable']//tr[2]//a"))
                            )
                            first_result.click()
                            print("Clicked first search result from message search (fallback method)")
                            search_successful = True
                            time.sleep(buffer)
                        except:
                            print("Message search also failed to find results")
                            search_successful = False

            # If email and message searches failed, try sender name as third fallback
            if not search_successful and sender_name and sender_name != "Unknown":
                print(f"Email and message searches failed, trying to search with sender name: '{sender_name}'")
                
                # Navigate back to main page for sender name search
                driver.get("https://app.thestudiodirector.com/danceink/admin.sd")
                time.sleep(buffer)
                
                # Find search field again
                search_field = None
                for selector_type, selector_value in search_selectors:
                    try:
                        search_field = driver.find_element(selector_type, selector_value)
                        print(f"Found search field for sender name search using: {selector_type}='{selector_value}'")
                        break
                    except:
                        continue
                
                if search_field:
                    search_field.clear()
                    search_field.send_keys(sender_name)
                    
                    # Click search button for sender name search
                    search_button = None
                    for selector_type, selector_value in search_button_selectors:
                        try:
                            search_button = driver.find_element(selector_type, selector_value)
                            break
                        except:
                            continue
                    
                    if search_button:
                        search_button.click()
                        print(f"Successfully searched for sender name: {sender_name}")
                    else:
                        search_field.send_keys("\n")
                        print(f"Tried Enter key for sender name search: {sender_name}")
                    
                    time.sleep(buffer)
                    
                    # Try to find results from sender name search
                    try:
                        search_result_div = driver.find_element(By.CLASS_NAME, "searchResultItem")
                        first_result_link = search_result_div.find_element(By.TAG_NAME, "a")
                        first_result_link.click()
                        print("Clicked first search result from sender name search (searchResultItem div)")
                        search_successful = True
                        time.sleep(buffer)
                    except Exception as sender_search_error:
                        print(f"Could not find search result with sender name search: {sender_search_error}")
                        try:
                            first_result = WebDriverWait(driver, 5).until(
                                EC.element_to_be_clickable((By.XPATH, "//table[@id='accountsTable']//tr[2]//a"))
                            )
                            first_result.click()
                            print("Clicked first search result from sender name search (fallback method)")
                            search_successful = True
                            time.sleep(buffer)
                        except:
                            print("Sender name search also failed to find results")
                            search_successful = False

            # If all three searches failed, skip this email
            if not search_successful:
                print("All searches failed (email, message, and sender name), skipping this email")
                continue

            # Check if we landed on a student page (no ledger tab) or family account page
            try:
                ledger_tab = driver.find_element(By.ID, "tab-ledger")
                # We have a ledger tab, so we're on a family account page
                ledger_tab.click()
                print("Clicked Ledger tab")
                time.sleep(buffer)
                
                # We'll parse unpaid charges later after clicking "Cash, check, trade"
                # For now, just set defaults
                payment_category = "Tuition"  # Will be updated after parsing charges
                payment_amount_to_use = amount  # Will be updated after allocation calculation
                all_allocations = []  # Will be populated after parsing
                    
            except Exception as ledger_error:
                print(f"Could not find Ledger tab: {ledger_error}")
                print("Looks like we're on a student page, trying to navigate to family account...")
                
                # Try to click the Family tab to get family information
                try:
                    family_tab = driver.find_element(By.ID, "tab-family")
                    family_tab.click()
                    print("Clicked Family tab")
                    time.sleep(buffer)
                    
                    # Look for Family Summary table and extract email
                    try:
                        family_summary_table = driver.find_element(By.XPATH, "//table[contains(@class, 'Family Summary') or contains(text(), 'Family Summary')]")
                        family_rows = family_summary_table.find_elements(By.TAG_NAME, 'tr')
                        
                        family_email = None
                        for row in family_rows:
                            cells = row.find_elements(By.TAG_NAME, 'td')
                            if len(cells) >= 2:
                                # Check if second cell contains an email (has @ symbol)
                                potential_email = cells[1].text.strip()
                                if '@' in potential_email:
                                    family_email = potential_email
                                    print(f"Found family email: {family_email}")
                                    break
                        
                        if family_email:
                            print(f"Searching for family account using email: {family_email}")
                            
                            # Navigate back to main page for family search
                            driver.get("https://app.thestudiodirector.com/danceink/admin.sd")
                            time.sleep(buffer)
                            
                            # Search for family using the extracted email
                            search_field = None
                            for selector_type, selector_value in search_selectors:
                                try:
                                    search_field = driver.find_element(selector_type, selector_value)
                                    break
                                except:
                                    continue
                            
                            if search_field:
                                search_field.clear()
                                search_field.send_keys(family_email)
                                
                                # Click search button
                                search_button = None
                                for selector_type, selector_value in search_button_selectors:
                                    try:
                                        search_button = driver.find_element(selector_type, selector_value)
                                        break
                                    except:
                                        continue
                                
                                if search_button:
                                    search_button.click()
                                    print(f"Successfully searched for family email: {family_email}")
                                else:
                                    search_field.send_keys("\n")
                                    print(f"Tried Enter key for family email search: {family_email}")
                                
                                time.sleep(buffer)
                                
                                # Try to find and click family account result with email verification
                                try:
                                    if find_correct_family_result(family_email):
                                        print("✅ Found and verified correct family from family email search")
                                        
                                        # Now try to click the Ledger tab on the family account
                                        ledger_tab = driver.find_element(By.ID, "tab-ledger")
                                        ledger_tab.click()
                                        print("Clicked Ledger tab on family account")
                                        time.sleep(buffer)
                                        
                                        # Set payment category to default since we're now on family account
                                        payment_category = "Tuition"
                                        print("Set payment category to Tuition (family account)")
                                    else:
                                        print("❌ Could not find family with matching family email")
                                        payment_category = "Tuition"
                                    
                                except Exception as family_search_error:
                                    print(f"Could not find/click family account: {family_search_error}")
                                    print("Using default Tuition category and continuing...")
                                    payment_category = "Tuition"
                            else:
                                print("Could not find search field for family email search")
                                payment_category = "Tuition"
                        else:
                            print("Could not find family email in Family Summary table")
                            payment_category = "Tuition"
                            
                    except Exception as family_table_error:
                        print(f"Could not find or read Family Summary table: {family_table_error}")
                        payment_category = "Tuition"
                        
                except Exception as family_tab_error:
                    print(f"Could not find Family tab: {family_tab_error}")
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

            # NOW parse unpaid charges and calculate payment allocation (after "Cash, check, trade" is clicked)
            print("Parsing unpaid charges after clicking 'Cash, check, trade'...")
            time.sleep(2)  # Give extra time for page to load with charge details
            
            unpaid_charges = parse_unpaid_charges(driver)
            payment_allocations = calculate_payment_allocation(amount, unpaid_charges)
            
            print(f"Payment allocations calculated: {payment_allocations}")
            
            # Update payment details based on parsed charges
            if payment_allocations:
                # For multiple allocations, we need to handle split payments differently
                if len(payment_allocations) > 1:
                    print(f"Multiple allocations detected: {payment_allocations}")
                    # Use the full payment amount for the form
                    payment_amount_to_use = amount
                    # We'll handle the splits after clicking Split Payment button
                else:
                    # Single allocation
                    payment_category = payment_allocations[0][0]  # Get the category from first allocation
                    payment_amount_to_use = payment_allocations[0][1]  # Get the amount
                    print(f"Single allocation: ${payment_amount_to_use} to {payment_category}")
            else:
                payment_category = "Tuition"  # Default fallback
                payment_amount_to_use = amount
                print("No allocations calculated, using default Tuition")
                
            # Store all allocations for processing splits
            all_allocations = payment_allocations

            # Set payment amount (using calculated allocation amount)
            amount_field = driver.find_element(By.NAME, "amount")
            amount_field.clear()
            amount_field.send_keys(str(payment_amount_to_use))
            print(f"Successfully set payment amount: ${payment_amount_to_use}")

            # Set reference in notes field (since there's no dedicated reference field)
            try:
                notes_field = driver.find_element(By.CSS_SELECTOR, '[name="notes"]')
                notes_field.clear()
                notes_field.send_keys(f"{reference_number}")
                print(f"Successfully set reference in notes field: {reference_number}")
            except Exception as e:
                print(f"Could not find notes field: {e}")
            
            # Set payment date using the correct field names - these are date input fields
            try:
                # Format date as YYYY-MM-DD for HTML date input
                formatted_date = f"{year}-{month_number:02d}-{day:02d}"
                
                # Set due_date field
                due_date_field = driver.find_element(By.NAME, "due_date")
                due_date_field.clear()
                due_date_field.send_keys(formatted_date)
                print(f"✅ Successfully set due_date: {formatted_date}")
                
            except Exception as date_error:
                print(f"Error setting due_date: {date_error}")
                
                # Try alternative method with deposit_date
                try:
                    formatted_date = f"{year}-{month_number:02d}-{day:02d}"
                    deposit_date_field = driver.find_element(By.NAME, "deposit_date")
                    deposit_date_field.clear()
                    deposit_date_field.send_keys(formatted_date)
                    print(f"✅ Successfully set deposit_date: {formatted_date}")
                except Exception as deposit_date_error:
                    print(f"Could not set deposit_date either: {deposit_date_error}")
                    
                    # Debug: List all select elements to find date fields
                    try:
                        all_selects = driver.find_elements(By.TAG_NAME, "select")
                        print(f"Available select fields on page:")
                        for select_field in all_selects:
                            name = select_field.get_attribute("name") or "no name"
                            select_id = select_field.get_attribute("id") or "no id"
                            if name != "no name" or select_id != "no id":
                                print(f"  Select: name='{name}', id='{select_id}'")
                    except:
                        pass

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
            
            # Handle split payments based on payment allocations
            print(f"Processing payment allocations: {all_allocations}")
            
            # Determine if we need split payments
            if all_allocations and len(all_allocations) > 1:
                print(f"Multiple categories detected - setting up split payments for {len(all_allocations)} categories")
                
                # Click Split Payment button multiple times to reveal all needed fields
                # Each click reveals one additional split pair (split_amt + paid_toward)
                # We need (len(all_allocations) - 1) clicks since the first pair is already visible
                clicks_needed = len(all_allocations) - 1
                print(f"Need to click Split Payment button {clicks_needed} times to reveal all fields")
                
                try:
                    for click_num in range(clicks_needed):
                        split_payment_button = driver.find_element(By.ID, 'splitpayment')
                        split_payment_button.click()
                        print(f"Clicked Split Payment button #{click_num + 1} of {clicks_needed}")
                        time.sleep(buffer)
                        
                        # Check if the expected field is now available
                        expected_field_num = click_num + 2  # After first click, we expect split_amt2, etc.
                        expected_field_name = f"split_amt{expected_field_num}"
                        try:
                            driver.find_element(By.NAME, expected_field_name)
                            print(f"✅ {expected_field_name} is now available")
                        except:
                            print(f"⚠️ {expected_field_name} not found after clicking")
                    
                    print("All Split Payment button clicks completed")
                    
                    # Process each allocation
                    for i, (category, allocation_amount) in enumerate(all_allocations):
                        field_number = i + 1  # paid_toward1, paid_toward2, etc.
                        
                        # Use the correct split amount field name
                        amount_field_name = f"split_amt{field_number}"
                        category_field_name = f"paid_toward{field_number}"
                        
                        print(f"Setting allocation {field_number}: ${allocation_amount} to {category}")
                        
                        # Check for any existing alerts before starting
                        try:
                            alert = driver.switch_to.alert
                            alert_text = alert.text
                            print(f"⚠️ Pre-existing alert found: {alert_text}")
                            alert.accept()
                            print("Pre-existing alert dismissed")
                            time.sleep(0.5)  # Wait after dismissing alert
                        except:
                            pass  # No alert, which is normal
                        
                        # Set the amount for this split
                        try:
                            # First, verify the field exists
                            split_amount_field = driver.find_element(By.NAME, amount_field_name)
                            print(f"✅ Found {amount_field_name} field")
                            
                            # Clear the field first
                            split_amount_field.clear()
                            print(f"Cleared {amount_field_name}")
                            
                            # Wait a moment for any validation to complete
                            time.sleep(0.5)
                            
                            # Set the allocation amount (ensure it's formatted properly)
                            amount_str = f"{allocation_amount:.2f}"
                            split_amount_field.send_keys(amount_str)
                            print(f"✅ Set {amount_field_name} to ${amount_str}")
                            
                            # Immediately verify the value was set correctly
                            time.sleep(0.2)
                            actual_value = split_amount_field.get_attribute("value")
                            print(f"Verification: {amount_field_name} contains: '{actual_value}'")
                            
                            # Check for alerts immediately after setting value
                            alert_handled = False
                            for attempt in range(3):  # Try up to 3 times to handle alerts
                                try:
                                    alert = driver.switch_to.alert
                                    alert_text = alert.text
                                    print(f"⚠️ Alert #{attempt + 1} appeared: {alert_text}")
                                    alert.accept()
                                    print(f"Alert #{attempt + 1} dismissed")
                                    alert_handled = True
                                    time.sleep(0.5)  # Wait after dismissing alert
                                    
                                    # Re-verify field value after alert
                                    current_value = split_amount_field.get_attribute("value")
                                    print(f"Field {amount_field_name} value after alert #{attempt + 1}: '{current_value}'")
                                    
                                    # If value was cleared by alert, re-set it
                                    if current_value != amount_str and current_value == "":
                                        print(f"Alert cleared the field! Re-setting {amount_field_name} to ${amount_str}")
                                        split_amount_field.clear()
                                        split_amount_field.send_keys(amount_str)
                                        time.sleep(0.3)
                                    
                                except:
                                    break  # No more alerts
                            
                            if alert_handled:
                                print(f"Finished handling alerts for {amount_field_name}")
                            else:
                                print(f"No alerts appeared for {amount_field_name}")
                                
                            # Final verification after all alert handling
                            final_value = split_amount_field.get_attribute("value")
                            print(f"Final verification: {amount_field_name} = '{final_value}'")
                            
                            if final_value != amount_str:
                                print(f"❌ Final value mismatch for {amount_field_name}! Expected: {amount_str}, Got: {final_value}")
                                # One more attempt to correct
                                split_amount_field.clear()
                                split_amount_field.send_keys(amount_str)
                                print(f"Made final correction attempt for {amount_field_name}")
                            else:
                                print(f"✅ {amount_field_name} value confirmed: ${final_value}")
                                
                        except Exception as amount_error:
                            print(f"❌ Could not find amount field {amount_field_name}: {amount_error}")
                            
                            # Debug: List available split amount fields
                            try:
                                print(f"Looking for available split amount fields...")
                                for field_num in range(1, 6):  # Check split_amt1 through split_amt5
                                    test_field_name = f"split_amt{field_num}"
                                    try:
                                        test_field = driver.find_element(By.NAME, test_field_name)
                                        print(f"  ✅ {test_field_name} exists")
                                    except:
                                        print(f"  ❌ {test_field_name} not found")
                            except:
                                pass
                                
                            # Check for alert in case of error
                            try:
                                alert = driver.switch_to.alert
                                alert_text = alert.text
                                print(f"⚠️ Alert appeared during amount setting: {alert_text}")
                                alert.accept()  # Dismiss the alert
                                print("Alert dismissed")
                            except:
                                pass
                                
                            # List all input fields to debug
                            try:
                                all_inputs = driver.find_elements(By.TAG_NAME, "input")
                                print(f"Available input fields on page:")
                                for input_field in all_inputs[:20]:  # Limit to first 20 to avoid spam
                                    name = input_field.get_attribute("name") or "no name"
                                    field_type = input_field.get_attribute("type") or "no type"
                                    if name != "no name":
                                        print(f"  Input: name='{name}', type='{field_type}'")
                            except:
                                pass
                        
                        # Set the category for this split
                        try:
                            category_select = Select(driver.find_element(By.NAME, category_field_name))
                            
                            # Get available options
                            all_options = category_select.options
                            available_options = [(opt.get_attribute('value'), opt.text) for opt in all_options]
                            print(f"Available options for {category_field_name}: {[text for value, text in available_options]}")
                            
                            # Try to find the matching category
                            selection_successful = False
                            
                            # First try exact match by visible text
                            try:
                                category_select.select_by_visible_text(category)
                                print(f"✅ Selected '{category}' by exact text for {category_field_name}")
                                selection_successful = True
                            except Exception as exact_error:
                                print(f"Exact text match failed for '{category}': {exact_error}")
                            
                            # If exact match failed, try by value
                            if not selection_successful:
                                for value, text in available_options:
                                    if text == category:
                                        try:
                                            category_select.select_by_value(value)
                                            print(f"✅ Selected '{category}' by value '{value}' for {category_field_name}")
                                            selection_successful = True
                                            break
                                        except Exception as value_error:
                                            print(f"Value selection failed for '{value}': {value_error}")
                                            continue
                            
                            # If still not successful, try partial matching
                            if not selection_successful:
                                print(f"Trying partial matching for '{category}'...")
                                for value, text in available_options:
                                    if category.lower() in text.lower() or text.lower() in category.lower():
                                        try:
                                            category_select.select_by_value(value)
                                            print(f"✅ Selected '{category}' option for {category_field_name}: '{value}' ('{text}')")
                                            selection_successful = True
                                            break
                                        except Exception as select_error:
                                            print(f"Partial match selection failed for '{value}': {select_error}")
                                            continue
                            
                            # If direct match failed, try partial matches
                            if not selection_successful:
                                if "tuition" in category.lower():
                                    for value, text in available_options:
                                        if "tuition" in text.lower():
                                            try:
                                                category_select.select_by_value(value)
                                                print(f"Selected Tuition option for {category_field_name}: '{value}' ('{text}')")
                                                selection_successful = True
                                                break
                                            except:
                                                continue
                                elif "costume" in category.lower():
                                    for value, text in available_options:
                                        if "costume" in text.lower():
                                            try:
                                                category_select.select_by_value(value)
                                                print(f"Selected Costume option for {category_field_name}: '{value}' ('{text}')")
                                                selection_successful = True
                                                break
                                            except:
                                                continue
                                elif "private" in category.lower():
                                    for value, text in available_options:
                                        if "private" in text.lower() or "lesson" in text.lower():
                                            try:
                                                category_select.select_by_value(value)
                                                print(f"Selected Private Lesson option for {category_field_name}: '{value}' ('{text}')")
                                                selection_successful = True
                                                break
                                            except:
                                                continue
                            
                            if not selection_successful:
                                print(f"⚠️ Could not find matching option for category: {category}")
                                
                        except Exception as category_error:
                            print(f"Error setting category for {category_field_name}: {category_error}")
                    
                    # Final verification: Check all split amounts are set correctly
                    print("=== Final Split Amount Verification ===")
                    for i, (category, expected_amount) in enumerate(all_allocations):
                        field_number = i + 1
                        amount_field_name = f"split_amt{field_number}"
                        try:
                            split_field = driver.find_element(By.NAME, amount_field_name)
                            actual_value = split_field.get_attribute("value")
                            expected_str = f"{expected_amount:.2f}"
                            
                            if actual_value == expected_str:
                                print(f"✅ {amount_field_name}: Expected ${expected_str}, Got ${actual_value}")
                            else:
                                print(f"❌ {amount_field_name}: Expected ${expected_str}, Got ${actual_value}")
                                print(f"Attempting to correct {amount_field_name}...")
                                split_field.clear()
                                split_field.send_keys(expected_str)
                                print(f"Corrected {amount_field_name} to ${expected_str}")
                                
                        except Exception as verify_error:
                            print(f"Could not verify {amount_field_name}: {verify_error}")
                    print("=== End Verification ===")
                    
                except Exception as multi_split_error:
                    print(f"Error setting up multiple split payments: {multi_split_error}")
                    
            elif all_allocations and len(all_allocations) == 1:
                # Single allocation - use traditional split payment method
                payment_category = all_allocations[0][0]
                print(f"Single allocation: Using {payment_category} for split payment")
                
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
                    print("Clicked Split Payment button for single allocation")
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
            else:
                print("No valid allocations - skipping split payment setup")
            
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
                                print(f"⚠️ Balance is NOT zero: {current_balance} - this was a partial payment")
                                print("✅ Payment was successfully applied - marking email as processed")
                                
                                # Mark email as processed since the payment was successfully applied
                                # Even though balance isn't zero, the e-transfer was processed
                                mark_email_processed(email_id, reference_number)
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
