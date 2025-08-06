#!/usr/bin/env python3

import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from config import studio_director_url, studio_director_username, studio_director_password, headless

# Set up the WebDriver
chrome_options = Options()
if headless:
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920x1080")

# Start the WebDriver
driver = webdriver.Chrome(options=chrome_options)
driver.implicitly_wait(10)

try:
    print("=== Debug Login Process ===")
    
    # Navigate to studio_director_url
    driver.get(studio_director_url)
    print(f"Navigated to: {driver.current_url}")
    print(f"Page title: {driver.title}")
    
    # Wait for page to load
    time.sleep(3)
    
    # Check what forms are available
    print("\n=== Available Forms ===")
    forms = driver.find_elements(By.TAG_NAME, 'form')
    for i, form in enumerate(forms):
        print(f"Form {i}: ID={form.get_attribute('id')}, Action={form.get_attribute('action')}")
    
    # Check what buttons are available
    print("\n=== Available Buttons ===")
    buttons = driver.find_elements(By.TAG_NAME, 'button')
    for i, button in enumerate(buttons):
        print(f"Button {i}: Text='{button.text}', Type={button.get_attribute('type')}, ID={button.get_attribute('id')}")
    
    # Check what input elements are available
    print("\n=== Available Input Elements ===")
    inputs = driver.find_elements(By.TAG_NAME, 'input')
    for i, inp in enumerate(inputs):
        print(f"Input {i}: Type={inp.get_attribute('type')}, ID={inp.get_attribute('id')}, Name={inp.get_attribute('name')}, Value={inp.get_attribute('value')}")
    
    # Try to find username field
    print("\n=== Finding Username Field ===")
    try:
        username_input = driver.find_element(By.XPATH, '//*[@id="idForm"]/div[2]/input')
        print("✅ Found username field")
        username_input.clear()
        username_input.send_keys(studio_director_username)
        print("✅ Entered username")
    except Exception as e:
        print(f"❌ Could not find username field: {e}")
        
    # Try to find password field
    print("\n=== Finding Password Field ===")
    try:
        password_input = driver.find_element(By.XPATH, '//*[@id="passwordField"]')
        print("✅ Found password field")
        password_input.clear()
        password_input.send_keys(studio_director_password)
        print("✅ Entered password")
    except Exception as e:
        print(f"❌ Could not find password field: {e}")
        
    # Try to find login button
    print("\n=== Finding Login Button ===")
    login_selectors = [
        '//input[@type="submit"]',
        '//button[@type="submit"]', 
        '//*[@id="idForm"]//input[@type="submit"]',
        '//*[@id="idForm"]//button[@type="submit"]',
        '//input[@value="Login"]',
        '//input[@value="Sign In"]',
        '//button[contains(text(), "Login")]',
        '//button[contains(text(), "Sign In")]',
        '//*[@id="idForm"]//input',
        '//*[@id="idForm"]//button'
    ]
    
    found_button = False
    for selector in login_selectors:
        try:
            elements = driver.find_elements(By.XPATH, selector)
            if elements:
                print(f"✅ Found {len(elements)} element(s) with selector: {selector}")
                for i, elem in enumerate(elements):
                    print(f"   Element {i}: Tag={elem.tag_name}, Type={elem.get_attribute('type')}, Value={elem.get_attribute('value')}, Text='{elem.text}'")
                found_button = True
            else:
                print(f"❌ No elements found with selector: {selector}")
        except Exception as e:
            print(f"❌ Error with selector '{selector}': {e}")
    
    if not found_button:
        print("❌ No login button found with any selector")
    
    # Get page source around the form
    print("\n=== Page Source Around Form ===")
    try:
        form_element = driver.find_element(By.ID, 'idForm')
        print("Form HTML:")
        print(form_element.get_attribute('outerHTML')[:1000] + "..." if len(form_element.get_attribute('outerHTML')) > 1000 else form_element.get_attribute('outerHTML'))
    except Exception as e:
        print(f"Could not get form HTML: {e}")
        
finally:
    driver.quit()
    print("\n=== Debug Complete ===")
