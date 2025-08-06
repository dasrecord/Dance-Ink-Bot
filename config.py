import os
from dotenv import load_dotenv

# Set safe mode to True to prevent any actions that could cause damage
safe_mode = True

# Set headless mode to True for headless operation (no GUI)
headless = False

# Load passwords from .env file
load_dotenv("./passwords.env")

# Set the username and password for the email account
email_username = os.getenv("EMAIL_USERNAME")
email_password = os.getenv("EMAIL_PASSWORD")

# Dance Ink Studio Director URL
studio_director_url = "https://app.thestudiodirector.com/danceink/login.sd"

#  Shotokan Karate Studio Director URL
shotokan_studio_director_url = "https://app.thestudiodirector.com/shotokankarateyxe/login.sd"

# Set username and password for the websites
studio_director_username = os.getenv("DANCE_INK_USERNAME")
studio_director_password = os.getenv("DANCE_INK_PASSWORD")
shotokan_studio_director_username = os.getenv("SHOTOKAN_USERNAME")
shotokan_studio_director_password = os.getenv("SHOTOKAN_PASSWORD")

# Set category hierarchy for payments
category_hierarchy = ("Registration", "Costume Deposit", "Tuition", "Exam Fee")


