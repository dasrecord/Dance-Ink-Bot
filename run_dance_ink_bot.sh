#!/bin/bash

# Dance Ink Bot Cron Wrapper Script
# This script ensures proper environment for running the bot via cron

# Set the working directory
cd "/Users/PD/PROJECTS/AUTOMATIONS/Dance Ink Bot"

# Add Python to PATH (adjust if your Python is in a different location)
export PATH="/usr/local/bin:/usr/bin:/bin:/opt/homebrew/bin:$PATH"

# Set display for headless Chrome (required for cron)
export DISPLAY=:0

# Log file for debugging
LOG_FILE="/Users/PD/PROJECTS/AUTOMATIONS/Dance Ink Bot/cron.log"

# Create log entry with timestamp
echo "=== Dance Ink Bot Cron Job Started at $(date) ===" >> "$LOG_FILE"

# Run the Python script and capture output
python3 dance_ink_bot.py >> "$LOG_FILE" 2>&1

# Log completion
echo "=== Dance Ink Bot Cron Job Completed at $(date) ===" >> "$LOG_FILE"
echo "" >> "$LOG_FILE"
