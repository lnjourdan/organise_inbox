# === HOW TO USE === #

# You will need the following installed on your PC:
# 1. Python (Can be dowbloaded from: https://www.python.org/downloads/)
# 2. pywin32 Library (Can be installed via powershell with the command: pip install pywin32)
# 3. Microsoft Outlook (Desktop App)
# 
# How to run:
# 1. Open your powershell/command line (can search for this from your task bar)
# 2. Direct yourself to where you have saved this, most likely your downloads (type this into the powershell: cd C:\Users\YOUR_NAME\downloads)
# 3. Type the following into your powershell in order to run the script: python organise_inbox.py
# 4. You should shortly see changes appear within your Outlook App.
#
# Note: I originally wrote this script to organise my own inbox, I had all emails organised into yearly folders, but wanted monthly subfolders for increased clarity.
# I have done my best to adaprt the script now so that it will work for users if they are organising directly from the inbox.
#
# Thanks for using!

# === DISCLAIMER === #

# You are using this script at your own risk, ensure important data is backed up or run a test if you feel unsure. 

import win32com.client
from datetime import datetime

# === CONFIGURATION ===
YEAR = "2024"  # The year of emails you want to organize (e.g., 2024, 2023, etc.)
# This determines which emails are selected and sorted into month folders
MAILBOX_NAME = "your_email"  # Your mailbox name (as shown in Outlook)

# === CONNECT TO OUTLOOK ===
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
mailbox = outlook.Folders[MAILBOX_NAME]
inbox = mailbox.Folders["Inbox"]

# === MESSAGE SELECTION ===
# ⚙️ This script supports two configurations:
# 
# OPTION 1 — If your emails are inside a subfolder like "Inbox > 2024":
# Uncomment the two lines below and make sure a folder named "2024" exists under Inbox.
year_folder = inbox.Folders[YEAR]
messages = year_folder.Items

# OPTION 2 — If your emails are all in the main Inbox (no "2024" folder):
# Comment out the two lines above (OPTION 1), and uncomment the line below:
# messages = inbox.Items

# === FILTER MESSAGES TO THE SELECTED YEAR ===
messages.Sort("[ReceivedTime]", Descending=False)
messages.IncludeRecurrences = False

# Only include messages from the selected year
messages_filtered = messages.Restrict(
    f"[ReceivedTime] >= '01/01/{YEAR}' AND [ReceivedTime] < '01/01/{int(YEAR) + 1}'"
)
print(f"Total MailItems in '{YEAR}': {messages_filtered.Count}")

# === HELPER FUNCTION TO CREATE MONTH SUBFOLDERS ===
def get_or_create_folder(parent, name):
    """
    Creates a folder if it doesn't exist and returns the folder object.
    """
    try:
        return parent.Folders[name]
    except:
        return parent.Folders.Add(name)

# === MOVE EMAILS TO CORRESPONDING MONTH SUBFOLDERS ===
moved = 0

# Choose the folder where month subfolders should be created:
# - If using OPTION 1: year_folder
# - If using OPTION 2 (Inbox only): inbox
destination_root = year_folder  # Change to "inbox" if using OPTION 2

for message in list(messages_filtered):
    try:
        if message.Class != 43:  # MailItem type
            continue

        # Determine month from the received date
        received = message.ReceivedTime
        month_name = received.strftime("%B")  # e.g., "January", "February"

        # Create or get the subfolder for that month
        month_folder = get_or_create_folder(destination_root, month_name)

        # Move the message
        message.Move(month_folder)
        moved += 1

    except Exception as e:
        print(f"Error moving message: {e}")

# === SUMMARY ===
print(f"\n Done. {moved} emails moved into subfolders under '{destination_root.Name}'.")
