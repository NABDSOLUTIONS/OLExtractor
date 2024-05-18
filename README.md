# OLExtractor

Email Extractor Outlook.Application This Python script extracts email addresses from Microsoft Outlook, excluding addresses specified in an exclude.txt file, and saves the unique, valid addresses to a CSV file.

--------------------------------------------------------------------


# Overview
This script uses the win32com.client library to access Outlook and extract email addresses from the messages stored in different folders. It checks the email addresses against a list of excluded addresses from the exclude.txt file and saves the unique, valid addresses to a CSV file.
Instructions
Ensure you have Python 3.x installed on your Windows machine.
Install the required library win32com.client. You can do this by running the following command in your terminal: pip install pywin32
Update the exclude.txt file with the email addresses that you want to exclude from extraction.
Run the script.
The script will create a CSV file containing unique, valid email addresses, excluding the addresses specified in the exclude.txt file. The CSV file name is dynamically generated with the current timestamp.
# How it Works
The script follows these steps:
- Loads email addresses from the exclude.txt file into a list.
- Accesses Outlook folders using the win32com.client library.
- Iterates through each Outlook folder and their items (messages).
- Extracts email addresses from the recipients of the messages.
- hecks each email address against the exclude list and other criteria to ensure uniqueness, validity, and length.
- Appends valid email addresses to a list.
- Saves the list of valid email addresses to a CSV file, with the file name generated based on the current timestamp.
# Note
You must run this script on a Windows machine with Microsoft Outlook installed.
The script requires the win32com.client library, which can be installed with pip install pywin32.
Ensure that you have read and write permissions for the directory containing the script and the exclude.txt file.
Always make sure to protect the privacy of the email addresses you extract and follow applicable data protection regulations.


Red Teaming & Marketing mindset by NABD Solutions 
