# SalesForce

This Python script automates assembling a spreadsheet based on information unique to an account in SalesForce and scheduling an appointment based on the customer's preferences. 
I'm a glutton for punishment so I wrote this using Safari on Mac OS. I'm certain at this point it would have been easier to write this using Google Chrome.

We use the Google Suite for spreadsheets and documents; [Cal.com](https://cal.com) for scheduling appointments.

## Setup
This script is written in [Python 3.10.8](https://www.python.org/downloads/release/python-3108/).
```
pip3 install requirements.txt
```
Your SafariDriver is located at /usr/bin/safaridriver. Be sure to "Allow Remote Automation" in Safari.

PyDrive2 uses OAuth created from Google Drive API. [Instructions here](https://docs.iterative.ai/PyDrive2/quickstart/).

GSpread uses credentials from Google Drive API as well. [Instructions here](https://docs.gspread.org/en/v5.12.1/oauth2.html).

-> Rename this file to "credentials.json"

BOTH CREDENTIALS NEED TO BE STORED INSIDE THE SAME FOLDER AS firewall_rules.py

You should now be ready to run the script with:
```
python3 salesforce.py
```
