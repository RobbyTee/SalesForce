# SalesForce

This Python script automates assembling a spreadsheet based on information unique to an account in SalesForce and scheduling an appointment based on the customer's preferences. 
I'm a glutton for punishment so I wrote this using Safari on Mac OS. I'm certain at this point it would have been easier to write this using Google Chrome.

We use the Google Suite for spreadsheets and documents; [Cal.com](https://cal.com) for scheduling appointments.

## Setup
### Python
This script is written in [Python 3.10.8](https://www.python.org/downloads/release/python-3108/). Once installed, navigate to the folder with requirements.txt and run:
```
pip3 install requirements.txt
```

### Safari 
Your SafariDriver is located at /usr/bin/safaridriver. Be sure to "Allow Remote Automation" in Safari. [This guide](https://developer.apple.com/documentation/webkit/testing_with_webdriver_in_safari#2957277) was very helpful for set up and testing.

### PyDrive2
PyDrive2 uses OAuth created from Google Drive API. [Instructions here](https://docs.iterative.ai/PyDrive2/quickstart/). Store this
file in the main root folder where SalesForce.py lives.

### GSpread
GSpread uses credentials from Google Drive API as well. [Instructions here](https://docs.gspread.org/en/v5.12.1/oauth2.html). Rename this file to "credentials.json" 
and store it in the main root folder as well.

### SalesForce
Your dashboard should have the Shipped/Arrived chart in the top left with the same columns I have listed in this image.


![Your dashboard should have the Shipped/Arrived chart in the top left with the same columns I have listed in this image.](https://i.imgur.com/r1nPiHh.jpeg)


## Global Variables
![Global variables necessary for script to run](https://i.imgur.com/LvZHjRH.png)
### Firewall Rules
**fw_rules_spreadsheet_id**

**fw_rules_pdf_filename**

**fw_rules_file_path**

### Gmail App Password

### Chat GPT API Key

### Cal.com API Key

## Global Controls
You should now be ready to run the script with:
```
python3 salesforce.py
```
