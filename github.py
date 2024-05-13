import logging

logging.basicConfig(filename='logfile.log', level=logging.INFO, 
                    format='%(asctime)s - %(levelname)s - %(message)s')

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.safari.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (TimeoutException, 
                                        NoSuchElementException)
from time import sleep
from oauth2client.service_account import ServiceAccountCredentials
from pydrive2.auth import GoogleAuth
from pydrive2.drive import GoogleDrive
from pydrive2.files import ApiRequestError
from email.mime.multipart import MIMEMultipart 
from email.mime.text import MIMEText 
from email.mime.base import MIMEBase 
from email import encoders 
from datetime import date, timedelta, datetime
from openai import OpenAI
from getpass import getpass
import os
import sys
import gspread
import smtplib
import requests
import pytz


#Global control
use_my_email = False
my_email = 'Your Email Address'
send_firewall_rules = True
attach_pdf = True
send_fw_email = True
schedule_install = False

#Global variables
fw_rules_spreadsheet_id = 'Spreadsheet ID Here'
fw_rules_pdf_filename = 'Firewall_Rules.pdf'
fw_rules_file_path = '/Full/Filepath/Here.pdf'
salesforce_username = input('SalesForce Username: ')
salesforce_password = getpass('SalesForce Password: ')
salesforce_website = 'SalesForce main login page'
salesforce_dashboard = 'Your Dashboard'
email_app_password = 'Your GMail App Password' #See Readme for more information
chatgpt_api_key = 'API Key'
cal_com_api_key = 'API Key'

class Salesforce_Objects:
    def __init__(self):
        self.website = salesforce_website
        self.dashboard = (salesforce_dashboard)
        self.header = ('//*[@class="slds-global-header slds-grid ' 
                       'slds-grid--align-spread"]') #DP logo        
        self.username = salesforce_username
        self.password = salesforce_password

        #Useful XPATHs
        self.x_account_tab_in_AU = ('//*[@class="tabBar slds-grid '
                                    'slds-tabs--default slds-sub-tabs"]'
                                    '/ul[2]/li[2]/a')
        self.x_account_tab_in_Case = ('//*[@class="tabBar slds-grid '
                                    'slds-tabs--default slds-sub-tabs"]'
                                    '/ul[2]/li[2]/a//*[@id="1323:0"]'
                                    '/div/div/div/div/ul[2]/li[2]/a')
        self.x_account_name = ('//*[@class=' #On main account tab
                            '"testonly-outputNameWithHierarchyIcon'
                            'slds-grid sfaOutputNameWithHierarchyIcon"]'
                            '/lightning-formatted-text')
        self.x_account_in_search = ('//*[@aria-label="Accounts||List View"]'
                                    '//table/tbody/tr/th/span/a')
        self.x_au_in_search = ('//*[@aria-label="Account Updates||List View"]'
                               '//table/tbody/tr/td[2]/span/a')
        self.x_case_in_search = ('//*[@aria-label="Cases||List View"]'
                                 '//table/tbody/tr/th/span/a')        
        self.x_asset_widget = ('//*[@class="slds-grid slds-wrap list"]'
                               '/li[3]//a') #On main account tab
        self.x_account_updates_widget = ('//*[@class="slds-grid slds-wrap '
                                         'list"]/li[6]//a')
        self.x_asset_list = ('//*[@role="grid"]/tbody/tr/th[1]//a//span')
        self.x_asset_mac_address = ('//*[@data-target-selection-name='
                                    '"sfdc:RecordField.Asset.MAC_Address__c"]'
                                    '/div/dl/dd/div/span/slot[1]'
                                    '/lightning-formatted-text')
        self.x_asset_url = ('//*[@data-target-selection-name='
                            '"sfdc:RecordField.Asset.Vow_Asset_URL__c"]'
                            '/div/dl/dd/div/span/slot[1]'
                            '/lightning-formatted-url/a')
        self.x_ob_account_update = ('//*[@title="VOW - Onboarding"]'
                                    '//ancestor::tr/td[3]//a')
        self.x_google_doc_link = ('//*[@data-target-selection-name="sfdc:'
                                  'RecordField.Account_Update__c.'
                                  'Customer_Account_Google_URL__c"]'
                                  '/div/dl/dd/div/span/slot[1]'
                                  '/lightning-formatted-url/a')
        self.x_dashboard_iframe = ('//frame[@title="dashboard"]')
        self.x_refresh = ('//*[@class="slds-button slds-button_'
                          'neutral refresh"]')
        self.x_shipped_table = ('//*[@title="Vow Equipment Shipped/Arrived"]')
        self.x_customer_name = ('//*[@data-target-selection-name='
                                '"sfdc:RecordField.Account_Update__c.Contact'
                                '__c"]/div/dl/dd/div/span/slot/force-lookup'
                                '/div/records-hoverable-link/div/a'
                                '/slot/slot/span')
        self.x_customer_email = ('//*[@data-target-selection-name='
                                 '"sfdc:RecordField.Account_Update__c.'
                                 'Contact_Email__c"]/div/dl/dd/div/span'
                                 '/slot/records-formula-output/slot'
                                 '/lightning-formatted-text/a')
        self.x_phone_number = ('//*[@data-target-selection-name='
                               '"sfdc:RecordField.Account.Phone"]'
                               '/div/dl/dd/div/span/slot/records-output-phone'
                               '/lightning-formatted-phone')
        self.x_firewall_edit = ('//*[@data-target-selection-name='
                                '"sfdc:RecordField.Account_Update__c.'
                                'Firewall_Rules_Required__c"]'
                                '/div/dl/dd/div/button')
        self.x_firewall_checkbox = ('//*[@data-target-selection-name='
                                    '"sfdc:RecordField.Account_Update__c.'
                                    'Firewall_Rules_Required__c"]'
                                    '/span/slot/records-record-layout-checkbox'
                                    '/lightning-input/lightning-primitive'
                                    '-input-checkbox/div/span/input')
        self.x_save_button = ('//button[@name="SaveEdit"]')
        self.x_done_button = ('//button[@class="slds-button '
                              'slds-button--neutral ok desktop uiButton--'
                              'default uiButton--brand uiButton"]')
        self.x_upload_files = ('//input'
                                    '[@class="slds-file-selector__input '
                                    'slds-assistive-text"]')
        self.x_local_time = ('//*[@class="highlights slds-clearfix '
                           'slds-page-header slds-page-header_record-home"]'
                           '/div[2]/slot/records-highlights-details-item/div'
                           '/p[2]/slot/records-formula-output/slot'
                           '/lightning-formatted-text')
        self.x_appt_instructions = ('//*[@data-target-selection-name='
                                     '"sfdc:RecordField.Account_Update__c.'
                                     'Setup_Instructions__c"]/div/dl/dd/div'
                                     '/span/slot/lightning-formatted-text')
        self.x_install_tier = ('//*[@data-target-selection-name="sfdc:'
                               'RecordField.Account_Update__c.'
                               'IVR_Install_Tier__c"]/div/dl/dd/div/span'
                               '/slot/lightning-formatted-text')
        self.x_address = ('//*[@class="highlights slds-clearfix '
                          'slds-page-header slds-page-header_record-home"]'
                          '/div[2]/slot/records-highlights-details-item[2]'
                          '/div/p[2]/slot/lightning-formatted-address/a/div[2]')
        self.x_edit_date = ('//*[@data-target-selection-name="sfdc:RecordField'
                            '.Account_Update__c.Install_Date_Time__c"]'
                            '/div/dl/dd/div/button')
        self.x_au_status = ('//*[@data-target-selection-name='
                            '"sfdc:RecordField.Account_Update__c.Status__c"]'
                            '/span/slot/records-record-picklist'
                            '/records-form-picklist/lightning-picklist'
                            '/lightning-combobox/div/div'
                            '/lightning-base-combobox/div/div/div/button')
        self.x_status_dropdown = ('//*[@data-target-selection-name='
                                  '"sfdc:RecordField.Account_Update__c'
                                  '.Status__c"]/span/slot/'
                                  'records-record-picklist'
                                  '/records-form-picklist/lightning-picklist'
                                  '/lightning-combobox/div/div'
                                  '/lightning-base-combobox/div/div/div[2]'
                                  '/lightning-base-combobox-item[6]')
        self.x_install_date_time = ('//*[@data-target-selection-name="sfdc:'
                                    'RecordField.Account_Update__c.Install_'
                                    'Date_Time__c"]/div/dl/dd/div/span/slot'
                                    '/lightning-formatted-text')


class Safari_Operations:

    def __init__(self):
        #Used only once to initialize Safari
        safari_driver = Service('/usr/bin/safaridriver') 
        #The primary function we use! safari.get('http://website.com')
        self.safari = webdriver.Safari(service=safari_driver) 
        #Sets the timer for how long to wait for pages to load
        self.safari.implicitly_wait(5) 
        #hold.until() is huge. Waits for element on screen before proceeding
        self.hold = WebDriverWait(self.safari, 5)
        #action.send_keys(Keys.RETURN).perform() is very handy.
        self.action = ActionChains(self.safari)

    def wait_until(self, XPATH) -> None:
        self.hold.until(EC.presence_of_element_located((By.XPATH, XPATH)))
        return
    
    def find_element(self, xpath):
        try:
            element = self.safari.find_element(By.XPATH, xpath)
        except NoSuchElementException:
            return 0
        return element
    
    def get_href(self, element) -> str:
        href = element.get_attribute('href') 
        return href
    
    def get_value(self, element) -> str:
        value = element.get_attribute('value') 
        return value

    def get_text(self, xpath: str) -> str:
        element = self.find_element(xpath)
        return element.text

    def login(self, website, username, password, test):
        self.safari.get(website)
        self.safari.find_element(By.ID, 'username').send_keys(username)
        self.safari.find_element(By.ID, 'password').send_keys(password,
                                                              Keys.RETURN)
        sleep(20)

        if test == None:
            return
        else:
            try:
                self.wait_until(test)
            except TimeoutException as e:
                logging.error('Error at Safari_Operations.login()\n', e)
                self.safari.close()
                return
    
    def get_asset_url(self, assets_list, asset_name):
        list = self.safari.find_elements(By.XPATH, assets_list)
        for name in list:
            if asset_name in name.text:
                return name.find_element\
                    (By.XPATH, "./ancestor::a").get_attribute('href')
            
    def get_doc_url(self, account_url, account_updates_widget, ob_au, doc):
        self.safari.get(account_url)
        au_url = self.get_href(self.find_element(account_updates_widget))
        self.safari.get(au_url)
        ob_url = self.get_href(self.find_element(ob_au))
        self.safari.get(ob_url)
        
        return self.get_href(self.find_element(doc))
    
    def google_auth(self):
        sys.stdout = open(os.devnull, 'w')     
        gauth = GoogleAuth()
        gauth.LocalWebserverAuth()
        self.drive = GoogleDrive(gauth)
        sys.stdout = sys.__stdout__

    def download_doc(self, pharmacy_name, id):
        filename = pharmacy_name + '.txt'
        try:
            foldered_list = self.drive.ListFile({'q':  "'" + id\
                                                  + "' in parents and trashed"
                                                  "=false"}).GetList()
            google_file_id = False
            for file in foldered_list:
                if(file['title']==pharmacy_name):
                    google_file_id = file['id']
            if google_file_id == False:
                return google_file_id, filename
              
            google_file = self.drive.CreateFile({'id': google_file_id})
            google_file.GetContentFile(filename)
            return google_file_id, filename
        
        except ApiRequestError:
            google_file_id = None
            os.remove(filename)
            return google_file_id, filename
        
    def get_folder_id(self, pharmacy_name, id):
        file_list = self.drive.ListFile({'q': "'" + id + "' in parents and"
                                         " trashed=false"}).GetList()
        for file1 in file_list:
            if file1['title'] == pharmacy_name:
                return file1['id']
    
    def read_google_doc(self, pharmacy_name):

        filename = pharmacy_name + '.txt'
        with open(filename, "r") as file:
            for line in file:
                if "IP Address:" in line and "Pharmacy" not in line \
                    and "Public" not in line:
                    ip_line = line
                elif "Pharmacy Software Vendor" in line:
                    pms_line = line
                elif "Contact Name:" in line:
                    name = line
                elif "Contact Email:" in line:
                    email = line
        
        #Separate Opie IP out of the line
        try:
            x = ip_line.split(sep=":")
            ip = x[1].split(sep='.')
            y = ip[0] + "." + ip[1] + "." + ip[2] + "." + "250"
            opie_ip = y.lstrip(" ")
        except:
            opie_ip = "DHCP"
        
        #Separate PMS vendor out of the line
        x = pms_line.split(sep=":")
        vendor = x[1].lstrip(" ")
        pms_vendor = vendor.strip("\n")

        #Separate the contact name out of the line
        try:
            temp_name = name.split(sep=":")[1]
            it_name = temp_name.strip("\n")
        except:
            it_name = None

        #Separate the email out of the line
        try:  
            temp_email = email.split(sep=":")[1]
            it_email = temp_email.strip("\n")
        except:
            it_email = ""

        return opie_ip, pms_vendor, it_name, it_email
    
    def update_sheet(self, hostname, mac, ip, pms):
        self.spreadsheet_id = fw_rules_spreadsheet_id #Global Variable
        pms_vendor_string = pms + " Server IPv4 Address"
        scope = ["https://spreadsheets.google.com/feeds",
                 "https://www.googleapis.com/auth/spreadsheets",
                 "https://www.googleapis.com/auth/drive.file",
                 "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name\
            ("credentials.json",scope)
        client = gspread.authorize(creds)
        sheet = client.open_by_key(self.spreadsheet_id).get_worksheet(0)
        sheet.update_cell(row=28, col=2, value=hostname)
        sheet.update_cell(row=30, col=2, value=mac)
        sheet.update_cell(row=31, col=2, value=ip)
        sheet.update_cell(row=33, col=2, value=pms_vendor_string)

    def clean_up(self, filename):
        os.remove(filename)

    def get_asset_info(self, assets_url: str, asset_names: list) -> tuple:
        o = Salesforce_Objects()
        
        #Open the Assets URL
        self.safari.get(assets_url)

        #Search for the asset based on different aliases
        sleep(5)
        for name in asset_names:
            try:
                asset_url = self.get_asset_url(o.x_asset_list, name)
            except:
                continue

        #Open the asset
        self.safari.get(asset_url)

        mac = None    
        try:
            mac = self.get_value(self.find_element(o.x_asset_mac_address))
        except NoSuchElementException:
            pass
        
        hostname = None
        try:
            hostname = self.get_href(self.find_element(o.x_asset_url))
        except NoSuchElementException:
            pass
        except AttributeError:
            pass

        return asset_url, mac, hostname 

    def update_au(self, url: str, date: str, time: str) -> None:
        o = Salesforce_Objects()
        self.safari.get(url)
        self.find_element(o.x_edit_date).send_keys(Keys.ENTER)
        sleep(1)
        input = self.find_element('//input[@name="Install_Date_Time__c"]')
        input.send_keys(date)
        self.action.send_keys(Keys.TAB, Keys.BACKSPACE).perform()
        self.action.send_keys(time, Keys.TAB).perform()
        self.find_element(o.x_au_status).send_keys(Keys.ENTER)
        option = self.find_element(o.x_status_dropdown)
        self.safari.execute_script('arguments[0].click()', option)
        element = self.find_element(o.x_save_button)
        self.safari.execute_script('arguments[0].click()', element)
        return


class Email_Operations:
    def email_fw_rules(self, account_name, customer_email, it_email):
        fromaddr = my_email

        # instance of MIMEMultipart 
        msg = MIMEMultipart() 
        
        # storing the senders email address   
        msg['From'] = fromaddr 
        
        # storing the receivers email address  
        recipients = [customer_email, it_email]
        msg['To'] = ", ".join(recipients)
        
        
        # storing the subject  
        msg['Subject'] = f"Firewall Rules to Implement - {account_name}"
        
        # string to store the body of the mail 
        body = """Hello,
    Please review the attached firewall rules and implement prior to the installation session.

    A DHCP pool is required for our phones and integration devices. The phones will remain DHCP, but we would like to statically assign the On Premise Interface Equipment (OPIE). Typically, the address at .250 is available on the network. We will statically assign the OPIE to .250 unless you have a conflict.

    If you have any questions, please reply to this email or call us 864-541-0650 and ask for the Installation Team.


    Thank you,
    Vow Installation Team
    109 Builders Ct.
    Boiling Springs, SC 29316
    Phone: (864) 541-0650 Ext 444
    Make sure you visit us on the web at https://www.vowinc.com"""
        
        # attach the body with the msg instance 
        msg.attach(MIMEText(body, 'plain')) 
        
        # open the file to be sent  
        filename = fw_rules_pdf_filename
        attachment = open(fw_rules_file_path, "rb") 
        
        # instance of MIMEBase and named as p 
        p = MIMEBase('application', 'octet-stream') 
        
        # To change the payload into encoded form 
        p.set_payload((attachment).read()) 
        
        # encode into base64 
        encoders.encode_base64(p) 
        
        p.add_header('Content-Disposition', "attachment; filename= %s" \
                     % filename) 
        
        # attach the instance 'p' to instance 'msg' 
        msg.attach(p) 
        
        # creates SMTP session 
        s = smtplib.SMTP('smtp.gmail.com', 587) 
        
        # start TLS for security 
        s.starttls() 
        
        # Authentication 
        s.login(fromaddr, email_app_password) 
        
        # Converts the Multipart msg into a string 
        text = msg.as_string() 
        
        # sending the mail 
        s.sendmail(fromaddr, recipients, text) 
        
        # terminating the session 
        s.quit()


class Time_Operations:
    def __init__(self) -> None:
        self.today = date.today()
        self.todays_week = self.today.isocalendar()[1]
        self.todays_year = int(self.today.strftime("%Y"))
        self.next_week = self.todays_week + 1
        self.next_fri = date.fromisocalendar(self.todays_year, 
                                             self.next_week, 5)
        
    def convert_to_time(self, int: int) -> str:
        # Create a datetime object with the given hour
        time = datetime(year=1994, month=4, day=14, hour=int)

        # Format the datetime object as HH:MM:SS
        time_str = time.strftime('%H:%M:%S')

        return time_str
    
    def convert_to_dates(self, days: list) -> list:
        next_week = self.todays_week + 1

        days_key = {"Monday": 1, "Tuesday": 2, "Wednesday": 3, 
            "Thursday": 4, "Friday": 5}

        dates = []
        for name, value in days_key.items():
            if name == days:
                dates.append(date.fromisocalendar(self.todays_year, 
                                next_week, value).strftime('%Y-%m-%d'))
                test = date.fromisocalendar(self.todays_year, 
                                self.todays_week, value)
                
                if self.today < test:
                    dates.append(test.strftime('%Y-%m-%d'))

        return dates

    def convert_to_est(self, time, local_timezone):
        pst_aliases = ["PT", "PDT", "PST"]
        if local_timezone in pst_aliases:
            return self.add_three_hours(time)
        
        cst_aliases = ["CT", "CDT", "CST"]
        if local_timezone in cst_aliases:
            return self.add_one_hour(time)

        est_aliases = ["ET", "EDT", "EST"]
        if local_timezone in est_aliases:
            return time

    def add_three_hours(self, time: str) -> str:
        time = datetime.strptime(time, '%H:%M:%S')
        new_time = time + timedelta(hours=3)
        new_time_str = new_time.strftime('%H:%M:%S')
        
        return new_time_str
    
    def add_one_hour(self, time: str) -> str:
        time = datetime.strptime(time, '%H:%M:%S')
        new_time = time + timedelta(hours=1)
        new_time_str = new_time.strftime('%H:%M:%S')
        
        return new_time_str

    def convert_timezone(self, short_timezone) -> str:
        timezones = {'PDT': "US/Pacific", 'PST': "US/Pacific", 
                     'CDT': "US/Central", 'CST': "US/Central",
                     'EDT': "US/Eastern", 'EST': "US/Eastern",
                     'MDT': "US/Mountain", 'MST': "US/Mountain"}
        return timezones.get(short_timezone)

    def combine_day_time(self, day, time, timezone) -> str:
        #Put the start date and start time together in one string
        date_obj = datetime.strptime(day, '%Y-%m-%d')
        time_obj = datetime.strptime(time, '%H:%M:%S').time()
        tz = pytz.timezone(timezone)
        combined_datetime = datetime.combine(date_obj, time_obj)
        combined_datetime = tz.localize(combined_datetime)
        return combined_datetime.strftime('%Y-%m-%dT%H:%M:%S%z')

    def compare_pref_to_available(self, pref_days: list, pref_hours: list, 
                                  avail_slots: dict) -> dict:
        best_slot = {}
        for day in pref_days: #Each day in customer's preference; find a match.

            if day in avail_slots.keys(): #Hoping to hit a match, otherwise
                                            #we'll schedule first available

                for time in pref_hours: #Customer's preference
                    
                    if time in avail_slots[day]: #Perfect match!
                        best_slot['day'] = day
                        best_slot['time'] = time
                        return best_slot

                    else: #Get within 2 hours of their preference
                        test_time = datetime.strptime(time, '%H:%M:%S')
                        for comp in avail_slots[day]: 
                            slot = datetime.strptime(comp, '%H:%M:%S')
                            n = abs((slot - test_time).total_seconds())
                            if n//3600 < 3 and slot > test_time: #Checks difference
                                best_slot['day'] = day
                                best_slot['time'] = comp
                                return best_slot

    def separate_date_time(self, date_time: str, \
                           timezone='US/Eastern') -> tuple:
        to_separate = datetime.strptime(date_time, "%Y-%m-%dT%H:%M:%S%z")

        # Convert the datetime object to Eastern Time
        convert_timezone = pytz.timezone(timezone)
        eastern_datetime = to_separate.astimezone(convert_timezone)

        # Format the date and time strings for Account Update
        day = eastern_datetime.strftime("%m/%d/%Y")
        time = eastern_datetime.strftime("%I:%M %p")
        return day, time


class Ai_Operations:
    def __init__(self) -> None:
        key = chatgpt_api_key
        self.client = OpenAI(api_key=key)

    def customers_best_days(self, prompt):
        #This function takes the preferences defined during Onboarding and
        #runs it through ChatGPT for a consistent, actionable reply.

        response = self.client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {
                f"role": "system",
                "content": "My input will be a customer's response to which days are best for install.\n\nRead the whole prompt and reply with the days that the customer has specified.\n\nYou are to reply with whole days spelt out and individual days separated by commas.\n\nDo not use any other words.\n\nA range of days will usually be identified by a hyphen. Reply with each day between the range. If it's not specified \"through\", \"thru\", or with a hyphen then reply with individual days.\n\nDo not repeat days."
                },
                { 
                "role": "user", "content": prompt
                }
            ],
            temperature=1,
            max_tokens=256,
            top_p=1,
            frequency_penalty=0,
            presence_penalty=0
            )

        days = []
        response = response.choices[0].message.content.split(sep=",")
        for item in response:
            days.append(item.strip(" "))

        cust_dates = []
        for day in days:
            cust_dates.extend(Time_Operations().convert_to_dates(day))

        cust_dates.sort()

        return cust_dates

    def customers_best_hours(self, prompt):
        #This function takes the preferences defined during Onboarding and
        #runs it through ChatGPT for a consistent, actionable reply.

        response = self.client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {
                f"role": "system",
                "content": "You are to reply with whole numbers separated by commas.\n\nI will feed you in a prompt that contains timeframes customers want to schedule an appointment.\n\nUse 24-hour format. For example, 1pm will be 13, 2pm will be 14, and so on.\n\nDo not use words. Only reply in numbers.\n\nSometimes my prompt will be a range of times. This will be signified by a hyphen. Please reply back with numbers in between the numbers I'm feeding you. For example, you may find \"8-10\" in my prompt. I want you to reply back with the numbers \"8, 9, 10\".\n\nNever reply with numbers less than 7.\n\nNever reply with numbers greater than 18.\n\nTreat these numbers as times, but I do not want \":00\" in your replies."
                },
                {
                "role": "user", "content": prompt
                }
            ],
            temperature=1,
            max_tokens=256,
            top_p=1,
            frequency_penalty=0,
            presence_penalty=0
            )
        
        hours = []
        response = response.choices[0].message.content.split(sep=",")
        for i in response:
            try:
                n = int(i)
                hours.append(Time_Operations().convert_to_time(n))
            except:
                continue

        return hours


class Cal_Operations:
    def __init__(self) -> None:
        self.api_key = cal_com_api_key 
        self.key_str = '?apiKey=' + self.api_key
        self.base_url = 'https://api.cal.com/v1/'
        #Event Types ID
        self.e_t3_network = 740786
        self.e_t2_network = 740772
        self.e_t1_phone = 740750
    
    def get_event_slots(self, event_id: int, start_date: str, 
                        end_date: str, timezone: str) -> list:
        payload = {"eventTypeId": event_id, 
                   "startTime": start_date, 
                   "endTime": end_date,
                   "timeZone": timezone}
        
        r = requests.get(f'{self.base_url}slots{self.key_str}',
                          params=payload)
        
        data = r.json()

        # Create a new dictionary to store cleaned data
        cleaned_data = {}

        # Iterate through the original data and restructure it
        for listed_date, times in data['slots'].items():
            cleaned_times = []
            for time_data in times:
                # Extract time using regular expression
                time_match = time_data['time'].split("T")[1].split("-")[0]
                cleaned_times.append(time_match)

            cleaned_data[listed_date] = cleaned_times

        # Sort the dates in ascending order
        # sorted_dates = sorted(cleaned_data.keys())

        return cleaned_data
        
    def create_payload(self, event_type: int, pharmacy: dict, 
                       customer: dict, date_time: str, timezone: str) -> dict:
        payload = {'eventTypeId': int, #Cal event type as integer
                    'start': str, #Formatted like: 2024-05-14T08:00:00-04:00 
                    'responses': { #Customer's information
                                'name': str, 
                                'email': str, 
                                'City': str,
                                'Attendee': str #Customer's name again
                                    }, 
                    'timeZone': str, #Long-form timezone = 'US/Eastern'
                    'language': 'en', 
                    'metadata': {}
                    }
    
        payload['eventTypeId'] = event_type
        payload['start'] = date_time
        payload['responses']['name'] = customer['name']
        payload['responses']['email'] = customer['email']
        payload['responses']['City'] = pharmacy['city']
        payload['responses']['Attendee'] = customer['name']
        payload['timeZone'] = timezone
        
        return payload

    def schedule_install(self, url: str, payload: dict) -> bool:
        r = requests.post(url, json=payload)
        
        if r.status_code == 200:
            return True
        else:
            return False

    def get_first_available(self, avail_slots: dict) -> str:
        for date, times in avail_slots.items():
            for time in times:
                if time >= '08:00:00':
                    return date, time


def main():
    sfo = Salesforce_Objects()
    saf = Safari_Operations()
    em = Email_Operations()
    cal = Cal_Operations()
    ai = Ai_Operations()
    tim = Time_Operations()

    logging.info("Script has started.")
    
    #The script fills in these variables as it runs
    opie = {"mac": "n/a", "ip": "dhcp", "asset_url": ""}
    pbx = {"hostname": "example.talkrxsky.com", "asset_url": ""}
    pharmacy = {"name": "", "account_url": "", "assets_url": "",
            "ob_au": "", "pms": "", "timezone": "", "google_doc": "", 
            "phone": "", "prompt": ""}
    customer = {"name": "", "email": ""}
    it_contact = {"name": "", "email": ""}

    #Authenticates only once
    saf.google_auth()
    saf.login(website=sfo.website, 
              username=sfo.username, 
              password=sfo.password, 
              test=sfo.header)
    
    logging.info("Authenticated Google and SalesForce successfully.")

    #Opens the dashboard
    saf.safari.get(sfo.dashboard)
    saf.safari.switch_to.frame(saf.find_element(sfo.x_dashboard_iframe))
    
    #Refreshes the dashboard
    saf.action.move_to_element_with_offset\
        (saf.find_element(sfo.x_refresh), -120, 0).click().perform()
    
    #Obtain the list of pharmacies to work on from the Shipped table
    saf.wait_until(sfo.x_shipped_table)
    shipped = saf.find_element(sfo.x_shipped_table)
    entries = shipped.find_elements(By.XPATH, './ancestor::h2'
                                '[@class="defaultHeading"]'
                                '/following-sibling::div//table/tbody/tr')

    for number, entry in enumerate(entries, start=0):
        # Must capture the variables again, because we'll leave this page
        table_x = saf.find_element(sfo.x_shipped_table)
        table = table_x.find_elements(By.XPATH, './ancestor::h2'
                                      '[@class="defaultHeading"]'
                                      '/following-sibling::div'
                                      '//table/tbody/tr')[number]
    
        #Now it starts acting on each entry in that table
        try:
            if table.find_element(By.XPATH, 'td[3]').text == "VOW Full":
                if table.find_element(By.XPATH, 'td[5]').text == "true":
                    pharmacy['fw_rules_req'] = True
                else:
                    pharmacy['fw_rules_req'] = False
                    pharmacy['networking'] = True
                
                pharmacy['ob_au'] = saf.get_href(table.find_element(By.XPATH, 
                                                                'td[1]//a'))
                pharmacy['name'] = table.find_element(By.XPATH, 'td[2]').text
                
                #Open the Onboarding Account Update
                saf.safari.get(pharmacy.get('ob_au'))    
                
                #Snag all the variables out of the AU
                customer['name'] = saf.find_element(sfo.x_customer_name).text
                customer['email'] = saf.find_element(sfo.x_customer_email).text
                pharmacy['google_doc'] = saf.get_href\
                    (saf.find_element(sfo.x_google_doc_link))
                pharmacy['account_url'] = saf.get_href\
                    (saf.find_element(sfo.x_account_tab_in_AU))
                pharmacy['prompt'] = saf.get_value(saf.find_element(\
                                            sfo.x_appt_instructions))
                pharmacy['tier'] = saf.get_value(saf.find_element\
                                                 (sfo.x_install_tier))
                if saf.get_value(saf.find_element(sfo.x_install_date_time)):
                    pharmacy['is_scheduled'] = True
                
                print(pharmacy)

                #Open the Account
                saf.safari.get(pharmacy['account_url'])

                #Snag all the variables out the Account page
                pharmacy['assets_url'] = saf.get_href\
                    (saf.find_element(sfo.x_asset_widget))
                pharmacy['phone'] = saf.get_value(saf.find_element\
                                                  (sfo.x_phone_number))
                local_time = saf.get_value(saf.find_element(sfo.x_local_time))
                pharmacy['timezone'] = local_time.split(sep=" ")[2]
                pharmacy['city'] = saf.find_element(sfo.x_address)\
                                                .text.split(sep=',')[0]


                #If True, send firewall rules
                #If False, skip down to scheduling the install
                if send_firewall_rules is True and \
                    pharmacy['fw_rules_req'] is True:
                    
                    #Prepare opie{}
                    opie_aliases = ["Libre", "Opie"]
                    opie['asset_url'], opie['mac'], opie['hostname'] = \
                        saf.get_asset_info(pharmacy.get('assets_url'), 
                                           opie_aliases)

                    #Prepare pbx{}
                    pbx_aliases = ["PBX", "TalkRx Sky"]
                    pbx['asset_url'], pbx['mac'], pbx['hostname'] = \
                        saf.get_asset_info(pharmacy.get('assets_url'), 
                                           pbx_aliases)

                    #Download the Google Doc
                    folder_id = pharmacy['google_doc'].split(sep='/')[5]
                    d = saf.download_doc(pharmacy.get('name'), folder_id)
                    if d[0] == False:
                        logging.error('Could not find Google Doc for '\
                                      + pharmacy.get('name')\
                                      +f'\n URL:{pharmacy.get("google_doc")}')
                        print("Check logfile.log for errors.")
                        break

                    else:
                        file_id, filename = d

                    #Sometimes the link leads to a folder of folders
                    if file_id == None:
                        id = saf.get_folder_id(pharmacy.get('name'), folder_id)
                        saf.download_doc(pharmacy.get('name'), id)

                    #Open the Google Doc and obtain Opie IP and PMS info
                    results = saf.read_google_doc(pharmacy.get('name'))
                    (opie['ip'], pharmacy['pms'], it_contact['name'], 
                     it_contact['email']) = results
                    
                    #Update firewall spreadsheet
                    saf.update_sheet(pbx.get('hostname'),opie.get('mac'), 
                                     opie.get('ip'), pharmacy.get('pms'))
                    sleep(2)

                    #Delete Google Doc from computer
                    saf.clean_up(filename)

                    #Download Spreadsheet
                    pdf_filename = 'Firewall_Rules.pdf'
                    dl_file = saf.drive.CreateFile({'id': saf.spreadsheet_id})
                    dl_file.GetContentFile(pdf_filename, 
                                           mimetype='application/pdf')
                    
                    if send_fw_email is True:
                        #Send email
                        if use_my_email is True:
                            customer['email'] = my_email
                        em.email_fw_rules(pharmacy.get('name'),
                                    customer_email=customer.get('email'),
                                    it_email=it_contact.get('email'))

                    #Attach firewall rules to AU
                    if attach_pdf is True:
                        saf.safari.get(pharmacy.get('ob_au'))
                        sleep(5)
                        file_path = ("/Users/robby/Documents/MySourceCode"
                                    "/FirewallRules/Firewall_Rules.pdf")
                        input = saf.find_element(sfo.x_upload_files)
                        input.send_keys(file_path)
                        sleep(5)
                        button = saf.find_element(sfo.x_done_button)
                        saf.safari.execute_script('arguments[0].click()', button)

                        #Uncheck "firewall rules required" checkbox
                        saf.safari.get(pharmacy['ob_au'])
                        saf.find_element(sfo.x_firewall_edit).send_keys(Keys.ENTER)
                        sleep(2)
                        element = saf.find_element(sfo.x_firewall_checkbox)
                        saf.safari.execute_script('arguments[0].click()', element)
                        saf.find_element(sfo.x_save_button).send_keys(Keys.ENTER)

                    #Delete the pdf, it's no longer needed.
                    os.remove("/Users/robby/Documents/MySourceCode/FirewallRules"
                            "/Firewall_Rules.pdf") 

                    #Log results
                    with open("firewall_rules_sent.log", "a") as file:
                        payload = (f"{tim.today}: Firewall Rules sent " +\
                        f"successfully to {pharmacy.get('name')}\n")
                        file.write(payload)
                        for x, y in pharmacy.items():
                            file.write(f"{x}: {y}\n")
                        for x, y in opie.items():
                            file.write(f"{x}: {y}\n")
                        for x, y in pbx.items():
                            file.write(f"{x}: {y}\n")
                        for x, y in customer.items():
                            file.write(f"{x}: {y}\n")
                        for x, y in it_contact.items():
                            file.write(f"{x}: {y}\n\n\n")
                    
                    print("Sent firewall rules successfully to " +\
                          pharmacy.get('name'))
                

                else:
                    if schedule_install is True and \
                        not pharmacy.get('is_scheduled'):
                        url = cal.base_url + 'bookings' + cal.key_str
                        timezone = tim.convert_timezone(pharmacy.get('timezone'))
                        days = ai.customers_best_days(pharmacy.get('prompt'))
                        hours = ai.customers_best_hours(pharmacy.get('prompt'))
                        
                        if pharmacy.get('tier') == "Tier III":
                            event_id = cal.e_t3_network

                        elif pharmacy.get('networking') is True\
                            and pharmacy.get('tier') == "Tier II":
                            event_id = cal.e_t2_network
                        
                        else: #Tier 1 & 2 Phone
                            event_id = cal.e_t1_phone

                        available_slots = cal.get_event_slots(\
                                                event_id=event_id, 
                                                start_date=tim.today, 
                                                end_date=tim.next_fri,
                                                timezone=timezone)
                        best_slot = tim.compare_pref_to_available(
                                                days,
                                                hours,
                                                available_slots)
                        
                        if best_slot is None:
                            day,time = cal.get_first_available(available_slots)
                            date_time = tim.combine_day_time(day=day,
                                                            time=time,
                                                            timezone=timezone)
                            
                        else:
                            date_time = tim.combine_day_time(
                                                day=best_slot['day'],
                                                time=best_slot['time'],
                                                timezone=timezone)
                        
                        if use_my_email is True:
                            customer['email'] = my_email

                        print(date_time)
                        payload = cal.create_payload(event_type=event_id,
                                                     pharmacy=pharmacy,
                                                     customer=customer,
                                                     date_time=date_time,
                                                     timezone=timezone)
                        
                        cal.schedule_install(url, payload)

                        print(f'Scheduled {pharmacy.get("name")} '
                              f'for {date_time}')

                        day, time = tim.separate_date_time(date_time)
                        saf.update_au(pharmacy.get('ob_au'), day, time)

                #Return to dashboard for the next pharmacy
                saf.safari.get(sfo.dashboard)
                saf.safari.switch_to.frame(saf.find_element\
                                           (sfo.x_dashboard_iframe))

        except NoSuchElementException:
            #The first two elements of the table are titles.
            #I'm not sure how else to handle this right now.
            continue


if __name__ == "__main__":
    main()