# Export current patients from WriteUpp and Import into Mailchimp replacing current audience
# V1.1win  Bec 15/12/22 contains path for Bec's computer

import csv
import os
import pymsgbox
import time
import ast
from configparser import ConfigParser
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.support.select import Select
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException


version_no = "v1.1win 15/12/2022"
wd = 'C:\\Billing\\Mailchimp'
downloadDirectory = wd
config_file = 'mailchimp.ini'

# Read the config file
os.chdir(wd)
config = ConfigParser()
config.read(config_file)
wu_loginURL = config.get('writeupp', 'URL')
driverpath = config.get('writeupp', 'driver_path')
wu_username = config.get('writeupp', 'user')
wu_password = config.get('writeupp', 'password')
wu_days = int(config.get('writeupp', 'days'))
writeUpp_open_URL = config.get('writeupp', 'open_url')
wu_export_fn  = config.get('writeupp', 'open_fn')

mc_loginURL = config.get('mailchimp', 'URL')
mc_username = config.get('mailchimp', 'mc_user')
mc_password = config.get('mailchimp', 'mc_password')
mc_import_file = config.get('mailchimp', 'import_fn')
mc_audience_URL = config.get('mailchimp', 'audience_URL')

today_date = datetime.now().strftime('%d/%m/%Y')
full_date = datetime.now().strftime('%d %B %Y')
past_date1 = datetime.now() - timedelta(days = wu_days)
from_date = past_date1.strftime('%d/%m/%Y')


def loginWriteupp():
    # Login to writeUpp
    driver.get(wu_loginURL)
    time.sleep(2)
    userNameField = driver.find_element_by_id('EmailAddress')
    userNameField.send_keys(wu_username)
    passwordField = driver.find_element_by_id('Password')
    passwordField.send_keys(wu_password)
    time.sleep(1)
    submitButton = driver.find_element_by_xpath('/html/body/div[2]/main/div/div[2]/div/form/div[3]/div/div/button')
    submitButton.click()
    time.sleep(3)
    return

def download_opencases():
    # get rid of any old files
    dir_name = wd
    old_csv_files = os.listdir(dir_name)
    for item in old_csv_files:
        if item.startswith(wu_export_fn) or item.startswith(mc_import_file):
            os.remove(os.path.join(dir_name, item))
    # get the latest open caseload
    driver.get(writeUpp_open_URL)
    time.sleep(2)
    driver.find_element_by_id('ctl00_ctl00_Content_ContentPlaceHolder1_btnExportCsv').click()
    time.sleep(5)
    return

def create_mailchimp_file():
    mailchimp_list = []
    mailchimp_list.append(['FNAME','Email'])
    with open(wu_export_fn, newline='') as f:
        patients = csv.reader(f)
        next(patients)  # skip header
        for patient in patients:
            patient_first_name = patient[3].split(' ')[0]
            patient_email = patient[5]
            last_appointment_date = patient[7]
            if last_appointment_date == "":
                print('Ignored ' + patient_first_name + " no date")
                continue
            last_date_obj = datetime.strptime(last_appointment_date, '%d/%m/%Y')
            if (patient_email == "") or (last_date_obj < past_date1):
                print('Ignored '+ patient_first_name + " " + last_appointment_date)
                continue # skip this on if there is no email address
            this_item = [patient_first_name,patient_email]
            mailchimp_list.append(this_item)
        f.close()
    # Write the mailchimp import file
    a = open(mc_import_file, 'w', newline='')
    writer = csv.writer(a)
    for patient_line in mailchimp_list:
        writer.writerow(patient_line)
    a.close()
    return

def login_mailchimp():
    driver.get(mc_loginURL)
    # time.sleep(4)
    try:
        userNameField = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, 'username')))
        # userNameField = driver.find_element_by_id('username')
        userNameField.send_keys(mc_username)
        passwordField = driver.find_element_by_id('password')
        passwordField.send_keys(mc_password)
    except TimeoutException:
        pymsgbox.alert('Timeout in line 117 - contact Justin')
        time.sleep(1)
    # driver.find_element_by_id('submit-btn').click()
    pymsgbox.alert('Complete 2FA and then click OK')
    return

def archive_contacts():
    driver.get('https://us3.admin.mailchimp.com/lists/members/archive-all?id=629937')
    driver.switch_to.frame(driver.find_element_by_id('fallback'))
    archive_field = driver.find_element_by_xpath('/html/body/div[2]/div[3]/main/div[3]/div/div/div/div/form/div[2]/div/input')
    time.sleep(1)
    archive_field.send_keys('ARCHIVE')
    time.sleep(1)
    driver.find_element_by_id('dijit_form_Button_0_label').click() # enable before you go live
    time.sleep(4)
    return

def upload_contacts():
    driver.get('https://us3.admin.mailchimp.com/lists/members/import?id=629937')
    # driver.switch_to.default_content()
    time.sleep(4)
    try:
        upload_link = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[3]/div/main/div/div/div/div/main/div/form/fieldset/div/div/div[2]/label/input')))
        upload_link.click()
        time.sleep(1)
    except TimeoutException:
        pymsgbox.alert('Timeout in line 166 - contact Justin')
        time.sleep(1)

    driver.find_element_by_xpath('/html/body/div[3]/div/main/div/div/div/div/main/div/form/div/button').click()
    time.sleep(2)
    browse_button = driver.find_element_by_xpath('/html/body/div[3]/div/main/div/div/div/div/main/div/form/div/div')
    browse_button.click()

    # The following two lines work but leave an open dialoge box
    file_input = driver.switch_to.active_element
    full_file_name = wd + "\\" + mc_import_file
    file_input.send_keys(full_file_name)
    time.sleep(4)
    pymsgbox.alert('Click Cancel  then click OK') # get rid of windows box
    driver.find_element_by_xpath('/html/body/div[3]/div/main/div/div/div/div/main/div/form/div/button').click()
    continue_to_tag = driver.find_element_by_xpath('/html/body/div[3]/div/main/div/div/div/div/main/div/form/div/div[2]/button')
    continue_to_tag.click()
    time.sleep(1)
    tag_input = driver.find_element_by_xpath('/html/body/div[3]/div/main/div/div/div/div/main/div/form/div/div/div[1]/div[3]/input')
    tag_input.send_keys(full_date)
    tag_input.send_keys(Keys.RETURN)
    time.sleep(1)
    continune_to_match = driver.find_element_by_xpath('/html/body/div[3]/div/main/div/div/div/div/main/div/form/div/button')
    continune_to_match.click()
    time.sleep(1)
    finalize_import = driver.find_element_by_xpath('/html/body/div[3]/div/main/div/div/div/div/main/div/div[2]/button')
    finalize_import.click()
    time.sleep(1)
    complete_import = driver.find_element_by_xpath('/html/body/div[3]/div/main/div/div/div/div/main/div/form/div/div/div/button')
    complete_import.click()
    return

firefox_options = Options()
firefox_options.binary_location = "C:/Users/BeccaThomas/AppData/Local/Mozilla Firefox/firefox.exe"
firefox_options.add_argument("--disable-infobars")
firefox_options.add_argument("--disable-extensions")
firefox_options.add_argument("--disable-popup-blocking")


profile = webdriver.FirefoxProfile()  # should I get rid of webdriver?
profile.set_preference('browser.download.folderList', 2)
profile.set_preference('browser.download.manager.showWhenStarting', False)
profile.set_preference('browser.download.dir', downloadDirectory)
profile.set_preference('browser.helperApps.neverAsk.saveToDisk', 'text/csv')
profile.set_preference('browser.download.alwaysOpenPanel', False)
profile.set_preference('layout.css.devPixelsPerPx', '0.8')

driver = webdriver.Firefox(executable_path=driverpath, firefox_profile=profile)
driver.set_window_size(1055,1300)
driver.implicitly_wait(3)

# Code starts here
loginWriteupp()
download_opencases()
create_mailchimp_file()
print('Export done')
login_mailchimp()
archive_contacts()
upload_contacts()
pymsgbox.alert('All done - press OK to close the browser')
driver.quit()





