from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.edge.service import Service
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
import os
import time
import datetime
import pandas as pd
import win32com.client as win32

##############################
# Purpose
# To cleanup the version history of the files stored in onedrive to save space from the drive
# 
# Note: Please manually input the root directory for version cleanup
#
# Requirement: Latest version of msedgedriver, 
# You can download it here: https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/?form=MA13LH
##############################

##############################
# Global Variables
##############################
staff_id = os.getlogin()
cUser = os.path.expanduser('~')
current_directory = os.path.dirname(os.path.realpath(__file__))
previous_directory = os.path.abspath(os.path.join(current_directory, os.pardir))
temp_directory = "C:\\temp"

now = datetime.datetime.now()
today = datetime.datetime.today()
now_string = now.strftime("%Y%m%d%H%M%S")
today_string = today.strftime("%Y%m%d")

driver_path = os.path.join(current_directory, "msedgedriver.exe")

###############################
# Functions
###############################
def delete_version_history(driver, url_path):
    current_url_path = driver.current_url

    # Click the delete all versions button
    driver.get(url_path)
    try:
        delete_version_button_element = driver.find_element(By.XPATH, '//a[@id="ctl00_PlaceHolderMain_MngVersionToolBar_RptControls_diidDeleteVersions_LinkText"]')
        delete_version_button_element.click()
        time.sleep(0.2)
        alert = driver.switch_to.alert
        alert.accept()
        time.sleep(0.3)
    except Exception as e:
        print(e)
        print("Exception do not cause a fatal error. Continue.")

    # Return to the previous directory
    driver.get(current_url_path)

def scan_directory(driver, url_path, return_url_path, scan_inner_directory_bool = True):
    directory_tensor = []
    
    # Scan the target directory
    driver.get(url_path)
    try:
        for i in range(1, 10000):
            # Scan page attributes
            is_target_file = False
            try:
                folder_name_element = driver.find_element(By.XPATH, '//table[@id="onetidUserRptrTable"]/tbody[1]/tr[' + str(1 + i) + ']/td[2]/a[1]')
                is_target_file = False
            except NoSuchElementException as e:
                file_name_element = driver.find_element(By.XPATH, '//table[@id="onetidUserRptrTable"]/tbody[1]/tr[' + str(1 + i) + ']/td[2]')
                try:
                    version_button_element = driver.find_element(By.XPATH, '//table[@id="onetidUserRptrTable"]/tbody[1]/tr[' + str(1 + i) + ']/td[9]/span[1]/a[1]')
                except Exception as e:
                    print(e)
                is_target_file = True
            total_size_element = driver.find_element(By.XPATH, '//table[@id="onetidUserRptrTable"]/tbody[1]/tr[' + str(1 + i) + ']/td[3]/span[1]')
            percentage_parent_element = driver.find_element(By.XPATH, '//table[@id="onetidUserRptrTable"]/tbody[1]/tr[' + str(1 + i) + ']/td[4]/span[1]')
            percentage_site_quota_element = driver.find_element(By.XPATH, '//table[@id="onetidUserRptrTable"]/tbody[1]/tr[' + str(1 + i) + ']/td[6]/span[1]')
            last_modified_date_element = driver.find_element(By.XPATH, '//table[@id="onetidUserRptrTable"]/tbody[1]/tr[' + str(1 + i) + ']/td[8]/span[1]')
            # Operations
            if not is_target_file:
                if scan_inner_directory_bool:
                    directory_link = folder_name_element.get_attribute('href')
                    inner_directory_tensor = scan_directory(driver, directory_link, url_path)
                    directory_tensor.append(inner_directory_tensor)
            else:
                try:
                    version_button_link = version_button_element.get_attribute('href')
                    print("Deleting the version history for: " + version_button_link)
                    directory_tensor.append(file_name_element.text)
                    delete_version_history(driver, version_button_link)
                except Exception as e:
                    print(e)
                    print("Exception do not cause a fatal error. Continue.")
    except NoSuchElementException:
        print("Reach the end at line: " + str(i))
    
    # Check "Next" button
    try:
        try:
            previous_button_element = driver.find_element(By.XPATH, '/html/body/form[1]/div[12]/div[1]/div[2]/div[2]/div[3]/table[2]/tbody[1]/tr[1]/td[1]/a[1]')
            next_button_element = driver.find_element(By.XPATH, '/html/body/form[1]/div[12]/div[1]/div[2]/div[2]/div[3]/table[2]/tbody[1]/tr[1]/td[1]/a[2]')
        except NoSuchElementException:
            next_button_element = driver.find_element(By.XPATH, '/html/body/form[1]/div[12]/div[1]/div[2]/div[2]/div[3]/table[2]/tbody[1]/tr[1]/td[1]/a[1]')
        if "Next" in next_button_element.text:
            next_directory_tensor = scan_directory(driver, next_button_element.get_attribute('href'), return_url_path)
            directory_tensor = directory_tensor + next_directory_tensor
            return directory_tensor
    except NoSuchElementException as e:
        print("Complete scanning the directory: " + url_path)
    
    # Return to the previous directory
    driver.get(return_url_path)
    time.sleep(0.5)
    return directory_tensor

###############################
# Main Process
###############################
# Input Root Directory
while True:
    print("Input the root URL path in Onedrive for version history cleanup: ")
    root_directory_input = input().rstrip()
    if root_directory_input != "":
        root_directory_url_path = root_directory_input
        break
    else:
        print("Make sure you input the root URL. Try again.")

# Input Root Directory
while True:
    scan_inner_directory_input = input("Include scanning inner directory (Y/N)? ").rstrip()
    print(scan_inner_directory_input)
    if scan_inner_directory_input == "Y":
        scan_inner_directory_bool = True
        break
    elif scan_inner_directory_input == "N":
        scan_inner_directory_bool = False
        break
    else:
        print("Make sure you input Y or N. Try again.")

try:
    # Get Driver
    try:
        driver = webdriver.Edge()
    except Exception as e:
        driver = webdriver.Edge(service=Service(driver_path))
    driver.maximize_window()

    # Process the deletion
    driver.get(root_directory_url_path)
    directory_tensor = scan_directory(driver, root_directory_url_path, driver.current_url, scan_inner_directory_bool)
    print("The directory structure is: ")
    print(directory_tensor)
except Exception as e:
    now_string_formatted = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print("Sending Email - Clear Version History in Storage Metrices Failed (" + now_string_formatted + ")")
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'chris.cm.lam@clp.com.hk;'
    mail.Subject = 'Clear Version History in Storage Metrices Failed (' + now_string_formatted + ')'
    mail.HTMLBody = '<html><body>' + \
                    '<p>Hi,<br><br>Encountered failure to clear version history at: ' + root_directory_url_path + '</p>' + \
                    '<p>Error Reason: ' + str(e) + '</p>' + \
                    '<p>Regards,<br>SMP Team</p>' + \
                    '</body></html>'
    mail.Send()