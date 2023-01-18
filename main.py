import time

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from time import sleep
from threading import Thread
from pynput.keyboard import Key, Controller
import os
import csv
from GSheet_class import GSheet_finance

''' This program has been written by Martin Carrier, ing., P. Eng. on January 6th, 2023 '''

# functions
def load_rpp(driver_var):
    # driver.get('https://rpp.tsl.telus.com/capitalmanagement/worksheet/')  # This load the webpage. There is a pop up for certificate and credentials
    driver_var.get('https://rpp2.tsl.telus.com/capitalmanagement/worksheet/')  # This load the webpage. There is a pop up for certificate and credentials


def send_keys(keyboard_var):
    # deal with certificate popup window -> hit return key
    keyboard_var.press(Key.enter)
    keyboard_var.release(Key.enter)
    sleep(1)
    keyboard_var.type('Your NIP')
    keyboard_var.press(Key.enter)
    keyboard_var.release(Key.enter)


def handle_rpp(driver_var, ii_var, profile_var):  # function that handle RPP page
    # driver_var.find_element_by_name('flt.keyword').send_keys(ii_var)  # insert II's
    driver_var.find_element(By.NAME, 'flt.keyword').send_keys(ii_var)  # insert II's
    # drop_down = driver_var.find_element_by_tag_name('b')  # find profile drop down button
    drop_down = driver_var.find_element(By.TAG_NAME,
                                        'b')  # find profile drop down button
    drop_down.click()  # click on the drop down button
    # print('Click on dropdown button')
    # sleep(1)
    # choice = driver_var.find_element_by_class_name('chosen-search-input')  # search field in the dropdown menu
    choice = driver_var.find_element(By.CLASS_NAME,
                                     'chosen-search-input')  # search field in the dropdown menu
    choice.send_keys(profile_var)  # push profile
    choice.send_keys(Keys.RETURN)  # hit enter
    # print('Choose actual profile')
    # submit = driver_var.find_element_by_xpath('/html/body/div[5]/form/div[2]/div[10]/input[1]')  # find Submit button
    submit = driver_var.find_element(By.XPATH,
                                     '/html/body/div[5]/form/div[2]/div[10]/input[1]')  # find Submit button
    submit.click()
    # print('Click on LOAD button')
    sleep(2)
    # export = driver_var.find_element_by_xpath('//*[@id="btnExport"]')  # find export button
    export = driver_var.find_element(By.XPATH,
                                     '//*[@id="btnExport"]')  # find export button
    export.click()
    # print('Click on EXPORT button')
    sleep(2)
    # print('Closing chromedriver')
    # driver_var.close()
    # print('Chromedriver is closed')


def delete_file(file):
    if os.path.exists(file):
        os.remove(file)
    else:
        print("The file doesn't not exist")


# Variables
driver = webdriver.Chrome()
envelopes = 'MM16.01.MNT0J,MM16.02.MNT0L'
profile = 'Actuals'
keyboard = Controller()  # variable to push keystrokes
download_file = 'C:\\Users\\t894031\\Downloads\\RPP Capital Management Compare Export.csv'
projects = {}  # dictionary of project codes
monthly_actual = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
date_loc = 'Finance Report!D2'
nb_of_proj_loc = 'Finance Report!D5'


start = time.time()
# load page
print('\nProgram starts now...', '{:.3f}'.format(time.time() - start), 's.')
thread_1 = Thread(target=load_rpp,
                  args=[driver])  # assign page loading to a trhead class
thread_1.start()  # start the thread and move on with next steps
sleep(2)
print('\nRPP has opened in a thread.', '{:.3f}'.format(time.time() - start), 's.')
send_keys(keyboard)  # send keystrokes
print('\nKeystrokes have been sent for credentials.', '{:.3f}'.format(time.time() - start), 's.')
sleep(5)  # wait for 5 seconds
# print('\nPush selenium commands to RPP page and export CSV file to "download folder"...', '{:.2f}'.format(time.time() - start), 's')
handle_rpp(driver,
           envelopes,
           profile)
# print('Sleep for 5 seconds...')
# sleep(5)
print('\nSelenium commands completed.', '{:.3f}'.format(time.time() - start), 's.')
with open(download_file,
          newline='') as csvfile:
    reader = csv.DictReader(csvfile)
    # print(reader.fieldnames)
    ''' reader.fieldnames contains headers.
    PCode = 4, Cost Element = 6, Jan = 7, Feb = 8, Mar = 9, Apr = 10, May = 11, 'Jun = 12, Jul = 13, Aug = 14, 
    Sep = 15, Oct = 16, Nov = 17, Dec = 18, Commitments = 19, Total Actuals = 24
    '''

    for row in reader:
        '''# print(row[reader.fieldnames[7]], row[reader.fieldnames[19]]) '''
        if row[reader.fieldnames[6]] == 'Total':
            projects[row[reader.fieldnames[4]]] = [row[reader.fieldnames[24]]], [row[reader.fieldnames[19]]]
        elif row[reader.fieldnames[6]] == 'Grand Total':
            for i in range(12):
                monthly_actual[i] = row[reader.fieldnames[i + 7]]
        else:
            continue

    # print(projects)
    # print(monthly_actual)

print('\nCSV file parsing completed.', '{:.3f}'.format(time.time() - start), 's.')
google_sheet_update = GSheet_finance(projects, monthly_actual)
print('\nGoogle Sheet class instantiation completed.', '{:.3f}'.format(time.time() - start), 's.\n')
google_sheet_update.disconnect_vpn()
# sleep(10)  # just for demo
sleep(2)
print('\nVPN is disconnected.', '{:.3f}'.format(time.time() - start), 's.\n')
google_sheet_update.gsheet_authenticatication()
print('\nAuthorization and authentication completed.', '{:.3f}'.format(time.time() - start), 's.')
google_sheet_update.nb_of_projects = int(google_sheet_update.get_cell_value(nb_of_proj_loc))
google_sheet_update.update_projects()
print('\nGoogle Sheet projects are updated.', '{:.3f}'.format(time.time() - start), 's.')
google_sheet_update.update_eqx()
print('\nGoogle Sheet EQX is updated.', '{:.3f}'.format(time.time() - start), 's.')
google_sheet_update.update_cell(date_loc, google_sheet_update.today)  # update date to today
print('\nGoogle Sheet date is updated.', '{:.3f}'.format(time.time() - start), 's.')
missing_projects = google_sheet_update.validate_missing_project()
if len(missing_projects) > 0:
    print('\nYou miss these project(s) in the GSheet: ')
    for project in missing_projects:
        print('\t - ' + project)
else:
    print('\nNo error found! All good!', '{:.3f}'.format(time.time() - start), 's.')
delete_file(download_file)
# print('\nCSV file from "download folder" is deleted.', '{:.3f}'.format(time.time() - start), 's.')
stop = time.time()
print('\nFinance report update is completed! This automation took', '{:.3f}'.format(stop - start), 'seconds to run.')


