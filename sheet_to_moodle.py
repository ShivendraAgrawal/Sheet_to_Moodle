"""
Author : Shivendra Agrawal
Email : shivendra.agrawal@colorado.edu

Instructions for Google Sheet API at :
https://developers.google.com/sheets/api/quickstart/python?authuser=1
"""

from __future__ import print_function

import json
from pprint import pprint
import httplib2
import os
import time
import pickle
from apiclient import discovery
from selenium import webdriver
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/sheets.googleapis.com-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/spreadsheets.readonly'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Google Sheets API Python Quickstart'

def write_file(content, filename):
    # output = output.strip() + "\n"
    with open(filename, mode='w', encoding='utf-8') as f:
        f.write(content)

def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'sheets.googleapis.com-python-quickstart.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

def prettify_feedbacks(values):
    """
    :param values: list of lists
    :return: dictionary with key as student name value as consolidated feedback
    """
    student_dict = {}
    for student in values[3][3:]:
        student_dict[student.lower()] = ""

    for column in range(3, len(values[0])):
        consolidated_feedback = ""
        for row in range(4, len(values)):
            try:
                a = values[row][column]
            except:
                print(row, column, len(values[row]), values[0][column])
                print(values[row][2] + " "  + "\n")

            if 'Score' in values[row][2] or 'typeset' in values[row][2] or 'credit' in values[row][2]:
                consolidated_feedback += values[row][2] + " "+ values[row][column] +'/'+values[row][0]+ "\n"
            else:
                consolidated_feedback += values[row][2] + " " + values[row][column] + "\n"
        # consolidated_feedback = consolidated_feedback.replace('\n', ' ')
        student_dict[values[3][column].lower()] = consolidated_feedback

    return student_dict


def sheet_to_txt(spreadsheetId, assignment_no, sheet_name):
    '''
    :param spreadsheetId: Google sheet ID
    :param assignment_no: Folder name to put all the .txt feedback in
    :param sheet_name: Google sheet name which has the feedback for a particular assignment
    :return: None (writes .txt files from Google sheet)

    Very basic structure of the google sheet with feedbacks
    https://docs.google.com/spreadsheets/d/1too-eTt-t79C1ZDYy6nJkaUxSy6fs6-eh_OM1vMQjhY/edit?usp=sharing
    '''

    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service = discovery.build('sheets', 'v4', http=http,
                              discoveryServiceUrl=discoveryUrl)

    rangeName = sheet_name + '!A:KV'
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheetId, range=rangeName).execute()
    values = result.get('values', [])

    if not os.path.exists(assignment_no):
        os.makedirs(assignment_no)

    if not values:
        print('No data found.')
    else:
        consolidated_feedbacks = prettify_feedbacks(values)
        count = 0
        for student in values[3][3:]:
            count += 1
            print('Writing feedback for : ' + student)
            # print("Filename : " + student.replace(".", "_") + ".txt")
            # print("File content : \n" + consolidated_feedbacks[student.lower()])
            write_file(consolidated_feedbacks[student.lower()], os.path.join(assignment_no, student.replace(".", "_")+".txt"))
        print("Total feed backs written = " + str(count))


def setup_browser(moodle_username, moodle_password):
    options = webdriver.ChromeOptions()
    # options.add_argument('--ignore-certificate-errors')
    # options.add_argument("--test-type")
    driver = webdriver.Chrome("./chromedriver")
    driver.get('https://moodle.cs.colorado.edu/login/index.php')

    driver.find_element_by_id("username").send_keys(moodle_username)
    driver.find_element_by_id("password").send_keys(moodle_password)

    driver.find_element_by_name("_eventId_proceed").click()
    return driver


def delete_feedback(driver, base_url, id_lookup, notify=False):
    with open("delete_feedback_for.txt", mode='r', encoding='utf-8') as f:
        delete_for = f.read().split('\n')

    for email in delete_for:
        try:
            moodle_id = id_lookup[email.lower()]
        except:
            print("Moodle ID lookup failed for : " + email)
            continue

        driver.get(base_url + moodle_id)

        print('Deleting feedback for : ' + email)
        if notify is False:
            while True:
                try:
                    driver.find_element_by_name('sendstudentnotifications').click()
                    break
                except:
                    print("Delete - Waiting for JS (1) calls to render completely!!")
                    time.sleep(2)
        while True:
            try:
                driver.find_element_by_class_name('fp-file').click()
                break
            except:
                print("Delete - Waiting for JS (2) calls to render completely!!")
                time.sleep(2)

        while True:
            try:
                driver.find_element_by_class_name('fp-file-delete').click()
                break
            except:
                print("Delete - Waiting for JS (3) calls to render completely!!")
                time.sleep(2)

        while True:
            try:
                driver.find_element_by_class_name('fp-dlg-butconfirm').click()
                break
            except:
                print("Delete - Waiting for JS (4) calls to render completely!!")
                time.sleep(2)

        while True:
            try:
                driver.find_element_by_name("savechanges").click()
                break
            except:
                print("Delete - Waiting for JS (5) calls to render completely!!")
                time.sleep(2)
        driver.maximize_window()  # For maximizing window
        while True:
            try:
                driver.find_element_by_css_selector("input[value='Ok']").click()
                break
            except:
                print("Delete - Waiting for JS (6) calls to render completely!!")
                time.sleep(2)


def is_already_uploaded(driver):
    try:
        time.sleep(3)
        driver.find_element_by_class_name('fp-file')
        return True
    except:
        return False


def upload_on_moodle(driver, id_lookup, folder, base_url):
    '''
    :param driver: Selenium driver with Moodle signed in
    :param id_lookup: Email to Moodle ID map
    :param folder: Folder to pick .txt feedback from and upload to Moodle
    :param base_url: Assignment Base URL
    :return: None
    '''
    directory = os.fsencode(folder)
    print("*" * 60 + "\n")

    with open("already_uploaded.txt", mode='r', encoding='utf-8') as f:
        already_uploaded = f.read().split('\n')
        # already_uploaded = []
    with open("upload_for.txt", mode='r', encoding='utf-8') as f:
        upload_for = f.read().split('\n')
    with open("do_not_upload_for.txt", mode='r', encoding='utf-8') as f:
        do_not_upload_for = f.read().split('\n')


    count = 0

    print(os.listdir(directory))

    for file in os.listdir(directory):
        filename = os.fsdecode(file)
        print(filename)

        email = filename.split('.')[0].replace('_', '.').lower()
        if filename.endswith(".txt") and email not in do_not_upload_for:

            print('Uploading feedback for : ' + filename)

            try:
                moodle_id = id_lookup[email.lower()]
            except:
                print("Moodle ID lookup failed for : " + email)
                continue


            driver.get(base_url + moodle_id)

            ##### IMPORTANT #######
            # To be used if we lost track of whom we already uploaded for
            # if is_already_uploaded(driver):
            #     continue

            while True:
                try:
                    driver.find_element_by_class_name("dndupload-arrow").click()
                    break
                except:
                    print("Waiting for JS calls (1) to render completely!!")
                    time.sleep(2)

            driver.maximize_window()  # For maximizing window

            while True:
                try:
                    driver.find_element_by_name("repo_upload_file").send_keys(
                        os.path.abspath(os.path.join(folder, filename)))
                    break
                except:
                    print("Waiting for JS calls (2) to render completely!!")
                    time.sleep(2)

            while True:
                try:
                    driver.find_element_by_class_name("fp-upload-btn").click()
                    break
                except:
                    print("Waiting for JS (3) calls to render completely!!")
                    time.sleep(2)

            while True:
                try:
                    driver.find_element_by_name("savechanges").click()
                    count += 1
                    break
                except:
                    print("Waiting for JS (4) calls to render completely!!")
                    time.sleep(2)
            driver.maximize_window()  # For maximizing window
            while True:
                try:
                    driver.find_element_by_css_selector("input[value='Ok']").click()
                    break
                except:
                    print("Waiting for JS (5) calls to render completely!!")
                    time.sleep(2)

    print("Total feed backs uploaded = " + str(count))

def email_to_moodleID():
    '''
    :return: A email (lower case and present in class roster) to Moodle ID dict
    '''
    with open('id_lookup.pickle', 'rb') as f:
        id_lookup = pickle.load(f)

        # Ugly way to append new email to Moodle ID mapping on the fly
        # Students keep on adding or leaving the course and
        # we need to update our Emails to Moodle ID dict somehow
        id_lookup.update({'first.last@colorado.edu' : '1234',
                            'first1.last1@colorado.edu' : '2345',
                            })
    return id_lookup

def generalize_url(example_url):
    return example_url[: example_url.find('userid=') + 7 ]

if __name__ == '__main__':
    # Google sheet ID which would have a unique tab for each assignment
    spreadsheetId = '1qIQy7Y05FDW-KUpA4Y4Jgli-ppeDtUxLmEeGeXBjhjY'

    # Provide folder name as 'assignment_no' and
    # an example url for the assignment feedback as 'example_url'
    assignment_no = 'Assignment 10'
    example_url = 'https://moodle.cs.colorado.edu/mod/assign/view.php?id=20717&rownum=0&action=grader&userid=700'

    # Writes consolidated feed backs to disk as .txt files
    sheet_to_txt(spreadsheetId, assignment_no, 'PS10')

    # Get Moodle ID from emails (email to id dict)
    id_lookup = email_to_moodleID()
    # pprint(id_lookup)

    # Logs into Moodle and create the driver object for selenium to control
    moodle_credentials = json.load(open('moodle_credentials.json'))
    driver = setup_browser(moodle_credentials["username"],
                           moodle_credentials["password"])

    # Upload the above generated .txt files to respective student
    # and assignment section on Moodle
    upload_on_moodle(driver, id_lookup, assignment_no, generalize_url(example_url))

    #### To be used if we want to delete already uploaded feedback
    # # # delete_feedback(driver, generalize_url(example_url), id_lookup)