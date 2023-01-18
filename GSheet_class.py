from datetime import datetime
import subprocess
import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


class GSheet_finance():
    def __init__(self, project_list, actual_by_month):  # constructor
        self.scope = ['https://www.googleapis.com/auth/spreadsheets']  # If modifying these scopes, delete the file token.json.
        # self.spreadsheet_id = '1ha_RJjB7cKpbBKIWQlfW4rTHE7c6mtDChOyTd83hzAk'  # ID of the Google Sheet Finance Report
        self.spreadsheet_id = '1ha_RJjB7cKpbBKIWQlfW4rTHE7c6mtDChOyTd83hzAk'  # ID of the Google Sheet Finance Report
        self.today = datetime.now().strftime('%Y-%m-%d  (%H:%M ET)')
        # self.today = datetime.today().strftime('%Y-%m-%d')
        self.project_list = project_list
        self.actual_by_month = actual_by_month
        self.creds = None
        self.sheet = None
        self.project_list_loc_prefix = 'Finance Report!C'
        self.actual_list_loc_prefix = 'Finance Report!H'
        self.commit_list_loc_prefix = 'Finance Report!I'
        self.first_row = 10
        self.nb_of_projects = 0
        self.eqx_list_loc_prefix = 'Finance Report!N'
        self.eqx_actual_loc_prefix = 'Finance Report!P'
        self.accrual_loc = 'Finance Report!M5'
        self.process_project = []


    def disconnect_vpn(self):
        cli_name = 'C:\\Program Files (x86)\\Cisco\\Cisco AnyConnect Secure Mobility Client\\vpncli.exe'
        command = 'disconnect'
        subprocess.Popen([cli_name, command])


    def gsheet_authenticatication(self):
        self.creds = None
        # The file token.json stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists('token.json'):
            creds = Credentials.from_authorized_user_file('token.json', self.scope)
        # If there are no (valid) credentials available, let the user log in.
        if not self.creds or not self.creds.valid:
            if self.creds and self.creds.expired and self.creds.refresh_token:
                self.creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file('credentials.json', self.scope)
                self.creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open('token.json', 'w') as token:
                token.write(self.creds.to_json())

        service = build('sheets', 'v4', credentials=self.creds)
        self.sheet = service.spreadsheets()


    def update_cell(self, cell_addr, value):
        values_var = [[str(value)]]
        body = {'values': values_var}
        self.sheet.values().update(spreadsheetId=self.spreadsheet_id,
                              range=cell_addr,
                              valueInputOption='USER_ENTERED',
                              body=body).execute()


    def get_cell_value(self, cell_addr):
        value = self.sheet.values().get(spreadsheetId=self.spreadsheet_id,
                                    range=cell_addr).execute()
        return value['values'][0][0]


    def update_projects(self):
        for i in range(self.first_row,self.first_row + self.nb_of_projects):
            project_id = self.get_cell_value(self.project_list_loc_prefix + str(i))
            if project_id in self.project_list.keys():  # if the project is part of RPP list
                if self.project_list[project_id][0][0] != '0':  # if actuals is different than 0$
                    self.update_cell(self.actual_list_loc_prefix + str(i), self.project_list[project_id][0][0])
                else:
                    self.update_cell(self.actual_list_loc_prefix + str(i),'')

                if self.project_list[project_id][1][0] != '0':  # if commitments is different than 0$
                    self.update_cell(self.commit_list_loc_prefix + str(i), self.project_list[project_id][1][0])
                else:
                    self.update_cell(self.commit_list_loc_prefix + str(i), '')
                self.process_project.append(project_id)
            else:
                self.update_cell(self.actual_list_loc_prefix + str(i), '')
                self.update_cell(self.commit_list_loc_prefix + str(i), '')


    def update_eqx(self):
        # accruals = self.get_cell_value(self.accrual_loc)
        for i in range(3, 6):
            month = self.get_cell_value(self.eqx_list_loc_prefix + str(i))  # figure out which month to update
            actual_that_month = self.actual_by_month[self.get_month_index(month)]
            if actual_that_month == '':
                actual_that_month = 0
            if self.is_actual_month(month): # if it's actual month, add accruals
                # self.update_cell(self.eqx_actual_loc_prefix + str(i), float(str(accruals).replace(',', '')) + float(actual_that_month))
                self.update_cell(self.eqx_actual_loc_prefix + str(i), '=M5+' + actual_that_month)
            else:  # if else: just update actuals
                self.update_cell(self.eqx_actual_loc_prefix + str(i), actual_that_month)


    def is_actual_month(self, month_str):  # return true if this is the actual month
        actual_month = datetime.now().month
        # actual_month = 8  # just for QA
        if actual_month == 1 and month_str == 'January':
            return True
        elif actual_month == 2 and month_str == 'February':
            return True
        elif actual_month == 3 and month_str == 'March':
            return True
        elif actual_month == 4 and month_str == 'April':
            return True
        elif actual_month == 5 and month_str == 'May':
            return True
        elif actual_month == 6 and month_str == 'June':
            return True
        elif actual_month == 7 and month_str == 'July':
            return True
        elif actual_month == 8 and month_str == 'August':
            return True
        elif actual_month == 9 and month_str == 'September':
            return True
        elif actual_month == 10 and month_str == 'October':
            return True
        elif actual_month == 11 and month_str == 'November':
            return True
        elif actual_month == 12 and month_str == 'December':
            return True
        else:
            return False


    def get_month_index(self, month_str):
        if month_str == 'January':
            return 0
        elif month_str == 'February':
            return 1
        elif month_str == 'March':
            return 2
        elif month_str == 'April':
            return 3
        elif month_str == 'May':
            return 4
        elif month_str == 'June':
            return 5
        elif month_str == 'July':
            return 6
        elif month_str == 'August':
            return 7
        elif month_str == 'September':
            return 8
        elif month_str == 'October':
            return 9
        elif month_str == 'November':
            return 10
        elif month_str == 'December':
            return 11


    def validate_missing_project(self):  # this function will flag if we miss projects in the google sheet
        unprocessed_projects = []
        for project in self.project_list:
            if project not in self.process_project:
                if self.project_list[project][0][0] != '0' or self.project_list[project][1][0] != '0':
                    unprocessed_projects.append(project)
        return unprocessed_projects


'''
proj = []
actu = []
test = GSheet_finance(proj, actu)
test.gsheet_authenticatication()
# test.update_cell('Finance Report!D2', test.today)
# print(test.get_cell_value('Finance Report!C10'))
test.update_projects()
'''