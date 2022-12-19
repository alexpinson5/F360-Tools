# ----------------------------------------------------------------------
#
# Compile Tooling Script
# Alex Pinson
# Saunders Machine Works / NYC CNC
# November 17, 2022
#
# ----------------------------------------------------------------------

# import needed modules
from __future__ import print_function
import sys
import subprocess
import pkg_resources
import os
import os.path

# install additional modules if not included
required = {'bs4', 'google-api-python-client', 'google-auth-httplib2', 'google-auth-oauthlib', 'xlsxwriter'}
installed = {pkg.key for pkg in pkg_resources.working_set}
missing = required - installed
if missing:
    python = sys.executable
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'bs4'])   #may have to change to 'beautifulsoup4' instead of 'bs4'
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'google-api-python-client'])
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'google-auth-httplib2'])
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'google-auth-oauthlib'])
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'xlsxwriter'])
from urllib.request import urlopen
from bs4 import BeautifulSoup
import xlsxwriter


from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from pprint import pprint
import google.auth


# set up directory name for setup sheets
directory = 'Setup Sheets'

# The ID and range of a sample spreadsheet.
SPREADSHEET_ID = 'REPLACE THIS TEXT WITH YOUR SPREADSHEET ID' # you can get this from the URL when the file is open in your browser
RANGE_NAME = 'REPLACE THIS TEXT WITH THE SHEET NAME!A2:B6' # this is in the bottom left corner with the Google Sheet open. Do not replace '!A2:B6'
 
# loop through files & folders in that directory to compile lists
masterTools = [] # prepare main tool list (single instance of each tool)
fileNames = [] # prepare list for each file name
toolIndex = [] #array for relating tool numbers to file names
for filename in os.listdir(directory):
    # get file path of current item
    f = os.path.join(directory, filename)
    if os.path.isfile(f):
        url = 'file:///' + os.path.abspath(f)
        html = urlopen(url).read()
        soup = BeautifulSoup(html, features="html.parser")

        # kill all script and style elements
        for script in soup(["script", "style"]):
            script.extract()    # rip it out

        # get text elements
        text = soup.get_text()

        # break into lines and remove leading and trailing space on each
        lines = (line.strip() for line in text.splitlines())
        # break multi-headlines into a line each
        chunks = (phrase.strip() for line in lines for phrase in line.split("  "))
        # drop blank lines
        text = '\n'.join(chunk for chunk in chunks if chunk)

        stripped = text.split('Tools:', 2)
        toolList = stripped[2].split('Maximum', 1)

        toolList = toolList[0].split()

        # strip 'T' off of each tool number
        i = 0
        for toolNumber in toolList:
            toolList[i] = toolNumber[1:] # strip off T char
            toolList[i] = int(toolList[i]) # convert to numeric integer
            masterTools.append(toolList[i]) #add tools to master tools
            i += 1
        toolIndex.append(toolList)
        fileNameString = filename.split('.html', 1)[0]
        fileNames.append(fileNameString)
        #print(toolList) # debug only

# process master tools list
masterTools = [*set(masterTools)]
masterTools.sort() # numerically sort tools
#print(masterTools) # debug only
print(fileNames) # debug only

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# get necessary google credentials
creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
if os.path.exists('Script/token.json'):
    creds = Credentials.from_authorized_user_file('Script/token.json', SCOPES)
# If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            'Script/credentials.json', SCOPES)
        creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open('Script/token.json', 'w') as token:
        token.write(creds.to_json())

def main():
    try:
        service = build('sheets', 'v4', credentials=creds)

        # Call the Sheets API
        sheet = service.spreadsheets()

        clear_values_request_body = {
        }

        # wipe previous data
        request = sheet.values().clear(spreadsheetId=SPREADSHEET_ID, range='2:1001', body=clear_values_request_body)
        response = request.execute()
        
        #print('Tool, Programs:')
        #for row in values:
            # Print columns A and E, which correspond to indices 0 and 4.
            #print('%s, %s' % (row[0], row[1]))
    except HttpError as err:
        print(err)

# function to write master tool list
def batch_update_values(spreadsheet_id, range_name, value_input_option, _values):
    try:
        service = build('sheets', 'v4', credentials=creds)

        values = _values
        data = [
            {
                'range': range_name,
                'values': values
            },
        ]
        body = {
            'valueInputOption': value_input_option,
            'data': data
        }
        result = service.spreadsheets().values().batchUpdate(
            spreadsheetId=spreadsheet_id, body=body).execute()
        print(f"{(result.get('totalUpdatedCells'))} cells updated.")
        return result
    except HttpError as error:
        print(f"An error occurred: {error}")
        return error

def append_values(spreadsheet_id, range_name, value_input_option, _values):
    """
    Creates the batch_update the user has access to.
    Load pre-authorized user credentials from the environment.
    TODO(developer) - See https://developers.google.com/identity
    for guides on implementing OAuth2 for the application.
        """
    # pylint: disable=maybe-no-member
    try:
        service = build('sheets', 'v4', credentials=creds)

        values = _values
        body = {
            'values': values
        }
        result = service.spreadsheets().values().append(
            spreadsheetId=spreadsheet_id, range=range_name,
            valueInputOption=value_input_option, body=body).execute()
        print(f"{(result.get('updates').get('updatedCells'))} cells appended.")
        return result

    except HttpError as error:
        print(f"An error occurred: {error}")
        return error

if __name__ == '__main__':
    main()

    # loop and write tool list values
    masterToolsInsert = []
    j = 0
    for tool in masterTools:
        masterToolsInsert.append([tool])
        k = 0
        setupInsert = []
        for setupSheet in toolIndex:
            if tool in setupSheet:
                #print("Tool " + str(tool) + " is in " + str(fileNames[k]))
                setupInsert.append(fileNames[k])
            k += 1  
        #print(setupInsert)
        batch_update_values(SPREADSHEET_ID, "B" + str(j + 2) + ":" + xlsxwriter.utility.xl_col_to_name(len(setupInsert) + 1) + str(j + 2), "USER_ENTERED", [setupInsert]) # write setup sheet list row
        j += 1
    batch_update_values(SPREADSHEET_ID, "A2:A" + str(len(masterTools) + 1), "USER_ENTERED", masterToolsInsert) # write tool list column