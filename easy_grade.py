
from __future__ import print_function
import httplib2
import os

from apiclient import discovery
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
SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
CLIENT_SECRET_FILE = 'YOUR_CLIENT_SECRET_FILE.json'
APPLICATION_NAME = 'YOUR_APPLICATION_NAME_HERE'
spreadsheetId = 'ID_OF_SPREADSHEET_TO_EDIT'


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


# Searches elements of each value for name/Id
def get_matches_for(name, values):
    matches = []

    for i, row in enumerate(values):
        for word in row:
            if name.lower() in word.lower():
                matches.append(i)

    return matches


def main():
    """
    Creates a Sheets API service object then
    searches for a student name or ID and
    allows entry of grade into specified column
    """

    # Creating Sheets API service object from credentials and spreadsheetId
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service = discovery.build('sheets', 'v4', http=http,
                              discoveryServiceUrl=discoveryUrl)

    rangeName = 'StandardOutput-CL-1474146027608!A:T'
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheetId, range=rangeName).execute()
    values = result.get('values', [])

    # Start User input

    input_column = input("What column is being entered? ")

    while True:
        search = input("Type in a name to search for or stop: ")

        if search == 'stop':
            break

        matches = get_matches_for(search, values)

        if not matches:
            print("No match found.")
        else:

            if len(matches) == 1:
                match = matches[0]
            else:
                print("There are multiple matches, which did you mean?")
                for i, mat in enumerate(matches):
                    print(str(i) + ") " + str(values[mat]))
                match = matches[int(input("Your choice: "))]

            print(match)
            print(values[match])

            grade = input("What is the grade? ")
            print("Writing " + grade + " to [" + input_column + ", " + str(match+1) + "]")

            input_values = [[grade], ]
            body = {'values': input_values}

            service.spreadsheets().values().update(
                spreadsheetId=spreadsheetId, range=input_column+str(match+1),
                valueInputOption="USER_ENTERED", body=body).execute()
            print(values[match])


if __name__ == '__main__':
    main()