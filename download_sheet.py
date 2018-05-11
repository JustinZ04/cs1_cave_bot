#!/usr/bin/env python

''' This code was taken from this stack overflow post:
https://stackoverflow.com/questions/44012089/google-drive-api-v3-download-google-spreadsheet-as-excel
which also uses code from the Google Drive API python quickstart guide found here:
https://developers.google.com/drive/v3/web/quickstart/python
'''

import httplib2
import os
import config

from twilio.rest import Client
from datetime import datetime
from apiclient import discovery
from googleapiclient.errors import HttpError
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

try:
    import argparse

    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/drive-python-quickstart.json
SCOPES = 'https://spreadsheets.google.com/feeds https://www.googleapis.com/auth/drive'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Drive API Python Quickstart'


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
                                   'drive-python-quickstart.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else:  # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials


def main():
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('drive', 'v3', http=http)

    file_id = config.cs1_file_id

    request = service.files().export_media(fileId=file_id,
                                           mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    try:
        with open('office_hours.xlsx', 'wb') as f:
            f.write(request.execute())

    except (ConnectionError, TimeoutError, HttpError) as e:
        # If the download fails log it to a file with the time and type of failure.
        # Also send a text message to the provided numbers.
        cur_time = datetime.now().time()
        file = open("logs/download.log", "a")
        file.write(str(cur_time) + " Download Failed because" + str(e) + "\n")
        file.close()
        c = Client(config.sms_sid, config.sms_token)
        c.messages.create(body='The spreadsheet did not download because' + str(e), from_=str(+14078900127),
                          to=config.m_phone)
        c.messages.create(body='The spreadsheet did not download because' + str(e), from_=str(+14078900127),
                          to=config.j_phone)


if __name__ == '__main__':
    main()
