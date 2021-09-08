from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import sys

#for data handling with SQL database
import mysql.connector as dbms
import pandas as pd

# Connecting to SQL server

mydb = dbms.connect(
  host="localhost",
  user="root",
  password="123456",
  database="pa2"
)

#data collection
mycursor = mydb.cursor()
dbcur = mydb.cursor()
dbcur.execute("SELECT * FROM pa2.datausage")
table_rows = dbcur.fetchall()
df = pd.DataFrame(table_rows)
names = df[1].tolist()
usage = df[2].tolist()
amount = df[3].tolist()

# Uncomment the below lines to use data from a CSV file instead of an SQL database

# df2 = pd.read_csv('pa2.csv')
# dfcsv = pd.DataFrame(df2)
# names1 = dfcsv["name"].tolist()
# usage1 = dfcsv["usage"].tolist()
# amount1 = dfcsv["amount"].tolist()

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/documents', 'https://www.googleapis.com/auth/drive']
#DOCUMENT_ID is the template document's id
DOCUMENT_ID = '15d5rhgnhJWAHOX2h3Ugmu7OZy0Q9VfnwvSYRZRhL5sY'


def main():

    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('docs', 'v1', credentials=creds)
    service2 = build('drive', 'v3', credentials=creds)
    # Retrieve the documents contents from the Docs service.
    document = service.documents().get(documentId=DOCUMENT_ID).execute()

    print('Using template doc with title : {}'.format(document.get('title')))

    field1 = 'field1'
    field2 = 'field2'
    field3 = 'field3'
    runs = len(names)
    # Fill all the requests to be made to the API
    # Iterate over all values in the data arrays to create multiple docs with entries of each user
    
    for i in range(runs):
        # Copy contents of the template  to a new document
        title = 'Invoice'
        doc = service2.files().copy(fileId=DOCUMENT_ID, body=document).execute()
        doc_id = doc.get('id')
        print('Created document with id: {0}'.format(doc.get('id')))
        print('https://docs.google.com/document/d/{0}'.format(doc.get('id')))
        #change names[i] to names1[i], usage[i] to usage1[i] and amount[i] to amount[i] in case of using CSV data
        name_entry = names[i]
        usage_entry = str(usage[i])
        amount_entry = str(amount[i])
        requests = [
            {
                'replaceAllText': {
                    'replaceText': name_entry,
                    'containsText': {
                        'text': field1,
                        'matchCase': False
                    }
                }
            },

            {
                'replaceAllText': {
                    'replaceText': usage_entry,
                    'containsText': {
                        'text': field2,
                        'matchCase': False
                    }
                }
            },

            {
                'replaceAllText': {
                    'replaceText': amount_entry,
                    'containsText': {
                        'text': field3,
                        'matchCase': False
                    }
                }
            }

        ]
        result = service.documents().batchUpdate(documentId=doc.get('id'), body={'requests': requests}).execute()

if __name__ == '__main__':
    main()
