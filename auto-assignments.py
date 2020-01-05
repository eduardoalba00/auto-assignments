from __future__ import print_function
import smtplib
import time
import imaplib
import email
import datetime
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

FROM_EMAIL  = "USER_EMAIL"
FROM_PWD    = "USER_PASSWORD"
SMTP_SERVER = "imap.gmail.com"
SMTP_PORT = 993

SCOPES = ['https://www.googleapis.com/auth/calendar']

def read_email_from_gmail():
    try:

        # Set up Calender API
        creds = None
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

        service = build('calendar', 'v3', credentials=creds)

        # ------------------------------------------------------------------------ #

        mail = imaplib.IMAP4_SSL(SMTP_SERVER)
        mail.login(FROM_EMAIL,FROM_PWD)
        mail.select('inbox')

        type, data = mail.search(None, '(UNSEEN)')
        mail_ids = data[0]

        id_list = mail_ids.split()   
        first_email_id = int(id_list[0])
        latest_email_id = int(id_list[-1])

        while True:
            for i in range(latest_email_id,first_email_id, -1):
                
                type, data = mail.fetch(i, '(RFC822)' )

                for response_part in data:
                    if isinstance(response_part, tuple):
                        msg = email.message_from_string(response_part[1])
                        email_subject = msg['subject']
                        email_from = msg['from']
                        email_date = msg['date']

                        date_sub = email_date[5:7]

                        class_name = email_subject[7:11]

                        # print (email_date)
                        # print (date_sub)

                        now = datetime.datetime.now()

                        year = now.year
                        month = now.month
                        day = str(now.day)

                        day_check = str(now.day-1)

                        # print (day)

                        start_hour = str(now.hour+1)
                        end_hour = str(now.hour+1)

                        if now.hour == 23:
                            day = str(now.day+1)
                            start_hour = "0"
                            end_hour = "0"

                        if now.hour < 12:
                            start_hour = "0" + str(now.hour)
                            end_hour = "0" + str(now.hour)

                        if now.hour == 0:
                            start_hour = "1"
                            end_hour = "1"


                        start_time = now.strftime("%Y-%m-"+day+"T"+start_hour + ":00:00")
                        end_time = now.strftime("%Y-%m-"+day+"T"+end_hour + ":00:00")

                        # If email subject matches that of a new assignment
                        if "Assignment: Open" in email_subject and day_check in date_sub:
                            # print ('From: ' + email_from)
                            # print ('Subject: ' + email_subject) # + '\n'

                            # print (start_time)
                            # print (end_time)

                            body={
                                 "summary": 'New (' + class_name +') Assignment',
                                 "description": 'Check New Assignment',
                                 "start": {
                                    "dateTime": start_time,
                                    "timeZone": 'America/New_York'},
                                 "end": {
                                    "dateTime": end_time, 
                                    "timeZone": 'America/New_York'},
                                "reminders": {
                                    "useDefault": False, 
                                    "overrides": [
                                        {'method': 'popup', 'minutes': 0},
                                    ],
                                },
                             }
                            event_result = service.events().insert(calendarId='primary', body=body).execute()
                            print ("New Event Has Been Created!")

    except Exception, e:
        print (str(e))

def main():
    read_email_from_gmail()

if __name__ == '__main__':
    main()
