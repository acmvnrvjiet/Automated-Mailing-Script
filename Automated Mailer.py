import pandas as pd
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow,Flow
from google.auth.transport.requests import Request
import os
import pickle
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from html.parser import HTMLParser
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SAMPLE_SPREADSHEET_ID_input = '''Insert Sheet ID of associated Google Sheet here'''
SAMPLE_RANGE_NAME = 'A1:AA1000'
def main():
    global values_input, service
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES) 
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()
    result_input = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID_input,
                                range=SAMPLE_RANGE_NAME).execute()
    values_input = result_input.get('values', [])
    if not values_input and not values_expansion:
        print('No data found.')
main()
email_list=pd.DataFrame(values_input[1:], columns=values_input[0])
your_name = '''Place name of the sender here'''
your_email = '''Enter Sender's Email ID'''
your_password = '''Enter Sender's Password'''
server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
server.ehlo()
server.login(your_email, your_password)
all_names = email_list['Name']
all_emails = email_list['Email']
all_subjects = email_list['Subject']
all_ids = email_list['Pdf']
'''Your Google Sheet must contain columns Name(Name of Receiver),Email(Email of Receiver),
   Subject(Custom for each recepient if needed) and PDF(File names of any attachments)'''
for idx in range(len(all_emails)):
    name = all_names[idx]
    email =  all_emails[idx]
    msg = MIMEMultipart() 
    msg['Subject'] = all_subjects[idx]
    msg['From'] = your_email
    msg['To'] = all_emails[idx]
    '''HTML Style Message Body, can add images or colours. Anytype of custom
    formatting. We can format any ACM Message as HTML, insert it here, and done.
    Further, any variables can also be used, as shown below with .format()'''
    message = """\
<html>
</html>
""".format(name=name)
    msg.attach(MIMEText(message,'html'))
    pdfname = all_ids[idx]
    binary_pdf = open(pdfname, 'rb') 
    payload = MIMEBase('application', 'octate-stream', Name=pdfname)
    payload.set_payload((binary_pdf).read()) 
    encoders.encode_base64(payload)
    payload.add_header('Content-Decomposition', 'attachment', filename=pdfname)
    msg.attach(payload)
    try:
        server.sendmail(your_email, [email], msg.as_string())
        print('Email to {} successfully sent!\n\n'.format(email))
    except Exception as e:
        print('Email to {} could not be sent :( because {}\n\n'.format(email, str(e)))
server.close()
