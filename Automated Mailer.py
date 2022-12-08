# --------------------------------- Importing Libraries ---------------------------------#

import os
import pickle
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from tkinter import *
from tkinter import messagebox

import pandas as pd
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# --------------------------------- Global Constants ---------------------------------#

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SPREADSHEET_ID = '''Insert Sheet ID of associated Google Sheet here'''
SAMPLE_RANGE_NAME = 'A1:AA1000'
MESSAGE = ""
HEADER = '''
        <html lang="en">
            <head>
            </head>
        <body>
        '''
FOOTER = '''
            <div class="signature">
                You can include your mail signature in this section
            </div>
        </body>
        </html>
        '''
TEST_MAIL = ''' Insert the mail ID of the sender for default test mails '''
SEND_TEST = False


# --------------------------------- Script Functions ---------------------------------#


def check_script():
    test_button_clicked()
    run_button.config(state=DISABLED)
    if not has_id.get():
        run_script()
    else:
        run_id_attach_script()


def test_button_clicked():
    if test_value.get():
        global TEST_MAIL, SEND_TEST
        if email_entry.get():
            TEST_MAIL = email_entry.get()
        SEND_TEST = True


# Code that'll run when the mails need customized attachments
def run_id_attach_script():
    email_list = pd.DataFrame(values_input[1:], columns=values_input[0])
    your_name = ''' Insert sender's name here '''
    your_email = ''' Insert sender's mail ID here '''
    your_password = ''' Insert sender's app password here '''
    server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
    server.ehlo()
    server.login(your_email, your_password)
    all_names = email_list['Name']
    all_emails = email_list['Email']
    all_subjects = email_list['Subject']
    all_attached = email_list['Attachments']

    '''Your Google Sheet must contain columns: Name(Names of Receivers),Email(Email IDs of Receivers),
       Subject(Customized subjects for recipients, if necessary) and Attachments(File names of any attachments)'''
    for idx in range(len(all_emails)):
        name = all_names[idx]
        email = all_emails[idx]

        attachment = all_attached[idx]
        msg = MIMEMultipart()

        msg['Subject'] = all_subjects[idx]
        msg['From'] = your_email
        msg['To'] = all_emails[idx]

        pdfname = attachment
        binary_pdf = open(pdfname, 'rb')
        payload = MIMEBase('application', 'octate-stream', Name=pdfname)
        payload.set_payload(binary_pdf.read())
        encoders.encode_base64(payload)
        payload.add_header('Content-Decomposition', 'attachment', filename=pdfname)
        msg.attach(payload)

        '''This is the HTML Message Body. You can add images, colors and format it to meet your requirements. Any 
        variables that are used will be wrapped with .format() '''
        global MESSAGE
        MESSAGE = HEADER + input_message.get("1.0", "end-1c").format(name=name) + FOOTER
        msg.attach(MIMEText(MESSAGE, 'html'))
        global SEND_TEST
        if idx == 0 and SEND_TEST:
            try:
                server.sendmail(your_email, [TEST_MAIL], msg.as_string())
                print('Test mail to {} successfully sent!\n\n'.format(TEST_MAIL))
                SEND_TEST = False
            except Exception as e:
                print('Test mail to {} could not be sent because {}\n\n'.format(TEST_MAIL, str(e)))
                window.quit()
                return
            send_all = messagebox.askyesno(title="Send All Mails", message="Are you ready to send all mails?")
            if not send_all:
                window.quit()
                return
        try:
            server.sendmail(your_email, [email], msg.as_string())
            print('Email to {} successfully sent!\n\n'.format(email))
        except Exception as e:
            print('Email to {} could not be sent :( because {}\n\n'.format(email, str(e)))
    server.close()


# Code that'll run when for mails without attachments
def run_script():
    email_list = pd.DataFrame(values_input[1:], columns=values_input[0])
    your_name = ''' Insert sender's name here '''
    your_email = ''' Insert sender's mail ID here '''
    your_password = ''' Insert sender's app password here '''
    server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
    server.ehlo()
    server.login(your_email, your_password)
    all_names = email_list['Name']
    all_emails = email_list['Email']
    all_subjects = email_list['Subject']

    for idx in range(len(all_emails)):
        name = all_names[idx]
        email = all_emails[idx]

        msg = MIMEMultipart()
        msg['Subject'] = all_subjects[idx]
        msg['From'] = your_email
        msg['To'] = email
        '''This is the HTML Message Body. You can add images, colors and format it to meet your requirements. Any 
        variables that are used will be wrapped with .format() '''
        global MESSAGE
        MESSAGE = HEADER + input_message.get("1.0", "end-1c").format(name=name) + FOOTER

        msg.attach(MIMEText(MESSAGE, 'html'))
        global SEND_TEST
        if idx == 0 and SEND_TEST:
            try:
                server.sendmail(your_email, [TEST_MAIL], msg.as_string())
                print('Test mail to {} successfully sent!\n\n'.format(TEST_MAIL))
                SEND_TEST = False
            except Exception as e:
                print('Test mail to {} could not be sent because {}\n\n'.format(TEST_MAIL, str(e)))
                window.quit()
                return
            send_all = messagebox.askyesno(title="Send All Mails", message="Are you ready to send all mails?")
            if not send_all:
                window.quit()
                return
        try:
            server.sendmail(your_email, [email], msg.as_string())
            print('Email to {} successfully sent!\n\n'.format(email))
        except Exception as e:
            print('Email to {} could not be sent :( because {}\n\n'.format(email, str(e)))
    server.close()


# --------------------------------- Authentication to access Spreadsheets ---------------------------------#

def main():
    global values_input, service
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        flow = InstalledAppFlow.from_client_secrets_file(
            'credentials.json', SCOPES)
        creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()
    result_input = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                      range=SAMPLE_RANGE_NAME).execute()
    values_input = result_input.get('values', [])


# --------------------------------- Building Tkinter App ---------------------------------#

main()

window = Tk()
window.title("ACM Mailing Project")
window.minsize(width=450, height=450)

message_label = Label(text="Message", font=("Courier", 20, "bold"))
message_label.pack()

input_message = Text(width=40, height=15)
input_message.pack()

test_value = IntVar()
test_mail_button = Radiobutton(text="Test mail", font=("Courier", 20, "bold"), variable=test_value, value=1)
test_mail_button.pack()

email_entry = Entry(width=25)
email_entry.pack()

# Only select this checkbox if the mails need customized attachments
has_id = IntVar()
has_id_button = Checkbutton(text="This mail needs customized attachments.", variable=has_id, onvalue=1,
                            offvalue=0)
has_id_button.pack()

run_button = Button(text="Run", command=check_script, width=30)
run_button.pack()

window.mainloop()
