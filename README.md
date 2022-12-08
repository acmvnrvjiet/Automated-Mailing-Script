# Automated-Mailing-Script
Python based Source code for an automated mailing system capable of sending customized bulk emails to varied recepients
Minimum Requirements to run the script:
- A Python IDE
- Internet Connectivity.

## Installation Guide:
- Install the following packages using pip install:
  - pip install google-auth-oauthlib
  - pip install --upgrade google-api-python-client
- Reference for API setup: https://developers.google.com/sheets/api/quickstart/python 

## Main Screen:
![image](https://user-images.githubusercontent.com/57917661/206483826-00e0b122-930c-4a74-aa72-bfdd4d2b9821.png)

## Prompt for sending all mails:
![image](https://user-images.githubusercontent.com/57917661/206484484-87de9554-45fe-4a8b-b544-6b8462696886.png)


### Note:
- The source code cannot be run in a standalone manner.
Two dependecies namely cresentials.json and token.pickle
must be setup on every system that the code is pulled to.
- The .json and .pickle files can be obtained by establishing
a link by means of the Google Sheets API associated
with the sender's email ID.
