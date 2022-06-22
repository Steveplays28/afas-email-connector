# Afas email connector
# Written by Darion Spaargaren (AEL stage 22-06-2022) for Storm Outsourcing B.V.
#
# REQUIREMENTS:
# pip install python-dotenv
# pip install requests
#
# Create a "secrets.env" file in script directory
# USERNAME = "address@domain"
# PASSWORD = "password"
# AFAS_UPDATECONNECTOR_API_TOKEN = "<token><version>N</version><data>TOKEN</data></token>"

import base64
import email
import email.contentmanager
from email.message import Message
import imaplib
import os
from dotenv import load_dotenv
import requests

__location__ = os.path.realpath(
    os.path.join(os.getcwd(), os.path.dirname(__file__)))
load_dotenv(os.path.join(__location__, ".env"))

USERNAME = os.getenv("USERNAME")
PASSWORD = os.getenv("PASSWORD")
IMAP_SERVER = "outlook.office365.com"
MESSAGE_FETCH_AMOUNT = 1
AFAS_UPDATECONNECTOR_API_ENDPOINT = "https://76080.resttest.afas.online/ProfitRestServices/connectors/KnSubject"
AFAS_UPDATECONNECTOR_API_TOKEN = os.getenv("AFAS_UPDATECONNECTOR_API_TOKEN")
SUBJECT_TYPE = 21
FEATURE_TYPE = 127


def send_updateconnector_post_request(From: str, subject: str, body: str):
    api_token_base64 = base64.b64encode(
        AFAS_UPDATECONNECTOR_API_TOKEN.encode("ascii")).decode("ascii")

    headers = {"Authorization": f'AfasToken {api_token_base64}'}
    data = '{ "KnSubject": { "Element": { "Fields": { "StId": 21, "Ds": "onderwerp", "SbTx": "toelichting", "Da": "2022-06-22T15:30:48", "FvF1": 127 }, "Objects": [ { "KnSubjectLink": { "Element": { "Fields": { "ToSR": true, "BcId": "999999" } } } } ] } } }'

    response = requests.post(AFAS_UPDATECONNECTOR_API_ENDPOINT, data,
                             headers=headers)
    print(response.status_code, response.text)


def get_text(msg: Message):
    if msg.is_multipart():
        return get_text(msg.get_payload(0))
    else:
        return msg.get_payload(None, True)


def main():
    imap = imaplib.IMAP4_SSL(IMAP_SERVER)
    imap.login(USERNAME, PASSWORD)

    status, messages = imap.select("INBOX", True)
    email_count = int(messages[0].decode("utf-8"))

    # for i in range(email_count - 1, email_count - MESSAGE_FETCH_AMOUNT - 1, -1):
    #     print(i)

    #     type, data = imap.fetch(str(i), "(RFC822)")
    #     raw_email = data[0][1]

    #     raw_email_string = raw_email.decode("utf-8")
    #     email_message = email.message_from_string(raw_email_string)
        
    #     print(email_message)

    for i in range(email_count - 1, email_count - MESSAGE_FETCH_AMOUNT - 1, -1):
        print(i)

        type, data = imap.fetch(str(2), "(RFC822)")
        raw_email = data[0][1]

        raw_email_string = raw_email.decode("utf-8")
        email_message = email.message_from_string(raw_email_string)
        
        print(get_text(email_message))

    imap.close()
    imap.logout()


# send_updateconnector_post_request(
#     "djspaargaren@outlook.com", "Python test", "Test")

main()
