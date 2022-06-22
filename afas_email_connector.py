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
from email.header import decode_header
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


def send_updateconnector_post_request(From: str, subject: str, body: str):
    api_token_base64 = base64.b64encode(
        AFAS_UPDATECONNECTOR_API_TOKEN.encode("ascii")).decode("ascii")

    headers = {"Authorization": f'AfasToken {api_token_base64}'}
    data = '{ "KnSubject": { "Element": { "Fields": { "StId": 21, "Ds": "onderwerp", "SbTx": "toelichting", "Da": "2022-06-22T15:30:48", "FvF1": 127 }, "Objects": [ { "KnSubjectLink": { "Element": { "Fields": { "ToSR": true, "BcId": "999999" } } } } ] } } }'

    response = requests.post(AFAS_UPDATECONNECTOR_API_ENDPOINT, data,
                             headers=headers)
    print(response.status_code, response.text)


def main():
    imap = imaplib.IMAP4_SSL(IMAP_SERVER)
    imap.login(USERNAME, PASSWORD)
    status, messages = imap.select("INBOX", True)

    for i in range(messages, messages - MESSAGE_FETCH_AMOUNT, -1):
        response_type, response_data = imap.fetch(str(i), "(RFC822)")

        for response in response_data:
            message = email.message_from_bytes(response[1])

            From, encoding = decode_header(message.get("From"))[0]
            if isinstance(From, bytes):
                From = From.decode(encoding)

            subject, encoding = decode_header(message["Subject"])[0]
            if isinstance(subject, bytes):
                subject = subject.decode(encoding)

            if message.is_multipart():
                for part in message.walk():
                    content_type = part.get_content_type()
                    content_disposition = str(
                        part.get("Content-Disposition"))

                    try:
                        body = part.get_payload(decode=True).decode()
                    except:
                        pass

                    if content_type == "text/plain" and "attachment" not in content_disposition:
                        # TODO: Send POST API request to Afas UpdateConnector KnSubject
                        send_updateconnector_post_request(From, subject, body)
                        print(body)
                    elif "attachment" in content_disposition:
                        filename = part.get_filename()

                        if filename:
                            # TODO: Check if e-mail subject is a valid folder name
                            email_directory = os.path.join(
                                __location__, subject)
                            if os.path.isdir(email_directory) == False:
                                os.mkdir(email_directory)

                            filepath = os.path.join(
                                email_directory, filename)

                            open(filepath, "wb").write(
                                part.get_payload(decode=True))

                            # TODO: Send POST API request to Afas UpdateConnector KnSubject

            else:
                content_type = message.get_content_type()
                body = message.get_payload(decode=True).decode()

                if content_type == "text/plain":
                    # TODO: Send POST API request to Afas UpdateConnector KnSubject
                    send_updateconnector_post_request(From, subject, body)
                    print(body)


send_updateconnector_post_request(
    "djspaargaren@outlook.com", "Python test", "Test")
