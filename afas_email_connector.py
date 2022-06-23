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
from email.message import Message
import imaplib
import json
import os
from dotenv import load_dotenv
import requests
import re
from datetime import datetime

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


def send_updateconnector_post_request(date: str, From: str, subject: str, body: str):
    # Convert Afas UpdateConnector API token to base64
    api_token_base64 = base64.b64encode(
        AFAS_UPDATECONNECTOR_API_TOKEN.encode("ascii")).decode("ascii")

    # Create header JSON
    headers = {"Authorization": f'AfasToken {api_token_base64}'}

    # Load default JSON
    input_file = open("default_post_payload.json", "r")
    data = json.load(input_file)
    input_file.close()

    # Parse date
    date_formatted: datetime = parse_date(date)

    # Change JSON property values
    data["KnSubject"]["Element"]["Fields"]["Da"] = date_formatted
    data["KnSubject"]["Element"]["Fields"]["UsId"] = From
    data["KnSubject"]["Element"]["Fields"]["Ds"] = subject
    data["KnSubject"]["Element"]["Fields"]["Ds"] = subject

    # Format JSON
    data_formatted: str = json.dumps(data)

    # Send JSON via POST request
    # response = requests.post(AFAS_UPDATECONNECTOR_API_ENDPOINT, data_formatted,
    #                          headers=headers)

    # print(data_formatted)
    # print(response.status_code, response.text)


def message_to_body_text(message: Message) -> str:
    # Skip non-text part
    if message.get_content_maintype() != "text":
        return None

    # Decode text
    content_charset = message.get_content_charset()
    text = bytes(message).decode(content_charset)

    # Strip header
    match: re.Match = re.search("<html>([\s\S]*?)<\/html>", text)

    # HTML to text
    # if message.get_content_subtype() == "html":
    #     try:
    #         text = html2text.html2text(text)
    #     except:
    #         print("Error occurred while converting HTML to text")

    return match[0]


def process_multipart_message(message: Message) -> str:
    body = ""
    if message.is_multipart():
        for part in message.walk():
            body = body + process_multipart_message(part)
    else:
        body = body + message_to_body_text(message)

    return body


def parse_date(date: str) -> datetime:
    # Format date to iso format (timezone gets stripped)
    # "Mon, 20 Jun 2022 10:43:17 +0200" ----> "2022-06-20T10:43:17"
    date_formatted: datetime = datetime.strptime(
        date, "%a, %d %b %Y %H:%M:%S %z")
    date_formatted = datetime.replace(date_formatted, tzinfo=None)
    date_formatted = date_formatted.isoformat()

    return date_formatted


def main():
    # Log in to the email server using IMAP
    imap = imaplib.IMAP4_SSL(IMAP_SERVER)
    imap.login(USERNAME, PASSWORD)

    # Select inbox folder
    status, messages = imap.select("INBOX", True)
    email_count = int(messages[0].decode("utf-8"))

    # Loop through N newest emails
    for i in range(email_count - 1, email_count - MESSAGE_FETCH_AMOUNT - 1, -1):
        type, data = imap.fetch(str(2), "(RFC822)")
        raw_email = data[0][1]
        email_data: Message = email.message_from_bytes(raw_email)

        # Decode date, from, and subject from e-mail header
        date = decode_header(email_data.get("Date"))[0][0]
        From = decode_header(email_data.get("From"))[0][0]
        subject = decode_header(email_data.get("Subject"))[0][0]

        # Get body from e-mail
        body = ""
        if email_data.is_multipart():
            for part in email_data.walk():
                text = message_to_body_text(part)
                if text is not None:
                    body = body + text
        else:
            text = message_to_body_text(email_data)
            body = text

        # Send UpdateConnector POST request
        send_updateconnector_post_request(date, From, subject, body)

    # Close inbox folder and log out
    imap.close()
    imap.logout()


main()
