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
from pickle import NONE
from dotenv import load_dotenv
import requests
import re
from datetime import datetime
import sys


__location__ = os.path.realpath(os.path.join(
    os.getcwd(), os.path.dirname(__file__)))
load_dotenv(os.path.join(__location__, ".env"))

USERNAME = os.getenv("USERNAME")
PASSWORD = os.getenv("PASSWORD")
IMAP_SERVER: str = "outlook.office365.com"
MESSAGE_FETCH_AMOUNT: int = 1
AFAS_UPDATECONNECTOR_API_ENDPOINT: str = (
    "https://76080.resttest.afas.online/ProfitRestServices/connectors/KnSubject"
)
AFAS_UPDATECONNECTOR_API_TOKEN = os.getenv("AFAS_UPDATECONNECTOR_API_TOKEN")
SUBJECT_TYPE: int = 21
FEATURE_TYPE: int = 127
PROCESSED_MESSAGES_FOLDER: str = "PROCESSED"

is_debug = sys.argv.__contains__("-debug")


def send_updateconnector_post_request(
    date: str, From: str, subject: str, body: str, files: list((str, bytes))
):
    # Convert Afas UpdateConnector API token to base64
    api_token_base64 = base64.b64encode(
        AFAS_UPDATECONNECTOR_API_TOKEN.encode("ascii")
    ).decode("ascii")

    # Create header JSON
    headers = {"Authorization": f"AfasToken {api_token_base64}"}

    # Load default JSON
    input_file = open("default_post_payload.json", "r")
    data = json.load(input_file)
    input_file.close()

    # Parse date
    date_formatted: datetime = parse_date(date)

    # Change JSON property values
    data["KnSubject"]["Element"]["Fields"]["Da"] = date_formatted
    #data["KnSubject"]["Element"]["Fields"]["UsId"] = From
    data["KnSubject"]["Element"]["Fields"]["Ds"] = subject
    data["KnSubject"]["Element"]["Fields"]["SbTx"] = body

    if len(files) > 2:
        # Load default file attachment JSON
        attachment_file = open("attachment.json", "r")
        attachment_json = json.load(attachment_file)
        attachment_file.close()

        # Iterate over files
        x = -1
        for file in files:
            x = x + 1

            if file is None:
                continue

            if x < 2:
                continue

            print(file[0])

            fileContentEncoded = base64.b64encode(
                file[1]
            ).decode("ascii")

            # Change JSON property values
            attachment_json["KnSubjectAttachment"]["Element"]["Fields"][
                "FileName"
            ] = file[0]
            attachment_json["KnSubjectAttachment"]["Element"]["Fields"][
                "FileStream"
            ] = fileContentEncoded

            # Add file attachment JSON to main JSON data
            data["KnSubject"]["Element"]["Objects"].append(attachment_json)

    # Format JSON
    data_formatted: str = json.dumps(data)

    if is_debug == False:
        # Send JSON via POST request
        response = requests.post(
            AFAS_UPDATECONNECTOR_API_ENDPOINT, data_formatted, headers=headers
        )
        if response != None:
            print(response.text)
    else:
        print(data_formatted)


def message_to_body_text(message: Message) -> str:
    result = ""

    # Skip non-text part
    if message.get_content_maintype() == "text":

        # Decode text
        content_charset = message.get_content_charset()
        text = bytes(message).decode(content_charset)

        # Strip header
        match: re.Match = re.search("<html>([\s\S]*?)<\/html>", text)

        if match is None:
            result = ""
        else:
            result = match[0]

    return result


def process_multipart_message(message: Message) -> tuple[str, list((str, bytes))]:
    body = ""
    files = list((str, bytes))
    i: int = 0

    if message.is_multipart():
        for part in message.walk():
            if part.get_content_maintype() == "text":
                body = body + message_to_body_text(part)
            else:
                if part.get("Content-Disposition") is None:
                    continue

                fileTuple = (part.get_filename(),
                             part.get_payload(decode=True))
                files.append(fileTuple)

                process_multipart_message(part)
                i = i + 1
    else:
        body = body + message_to_body_text(message)

    if i == 0:
        files = (None, None)

    return (body, files)


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

    # Check status of select
    if status != "OK":
        print("Error while selecting inbox folder, returned: " + status)
        return

    # Loop through N newest emails
    for i in range(email_count - 1, email_count - MESSAGE_FETCH_AMOUNT - 1, -1):
        type, data = imap.fetch(str(i), "(RFC822)")
        raw_email = data[0][1]
        email_data: Message = email.message_from_bytes(raw_email)

        # Decode date, from, and subject from e-mail header
        date = decode_header(email_data.get("Date"))[0][0]
        From = decode_header(email_data.get("From"))[0][0]
        subject = decode_header(email_data.get("Subject"))[0][0]

        # Get body from e-mail
        body, files = process_multipart_message(email_data)

        # Send UpdateConnector POST request
        send_updateconnector_post_request(date, From, subject, body, files)

        # Move e-mails to PROCESSED_MESSAGES_FOLDER
        # if is_debug == False:
        # imap.uid("MOVE", email_data.get("Message-ID"), PROCESSED_MESSAGES_FOLDER)

    # Close inbox folder and log out
    imap.close()
    imap.logout()


main()
