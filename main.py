import datetime
import json
import os.path
import subprocess
import time

import schedule
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build


def run_speedtest():
    date = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    cmd = ["speedtest", "--format=json"]
    try:
        result = subprocess.run(
            cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True
        )
        if result.returncode != 0:
            raise Exception(result.stderr)
        output = json.loads(result.stdout)
        download_speed = int(output["download"]["bandwidth"] / 125_000)
        upload_speed = int(output["upload"]["bandwidth"] / 125_000)
        print(f"{date}, {download_speed}, {upload_speed}")
        write_to_sheets(date, download_speed, upload_speed)
    except Exception as e:
        print(f"{date} - failed to perform speedtest. {e}")


def get_google_sheets_service():
    # If modifying these SCOPES, delete the file token.json.
    SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

    creds = None
    # The file token.json stores the user's access and refresh tokens and is
    # created automatically when the authorization flow completes for the first time.
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json")
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    service = build("sheets", "v4", credentials=creds)

    return service


def write_to_sheets(date, download_speed, upload_speed):
    service = (
        get_google_sheets_service()
    )  # Make sure you've defined this function to get your Sheets API service

    # The ID of the spreadsheet to update
    spreadsheet_id = "1joCoGDGrL9ltwTB0J35tqKo7HI9-Q4cM1cLED1ZXWTc"

    # The A1 notation of the values to update
    range_name = "results!A1"

    # Values to append: a list of lists, each inner list representing a row
    values = [[date, download_speed, upload_speed]]

    response = append_data_to_sheet(service, spreadsheet_id, range_name, values)


def append_data_to_sheet(service, spreadsheet_id, range_name, values):
    # How the input data should be interpreted.
    value_input_option = (
        "USER_ENTERED"  # The data will be parsed as if the user typed it into the UI
    )

    # How the input data should be inserted.
    insert_data_option = "INSERT_ROWS"  # The new data will be inserted into new rows

    value_range_body = {"values": values}  # The actual data to append

    request = (
        service.spreadsheets()
        .values()
        .append(
            spreadsheetId=spreadsheet_id,  # The ID of the spreadsheet to update
            range=range_name,  # The A1 notation of a range to search for a logical table of data
            valueInputOption=value_input_option,  # How input data should be interpreted
            insertDataOption=insert_data_option,  # How input data should be inserted
            body=value_range_body,
        )
    )

    response = request.execute()  # Execute the request

    return response


if __name__ == "__main__":
    schedule.every().hour.at(":00").do(run_speedtest)
    schedule.every().hour.at(":15").do(run_speedtest)
    schedule.every().hour.at(":30").do(run_speedtest)
    schedule.every().hour.at(":45").do(run_speedtest)
    while True:
        schedule.run_pending()
        time.sleep(1)
