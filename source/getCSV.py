import csv
import os

import gspread
from dotenv import load_dotenv
from oauth2client.service_account import ServiceAccountCredentials

load_dotenv()

# secret_key = os.path.join("key", "botAccount.json")
secret_key = os.getenv("SERVICE_ACCOUNT")

scopes = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

creds = ServiceAccountCredentials.from_json_keyfile_name(secret_key, scopes=scopes)


def getCSV(
    retrieve="local",
    filePath=None,
    googleSheetURL=None,
    sheetTitle=None,
):
    if retrieve == "local":
        while True:
            try:
                dataList = []
                # Attempt to open the file
                with open(filePath, mode="r", encoding="utf-8") as file:
                    csv_reader = csv.reader(file)
                    # Successfully opened the file, returning the csv_reader
                    print("File opened successfully!")
                    for row in csv_reader:
                        dataList.append(row)
                    return dataList
            except FileNotFoundError:
                print(f"Error: the file '{filePath}' does not exist.")
            except PermissionError:
                print(f"Error: Permission denied for file '{filePath}'.")
            except IsADirectoryError:
                print(f"Error: Expected a file but found a directory at '{filePath}'.")
            except UnicodeDecodeError:
                print(
                    f"Error: Could not decode the file '{filePath}'. Please check the file encoding."
                )
            except csv.Error as e:
                print(f"CSV error: {e}")
            except Exception as e:
                print(f"An unexpected error occurred: {e}")

    elif retrieve == "online":
        while True:
            try:
                dataList = []
                # the client is now the bot that does the job for us
                client = gspread.authorize(creds)
                spreadsheet = client.open_by_url(googleSheetURL)
                sheet = spreadsheet.worksheet(sheetTitle)
                return sheet.get_all_values()
            except gspread.exceptions.SpreadsheetNotFound:
                print("Error: The Google Sheet was not found. Please check the URL.")
            except gspread.exceptions.WorksheetNotFound:
                print(
                    f"Error: The worksheet '{sheetTitle}' does not exist in the Google Sheet."
                )
            except gspread.exceptions.APIError as e:
                print(f"API error occurred: {e}")
            except gspread.exceptions.GSpreadException as e:
                print(f"GSpread error occurred: {e}")
            except Exception as e:
                print(f"An unexpected error occurred: {e}")
    else:
        print('Please pass a correct argument either "local" or "online"')
        return None


# READ BEFORE USAGE
# This module returns the CSV in array format: [[row1], [row2], [row3]]
# getCSV(retrieve = "online/local", filePath = "path-to-file-on-computer", googleSheetURL = "URL-to-the-sheet", sheetTitle = "sheet-title-you-want-to-extract"
print(
    getCSV(
        retrieve="online",
        googleSheetURL="https://docs.google.com/spreadsheets/d/16QwhaNCExojCQCUNijhFhNRugP1PeuhYE1af4idms0o/edit?usp=sharing",
        sheetTitle="Sheet1",
    )
)
