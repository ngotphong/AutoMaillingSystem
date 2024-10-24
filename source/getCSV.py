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


# returns time from the format 15h45 --> 15.45
def parse_time_range(time_range):
    [start, end] = time_range.split("-")
    start_time = start.replace("h", ".")
    end_time = end.replace("h", ".")

    return (float(start_time), float(end_time))


# list/arrays are stored as reference pointers so when you use it in an array they are not duplicated like primitive datatypes, but rather if we are changing it, we are directly manipulating it
def insertionSort(inputList, value):
    # if inputList is empty, we populate it with the first value
    if not inputList:
        inputList.append(value)
    else:
        # traversing through every values to check if the value us smaller than it, and will insert it as soon as its smaller than a value
        for i in range(len(inputList)):
            if parse_time_range(value[4]) < parse_time_range(inputList[i][4]):
                inputList.insert(i, value)
                return
                # done

        # if the value is large than all, we add it to the end
        inputList.append(value)


def getCSV(
    filePath=None,
    googleSheetURL=None,
    sheetTitle=None,
):
    if filePath:
        try:
            dataList = []
            # Attempt to open the file
            with open(filePath, mode="r", encoding="utf-8") as file:
                csv_reader = csv.reader(file)
                # Successfully ope       ned the file, returning the csv_reader
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

    elif googleSheetURL:
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


# Function to find a row by a primary key
def findRowByPrimaryKey(sheet, column, primary_key):
    try:
        # Fetch all values from the specified column
        col_values = sheet.col_values(column)

        # Find the index of the primary key in the column
        if primary_key in col_values:
            row_index = col_values.index(primary_key) + 1  # Sheets is 1-indexed
            return row_index
        else:
            print(f"Primary key '{primary_key}' not found.")

    except Exception as e:
        print(f"An error occurred: {e}")


def writeToCell(
    googleSheetURL,
    sheetTitle,
    row,
    col,
    value,
):
    try:
        # Authorize Google Sheets
        client = gspread.authorize(creds)
        spreadsheet = client.open_by_url(googleSheetURL)
        sheet = spreadsheet.worksheet(sheetTitle)

        # Update the cell at (row, col) with the value
        sheet.update_cell(row, col, value)
    except Exception as error:
        print(f"An error occured: {error}")


# READ BEFORE USAGE
# This module returns the CSV in array format: [[row1], [row2], [row3]]
# getCSV(filePath = "path-to-file-on-computer", googleSheetURL = "URL-to-the-sheet", sheetTitle = "sheet-title-you-want-to-extract"
