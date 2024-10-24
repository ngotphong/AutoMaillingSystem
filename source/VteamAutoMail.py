import base64
import csv
import os
import re
from email import encoders
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import customtkinter as ctk
import gspread
from bs4 import BeautifulSoup
from google.auth.exceptions import (  # For authentication errors
    GoogleAuthError,
    RefreshError,
)
from google.oauth2.service_account import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError  # For HTTP errors
from oauth2client.service_account import ServiceAccountCredentials
from PIL import Image


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


def getCSV(filePath=None, googleSheetURL=None, sheetTitle=None, creds_sheet=None):
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
            client = gspread.authorize(creds_sheet)
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
            print(f"API Sheet error occurred: {e}")
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
        print(f"An error occurred with primary key: {e}")


def writeToCell(googleSheetURL, sheetTitle, row, col, value, creds_sheet=None):
    try:
        # Authorize Google Sheets
        client = gspread.authorize(creds_sheet)
        spreadsheet = client.open_by_url(googleSheetURL)
        sheet = spreadsheet.worksheet(sheetTitle)

        # Update the cell at (row, col) with the value
        sheet.update_cell(row, col, value)
    except Exception as error:
        print(f"An error occured writing to cell: {error}")


# Function to extract the document ID from the URL
def getDocumentID(url):
    return url.split("/d/")[1].split("/")[0]


def removeBodyStyling(htmlContent):
    # parse the html content with BeautifulSoup
    soup = BeautifulSoup(htmlContent, "lxml")

    # finding the body tag and remove its style attribute if it exists
    bodyTag = soup.find("body")
    if bodyTag and bodyTag.has_attr("style"):
        del bodyTag["style"]

    # return the modified html
    return str(soup)


# Function to export Google Doc as HTML
def downloadGoogleDocAsHTML(
    docID=None, docURL=None, outputFile="output.html", file=False, service_doc=None
):
    try:
        if docURL:
            id = getDocumentID(docURL)
            request = service_doc.files().export_media(fileId=id, mimeType="text/html")
        elif docID:
            request = service_doc.files().export_media(
                fileId=docID, mimeType="text/html"
            )

        htmlContent = request.execute().decode("utf-8")

        if file:
            with open(outputFile, "w", encoding="utf-8") as file:
                file.write(removeBodyStyling(htmlContent))

        return removeBodyStyling(htmlContent)

    except HttpError as error:
        print(f"An error occurred downloading Google Doc: {error}")

    except Exception as e:
        print(f"An unexpected error occurred with downloading Google Doc: {e}")


def getPlaceholders(htmlContent):
    # Parse the HTML content with BeautifulSoup
    soup = BeautifulSoup(htmlContent, "html.parser")

    # Find all text in the HTML
    all_text = soup.get_text()

    # Regular expression to find $[...] patterns
    pattern = re.compile(r"\$\[([^\]]+)\]")
    # Find all matches
    matches = pattern.findall(all_text)

    # Print the matches
    return matches


# function to replace $[...] with values from the dictionary
def replacePlaceholders(htmlContent, replacementDict):
    pattern = re.compile(r"\$\[([^\]]+)\]")

    def replace_match(match):
        identifier = match.group(1)
        return replacementDict.get(identifier, match.group(0))

    # Use re.sub() to replace all occurrences of $[...] with the corresponding dictionary value
    return re.sub(pattern, replace_match, htmlContent)


# authenticate with OAuth2 and get the credentials/permission(opens a browser that allows us to get the permission)
def authenticateGmail(oauth2_client_id, scopes_gmail):
    try:
        flow = InstalledAppFlow.from_client_secrets_file(
            oauth2_client_id, scopes=scopes_gmail
        )
        # open up the browser
        creds = flow.run_local_server(port=0)
        return creds
    except FileNotFoundError:
        print("Error: The credentials file was not found(Authentication).")
    except ValueError:
        print("Error: The credentials file is improperly formatted(Authentication).")
    except GoogleAuthError as e:
        print(f"Authentication failed: {e}")
    except Exception as error:
        print(f"An unexpected error occurred during authentication: {error}")


def createEmail(
    to=None,
    cc=None,
    bcc=None,
    subject=None,
    messageText=None,
    htmlContent=None,
    imagePath=None,
    filePath=None,
):
    # create a multipart email
    message = MIMEMultipart()
    message["to"] = to
    message["subject"] = subject

    if isinstance(to, list):
        to = ", ".join(to)
    if isinstance(cc, list):
        cc = ", ".join(cc)
    if isinstance(bcc, list):
        bcc = ", ".join(bcc)

    if cc:
        message["cc"] = cc
    if bcc:
        message["bcc"] = bcc

    # attach the message/html content to the body
    if messageText:
        message.attach(MIMEText(messageText, "plain"))
    if htmlContent:
        message.attach(MIMEText(htmlContent, "html"))

    if imagePath:
        # open the image file and attach it
        with open(imagePath, "rb") as imgFile:
            image = MIMEImage(imgFile.read())
            image.add_header("Content-ID", "<image1>")
            image.add_header(
                "Content-Disposition", "inline", filename=os.path.basename(imagePath)
            )
            message.attach(image)

    if filePath:
        with open(filePath, "rb") as fileAttachment:
            attachment = MIMEBase("application", "octet-stream")
            attachment.set_payload(fileAttachment.read())
            encoders.encode_base64(attachment)
            attachment.add_header(
                "Content-Disposition",
                f'attachment; filename="{os.path.basename(filePath)}"',
            )
            message.attach(attachment)

    # encrypting
    rawMessage = base64.urlsafe_b64encode(message.as_bytes()).decode()
    return {"raw": rawMessage}


# build the email and send it
def sendEmail(
    creds=None,
    to=None,
    cc=None,
    bcc=None,
    subject=None,
    messageText=None,
    htmlContent=None,
    imagePath=None,
    filePath=None,
):
    # try:
    # the build function creates a connection(service) with the Gmail API
    service = build("gmail", "v1", credentials=creds)
    # creating the email message
    emailMessage = createEmail(
        to, cc, bcc, subject, messageText, htmlContent, imagePath, filePath
    )
    # sending the message
    sentMessage = (
        # userID = "me" states the sender of the email, in this case "me" is recognized as the authenticated Gmail account, you can specify "me" with a hardcoded email but we want to minimize fully
        service.users()
        .messages()
        .send(userId="me", body=emailMessage)
        .execute()
    )
    print(f"Email sent successfully! Message ID: {sentMessage['id']}")


# except HttpError as error:
#     print(f"HTTP error occurred: {error.resp.status} - {error._get_reason()}")
# except RefreshError:
#     print("Error: Unable to refresh the authentication token.")
# except Exception as error:
#     print(f"An unexpected error occurred while sending the email: {error}")


def findEmailRow(data):
    for row in data:
        for item in row:
            if item == "Email" or item == "email":
                return data.index(row)


def customHTMLEmail(
    docURL=None,
    sheetURL=None,
    sheetTitle=None,
    filePath=None,
    creds_sheet=None,
    oauth2_client_id=None,
    scopes_gmail=None,
    service_doc=None,
):
    data = getCSV(
        googleSheetURL=sheetURL, sheetTitle=sheetTitle, creds_sheet=creds_sheet
    )
    htmlContent = downloadGoogleDocAsHTML(docURL=docURL, service_doc=service_doc)

    emailRow = findEmailRow(data=data)
    writeToCell(
        sheetURL,
        sheetTitle,
        emailRow + 1,
        len(data[emailRow]) + 1,
        "Sent",
        creds_sheet=creds_sheet,
    )

    emailColumnIndex = (
        data[emailRow].index("Email")
        if "Email" in data[emailRow]
        else (
            data[emailRow].index("email") if "email" in data[emailRow] else None
        )  # Handle the case where neither "Email" nor "email" is found
    )

    subjectColumnIndex = (
        data[emailRow].index("Subject")
        if "Subject" in data[emailRow]
        else (data[emailRow].index("subject") if "subject" in data[emailRow] else None)
    )
    columnIndexes = {"Email": emailColumnIndex, "Subject": subjectColumnIndex}
    for fillIn in getPlaceholders(htmlContent=htmlContent):
        index = (
            data[emailRow].index(fillIn)
            if fillIn in data[emailRow]
            else (
                data[emailRow].index(fillIn.lower())
                if fillIn.lower() in data[emailRow]
                else None
            )  # Handle the case where neither is found
        )

        columnIndexes[fillIn] = index

    creds = authenticateGmail(oauth2_client_id, scopes_gmail)
    count = 0

    for row in data[(emailRow + 1) :]:  # Skipping the header row
        userData = columnIndexes.copy()
        for key, index in columnIndexes.items():
            userData[key] = row[index]

        print(userData)

        # Make a fresh copy of the original HTML content for each email
        emailContent = htmlContent

        # Replace placeholders in the copied emailContent, not the original htmlContent
        emailContent = replacePlaceholders(
            htmlContent=emailContent, replacementDict=userData
        )

        try:
            # Send the email with the modified emailContent
            if filePath:
                sendEmail(
                    creds=creds,
                    to=userData["Email"],
                    subject=userData["Subject"],
                    htmlContent=emailContent,
                    filePath=filePath,
                )
            if not filePath:
                sendEmail(
                    creds=creds,
                    to=userData["Email"],
                    subject=userData["Subject"],
                    htmlContent=emailContent,
                )

            writeToCell(
                sheetURL,
                sheetTitle,
                data.index(row) + 1,
                len(row) + 1,
                "Yes",
                creds_sheet=creds_sheet,
            )
            count += 1
            print(f"Emails sent: {count}")
        except Exception as error:
            writeToCell(
                sheetURL,
                sheetTitle,
                data.index(row) + 1,
                len(row) + 1,
                "No",
                creds_sheet=creds_sheet,
            )
            print(f"An error orcurred in sending: {error}")
    print("All Done!")


# Initialize the customtkinter theme
ctk.set_appearance_mode("System")  # or "Dark"
ctk.set_default_color_theme("blue")  # or any other color theme


# Function to send email logic
def send_email():
    # secret_key = os.path.join("key", "botAccount.json")
    service_account = service_account_entry.get().replace('"', "")
    oauth2_client_id = oauth_client_id_entry.get().replace('"', "")

    scopes_sheet = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    scopes_doc = ["https://www.googleapis.com/auth/drive.readonly"]

    scopes_gmail = ["https://www.googleapis.com/auth/gmail.send"]

    creds_sheet = ServiceAccountCredentials.from_json_keyfile_name(
        service_account, scopes=scopes_sheet
    )

    # Authenticate and build the service
    creds_doc = Credentials.from_service_account_file(
        service_account, scopes=scopes_doc
    )
    service_doc = build("drive", "v3", credentials=creds_doc)

    docURL = doc_url_entry.get()
    sheetURL = sheet_url_entry.get()
    sheetTitle = sheet_title_entry.get()
    filePath = file_path_entry.get()

    if (
        not docURL
        or not sheetURL
        or not sheetTitle
        or not service_account
        or not oauth2_client_id
    ):
        return

    # Calling the customHTMLEmail function and getting the result message
    customHTMLEmail(
        docURL=docURL,
        sheetURL=sheetURL,
        sheetTitle=sheetTitle,
        filePath=filePath,
        creds_sheet=creds_sheet,
        oauth2_client_id=oauth2_client_id,
        scopes_gmail=scopes_gmail,
        service_doc=service_doc,
    )


# Function to clear all input fields
def clear_fields():
    doc_url_entry.delete(0, ctk.END)
    sheet_url_entry.delete(0, ctk.END)
    sheet_title_entry.delete(0, ctk.END)
    file_path_entry.delete(0, ctk.END)


ctk.set_appearance_mode("dark")  # Options: "dark", "light", "system"
ctk.set_default_color_theme("blue")  # Options: "blue", "green", "dark-blue"

# Main window
root = ctk.CTk()
root.title("Email Sender")
root.geometry("550x450")  # Adjusted height for footer space
root.configure(bg="#000000")  # Set background to blacks

# Configure grid to be stretchable
for i in range(8):
    root.grid_rowconfigure(i, weight=2)

for j in range(2):
    root.grid_columnconfigure(j, weight=2)

# Input fields for SERVICE ACCOUNT and OAUTH 2.0 CLIENT ID
ctk.CTkLabel(root, text="SERVICE ACCOUNT:", font=ctk.CTkFont(size=13)).grid(
    row=0, column=0, sticky=ctk.W, padx=20, pady=10
)
service_account_entry = ctk.CTkEntry(root, width=400, corner_radius=10)
service_account_entry.grid(row=0, column=1, padx=20, pady=10, sticky=ctk.EW)

ctk.CTkLabel(root, text="OAUTH 2.0 CLIENT ID:", font=ctk.CTkFont(size=13)).grid(
    row=1, column=0, sticky=ctk.W, padx=20, pady=10
)
oauth_client_id_entry = ctk.CTkEntry(root, width=400, corner_radius=10)
oauth_client_id_entry.grid(row=1, column=1, padx=20, pady=10, sticky=ctk.EW)

# Labels and input fields for remaining fields
ctk.CTkLabel(root, text="Google Doc URL:", font=ctk.CTkFont(size=13)).grid(
    row=2, column=0, sticky=ctk.W, padx=20, pady=10
)
doc_url_entry = ctk.CTkEntry(root, width=400, corner_radius=10)
doc_url_entry.grid(row=2, column=1, padx=20, pady=10, sticky=ctk.EW)

ctk.CTkLabel(root, text="Google Sheet URL:", font=ctk.CTkFont(size=13)).grid(
    row=3, column=0, sticky=ctk.W, padx=20, pady=10
)
sheet_url_entry = ctk.CTkEntry(root, width=400, corner_radius=10)
sheet_url_entry.grid(row=3, column=1, padx=20, pady=10, sticky=ctk.EW)

ctk.CTkLabel(root, text="Sheet Title:", font=ctk.CTkFont(size=13)).grid(
    row=4, column=0, sticky=ctk.W, padx=20, pady=10
)
sheet_title_entry = ctk.CTkEntry(root, width=400, corner_radius=10)
sheet_title_entry.grid(row=4, column=1, padx=20, pady=10, sticky=ctk.EW)

ctk.CTkLabel(root, text="File Path:", font=ctk.CTkFont(size=13)).grid(
    row=5, column=0, sticky=ctk.W, padx=20, pady=10
)
file_path_entry = ctk.CTkEntry(root, width=400, corner_radius=10)
file_path_entry.grid(row=5, column=1, padx=20, pady=10, sticky=ctk.EW)

# Modern Send and Clear buttons with rounded corners and orange color
send_button = ctk.CTkButton(
    root,
    text="Clear",
    command=clear_fields,
    width=180,
    corner_radius=20,
    fg_color="#FFA500",
)
send_button.grid(row=6, column=0, padx=20, pady=15, sticky=ctk.E)

clear_button = ctk.CTkButton(
    root,
    text="Send",
    command=send_email,
    width=180,
    corner_radius=20,
    fg_color="#FFA500",
)
clear_button.grid(row=6, column=1, padx=20, pady=15, sticky=ctk.W)

# Log area configuration
root.grid_rowconfigure(7, weight=2)  # Allow log area to take more space when resizing

# Start the customtkinter event loop
root.mainloop()
