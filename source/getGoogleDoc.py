import os
import re

from bs4 import BeautifulSoup
from dotenv import load_dotenv
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

load_dotenv(override=True)

# Path to your service account key file
# secret_key = os.path.join("key", "botAccount.json")
secret_key = os.getenv("SERVICE_ACCOUNT")

# Define the scopes
scopes = ["https://www.googleapis.com/auth/drive.readonly"]

# Authenticate and build the service
creds = Credentials.from_service_account_file(secret_key, scopes=scopes)
service = build("drive", "v3", credentials=creds)


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
    docID=None, docURL=None, outputFile="output.html", file=False
):
    try:
        if docURL:
            id = getDocumentID(docURL)
            request = service.files().export_media(fileId=id, mimeType="text/html")
        elif docID:
            request = service.files().export_media(fileId=docID, mimeType="text/html")

        htmlContent = request.execute().decode("utf-8")

        if file:
            with open(outputFile, "w", encoding="utf-8") as file:
                file.write(removeBodyStyling(htmlContent))

        return removeBodyStyling(htmlContent)

    except HttpError as error:
        print(f"An error occurred: {error}")

    except Exception as e:
        print(f"An unexpected error occurred: {e}")


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


# Getting the Google Doc ID
# doc_url = "https://docs.google.com/document/d/1mWFUuHTYzyOozT3kA-kL0H_IP3VH1p64o4_QKwK44RI/edit?usp=sharing"
# docID = getDocumentID(doc_url)

# Export the Google Doc as HTML
# htmlContent = downloadGoogleDocAsHTML(docID="pass-Doc-ID", docURL="pass-Doc-URL", file="whether-you-want-an-html-file-or-not(False as default)")

# htmlContent = downloadGoogleDocAsHTML(
#     docURL="https://docs.google.com/document/d/1D0bZ22qu7qxx6pnXq5bPBhRLKHkdOsPnTP2ziDF_J74/edit",
#     file=True,
# )
# return getPlaceHolders(htmlContent))


# with open("output.html", "w", encoding="utf-8") as file:
#     # Replace placeholders in the HTML content
#     file.write(
#         replace_placeholders(
#             htmlContent,
#             {
#                 "Tên": "Chu Lê Nhật Linh",
#                 "Tên người điền": "Mỹ Anh",
#                 "Lớp (vd: 10B4)": "7B1",
#             },
#         )
#     )
