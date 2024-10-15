import os

from dotenv import load_dotenv
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

load_dotenv()

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


def formatHTML(htmlContent):
    return htmlContent.replace(
        "padding:72pt 72pt 72pt 72pt;max-width:451.4pt",
        "/*padding:72pt 72pt 72pt 72pt;max-width:451.4pt*/",
    )


# Function to export Google Doc as HTML
def downloadGoogleDocAsHTML(docID, outputFile="output.html"):
    try:
        request = service.files().export_media(fileId=docID, mimeType="text/html")
        htmlContent = request.execute().decode("utf-8")

        # with open(outputFile, "w", encoding="utf-8") as file:
        #     file.write(formatHTML(htmlContent))

        print(f"Document exported and saved as '{outputFile}'.")
        return formatHTML(htmlContent)

    except HttpError as error:
        print(f"An error occurred: {error}")
        return None
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        return None


# Replace with your Google Doc URL
# doc_url = "(url to documentation)"
# docID = getDocumentID(doc_url)

# Export the Google Doc as HTML
# htmlContent = downloadGoogleDocAsHTML(docID)
