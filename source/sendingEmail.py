import base64
import os
from email import encoders
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from dotenv import load_dotenv
from google.auth.exceptions import (  # For authentication errors
    GoogleAuthError,
    RefreshError,
)
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError  # For HTTP errors

from source.getGoogleDoc import downloadGoogleDocAsHTML, getDocumentID

load_dotenv()

# with the OAuth 2.0 Client IDs feature, this basically helps you create a Client ID(Client = Your Application) with this ID helps your app to present itself and ask for the permission/credentials to send emails
# secret_key = os.path.join("key", "emailSender.json")
secret_key = os.getenv("OAUTH_2.0_CLIENT_ID")

# limits the access that the client/app is asking for
scopes = ["https://www.googleapis.com/auth/gmail.send"]


# authenticate with OAuth2 and get the credentials/permission(opens a browser that allows us to get the permission)
def authenticateGmail():
    try:
        flow = InstalledAppFlow.from_client_secrets_file(secret_key, scopes=scopes)
        # open up the browser
        creds = flow.run_local_server(port=0)
        return creds
    except FileNotFoundError:
        print("Error: The credentials file was not found.")
    except ValueError:
        print("Error: The credentials file is improperly formatted.")
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
    try:
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
    except HttpError as error:
        print(f"HTTP error occurred: {error.resp.status} - {error._get_reason()}")
    except RefreshError:
        print("Error: Unable to refresh the authentication token.")
    except Exception as error:
        print(f"An unexpected error occurred while sending the email: {error}")


# READ BEFORE USE
# Make sure to authenticate before usage
# creds = authenticateGmail()


# If you want to send content from a Google doc, you have to retrieve the html contents of it
# docID = getDocumentID(url="(url of documentation)")
# htmlContent = downloadGoogleDocAsHTML(docID)


# Sending the emaiL
# sendEmail(
#   creds=creds,
#   to="reciever-email",
#   cc=["mail1", "mail2"],
#   bcc=["mail1", "mail2"],
#   subject="subject-of-email",
#   messageText="plain-text",
#   htmlContent="html-build",
#   imagePath="path-of-image",
#   filePath="path-of-attachment")
