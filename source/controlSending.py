from getCSV import getCSV, writeToCell
from getGoogleDoc import downloadGoogleDocAsHTML, getPlaceholders, replacePlaceholders
from sendingEmail import authenticateGmail, sendEmail


def findEmailRow(data):
    for row in data:
        for item in row:
            if item == "Email" or item == "email":
                return data.index(row)


def customHTMLEmail(docURL=None, sheetURL=None, sheetTitle=None, filePath=None):
    try:
        data = getCSV(googleSheetURL=sheetURL, sheetTitle=sheetTitle)
        htmlContent = downloadGoogleDocAsHTML(docURL=docURL)

        emailRow = findEmailRow(data=data)
        writeToCell(sheetURL, sheetTitle, emailRow + 1, len(data[emailRow]) + 1, "Sent")

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
            else (
                data[emailRow].index("subject") if "subject" in data[emailRow] else None
            )
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

        creds = authenticateGmail()

        for row in data[(emailRow + 1) :]:  # Skipping the header row
            userData = columnIndexes.copy()
            for key, index in columnIndexes.items():
                userData[key] = row[index]

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
                    sheetURL, sheetTitle, data.index(row) + 1, len(row) + 1, "Yes"
                )
            except Exception as error:
                writeToCell(
                    sheetURL, sheetTitle, data.index(row) + 1, len(row) + 1, "No"
                )
    except Exception as error:
        print(f"An error orcurred: {error}")


# docURL = "https://docs.google.com/document/d/1D0bZ22qu7qxx6pnXq5bPBhRLKHkdOsPnTP2ziDF_J74/edit?usp=sharing"
# sheetURL = "https://docs.google.com/spreadsheets/d/1vaKcL1u4031p7CbywG6_Yg4mLDGyU4AY3rUWNErj05E/edit?usp=sharing"
# sheetTitle = "Sheet4"
# filePath = r"C:\Users\Admin\Downloads\bruh.html"

# docURL, sheetURL, sheetTitle, filePath = "", "", "", ""


# while True:
#     docURL = input("Please enter the URL of the Google Doc (or -1 to exit): ")
#     if docURL == "-1":
#         break

#     sheetURL = input("Please enter the URL of the Google Sheet (or -1 to exit): ")
#     if sheetURL == "-1":
#         break

#     sheetTitle = input("Please enter the title of the sheet (or -1 to exit): ")
#     if sheetTitle == "-1":
#         break

#     filePath = input("Please enter the filePath to your attachment (or -1 to exit): ")
#     if filePath == "-1":
#         break

#     customHTMLEmail(
#         docURL=docURL,
#         sheetURL=sheetURL,
#         sheetTitle=sheetTitle,
#         filePath=filePath,
#     )
