from tkinter import messagebox, scrolledtext

import customtkinter as ctk
from controlSending import customHTMLEmail

# Initialize the customtkinter theme
ctk.set_appearance_mode("System")  # or "Dark"
ctk.set_default_color_theme("blue")  # or any other color theme


# Function to send email logic
def send_email():
    docURL = doc_url_entry.get()
    sheetURL = sheet_url_entry.get()
    sheetTitle = sheet_title_entry.get()
    filePath = file_path_entry.get()
    serviceAccount = service_account_entry.get()
    oauthClientID = oauth_client_id_entry.get()

    if (
        not docURL
        or not sheetURL
        or not sheetTitle
        or not serviceAccount
        or not oauthClientID
    ):
        log_area.insert(
            ctk.END,
            "Error: All fields are required.\n",
        )
        return

    # Calling the customHTMLEmail function and getting the result message
    message = customHTMLEmail(
        docURL=docURL,
        sheetURL=sheetURL,
        sheetTitle=sheetTitle,
        filePath=filePath,
        serviceAccount=serviceAccount,
        oauthClientID=oauthClientID,
    )

    log_area.insert(ctk.END, f"{message}\n")


# Function to clear all input fields
def clear_fields():
    doc_url_entry.delete(0, ctk.END)
    sheet_url_entry.delete(0, ctk.END)
    sheet_title_entry.delete(0, ctk.END)
    file_path_entry.delete(0, ctk.END)
    log_area.delete(1.0, ctk.END)


# Main window
root = ctk.CTk()
root.title("Email Sender")
root.geometry("600x600")

# Configure grid to be stretchable
for i in range(8):  # Updated to configure all rows
    root.grid_rowconfigure(i, weight=1)

for j in range(2):
    root.grid_columnconfigure(j, weight=1)

# New input fields for SERVICE ACCOUNT and OAUTH 2.0 CLIENT ID (moved to the top)
ctk.CTkLabel(root, text="SERVICE ACCOUNT:").grid(
    row=0, column=0, sticky=ctk.W, padx=10, pady=5
)
service_account_entry = ctk.CTkEntry(root, width=400)
service_account_entry.grid(row=0, column=1, padx=10, pady=5, sticky=ctk.EW)

ctk.CTkLabel(root, text="OAUTH 2.0 CLIENT ID:").grid(
    row=1, column=0, sticky=ctk.W, padx=10, pady=5
)
oauth_client_id_entry = ctk.CTkEntry(root, width=400)
oauth_client_id_entry.grid(row=1, column=1, padx=10, pady=5, sticky=ctk.EW)

# Labels and input fields for the remaining fields
ctk.CTkLabel(root, text="Google Doc URL:").grid(
    row=2, column=0, sticky=ctk.W, padx=10, pady=5
)
doc_url_entry = ctk.CTkEntry(root, width=400)
doc_url_entry.grid(row=2, column=1, padx=10, pady=5, sticky=ctk.EW)

ctk.CTkLabel(root, text="Google Sheet URL:").grid(
    row=3, column=0, sticky=ctk.W, padx=10, pady=5
)
sheet_url_entry = ctk.CTkEntry(root, width=400)
sheet_url_entry.grid(row=3, column=1, padx=10, pady=5, sticky=ctk.EW)

ctk.CTkLabel(root, text="Sheet Title:").grid(
    row=4, column=0, sticky=ctk.W, padx=10, pady=5
)
sheet_title_entry = ctk.CTkEntry(root, width=400)
sheet_title_entry.grid(row=4, column=1, padx=10, pady=5, sticky=ctk.EW)

ctk.CTkLabel(root, text="File Path:").grid(
    row=5, column=0, sticky=ctk.W, padx=10, pady=5
)
file_path_entry = ctk.CTkEntry(root, width=400)
file_path_entry.grid(row=5, column=1, padx=10, pady=5, sticky=ctk.EW)

# Send and Clear buttons
send_button = ctk.CTkButton(root, text="Send", command=send_email, width=100)
send_button.grid(row=6, column=0, padx=10, pady=10)

clear_button = ctk.CTkButton(root, text="Clear", command=clear_fields, width=100)
clear_button.grid(row=6, column=1, padx=10, pady=10, sticky=ctk.W)

# Log area using scrolledtext from tkinter (it integrates well with customtkinter)
log_area = scrolledtext.ScrolledText(root, width=58, height=10, wrap=ctk.WORD)
log_area.grid(row=7, column=0, columnspan=2, padx=10, pady=10, sticky=ctk.EW)

# Make the log area stretch vertically
root.grid_rowconfigure(7, weight=2)  # Allow log area to take more space when resizing

# Start the customtkinter event loop
root.mainloop()
