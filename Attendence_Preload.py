import gspread
from oauth2client.service_account import ServiceAccountCredentials
from tqdm import tqdm, trange
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
import threading



print("ATTENDANCE OVERRIDER\n\n   Connecting to Google Spreadsheets...\n   |")
# Set up Google Sheets API credentials
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('I:/Alumni Meet/QRAuto/consummate-sled-415609-92c5f711e587.json', scope)

# Open the Google Spreadsheet using its title
spreadsheet_title = 'Alumni1'
worksheet_title = 'Alumni1'
overall_progress = 0
total_threads_to_process = 0


try:
    gc = gspread.authorize(credentials)
    spreadsheet = gc.open(spreadsheet_title)
    worksheet = spreadsheet.worksheet(worksheet_title)
    total_records = len(worksheet.get_all_records())
    
except gspread.SpreadsheetNotFound as e:
    print(f"Spreadsheet '{spreadsheet_title}' or worksheet '{worksheet_title}' not found.")
    print(e)
    exit()

# Function to send email with QR code

# Function to process and send emails for even and odd halves
progress_bar = tqdm(total=total_records - 1, desc="   Overriding Entries")
# Determine the entries to be sent emails
for row_num in range(2, total_records + 2):
    entry_status = worksheet.cell(row_num, 8).value  # Entry status from the 'Entry' column
    if entry_status is None or not entry_status.strip():
        worksheet.update_cell(row_num, 8, 'NOT CHECKED')
    progress_bar.update(1)  # Update tqdm progress bar
# Close tqdm progress bar
progress_bar.close()
print("Done")
    # Distribute entries among threads
