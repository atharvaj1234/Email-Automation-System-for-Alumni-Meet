<h1>Alumni Registration Email Sender</h1>
This script is designed to send registration emails with QR codes to alumni for a reunion event. The email sending process is managed using the Gmail SMTP server and the registration data is fetched from a Google Spreadsheet.

<h3>Prerequisites</h3>
Install the required libraries by running pip install gspread oauth2client-service-account tqdm
Create a Google Cloud Project, enable the Google Sheets API, and download the JSON key file for the service account. Update the path in the script with the actual path of the JSON key file.
Update the smtp_username and smtp_password variables in the script with your Gmail address and password.
<h3>How it works</h3>
The script connects to the Google Spreadsheet using the Google Sheets API.
It reads the alumni data from the spreadsheet and filters out the entries that haven't received an email yet.
The email sending process is divided into multiple threads to speed up the execution.
The script sends an email with a QR code to each alumnus in the filtered list.
After sending an email, it updates the 'Status_Mail_Sent' column in the Google Spreadsheet to 'YES' to mark the entry as processed.
<h3>Code Overview</h3>
send_email: Function to send an email with a QR code to a specific alumnus.
process_and_send_emails: Function to process and send emails for a specific range of entries.
distribute_entries: Function to distribute the entries among different threads.
<h3>Running the script</h3>
Update the required credentials and paths in the script.
Run the script using python alumni_email_sender.py.
<h2>Note</h2>
The script uses the tqdm library for progress bars, which helps in tracking the progress during the execution.



