import gspread
from oauth2client.service_account import ServiceAccountCredentials
import qrcode
from qrcode.image.styledpil import StyledPilImage
from qrcode.image.styles.moduledrawers import RoundedModuleDrawer
from qrcode.image.styles.colormasks import VerticalGradiantColorMask
from PIL import Image, ImageDraw, ImageOps
from tqdm import tqdm, trange
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
import threading



print("QR CODE MAIL SENDER\n\n   Connecting to Google Spreadsheets...\n   |")
# Set up Google Sheets API credentials
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('I:/Alumni Meet/QRAuto/consummate-sled-415609-92c5f711e587.json', scope)

# Open the Google Spreadsheet using its title
spreadsheet_title = 'Alumni1'
worksheet_title = 'Alumni1'
overall_progress = 0
total_threads_to_process = 0

# Email setup
smtp_server = 'smtp.gmail.com'
smtp_port = 587
smtp_username = 'reunion.vbvp@gmail.com'
smtp_password = 'jwaf pgwa qjio ytaj'
from_email = 'reunion.vbvp@gmail.com'

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
def send_email(email, alumnus_name, unique_data):
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=10,
        border=4,
    )
    qr.add_data(unique_data)
    qr.make(fit=True)

    # Create image styles
    img = qr.make_image(
    image_factory=StyledPilImage, # Use the styled image factory
    module_drawer=RoundedModuleDrawer(), # Use the rounded module drawer
    )
    img = img.convert("RGBA")

    # Create a mask with rounded corners
    mask = Image.new("L", (img.width, img.height), 0)
    draw = ImageDraw.Draw(mask)
    draw.rounded_rectangle([(0, 0), (img.width, img.height)], radius=30, fill=255)
    mask = mask.resize(img.size, Image.LANCZOS)

    # Apply the mask to the image
    img.putalpha(mask)
    img_path = f'C:/Users/Atharva/Documents/python/QRAuto/QRAuto/QRtemp/QRCode_{email}.png'
    img.save(img_path)

    subject = 'Your Unique QR Code'

    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = email
    msg['Subject'] = subject

    # HTML body with embedded image
    body = """
<!DOCTYPE html>
<html lang="en">

<body style="align-content:center;" >

    <div style="width:96%; max-width:650px; padding-bottom: 30px; border-radius: 25px; border: 2px solid #fcfeff; margin: 10px auto; position: relative; background-color: #fff; padding: 5px; box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);">
        <div style="border: 2px solid #1e57c9; width: 99%; border-radius: 25px;"
        <div style="width: 100%; border-radius: 25px; overflow: hidden;">
        <div style="width: 100%; border-radius: 25px; overflow: hidden;">
            <center>
                   <img src="cid:AL" alt="Alumni Image" width="100%">
           </center>
        </div>
            <br><br>
            <div style="padding: 10px; width: 90%; border-color: #f8f8f8; font-size: 12px; font-family: 'EB Garamond', serif; text-align: left;">
                <p align="justify" size="14px">Dear {alumnus_name},<br><br>

                    &nbsp;&nbsp; We are excited to have you join us at the upcoming Alumni Meet. <b>Your registration and payment have been
                    successfully processed</b>, and we appreciate your commitment to making this event a success.
                    <br>
                    &nbsp;&nbsp; To streamline your entry and ensure a smooth experience at the venue, we are providing you with a unique
                    QR code. <b>This code will serve as your electronic ticket and can only be scanned once for entry</b>.<br><br>
                </p>
                    <center>
                        <img src="cid:qrcode" alt="QR Code" style="width:75%; max-width: 290px;">
                    </center>
                    <br><br>
                <p align="justify" size="14px">
                    <b>Please make sure to have this QR code accessible on your mobile device or in print when you arrive at
                    the event</b>. If you encounter any issues or have questions, our event staff will be readily available to
                    assist you.
                    <br>
                    &nbsp;&nbsp;Thank you for your participation, and we look forward to creating wonderful memories together at the
                    Alumni Meet.<br><br>
                </p>
                <p>
                    Best Regards,<br>
                    Vidyavardhini's Bhausaheb Vartak Polytechnic<br>
                </p>
            </div>
        </div>
        </div>
    </div>
</body>

</html>            """.format(alumnus_name=alumnus_name)

    msg.attach(MIMEText(body, 'html'))

    # Attach the image
    with open(img_path, 'rb') as img_file:
        img_data = img_file.read()
        img_mime = MIMEImage(img_data, name='QRCode.png')
        img_mime.add_header('Content-ID', '<qrcode>'.format(img_path))
        img_mime.add_header('Content-Disposition', 'inline', filename='QRCode.png')
        msg.attach(img_mime)

    # Attach the second image
    with open('I:\Alumni Meet\QRAuto\AL.png', 'rb') as img_file:
        img_data = img_file.read()
        img_mime = MIMEImage(img_data, name='AL.png')
        img_mime.add_header('Content-ID', '<AL>')  # Unique Content-ID
        img_mime.add_header('Content-Disposition', 'inline', filename='AL.png')
        msg.attach(img_mime)

    # Send the email
    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(smtp_username, smtp_password)
        server.sendmail(from_email, email, msg.as_string())

# Function to process and send emails for even and odd halves
def process_and_send_emails(emails_range, pbar):
    global overall_progress
    for row_num in emails_range['range']:
        entry_status = worksheet.cell(row_num, 7).value  # Entry status from the 'Status_Mail_Sent' column
        if entry_status is None or not entry_status.strip():
            # Extract data from respective columns
            email = worksheet.cell(row_num, 2).value
            unique_data = worksheet.cell(row_num, 5).value
            alumnus_name = worksheet.cell(row_num, 3).value
            
            # Call send_email function
            send_email(email, alumnus_name, unique_data)
            
            # Update the 'Status_Mail_Sent' column to 'YES'
            worksheet.update_cell(row_num, 7, 'YES')
            
            # Update progress counters
            overall_progress += 1
            pbar.update(1)


def distribute_entries(email_ranges):
    # Initialize tqdm progress bar
    progress_bar = tqdm(total=total_records - 1, desc="   Searching & Distributing Entries")

    # Determine the entries to be sent emails
    entries_to_send = []
    for row_num in range(2, total_records + 2):
        entry_status = worksheet.cell(row_num, 7).value  # Entry status from the 'Entry' column
        if entry_status is None or not entry_status.strip():
            entries_to_send.append(row_num)
        progress_bar.update(1)  # Update tqdm progress bar

    # Close tqdm progress bar
    progress_bar.close()

    # Distribute entries among threads
    num_threads = len(email_ranges)
    entries_per_thread = len(entries_to_send) // num_threads
    remaining_entries = len(entries_to_send) % num_threads

    distributed_entries = []
    start_idx = 0
    for i in range(num_threads):
        thread_entries = entries_to_send[start_idx:start_idx + entries_per_thread]
        if remaining_entries > 0:
            thread_entries.append(entries_to_send[start_idx + entries_per_thread])
            remaining_entries -= 1
            start_idx += entries_per_thread + 1
        else:
            start_idx += entries_per_thread
        distributed_entries.append(thread_entries)

    # Update email_ranges with distributed entries
    for i in range(num_threads):
        email_ranges[i]['range'] = distributed_entries[i]

    return email_ranges

if __name__ == "__main__":
    email_ranges = [
        {'range': [], 'mod': 1, 'desc': 'Thread 1'},  # Odd half, first thread
        {'range': [], 'mod': 0, 'desc': 'Thread 2'},  # Even half, second thread
        {'range': [], 'mod': 1, 'desc': 'Thread 3'},  # Odd half, third thread
        {'range': [], 'mod': 0, 'desc': 'Thread 4'}   # Even half, fourth thread
    ]

    email_ranges = distribute_entries(email_ranges)

    # Count the number of threads with non-empty ranges
    total_threads_to_process = sum(1 for email_range in email_ranges if email_range['range'])

    # Initialize tqdm progress bar for sending emails only if there are threads to process
    print("   |")
    if total_threads_to_process > 0:
        with tqdm(total=total_threads_to_process, desc='   Sending Emails to Alumnus') as pbar:
            threads = []
            for email_range in email_ranges:
                if email_range['range']:  # Only start threads with non-empty ranges
                    t = threading.Thread(target=process_and_send_emails, args=(email_range, pbar))
                    threads.append(t)
                    t.start()

            for thread in threads:
                thread.join()
    else:
        print("   No email entires found to process and send mails")

    print("Done!")