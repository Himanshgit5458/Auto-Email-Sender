from flask import Flask, request
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import smtplib

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formatdate
# # Initialize Flask app
app = Flask(__name__)

# # Configure the Google Sheets API
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('gmail-test-394006-2dbe2bed0a21.json', scope)
client = gspread.authorize(credentials)

# # Your Google Spreadsheet details
SPREADSHEET_NAME = 'Gmail - Test'
WORKSHEET_NAME = 'Your Worksheet Name'

# # Function to get data from the Google Spreadsheet
def get_spreadsheet_data():
    sheet = client.open(SPREADSHEET_NAME).get_worksheet(0)
    return sheet.get_all_records()


# # Function to update the status in the Google Spreadsheet
def update_status(row_index, status):
    sheet = client.open(SPREADSHEET_NAME).get_worksheet(0)
    cell = sheet.cell(row_index, 3)  # Assuming status column is the 3rd column (C)
    cell.value = status
    sheet.update_cell(row_index, 3, status)


def send_email(recipient_email, template):
    smtp_server = "smtp.gmail.com"  # Use your email provider's SMTP server if not using Gmail
    smtp_port = 25  # Use the appropriate SMTP port for your email provider

    # Replace the following with your own email credentials
    sender_email = "Himanshuiitm.3011@gmail.com"
    sender_password = "ifzpoqvvztcqxvuj"

    msg = MIMEMultipart("alternative")
    msg["From"] = sender_email
    msg["To"] = recipient_email
    msg["Date"] = formatdate(localtime=True)
    msg["Subject"] = f"Task Update Request:"

    # with open("/content/a.html") as f:
    #     email_template = f.read()

    # html = email_template.format(task=task_name)
    msg.attach(MIMEText(template, "html"))

    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, recipient_email, msg.as_string())
        server.quit()
        print(f"Email sent successfully to {recipient_email}")
    except Exception as e:
        print(f"Failed to send email to {recipient_email}. Error: {e}")
# # Route to render the HTML interface
@app.route('/')
def index():
    data = get_spreadsheet_data()
    rows_html = ""
    for index, row in enumerate(data, start=2):  # Start from row 2 (assuming header in row 1)
        template = """
        <p>Task: {task_name}</p>
        <p>Email: {email}</p>
        <form action="/update_status" method="post">
            <input type="hidden" name="row_index" value="{row_index}">
            <input type="submit" name="status" value="Done">
            <input type="submit" name="status" value="Not Done Yet">
        </form>
        """.format(task_name=row['Task Name'], email=row['Email'], row_index=index)

        recipient_email=row['Email']
        send_email(recipient_email, template)
        

# # Route to handle status update
@app.route('/update_status', methods=['POST'])
def update_status_route():
    row_index = int(request.form['row_index'])
    status = request.form['status']
    update_status(row_index, status)
    return "Status updated successfully!"

if __name__ == '__main__':
    app.run()
