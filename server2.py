from flask import Flask, request
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import smtplib

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formatdate

# Initialize Flask app
app = Flask(__name__)

# Configure the Google Sheets API
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('gmail-test-394006-2dbe2bed0a21.json', scope)
client = gspread.authorize(credentials)

# Your Google Spreadsheet details
SPREADSHEET_NAME = 'Gmail - Test'
WORKSHEET_NAME = 'Your Worksheet Name'

# Function to get data from the Google Spreadsheet
def get_spreadsheet_data():
    sheet = client.open(SPREADSHEET_NAME).get_worksheet(0)
    return sheet.get_all_records()

# Function to update the status in the Google Spreadsheet
def update_status(row_index, status):
    sheet = client.open(SPREADSHEET_NAME).get_worksheet(0)
    cell = sheet.cell(row_index, 3)  # Assuming status column is the 3rd column (C)
    cell.value = status
    sheet.update_cell(row_index, 3, status)

# Function to send the email
def send_email(recipient_email, template, row_index):
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

    msg.attach(MIMEText(template, "html"))

    # Append the row_index as a query parameter in the feedback link
    feedback_link = f"http://localhost:5000/feedback/{row_index}"
    template += f"<p>Provide your feedback: <a href='{feedback_link}'>Click here</a></p>"

    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, recipient_email, msg.as_string())
        server.quit()
        print(f"Email sent successfully to {recipient_email}")
    except Exception as e:
        print(f"Failed to send email to {recipient_email}. Error: {e}")

# Route to render the HTML interface
@app.route('/')
def index():
    data = get_spreadsheet_data()
    rows_html = ""
    for index, row in enumerate(data, start=2):
        template = """
        <p>Task: {task_name}</p>
        <p>Email: {email}</p>
        <form action="/update_status" method="post">
            <input type="hidden" name="row_index" value="{row_index}">
            <input type="submit" name="status" value="Done">
            <input type="submit" name="status" value="Not Done Yet">
        </form>
        """.format(task_name=row['Task Name'], email=row['Email'], row_index=index)

        # Send the email with the feedback link
        send_email(row['Email'], template, index)

        rows_html += template

    return rows_html

# Route to handle status update
@app.route('/update_status', methods=['POST'])
def update_status_route():
    row_index = int(request.form['row_index'])
    status = request.form['status']
    update_status(row_index, status)

    # Get the recipient email for this row
    data = get_spreadsheet_data()
    recipient_email = data[row_index - 2]['Email']  # Subtract 2 to adjust for enumeration starting from 2

    # Send the email
    send_email(recipient_email, f"Task status updated to: {status}", row_index)

    return "Status updated successfully!"

# Route to handle user feedback
@app.route('/feedback/<int:row_index>', methods=['GET', 'POST'])
def feedback(row_index):
    if request.method == 'POST':
        # Get the user's feedback from the form submission
        feedback = request.form['feedback']

        # Update the status in the Google Spreadsheet based on the feedback
        if feedback.lower() == 'done':
            update_status(row_index, 'Done')
        elif feedback.lower() == 'not done yet':
            update_status(row_index, 'Not Done Yet')

        # You can display a confirmation message to the user if desired.

    # Your HTML template for the feedback form goes here
    feedback_template = """
    <h1>Provide Your Feedback</h1>
    <form action="" method="post">
        <input type="radio" name="feedback" value="Done"> Done<br>
        <input type="radio" name="feedback" value="Not Done Yet"> Not Done Yet<br>
        <input type="submit" value="Submit">
    </form>
    """
    
    return feedback_template

if __name__ == '__main__':
    app.run()
