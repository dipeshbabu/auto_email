import smtplib
import csv
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from pathlib import Path

# Inputs
data_file_path = "data.csv"
attachment_path = "Event Invitation Sample.pdf"
subject_path = "subject.txt"
message_path = "message.txt"
sender = "your_outlook_email@example.com"
password = "your_outlook_password"


# Get text from txt file
def get_text(file_path):
    with open(file_path, "r") as file:
        return file.read()


# Function to send email
def send_email(subject, body, sender, name, recipient, password):
    msg = MIMEMultipart()
    msg["Subject"] = subject
    msg["From"] = sender
    msg["To"] = recipient
    body = f"Dear {name},\n\n{body}"
    msg.attach(MIMEText(body, "plain"))

    # Attach the file
    part = MIMEBase("application", "octet-stream")
    with open(attachment_path, "rb") as file:
        part.set_payload(file.read())
    encoders.encode_base64(part)
    part.add_header(
        "Content-Disposition", f"attachment; filename={Path(attachment_path).name}"
    )
    msg.attach(part)

    # Use Outlook's SMTP server and port
    with smtplib.SMTP("smtp.office365.com", 587) as smtp_server:
        smtp_server.starttls()  # Upgrade the connection to use TLS

        # Log in to your Outlook account
        smtp_server.login(sender, password)

        # Send the email
        smtp_server.sendmail(sender, recipient, msg.as_string())
    print(f"Message sent to {recipient}")


# Read emails from csv file and send emails
with open(data_file_path, "r") as csvfile:
    reader = csv.reader(csvfile, delimiter=",")
    next(reader)  # Skip the header row
    for row in reader:
        # Check if there are at least 3 values in the row
        if len(row) >= 3:
            recipient_name, _, recipient_email = row[
                :3
            ]  # Only unpack the first 3 values
            email_body = get_text(message_path)
            email_subject = get_text(subject_path)
            send_email(
                email_subject,
                email_body,
                sender,
                recipient_name,
                recipient_email,
                password,
            )
        else:
            print("Invalid data format in CSV row. Skipping the row.")
