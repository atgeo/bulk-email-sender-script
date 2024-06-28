import configparser
import pandas as pd
import json
from email_validator import validate_email, EmailNotValidError
import smtplib
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.mime.multipart import MIMEMultipart
from string import Template

# Load configuration from ini file
config = configparser.ConfigParser()
config.read('config.ini')

smtp_host = config['SMTP']['SMTP_HOST']
smtp_port = config['SMTP'].getint('SMTP_PORT')
smtp_username = config['SMTP']['SMTP_USERNAME']
smtp_password = config['SMTP']['SMTP_PASSWORD']

subject = config['EmailSettings']['SUBJECT']
from_email = smtp_username


def send_email(receiver_email, receiver_name):
    message = MIMEMultipart()
    message["From"] = from_email
    message["To"] = receiver_email
    message["Subject"] = subject

    placeholders = {
        'name': receiver_name
    }

    # Create the email body by substituting placeholders
    body_template = config['EmailSettings']['BODY_TEMPLATE']
    template = Template(body_template)
    body = template.substitute(placeholders)

    message.attach(MIMEText(body, "html"))

    # Attach a file (example: resume.pdf)
    filename = config['EmailSettings']['ATTACHMENT_FILENAME']
    attachment = open(filename, 'rb')

    part = MIMEBase('application', 'octet-stream')
    part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename= {filename}')
    message.attach(part)

    server = None

    try:
        # Start TLS encryption
        server = smtplib.SMTP(smtp_host, smtp_port)
        server.starttls()

        # Login to Gmail
        server.login(smtp_username, smtp_password)

        # Send email
        server.sendmail(from_email, receiver_email, message.as_string())
        print("Email sent successfully!")

    except Exception as e:
        print(f"Failed to send email. Error: {str(e)}")

    finally:
        # Quit the SMTP session
        if server:
            server.quit()


# Function to save the current position to a file
def save_position(row, col):
    with open('position.json', 'w') as f:
        json.dump({'row': row, 'col': col}, f)


# Function to load the last saved position from a file
def load_position():
    try:
        with open('position.json', 'r') as f:
            position = json.load(f)
    except FileNotFoundError:
        position = {'row': 0, 'col': 0}
        save_position(position['row'], position['col'])

    return position.get('row', 0), position.get('col', 0)


def main():
    input_filename = config['INPUT']['FILENAME']
    df = pd.read_excel(input_filename, header=None)

    row, col = load_position()

    # Iterate over each row and column in the DataFrame
    for i in range(row, len(df)):
        for j in range(col, len(df.columns)):
            save_position(i, j)

            email_address = df.iloc[i, j]
            try:
                validate_email(email_address)
                print(f"Row {i + 1}, Column {chr(ord('A') + j)}: {email_address} is a valid email address.")
            except EmailNotValidError as e:
                print(f"Row {i + 1}, Column {chr(ord('A') + j)}: {email_address} is not a valid email address.")

            receiver_email = email_address
            receiver_name = input("Enter the receiver's name: ")

            input("Press Enter to send the email...")

            send_email(receiver_email, receiver_name)


if __name__ == '__main__':
    main()
