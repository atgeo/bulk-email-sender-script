import configparser
import json
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from string import Template

import pandas as pd
from email_validator import validate_email, EmailNotValidError

config = configparser.ConfigParser()
config.read("config.ini")

smtp_host = config["SMTP"]["SMTP_HOST"]
smtp_port = config["SMTP"].getint("SMTP_PORT")
smtp_username = config["SMTP"]["SMTP_USERNAME"]
smtp_password = config["SMTP"]["SMTP_PASSWORD"]

subject = config["EmailSettings"]["SUBJECT"]
from_email = smtp_username


def send_email(receiver_email: str, receiver_name: str) -> None:
    message = MIMEMultipart()
    message["From"] = from_email
    message["To"] = receiver_email
    message["Subject"] = subject

    body_template = config["EmailSettings"]["BODY_TEMPLATE"]
    body = Template(body_template).substitute(name=receiver_name)
    message.attach(MIMEText(body, "html"))

    filename = config["EmailSettings"]["ATTACHMENT_FILENAME"]
    with open(filename, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f"attachment; filename= {filename}")
    message.attach(part)

    with smtplib.SMTP(smtp_host, smtp_port) as server:
        server.starttls()
        server.login(smtp_username, smtp_password)
        server.sendmail(from_email, receiver_email, message.as_string())
        print("Email sent successfully!")


def save_position(row: int, col: int) -> None:
    with open("position.json", "w") as f:
        json.dump({"row": row, "col": col}, f)


def load_position() -> tuple[int, int]:
    try:
        with open("position.json", "r") as f:
            position = json.load(f)
    except FileNotFoundError:
        position = {"row": 0, "col": 0}
        save_position(position["row"], position["col"])

    return position.get("row", 0), position.get("col", 0)


def main() -> None:
    input_filename = config["INPUT"]["FILENAME"]
    df = pd.read_excel(input_filename, header=None)

    row, col = load_position()

    for i in range(row, len(df)):
        for j in range(col, len(df.columns)):
            save_position(i, j)

            email_address = df.iloc[i, j]
            try:
                validate_email(email_address)
                print(f"Row {i + 1}, Column {chr(ord('A') + j)}: {email_address} is a valid email address.")
            except EmailNotValidError:
                print(f"Row {i + 1}, Column {chr(ord('A') + j)}: {email_address} is not a valid email address.")

            receiver_email = email_address
            receiver_name = input("Enter the receiver's name: ")

            input("Press Enter to send the email...")

            send_email(receiver_email, receiver_name)


if __name__ == "__main__":
    main()
