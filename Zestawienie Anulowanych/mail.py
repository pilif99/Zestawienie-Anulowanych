
import email, smtplib, ssl
import getpass
import data

from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

class Mail:

    def __init__(self):

        subject = "Zestawienie anulowanych"
        sender_email = "pilifznerol.praca@gmail.com"
        #receiver_email = "pilifznerol@gmail.com"
        receiver_email = "j.wachalski@powerbike.pl"
        password = getpass.getpass("Wprowadź hasło:")

        message = MIMEMultipart()
        message["From"] = sender_email
        message["To"] = receiver_email
        message["Subject"] = subject

        filename = "Zestawienie_anulowanych_" + str(data.Data()) + ".xlsx"

        with open(filename, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())

        encoders.encode_base64(part)

        part.add_header(
            "Content-Disposition",
            f"attachment; filename= {filename}",
        )

        message.attach(part)
        text = message.as_string()

        context = ssl.create_default_context()
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
            server.login(sender_email, password)
            server.sendmail(sender_email, receiver_email, text)