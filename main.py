from flask import Flask, render_template, request
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd
from werkzeug.utils import secure_filename
import os

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'

# Ensure the 'uploads' directory exists
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

# Variabili per l'indirizzo email e la password del mittente
SENDER_EMAIL = 'nicolaheavy@gmail.com'
SENDER_PASSWORD = 'fislfmdygnraoayb'


def send_email(recipient_email, subject, body):
    msg = MIMEMultipart()
    msg['From'] = SENDER_EMAIL
    msg['To'] = recipient_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.send_message(msg)


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        recipient_email = request.form['recipient_email']
        subject = request.form['subject']
        body = request.form['body']
        textarea_emails = request.form['textarea_emails']

        # Converti gli indirizzi email dalla textarea in una lista
        emails_from_textarea = [email.strip() for email in textarea_emails.split('\n') if email.strip()]

        # Unisci gli indirizzi email dalla textarea con quelli dal file (se presente)
        recipient_emails = set(emails_from_textarea)

        file = request.files.get('file')
        if file:
            filename = secure_filename(file.filename)
            file_path = os.path.join(UPLOAD_FOLDER, filename)
            file.save(file_path)

            if file.filename.endswith('.xlsx'):
                df = pd.read_excel(file_path)
            elif file.filename.endswith('.csv'):
                df = pd.read_csv(file_path)
            else:
                return "Invalid file format!"

            # Aggiungi gli indirizzi email dal file alla lista
            emails_from_file = set(df.iloc[:, 0].dropna())
            recipient_emails.update(emails_from_file)

        # Invia email a tutti gli indirizzi email raccolti
        for recipient_email in recipient_emails:
            send_email(recipient_email, subject, body)

        return "Email sent successfully!"
    return render_template('index.html')


if __name__ == '__main__':
    app.run(debug=True)
