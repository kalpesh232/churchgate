from flask import Flask, render_template, request
import pandas as pd
import smtplib
from email.message import EmailMessage
import time
import threading
import schedule

app = Flask(__name__)

email_counts = {}  # Dictionary to store email counts

def send_emails(sender_email, sender_pass):
    count_sent_emails = 0
    global email_counts
    df = pd.read_excel("For email.xlsx")
    receivers_email = df["EMAIL_ID"].values
    sub = "Test Mail"
    attach_files = df["Files to be attached"]
    name = df["NAME"].values

    for (a, b, c) in zip(receivers_email, attach_files, name):
        msg = EmailMessage()
        files = [f"{b}.xlsx"]

        for file in files:
            with open(file, 'rb') as f:
                file_data = f.read()

            msg['From'] = sender_email
            msg['To'] = a
            msg['Subject'] = sub
            msg.set_content(f"hello {c}! I have something for you.")
            msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=f"{b}.pdf")

            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                smtp.login(sender_email, sender_pass)
                smtp.send_message(msg)

            count_sent_emails += 1  # Increment the count for each email sent

        # Update email counts
        key = f"{sender_email} -> {a}"
        email_counts[key] = email_counts.get(key, 0) + 1

    print("All mail sent!")
    return count_sent_emails  # Return the count

def schedule_email_task(sender_email, sender_pass):
    # Schedule the send_emails function every 2 minutes
    # schedule.every(2).minutes.do(send_emails, sender_email=sender_email, sender_pass=sender_pass)
    while True:
        count_sent_emails = send_emails(sender_email, sender_pass)
        print(f"{count_sent_emails} email(s) sent in the last 2 minutes")
        schedule.run_pending()
        time.sleep(120)  # Sleep for 2 minutes
        # time.sleep(1)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        sender_email = request.form['email']
        sender_pass = request.form['password']
        
        # Start the schedule_email_task function in a separate thread with dynamic inputs
        threading.Thread(target=schedule_email_task, args=(sender_email, sender_pass)).start()

        # You can also call send_emails here directly if you want to send emails immediately
        # send_emails(sender_email, sender_pass)

    return render_template('index.html')

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
