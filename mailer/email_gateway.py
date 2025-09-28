from fastapi import FastAPI, HTTPException
from pydantic import BaseModel, EmailStr
from typing import List, Optional
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import json

app = FastAPI()

# Define your SMTP configurations
f = open("configs.json", "rb")
SMTP_CONFIGS = json.loads(f.read())
f.close()

class EmailRequest(BaseModel):
    project: str  # which project SMTP config to use
    subject: str11
    recipients: List[EmailStr]
    body: str
    sender: Optional[EmailStr] = None  # override default sender if needed


def send_email(config: dict, subject: str, recipients: List[str], body: str, sender: Optional[str] = None):
    try:
        msg = MIMEMultipart()
        msg["From"] = sender or config["MAIL_DEFAULT_SENDER"]
        msg["To"] = ", ".join(recipients)
        msg["Subject"] = subject
        msg.attach(MIMEText(body, "html"))

        if config["MAIL_USE_SSL"]:
            server = smtplib.SMTP_SSL(config["MAIL_SERVER"], config["MAIL_PORT"])
        else:
            server = smtplib.SMTP(config["MAIL_SERVER"], config["MAIL_PORT"])
            server.starttls()

        server.login(config["MAIL_USERNAME"], config["MAIL_PASSWORD"])
        server.sendmail(msg["From"], recipients, msg.as_string())
        server.quit()
        return True
    except Exception as e:
        print(f"Error sending email: {e}")
        return False


@app.post("/send-email")
def send_email_api(request: EmailRequest):
    project = request.project
    if project not in SMTP_CONFIGS:
        raise HTTPException(status_code=400, detail="Invalid project name")

    config = SMTP_CONFIGS[project]

    success = send_email(
        config=config,
        subject=request.subject,
        recipients=request.recipients,
        body=request.body,
        sender=request.sender,
    )

    if success:
        return {"status": "Email sent successfully!"}
    else:
        raise HTTPException(status_code=500, detail="Failed to send email")
