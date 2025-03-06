from fastapi import FastAPI, File, UploadFile, Form, HTTPException
from fastapi.responses import StreamingResponse
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
import shutil
import time
from dotenv import load_dotenv
import yaml
from io import BytesIO
from typing import Optional
from utils import convert_to_str

app = FastAPI()
load_dotenv()

config_path = "config.yaml"
with open(config_path, "r") as file:
    config = yaml.safe_load(file)

@app.post("/")
async def send_emails(
    sender_email: str = Form(...,description="Senders email ID"),
    sender_pass: str = Form(...,description="Senders Password"),
    photo_folder: str = Form(...,description="Path to your folder containing the files"),
    photo_suffix: Optional[str] = Form(...,description="default suffix to your file name"),
    filetype : str = Form(...,description="Example: jpg, png, pdf etc."),
    file: UploadFile = File(...),
):
    if sender_email in config["email"]:
        sender_password_env = config["email"][sender_email]
        sender_password = os.getenv(sender_password_env)
        if not sender_password:
            raise HTTPException(status_code=401, detail="Email password not found in environment variables")
    else:
        sender_password = sender_pass

    successful = []
    failed = []
    start_time = time.time()

    temp_file_path = f"temp_{file.filename}"
    with open(temp_file_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    try:
        df = pd.read_excel(temp_file_path)
    except Exception as e:
        os.remove(temp_file_path)
        raise HTTPException(status_code=400, detail=f"Error reading Excel file: {str(e)}")

    name_column = "Name"
    email_column = "Email"
    photo_columns = [col for col in df.columns if col.startswith("PhotoID")]

    try:
        session = smtplib.SMTP("smtp.gmail.com", 587)
        session.starttls()
        session.login(sender_email, sender_password)
    except smtplib.SMTPAuthenticationError:
        os.remove(temp_file_path)
        raise HTTPException(status_code=401, detail="Invalid email credentials")
    except Exception as e:
        os.remove(temp_file_path)
        raise HTTPException(status_code=500, detail=f"SMTP error: {str(e)}")

    for _, row in df.iterrows():
        name = row[name_column]
        email = row[email_column]
        msg = MIMEMultipart()
        msg["From"] = sender_email
        msg["To"] = email
        msg["Subject"] = "Your Photos"
        body = f"Hello {name},\n\nPlease find your photos attached."
        msg.attach(MIMEText(body, "plain"))

        attached_photos = []
        for column in photo_columns:
            photo_id = row[column]
            if not pd.isna(photo_id):
                numeric_photo_id = convert_to_str(photo_id)
                print(numeric_photo_id)
                actual_photo_id = f"{photo_suffix}{numeric_photo_id}.{filetype}"
                photo_file_path = os.path.join(photo_folder, actual_photo_id)
                
                if os.path.exists(photo_file_path):
                    with open(photo_file_path, "rb") as attachment:
                        part = MIMEBase("application", "octet-stream")
                        part.set_payload(attachment.read())
                        encoders.encode_base64(part)
                        part.add_header("Content-Disposition", f"attachment; filename={actual_photo_id}")
                        msg.attach(part)
                        attached_photos.append(numeric_photo_id)

        if attached_photos:
            try:
                session.sendmail(sender_email, email, msg.as_string())
                successful.append([name, email] + attached_photos + [None] * (len(photo_columns) - len(attached_photos)))
                print(f'Email successfully sent to {email} with attachments {attached_photos}')
            except Exception:
                failed.append([name, email] + attached_photos + [None] * (len(photo_columns) - len(attached_photos)))
                print(f'Failed to send mail to {email} with attachments {attached_photos}')
        else:
            failed.append([name, email] + [None] * len(photo_columns))
    
    session.quit()
    os.remove(temp_file_path)

    success_df = pd.DataFrame(successful, columns=[name_column, email_column] + photo_columns)
    failed_df = pd.DataFrame(failed, columns=[name_column, email_column] + photo_columns)
    
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        success_df.to_excel(writer, index=False, sheet_name="Successful")
        failed_df.to_excel(writer, index=False, sheet_name="Failed")
    output.seek(0)
    
    return StreamingResponse(output, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                             headers={"Content-Disposition": "attachment; filename=email_reports.xlsx"})

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, port=8000)
