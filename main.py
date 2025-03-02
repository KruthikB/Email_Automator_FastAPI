from fastapi import FastAPI, File, UploadFile, Form, HTTPException
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import time
import os
from typing import List
import shutil
import uvicorn
from io import BytesIO
import json

app = FastAPI()

class EmailInput(BaseModel):
    sender_email: str
    sender_password: str
    photo_folder: str
    failed_records: dict

@app.post("/send-emails/")
async def send_emails(
    sender_email: str = Form(...),
    sender_password: str = Form(...),
    photo_folder: str = Form(...),
    photo_suffix: str = Form(...),
    file: UploadFile = File(...),

):
    successful = 0
    failed = 0
    failed_ids = {}
    start_time = time.time()

    temp_file_path = f"temp_{file.filename}"
    with open(temp_file_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    try:
        df = pd.read_excel(temp_file_path)
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Error reading Excel file: {str(e)}")

    name_column = "Name"
    email_column = "Email"

    try:
        session = smtplib.SMTP("smtp.gmail.com", 587)
        session.starttls()
        session.login(sender_email, sender_password)
    except smtplib.SMTPAuthenticationError:
        raise HTTPException(status_code=401, detail="Invalid email credentials")
    except Exception as e:
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

        attached_any = False
        for column in df.columns:
            if column.startswith("Photo"):
                photo_id = row[column]
                if not pd.isna(photo_id):
                    numeric_photo_id = "".join(filter(str.isdigit, str(int((photo_id)))))
                    print(numeric_photo_id)
                    actual_photo_id = f"{photo_suffix}{numeric_photo_id}.jpg"
                    photo_file_path = os.path.join(photo_folder, actual_photo_id)

                    if os.path.exists(photo_file_path):
                        with open(photo_file_path, "rb") as attachment:
                            part = MIMEBase("application", "octet-stream")
                            part.set_payload(attachment.read())
                            encoders.encode_base64(part)
                            part.add_header("Content-Disposition", f"attachment; filename={actual_photo_id}")
                            msg.attach(part)
                            attached_any = True
                    else:
                        failed += 1
                        if email not in failed_ids:
                            failed_ids[email]=list()
                            failed_ids[email].append(actual_photo_id)
                        else:
                            failed_ids[email].append(actual_photo_id)

        if attached_any:
            try:
                session.sendmail(sender_email, email, msg.as_string())
                successful += 1
            except Exception as e:
                failed += 1

    session.quit()
    end_time = time.time()
    execution_time = end_time - start_time
    os.remove(temp_file_path)
    
    print(
        "Successful: ", successful,"Failed: ", failed
    )
    if not failed_ids:
        raise HTTPException(status_code=404, detail="No failed records found")
    
    records = []
    for email, photos in failed_ids.items():
        for photo in photos:
            records.append({"Email": email, "Failed Photo": photo})
    
    df_failed = pd.DataFrame(records)
    output = BytesIO()
    df_failed.to_excel(output, index=False, engine="openpyxl")
    output.seek(0)
    return StreamingResponse(output, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": "attachment; filename=failed_report.xlsx"})
    # return {
    #     "Successful": successful,
    #     "Failed": failed,
    #     "Failed_Photos": failed_ids,
    #     "Execution_Time": execution_time,
    # }

@app.get("/download-failed-report/")
async def download_failed_report(failed_records:dict):
    if not failed_records:
        raise HTTPException(status_code=404, detail="No failed records found")
    
    records = []
    for email, photos in failed_records.items():
        for photo in photos:
            records.append({"Email": email, "Failed Photo": photo})
    
    df_failed = pd.DataFrame(records)
    output = BytesIO()
    df_failed.to_excel(output, index=False, engine="openpyxl")
    output.seek(0)
    return StreamingResponse(output, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": "attachment; filename=failed_report.xlsx"})

if __name__ == "__main__":
    uvicorn.run(app, port=8000)
