import xlrd 
import email, smtplib, ssl

from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def main():
    print("----- Certi-----")
    print("-----Sending Certificates------")

    path="Input.xlsx"
    output_dir_path = "/home/ash/Projects/13-Certificate-Generator-with-QR-code-in-python/Output/"
    #open Excel File
    # To open Workbook 
    wb = xlrd.open_workbook(path) 
    sheet = wb.sheet_by_index(0)
    #ROWS and Columns
    R=sheet.nrows
    C=sheet.ncols

    for i in range(1,R):
        comp=sheet.cell_value(i,0)
        name=sheet.cell_value(i,1)
        email=sheet.cell_value(i,2)
        subject = "An email with attachment from Python"
        body = "This is an email sent by certi"
        sender_email = ""
        receiver_email = email
        password =""

        message = MIMEMultipart()
        message["From"] = sender_email
        message["To"] = receiver_email
        message["Subject"] = subject
        message["Bcc"] = receiver_email

        message.attach(MIMEText(body, "plain"))
        filename = output_dir_path+name+str(i)+".pdf"

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
        
       

    print("Sucessfully sent certificates!")
    input("Please Enter to Exit")

if __name__ == "__main__":
    main()