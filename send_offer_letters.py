import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import pandas as pd
import os

# Paths
excel_path = "/Users/pavankumar.b/Library/Containers/com.microsoft.Excel/Data/Downloads/offer-ParvaM.xlsx"  # Path to your Excel file
output_dir = "Generated_Documents"  # Directory with generated PDFs

# Email configuration
smtp_server = "smtp.gmail.com"
smtp_port = 587
email_address = "pavan.parvam@gmail.com"  # Your email
email_password = "fxdvxbjinxphrirz"  # Your email password
cc_addresses = ["hr@parvamm.com", "directors@parvamm.com"]  # Default CC recipients


# Read Excel data
data = pd.read_excel(excel_path)

# Send emails
server = smtplib.SMTP(smtp_server, smtp_port)
server.starttls()
server.login(email_address, email_password)

for index, row in data.iterrows():
    name = row["Name"]
    email = row["Email"]
    pdf_path = os.path.join(output_dir, f"{name}_Offer_Letter.pdf")
    
    if not os.path.exists(pdf_path):
        print(f"PDF not found for {name}, skipping...")
        continue
    
    # Create email
    msg = MIMEMultipart()
    msg["From"] = email_address
    msg["To"] = email
    msg["Cc"] = ", ".join(cc_addresses)  # Add CC recipients
    msg["Subject"] = "Conditional Offer Letter for Trainee Engineer Position"
    
    # Email body
    body = f"Dear {name},\n\nPlease find attached your Conditional Offer Letter.\n\nBest regards,\nParvam HR Team"
    msg.attach(MIMEText(body, "plain"))
    
    # Attach PDF
    with open(pdf_path, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(pdf_path)}")
        msg.attach(part)
    
    recipients = [email] + cc_addresses  # Include CC recipients
    server.sendmail(email_address, recipients, msg.as_string())

    print(f"Email sent to {name} ({email})")

server.quit()
print("All emails sent successfully!")
