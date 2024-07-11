import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# Email server configuration
SMTP_SERVER = 'smtp.office365.com'  # e.g., smtp.gmail.com for Gmail
SMTP_PORT = 587  # For TLS
EMAIL_ADDRESS = 'you mail id'
EMAIL_PASSWORD = 'Password'  #if you are using gmail then Generate App Password

# Load the list of employees
df = pd.read_csv(r'D:\Bulk Mail\employees.csv')   #csv File Path

def send_email(to_email, cc_email, recipient_name, new_email, password, attachment_path):
    # Create the email
    msg = MIMEMultipart()
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = to_email
    msg['Cc'] = cc_email  # Add CC email
    msg['Subject'] = 'Your Login Credentials'
    
    # Email body with recipient's name
    body = f"""Hello {recipient_name},

    
    Here are your login credentials:

New Email: {new_email}
Password: {password}


For any queries or assistance, you can reach out to us via phone at ************  Alternatively, you can also contact us by writing to the aforementioned email address.

Best regards,
Your Name  

    
    msg.attach(MIMEText(body, 'plain'))
    
    # Attachment
    if attachment_path:
        part = MIMEBase('application', 'octet-stream')
        with open(attachment_path, 'rb') as attachment:
            part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header(
            'Content-Disposition',
            f'attachment; filename= {attachment_path.split("/")[-1]}',
        )
        msg.attach(part)
    
    # Send the email
    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
            server.sendmail(EMAIL_ADDRESS, [to_email] + [cc_email], msg.as_string())
            print(f"Email sent to {to_email} with CC to {cc_email}")
    except Exception as e:
        print(f"Error sending email to {to_email}: {str(e)}")

# Path to the attachment file
attachment_path = 'D:/Bulk Mail/f'  #Attachment Path if any

# Send emails to all employees
for index, row in df.iterrows():
    send_email(row['ToEmail'], row['CcEmail'], row['Name'], row['NewEmail'], row['Password'], attachment_path)