import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd

def send_email(subject, body, to_emails):
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587
    email = ''
    password = ''
    
    try:
        # Establish a secure session with the server
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(email, password)
        
        for recipient in to_emails:
            msg = MIMEMultipart()
            msg['From'] = email
            msg['To'] = recipient
            msg['Subject'] = subject
            msg.attach(MIMEText(body, 'html')) 
            
            server.sendmail(email, recipient, msg.as_string())
            print(f"Email sent to {recipient}")
        
        
        server.quit()
    
    except smtplib.SMTPAuthenticationError as e:
        print(f"Authentication error: {e.smtp_code} - {e.smtp_error.decode()}")
        print("Ensure that your email and password are correct, and that you've enabled access for less secure apps.")
    
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":

    portfolioLink = "Portfolio" 
    url = "" 
    portfolio = f'<a href="{url}">{portfolioLink}</a>'

    instagramLink = "Instagram" 
    url = "https://instagram.com/edsthecreator" 
    instagram = f'<a href="{url}">{instagramLink}</a>'

    subject = 'Looking to work with you!'
    body = f"""
    <html>
    <body>
        <p>Hi,</p>

        <p>I hope this message finds you well.</p>

        <p>My name is Edward Jones, and I’m a passionate graphic designer with a proven track record of delivering quality work to over 200 clients. I take pride in my ability to provide stunning visuals with a quick turnaround time.</p>

        <p>I would love the opportunity to collaborate with you on any upcoming projects where you might need creative artwork. Whether it's for branding, marketing materials, or digital content, I'm confident that my skills and experience can add value to your initiatives.</p>

        <p>Looking forward to the possibility of working together!</p>

        <p>{portfolio}<br>{instagram}</p>

        <p>Best regards,<br> 
        <br>
        (856)<br>
        <a href="mailto:"></a></p>
    </body>
    </html>
    """

    # Read email addresses from Excel file
    excel_file_path = 'emails.xlsx' 
    df = pd.read_excel(excel_file_path, header=None) 
    to_emails = df[1][189:].tolist()  # df[column][row:].tolist()
    # Send emails
    send_email(subject, body, to_emails)
    print("Emails sent.")
