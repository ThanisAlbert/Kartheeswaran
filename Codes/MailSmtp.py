import os
import smtplib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime

class MailSmtp:

    def __init__(self,customercode,customername,cnrefno,partnermail,ccdmail,filepath):

        self.customercode = customercode
        self.customername = customername
        self.cnrefno = cnrefno
        self.partnermail = partnermail
        self.ccdmail = ccdmail
        self.filepath = filepath

    def sendmailtocustomers(self):

        msg = MIMEMultipart('related')
        msg['From'] = 'claims.android@automail.redingtongroup.com'
        msg['To'] = self.partnermail
        msg['Subject'] = f"{self.customercode} {self.customername} {self.cnrefno} {self.partnermail} {self.ccdmail}"

        # Add recipients
        to_addr = [self.partnermail]
        cc_addr = ["thanis.albert@redingtongroup.com"]

        recipients = to_addr + cc_addr

        # Attach HTML content
        html_content = """
                    <html>
                        <body>
                            <h1>This is a test email with HTML content</h1>
                            <p>Hello, This is a test email sent from Python!</p>
                        </body>
                    </html>
                """
        msg.attach(MIMEText(html_content, 'html'))

        # Attach files
        file1 = self.customercode + ".xlsx"
        file2 = self.cnrefno + ".pdf"

        file_paths = [
            os.path.join('D:\\Macro\\RedIndia\\Redbot_2422_CN_SendMail_Customer\\Communication\\', file1),
            os.path.join('D:\\Macro\\RedIndia\\Redbot_2422_CN_SendMail_Customer\\Communication\\', file2)
        ]

        for file_path in file_paths:
            try:
                fname = os.path.basename(file_path)
                with open(file_path, 'rb') as file:
                    part = MIMEApplication(file.read(), Name=fname)
                part['Content-Disposition'] = f'attachment; filename="{fname}"'
                msg.attach(part)
            except Exception as e:
                print(e)

        try:
            # SMTP configuration
            smtp_server = 'smtp.automail.redingtongroup.com'
            smtp_port = 25
            smtp_username = 'claims.android'
            smtp_password = 'tf#31+P40gXZ'

            # Connect to SMTP server
            smtp_conn = smtplib.SMTP(smtp_server, smtp_port)
            smtp_conn.starttls()
            smtp_conn.login(smtp_username, smtp_password)

            # Send email
            smtp_conn.sendmail(msg['From'], recipients, msg.as_string())
            smtp_conn.quit()
            print("Email sent successfully!")

        except Exception as e:
            print(f"Error: {e}")









