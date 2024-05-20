import os
import smtplib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime
from datetime import date


class MailSmtp:

    def __init__(self,customercode,customername,cnrefno,partnermail,ccdmail,filepath,scheme_details):
        self.customercode = customercode
        self.customername = customername
        self.cnrefno = cnrefno
        self.partnermail = partnermail
        self.ccdmail = ccdmail
        self.filepath = filepath+"\\"
        self.schemedetails = scheme_details

    def sendmailtocustomers(self):

        msg = MIMEMultipart('related')
        msg['From'] = 'claims.android@automail.redingtongroup.com'
        msg['To'] = self.partnermail
        cc_list = self.ccdmail.split(';')
        msg['Cc'] = ", ".join(cc_list)
        msg['Subject'] = f"CN Communication - {self.schemedetails}"

        # Get today's date as a datetime object
        today = date.today()
        # Create a new date object with the desired format (year, month, day)
        desired_date = date(year=today.year, month=today.month, day=today.day)
        # Format the desired date as a string in the specified format (DD/MM/YYYY)
        formatted_date = desired_date.strftime("%d/%m/%Y")

        # Add recipients
        to_addr = [self.partnermail]
        cc_addr = cc_list

        recipients = to_addr + cc_addr

        # Attach HTML content
        html_content = f"""
    <html>
        <body>
            Dear Partner, 
            <br><br>
            For {self.customername}
            <br><br>
            Pleased to release the CN for {self.schemedetails}
            <br><br>
            The Credit note has been generated on {formatted_date} in our system, the details for the same are attached in the file.
            <br><br>
            Kindly contact Redington account manager for any clarification. 
            <br><br>
            *If any discrepancies are found, Please report with in 30 days from the date of communication, If not reported same will not be considered later. 
            <br><br>
            Regards, 
            <br>
            Team Redington 
        </body>
    </html>
    """

        msg.attach(MIMEText(html_content, 'html'))

        # Attach files
        file1 = self.customercode + ".xlsx"
        file2 = self.cnrefno + ".pdf"
        file3 = self.schemedetails +".pdf"
        file3 = str(file3).replace("\n","")
        print(file3)

        file_paths = [
            os.path.join(self.filepath, file1),
            os.path.join(self.filepath, file2),
            os.path.join(self.filepath, file3)
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









