import smtplib
import os
import logging
import jdatetime
from datetime import datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.utils import formataddr

logger = logging.getLogger(__name__)

class EmailSender:
    """
    Handles sending emails with attachments.
    """
    def __init__(self, sender_email, sender_password, smtp_server, smtp_port,
                 receiver_to, receiver_cc=None):
        self.sender_email = sender_email
        self.sender_password = sender_password
        self.smtp_server = smtp_server
        self.smtp_port = smtp_port
        self.receiver_to = [email.strip() for email in receiver_to.split(',')] if receiver_to else []
        self.receiver_cc = [email.strip() for email in receiver_cc.split(',')] if receiver_cc else []

    def send_email_report(self, excel_file_path):
        """Sends an email with the generated Excel report as an attachment."""
        
        if not self.sender_email or not self.sender_password or not self.smtp_server or not self.smtp_port:
            logger.warning("WARNING: Email sending skipped due to missing sender/server credentials.")
            return
        
        if not self.receiver_to and not self.receiver_cc:
            logger.warning("WARNING: Email sending skipped as no TO or CC recipients are specified.")
            return

        msg = MIMEMultipart()
        msg['From'] = formataddr(('WooCommerce Report', self.sender_email))
        
        if self.receiver_to:
            msg['To'] = ", ".join(self.receiver_to)
        if self.receiver_cc:
            msg['Cc'] = ", ".join(self.receiver_cc)

        msg['Subject'] = f"گزارش سفارشات ووکامرس - {datetime.now().strftime('%Y-%m-%d')}"

        yesterday_datetime_obj = datetime.now() - timedelta(days=1)
        yesterday_jalali = jdatetime.datetime.fromtimestamp(yesterday_datetime_obj.timestamp()).strftime('%Y/%m/%d')
        
        body = f"با سلام،\n\nفایل اکسل گزارش سفارشات ووکامرس برای روز گذشته ({yesterday_jalali}) پیوست شده است.\n\nبا احترام - واحد انفورماتیک"
        msg.attach(MIMEText(body, 'plain'))

        if excel_file_path and os.path.exists(excel_file_path):
            try:
                with open(excel_file_path, 'rb') as f:
                    attach = MIMEApplication(f.read(), _subtype="xlsx")
                    attach.add_header('Content-Disposition', 'attachment', filename=os.path.basename(excel_file_path))
                    msg.attach(attach)
                logger.info(f"INFO: Excel file '{os.path.basename(excel_file_path)}' successfully attached to email.")
            except Exception as e:
                logger.error(f"ERROR: Error attaching Excel file '{excel_file_path}': {e}. Email will be sent without attachment.")
        else:
            logger.warning(f"WARNING: Excel file '{excel_file_path}' not found or invalid path. Email will be sent without attachment.")

        try:
            with smtplib.SMTP(self.smtp_server, self.smtp_port) as server:
                server.starttls()
                server.login(self.sender_email, self.sender_password)
                
                all_recipients = self.receiver_to + self.receiver_cc
                server.send_message(msg, from_addr=self.sender_email, to_addrs=all_recipients)
                logger.info(f"INFO: Email successfully sent to '{', '.join(self.receiver_to) if self.receiver_to else 'N/A'}' and CC to '{', '.join(self.receiver_cc) if self.receiver_cc else 'N/A'}'.")
        except smtplib.SMTPAuthenticationError:
            logger.error("ERROR: SMTP Authentication Error: Check your email username and password in .env.")
        except smtplib.SMTPConnectError:
            logger.error(f"ERROR: SMTP Connection Error: Could not connect to '{self.smtp_server}' on port '{self.smtp_port}'. Check server address, port, or network.")
        except Exception as e:
            logger.error(f"ERROR: An unexpected error occurred while sending email: {e}")