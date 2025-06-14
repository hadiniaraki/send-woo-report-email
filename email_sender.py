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

    def send_email_report(self, excel_file_paths):
        """
        Sends an email with the generated Excel reports as attachments.
        """
        
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

        yesterday_dt_obj = datetime.now() - timedelta(days=1)
        yesterday_jalali_dt = jdatetime.datetime.fromgregorian(datetime=yesterday_dt_obj)
        subject_date_str = yesterday_jalali_dt.strftime('%Y-%m-%d')
        body_date_str = yesterday_jalali_dt.strftime('%Y/%m/%d')
        

        msg['Subject'] = f"گزارش سفارشات سایت - {subject_date_str}"
        
        body = f"با سلام،\n\nفایل اکسل گزارش سفارشات سایت برای روز گذشته ({body_date_str}) پیوست شده است.\n\nبا احترام - واحد انفورماتیک"
        msg.attach(MIMEText(body, 'plain'))
        # <--- پایان تغییرات اصلی

        attached_files_count = 0
        for file_path in excel_file_paths:
            if file_path and os.path.exists(file_path):
                try:
                    with open(file_path, 'rb') as f:
                        attach = MIMEApplication(f.read(), _subtype="xlsx")
                        attach.add_header('Content-Disposition', 'attachment', filename=os.path.basename(file_path))
                        msg.attach(attach)
                    logger.info(f"INFO: Excel file '{os.path.basename(file_path)}' successfully attached to email.")
                    attached_files_count += 1
                except Exception as e:
                    logger.error(f"ERROR: Error attaching Excel file '{file_path}': {e}. This file will not be attached.")
            else:
                logger.warning(f"WARNING: Excel file '{file_path}' not found or invalid path. It will not be attached to email.")
        
        if attached_files_count == 0:
            logger.warning("WARNING: No valid Excel files were attached to the email.")

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