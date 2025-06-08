import logging
import sys
from config import Config
from woocommerce_client import WooCommerceClient
from excel_reporter import ExcelReporter
from email_sender import EmailSender
from datetime import datetime, timedelta
import jdatetime

# --- Logging Configuration ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def main():
    """
    Main function to orchestrate fetching orders, generating report, and sending email.
    Handles critical errors during script execution.
    """
    logger.info("INFO: Script execution started.")

    # 1. Validate Configurations
    if not Config.validate_woo_config():
        sys.exit(1) # Exit if critical WooCommerce config is missing

    email_config_valid = Config.validate_email_config() # Email config is warned, not critical exit

    try:
        # Calculate yesterday's date in Gregorian for fetching orders and
        # convert to Jalali for the Excel filename.
        yesterday_dt = datetime.now() - timedelta(days=1)
        
        # Convert Gregorian yesterday's date to Jalali for the filename
        jalali_date_for_filename = jdatetime.datetime.fromtimestamp(yesterday_dt.timestamp())
        formatted_jalali_date_for_filename = jalali_date_for_filename.strftime('%Y-%m-%d')

        # 2. Initialize Clients/Components
        woo_client = WooCommerceClient(
            base_url=Config.WOO_BASE_URL,
            consumer_key=Config.WOO_CONSUMER_KEY,
            consumer_secret=Config.WOO_CONSUMER_SECRET
        )
        excel_reporter = ExcelReporter()
        email_sender = EmailSender(
            sender_email=Config.EMAIL_SENDER,
            sender_password=Config.EMAIL_PASSWORD,
            smtp_server=Config.SMTP_SERVER,
            smtp_port=Config.SMTP_PORT,
            receiver_to=Config.EMAIL_RECEIVER_TO,
            receiver_cc=Config.EMAIL_RECEIVER_CC
        )

        # 3. Execute Workflow
        orders = woo_client.get_orders_from_yesterday()

        if orders:
            excel_file = excel_reporter.create_excel_report(orders)
            if excel_file:
                if email_config_valid: # Only try to send email if config is valid
                    # Pass the list of all file paths to the email sender
                    email_sender.send_email_report(all_excel_files_to_attach)
                else:
                    logger.warning("WARNING: Email configuration incomplete. Reports generated but not sent via email.")
            else:
                logger.error("ERROR: No Excel files were created, skipping email sending.")
        else:
            logger.info("INFO: No orders found for the previous day. No Excel report or email will be generated.")

    except Exception as e:
        logger.critical(f"CRITICAL: Script terminated due to a critical error: {e}", exc_info=True)
        sys.exit(1) # Exit with an error code

if __name__ == "__main__":
    main()