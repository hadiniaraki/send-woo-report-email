import logging
import sys
from config import Config
from woocommerce_client import WooCommerceClient
from excel_reporter import ExcelReporter
from email_sender import EmailSender

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
            # excel_reporter.create_excel_report now returns two values:
            # main_report_path: path to WooCommerce_Orders_YYYY-MM-DD.xlsx
            # templated_report_paths: a list containing path(s) to tis-YYYY-MM-DD.xlsx
            main_report_path, templated_report_paths = excel_reporter.create_excel_report(orders)
            
            # Collect all generated file paths for email attachment
            all_excel_files_to_attach = []
            if main_report_path:
                all_excel_files_to_attach.append(main_report_path)
            if templated_report_paths: # This will be a list, iterate to add
                all_excel_files_to_attach.extend(templated_report_paths)

            if all_excel_files_to_attach:
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