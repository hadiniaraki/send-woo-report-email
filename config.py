import os
from dotenv import load_dotenv
import logging

logger = logging.getLogger(__name__)

load_dotenv()

class Config:
    """
    Manages application configuration loaded from environment variables.
    """
    WOO_BASE_URL = os.getenv("WOOCOMMERCE_BASE_URL")
    WOO_CONSUMER_KEY = os.getenv("WOOCOMMERCE_CONSUMER_KEY")
    WOO_CONSUMER_SECRET = os.getenv("WOOCOMMERCE_CONSUMER_SECRET")

    EMAIL_SENDER = os.getenv("EMAIL_SENDER")
    EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")
    EMAIL_RECEIVER_TO = os.getenv("EMAIL_RECEIVER_TO")
    EMAIL_RECEIVER_CC = os.getenv("EMAIL_RECEIVER_CC")
    SMTP_SERVER = os.getenv("SMTP_SERVER")
    SMTP_PORT = int(os.getenv("SMTP_PORT", 587)) # Default to 587 if not set

    @classmethod
    def validate_woo_config(cls):
        """Validates WooCommerce API configuration."""
        if not all([cls.WOO_BASE_URL, cls.WOO_CONSUMER_KEY, cls.WOO_CONSUMER_SECRET]):
            logger.critical("FATAL: Missing one or more WooCommerce API environment variables. Please check .env file.")
            return False
        return True

    @classmethod
    def validate_email_config(cls):
        """Validates Email sending configuration."""
        if not all([cls.EMAIL_SENDER, cls.EMAIL_PASSWORD, cls.SMTP_SERVER, cls.SMTP_PORT]):
            logger.warning("WARNING: One or more email environment variables (sender or server) are missing. Email sending might not function correctly.")
            return False # Return False to indicate potential issue, but don't exit
        if not cls.EMAIL_RECEIVER_TO and not cls.EMAIL_RECEIVER_CC:
            logger.warning("WARNING: Email sending skipped as no TO or CC recipients are specified in .env.")
            return False
        return True

# Initialize Config at import time for easy access
# Config.validate_woo_config() # Optional: validate immediately on import
# Config.validate_email_config() # Optional: validate immediately on import