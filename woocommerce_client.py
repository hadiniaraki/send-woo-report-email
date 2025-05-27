from woocommerce import API
from datetime import datetime, timedelta
import logging

logger = logging.getLogger(__name__)

class WooCommerceClient:
    """
    Handles interactions with the WooCommerce API.
    """
    def __init__(self, base_url, consumer_key, consumer_secret):
        try:
            self.wcapi = API(
                url=base_url,
                consumer_key=consumer_key,
                consumer_secret=consumer_secret,
                version="wc/v3"
            )
            logger.info("INFO: WooCommerce API client initialized.")
        except Exception as e:
            logger.critical(f"FATAL: Error configuring WooCommerce API: {e}")
            raise  # Re-raise the exception to be caught in main or higher level

    def get_orders_from_yesterday(self):  # اصلاح به متد نمونه (self)
        """Fetches all completed and processing orders from the previous day from the WooCommerce API."""
        yesterday = (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')
        
        all_orders = []
        page = 1
        per_page = 100

        try:
            while True:
                response_json = self.wcapi.get("orders", params={  # استفاده از self.wcapi
                    "status": "completed,processing",  # رشته با مقادیر جدا شده توسط کاما
                    "after": f"{yesterday}T00:00:00",
                    "before": f"{yesterday}T23:59:59",
                    "per_page": per_page,
                    "page": page
                }).json()

                if not isinstance(response_json, list) or not response_json:
                    break
                
                all_orders.extend(response_json)
                page += 1

        except Exception as e:
            logger.error(f"ERROR: Error fetching orders for {yesterday}: {e}")
            raise

        if not all_orders:
            logger.info(f"INFO: No completed or processing orders found for {yesterday}.")
        else:
            completed_count = len([order for order in all_orders if order.get('status') == 'completed'])
            processing_count = len([order for order in all_orders if order.get('status') == 'processing'])
            logger.info(f"INFO: Found {completed_count} completed orders and {processing_count} processing orders for {yesterday}.")

        return all_orders