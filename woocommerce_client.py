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
            raise # Re-raise the exception to be caught in main or higher level

    def get_orders_from_yesterday(self):
        """Fetches all completed orders from the previous day."""
        yesterday = (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')
        
        all_orders = []
        page = 1
        per_page = 100

        logger.info(f"INFO: Fetching completed orders for {yesterday}...")
        try:
            while True:
                response_json = self.wcapi.get("orders", params={
                    "status": "completed",
                    "after": f"{yesterday}T00:00:00",
                    "before": f"{yesterday}T23:59:59",
                    "per_page": per_page,
                    "page": page
                }).json()

                if not isinstance(response_json, list) or not response_json:
                    break
                
                all_orders.extend(response_json)
                page += 1
            logger.info(f"INFO: Successfully fetched {len(all_orders)} orders for {yesterday}.")

        except Exception as e:
            logger.error(f"ERROR: Error fetching orders for {yesterday}: {e}")
            raise # Re-raise to signal a failure in data fetching

        if not all_orders:
            logger.info(f"INFO: No completed orders found for {yesterday}.")
        
        return all_orders