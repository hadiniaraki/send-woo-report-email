# excel_reporter.py
import pandas as pd
from datetime import datetime
import os
import logging
import jdatetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

logger = logging.getLogger(__name__)

class ExcelReporter:
    """
    Handles the creation and styling of Excel reports from order data.
    """
    def create_excel_report(self, orders_data, report_date_jalali_str):
        """Processes order data and generates an Excel report with styling."""
        processed_data = []
        for order in orders_data:
            try:
                item_details = []
                for item in order.get('line_items', []):
                    item_name = item['name']
                    quantity = item.get('quantity', 0)
                    unit_price = float(item.get('price', 0))

                    refunded_qty_for_this_item = 0
                    for refund in order.get('refunds', []):
                        for refunded_item in refund.get('line_items', []):
                            if refunded_item.get('product_id') == item.get('product_id') and \
                               refunded_item.get('variation_id', 0) == item.get('variation_id', 0):
                                refunded_qty_for_this_item += refunded_item.get('qty', 0)
                    
                    final_quantity = quantity - refunded_qty_for_this_item
                    item_details.append({
                        "name": item_name,
                        "quantity": final_quantity,
                        "unit_price": unit_price
                    })

                item_names_str = "\n".join([detail['name'] for detail in item_details])
                item_quantities_str = "\n".join([str(detail['quantity']) for detail in item_details])
                item_unit_prices_str = "\n".join([str(detail['unit_price']) for detail in item_details])

                order_refund_total = sum(float(refund.get('total', 0)) for refund in order.get('refunds', []))

                created_datetime_obj = datetime.strptime(order['date_created'], '%Y-%m-%dT%H:%M:%S')
                jalali_date = jdatetime.datetime.fromtimestamp(created_datetime_obj.timestamp())
                formatted_jalali_date = jalali_date.strftime('%Y/%m/%d %H:%M:%S')

                custom_order_number = next(
                    (meta['value'] for meta in order.get('meta_data', []) if meta['key'] == '_wc_order_number' or meta['key'] == '_order_number'),
                    order.get('id')
                )

                order_row = {
                    "شماره سفارش": custom_order_number,
                    "تاریخ سفارش (شمسی)": formatted_jalali_date,
                    "نام": order.get('billing', {}).get('first_name', ''),
                    "نام خانوادگی": order.get('billing', {}).get('last_name', ''),
                    "آدرس": f"{order.get('billing', {}).get('address_1', '')} {order.get('billing', {}).get('address_2', '')}".strip(),
                    "شهر": order.get('billing', {}).get('city', ''),
                    "کد پستی": order.get('billing', {}).get('postcode', ''),
                    "تلفن": order.get('billing', {}).get('phone', ''),
                    "عنوان روش پرداخت": order.get('payment_method_title', ''),
                    "مبلغ تخفیف": float(order.get('discount_total', 0)),
                    "مجموع مبلغ سفارش": float(order.get('total', 0)),
                    "روش حمل و نقل": order.get('shipping_lines', [{}])[0].get('method_title', '') if order.get('shipping_lines') else '',
                    "مبلغ حمل و نقل": float(order.get('shipping_lines', [{}])[0].get('total', 0)) if order.get('shipping_lines') else 0,
                    "مبلغ استرداد کل سفارش": order_refund_total,
                    "مجموع نهایی سفارش (پس از کسر استرداد)": float(order.get('total', 0)) - order_refund_total,
                    "نام آیتم‌ها": item_names_str,
                    "تعداد آیتم‌ها (- استرداد)": item_quantities_str,
                    "قیمت واحد آیتم‌ها": item_unit_prices_str,
                    "مجموع هزینه آیتم‌ها": sum(float(item.get('total', 0)) for item in order.get('line_items', []))
                }
                processed_data.append(order_row)
            except Exception as e:
                logger.error(f"ERROR: Error processing order {order.get('id', 'N/A')}: {e}. This order was skipped.")
                continue

        df = pd.DataFrame(processed_data)
        
        df = df.sort_values(by="تاریخ سفارش (شمسی)", ascending=True)

        excel_filename = f"site_Orders_{report_date_jalali_str}.xlsx"
        try:
            df.to_excel(excel_filename, index=False, engine='openpyxl')

            workbook = load_workbook(excel_filename)
            sheet = workbook.active

            header_fill = PatternFill(start_color="CCE0F0", end_color="CCE0F0", fill_type="solid")
            header_font = Font(bold=True)
            item_name_fill = PatternFill(start_color="F0FFF0", end_color="F0FFF0", fill_type="solid")
            wrap_text_alignment = Alignment(wrapText=True, vertical='top')

            for col_idx, cell in enumerate(sheet[1], 1):
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal='center', vertical='center')

                max_length = 0
                column = get_column_letter(col_idx)
                for cell_in_col in sheet[column]:
                    try:
                        if cell_in_col.value is not None:
                            content_lines = str(cell_in_col.value).split('\n')
                            for line in content_lines:
                                if len(line) > max_length:
                                    max_length = len(line)
                    except:
                        pass
                adjusted_width = (max_length + 2)
                if adjusted_width > 70:
                    adjusted_width = 70
                sheet.column_dimensions[column].width = adjusted_width

            sheet.freeze_panes = sheet['A2']

            for row in sheet.iter_rows(min_row=2):
                for cell in row:
                    if cell.column_letter in [
                        get_column_letter(df.columns.get_loc("نام آیتم‌ها") + 1),
                        get_column_letter(df.columns.get_loc("تعداد آیتم‌ها (- استرداد)") + 1),
                        get_column_letter(df.columns.get_loc("قیمت واحد آیتم‌ها") + 1),
                        get_column_letter(df.columns.get_loc("آدرس") + 1)
                    ]:
                        cell.alignment = wrap_text_alignment

                    if cell.column_letter == get_column_letter(df.columns.get_loc("نام آیتم‌ها") + 1):
                        cell.fill = item_name_fill
            
            workbook.save(excel_filename)
            logger.info(f"INFO: Excel file '{excel_filename}' generated and styled successfully.")
            return excel_filename
        except Exception as e:
            logger.error(f"ERROR: Error creating or styling Excel file '{excel_filename}': {e}")
            return None