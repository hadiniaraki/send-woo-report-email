# excel_reporter.py
import pandas as pd
from datetime import datetime
import os
import logging
import jdatetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
import shutil # For copying the template file

logger = logging.getLogger(__name__)

class ExcelReporter:
    """
    Handles the creation and styling of Excel reports from order data.
    """
    def create_excel_report(self, orders_data):
        """Processes order data and generates an Excel report with styling."""
        processed_data = []
        
        # --- Prepare the templated Excel file (tis-{تاریخ شمسی}.xlsx) ONCE ---
        template_file = "tis.xlsx"
        templated_output_filename = None # Initialize to None

        # Get current Jalali date for the templated filename
        current_jalali_date_str = jdatetime.datetime.now().strftime('%Y-%m-%d')
        templated_output_filename_base = f"tis-{current_jalali_date_str}.xlsx"
        
        workbook_template = None
        sheet_body = None
        current_row_for_template = 0 

        if os.path.exists(template_file):
            try:
                # Ensure the output filename is unique if multiple runs happen on the same day
                counter = 1
                unique_templated_output_filename = templated_output_filename_base
                while os.path.exists(unique_templated_output_filename):
                    unique_templated_output_filename = f"tis-{current_jalali_date_str}_{counter}.xlsx"
                    counter += 1
                templated_output_filename = unique_templated_output_filename

                shutil.copy(template_file, templated_output_filename) # Copy template
                workbook_template = load_workbook(templated_output_filename)
                
                if "بدنه" in workbook_template.sheetnames:
                    sheet_body = workbook_template["بدنه"]
                    
                    # --- IMPORTANT: Adjust START_ROW_BODY based on your tis.xlsx template ---
                    # If your headers are in row 1, data starts from row 2.
                    START_ROW_BODY = 2 
                    current_row_for_template = START_ROW_BODY

                else:
                    logger.warning(f"WARNING: Sheet 'بدنه' not found in '{template_file}'. Templated report will not be generated.")
                    workbook_template = None # Prevent further operations on missing sheet
            except Exception as e:
                logger.error(f"ERROR: Could not copy or load template file '{template_file}': {e}. Templated report will not be generated.")
                workbook_template = None # Ensure it's None if loading failed
        else:
            logger.error(f"ERROR: Template file '{template_file}' not found. Cannot create templated report.")


        # --- Define column mappings for the 'بدنه' sheet (based on your provided order from A to R) ---
        # These MUST be verified with the actual layout of your 'tis.xlsx' template.
        # Column A (1): شماره صورتحساب
        # Column B (2): شناسه کالا/خدمت
        # Column C (3): شرح کالا/خدمت
        COL_DESCRIPTION = 3 
        # Column D (4): تعداد/مقدار
        COL_QUANTITY = 4    
        # Column E (5): واحد اندازه گیری
        COL_UNIT = 5        
        # Column F (6): مبلغ واحد
        COL_UNIT_PRICE = 6  
        # Column G (7): میزان ارز
        # Column H (8): نوع ارز
        # Column I (9): نرخ برابری ارز با ریال
        # Column J (10): مبلغ تخفیف
        COL_DISCOUNT = 10   
        # Column K (11): نرخ مالیات بر ارزش افزوده
        COL_VAT_RATE = 11   
        # Column L (12): موضوع سایرمالیات و عوارض
        COL_OTHER_TAX_SUBJECT = 12 
        # Column M (13): نرخ سایرمالیات و عوارض
        # Column N (14): مبلغ سایرمالیات و عوارض
        # Column O (15): موضوع سایر وجوه قانونی
        # Column P (16): نرخ سایر وجوه قانونی
        # Column Q (17): مبلغ سایر وجوه قانونی
        # Column R (18): شناسه یکتای ثبت قرارداد حق العمل کاری


        for order in orders_data:
            try:
                item_names = [item['name'] for item in order.get('line_items', [])]

                item_quantities = []
                for item in order.get('line_items', []):
                    item_name = item['name']
                    quantity = item.get('quantity', 0)
                    refunded_qty_for_this_item = 0
                    for refund in order.get('refunds', []):
                        for refunded_item in refund.get('line_items', []):
                            if refunded_item.get('product_id') == item.get('product_id') and \
                               refunded_item.get('variation_id', 0) == item.get('variation_id', 0):
                                refunded_qty_for_this_item += refunded_item.get('qty', 0)

                item_quantities.append(str(quantity - refunded_qty_for_this_item))

                order_refund_total = sum(float(refund.get('total', 0)) for refund in order.get('refunds', []))

                created_datetime_obj = datetime.strptime(order['date_created'], '%Y-%m-%dT%H:%M:%S')
                jalali_date = jdatetime.datetime.fromtimestamp(created_datetime_obj.timestamp())
                formatted_jalali_date = jalali_date.strftime('%Y/%m/%d %H:%M:%S')

                custom_order_number = next(
                    (meta['value'] for meta in order.get('meta_data', []) if meta['key'] == '_wc_order_number' or meta['key'] == '_order_number'),
                    order.get('id')
                )

                # --- Create data row for the main comprehensive report (for all order types) ---
                order_row = {
                    "شماره سفارش": custom_order_number,
                    "تاریخ سفارش (شمسی)": formatted_jalali_date,
                    "نام": order.get('billing', {}).get('first_name', ''),
                    "نام خانوادگی": order.get('billing', {}).get('last_name', ''),
                    "نام شرکت": company_name,
                    "شناسه ملی": national_id,
                    "شماره ثبت": register_id,
                    "آدرس": f"{order.get('billing', {}).get('address_1', '')} {order.get('billing', {}).get('address_2', '')}".strip(),
                    "شهر": order.get('billing', {}).get('city', ''),
                    "کد پستی": order.get('billing', {}).get('postcode', ''),
                    "تلفن": order.get('billing', {}).get('phone', ''),
                    "عنوان روش پرداخت": order.get('payment_method_title', ''),
                    "مبلغ تخفیف": float(order.get('discount_total', 0)),
                    "مجموع مبلغ سفارش (با مالیات)": float(order.get('total', 0)),
                    "مجموع نهایی سفارش (بدون مالیات)": total_items_price_no_tax,
                    "مجموع مالیات بر ارزش افزوده": total_items_vat,
                    "روش حمل و نقل": order.get('shipping_lines', [{}])[0].get('method_title', '') if order.get('shipping_lines') else '',
                    "مبلغ حمل و نقل": float(order.get('shipping_lines', [{}])[0].get('total', 0)) if order.get('shipping_lines') else 0,
                    "مبلغ استرداد کل سفارش": order_refund_total,
                    "مجموع نهایی سفارش (پس از کسر استرداد)": float(order.get('total', 0)) - order_refund_total,
                    "نام آیتم‌ها": "\n".join(item_names),
                    "تعداد آیتم‌ها (- استرداد)": "\n".join(item_quantities),
                    "مجموع هزینه آیتم‌ها": sum(float(item.get('total', 0)) for item in order.get('line_items', []))
                }
                processed_data.append(order_row)

            except Exception as e:
                logger.error(f"ERROR: Error processing order {order.get('id', 'N/A')}: {e}. This order was skipped for main report generation.")
                continue

        # --- Save the consolidated templated workbook (if it was opened and valid) ---
        if workbook_template and templated_output_filename: 
            try:
                workbook_template.save(templated_output_filename)
                logger.info(f"INFO: Consolidated templated Excel file '{templated_output_filename}' generated successfully.")
            except Exception as e:
                logger.error(f"ERROR: Error saving consolidated templated Excel file '{templated_output_filename}': {e}")
                templated_output_filename = None # Mark as not successfully saved if saving failed
        else:
            templated_output_filename = None # Ensure it's None if not generated or saved

        # --- Generate the main comprehensive report ---
        df = pd.DataFrame(processed_data)
        
        # --- Sort the DataFrame by 'تاریخ سفارش (شمسی)' (Jalali Order Date) ---
        df = df.sort_values(by="تاریخ سفارش (شمسی)", ascending=True)

        excel_filename = f"WooCommerce_Orders_{datetime.now().strftime('%Y-%m-%d')}.xlsx"
        try:
            df.to_excel(main_excel_filename, index=False, engine='openpyxl')

            # --- Apply Excel Styling ---
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
                    if cell.column_letter in [get_column_letter(df.columns.get_loc("نام آیتم‌ها") + 1),
                                              get_column_letter(df.columns.get_loc("تعداد آیتم‌ها (- استرداد)") + 1),
                                              get_column_letter(df.columns.get_loc("آدرس") + 1)]:
                        cell.alignment = wrap_text_alignment

                    if cell.column_letter == get_column_letter(df.columns.get_loc("نام آیتم‌ها") + 1):
                        cell.fill = item_name_fill
            
            workbook.save(main_excel_filename)
            logger.info(f"INFO: Main Excel file '{main_excel_filename}' generated and styled successfully.")
            
            # Return main report path and the consolidated templated report path (if successfully created)
            return main_excel_filename, [templated_output_filename] if templated_output_filename and os.path.exists(templated_output_filename) else []

        except Exception as e:
            logger.error(f"ERROR: Error creating or styling main Excel file '{main_excel_filename}': {e}")
            return None, [templated_output_filename] if templated_output_filename and os.path.exists(templated_output_filename) else []