from pkg_resources import safe_extra
from odoo import models



class InvoiceFormateXlsx(models.AbstractModel):
    _name = 'report.invoice_excel.report_invoice_xlsx'
    _inherit = ['report.report_type_xlsx.abstract']

# design section 
    def generate_xlsx_report(self, workbook, data, invoice):
        sheet = workbook.add_worksheet('Invoice')
        bold = workbook.add_format({'bold': True})
        header_formate = workbook.add_format({'bold': 1, 'border': 1, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#c7bfbf',})
        header_formate.set_font_size(17)
        sub_header_formate = workbook.add_format({'bold': 1, 'border': 1, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#c7bfbf',})
        demo_data_formate = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter',})
        demo_data_formate2 = workbook.add_format({'border': 1, 'align': 'left', 'valign': 'vcenter',})
        sub_header_small_formate = workbook.add_format({'bold': 1, 'border': 1, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#c7bfbf',})
        sub_header_small_formate.set_font_size(10)
        

        
# row size modification

        sheet.set_row(0, 45)
        sheet.set_row(1, 8)
        sheet.set_row(2, 20)
        sheet.set_row(10, 4)
        sheet.set_row(24, 3)

        dynamic_row = 0
        dynamic_col = 0

        dynamic_row2 = 2
        
        for obj in invoice:
            sheet.merge_range(dynamic_row , dynamic_col, dynamic_row,  dynamic_col+13, 'Invoice', header_formate)
            sheet.merge_range(dynamic_row2, dynamic_col, dynamic_row2, dynamic_col+5, 'Exporter', sub_header_formate)
            sheet.merge_range(dynamic_row2, dynamic_col+7, dynamic_row2, dynamic_col+10, 'Invoice No. & Date', sub_header_formate)
            dynamic_row2 +=1
            
            # sheet.merge_range(3, 0, 9, 5, 'Data here', demo_data_formate)
            
            # Exporter data here
            sheet.merge_range(dynamic_row2, dynamic_col, dynamic_row2, dynamic_col+5, obj.company_id.name, demo_data_formate)
            sheet.merge_range(dynamic_row2,  dynamic_col+7, dynamic_row2, dynamic_col+10, obj.name, demo_data_formate)
            dynamic_row2 += 1

            sheet.merge_range(dynamic_row2, dynamic_col, dynamic_row2, dynamic_col+5, obj.company_id.street, demo_data_formate)
            sheet.merge_range(dynamic_row2, dynamic_col+7, dynamic_row2, dynamic_col+10, obj.invoice_date, demo_data_formate)
            dynamic_row2 += 1

            sheet.merge_range(dynamic_row2, dynamic_col, dynamic_row2, dynamic_col+5, obj.company_id.street2, demo_data_formate)
            dynamic_row2 += 1

            sheet.merge_range(dynamic_row2, dynamic_col, dynamic_row2, dynamic_col+5, obj.company_id.city, demo_data_formate)
            sheet.merge_range(dynamic_row2, dynamic_col+7, dynamic_row2, dynamic_col+10, 'Buyers Order No. & Date', sub_header_formate)
            dynamic_row2 += 1

            sheet.merge_range(dynamic_row2, dynamic_col, dynamic_row2, dynamic_col+5, obj.company_id.state_id.name, demo_data_formate)
            sheet.write(dynamic_row2, dynamic_col+7, 'IEC', sub_header_small_formate)
            sheet.merge_range(dynamic_row2, dynamic_col+8, dynamic_row2, dynamic_col+10, obj.company_id.iec_no, demo_data_formate)
            dynamic_row2 += 1

            sheet.merge_range(dynamic_row2, dynamic_col, dynamic_row2, dynamic_col+5, obj.company_id.zip, demo_data_formate)
            sheet.write(dynamic_row2, dynamic_col+7, 'PAN', sub_header_small_formate)
            sheet.merge_range(dynamic_row2, dynamic_col+8, dynamic_row2, dynamic_col+10, obj.company_id.pan_no, demo_data_formate)
            dynamic_row2 += 1

            sheet.merge_range(dynamic_row2, dynamic_col, dynamic_row2, dynamic_col+5, obj.company_id.country_id.name, demo_data_formate)
            sheet.write(dynamic_row2, dynamic_col+7, 'CIN', sub_header_small_formate)
            sheet.merge_range(dynamic_row2, dynamic_col+8, dynamic_row2, dynamic_col+10, obj.company_id.cin_no, demo_data_formate)
            dynamic_row2 += 2

            sheet.merge_range(dynamic_row2, dynamic_col, dynamic_row2, dynamic_col+5, 'Consignee', sub_header_formate)
            sheet.merge_range(dynamic_row2, dynamic_col+7, dynamic_row2, dynamic_col+13, 'Bill to', sub_header_formate)
            dynamic_row2 += 1

            sheet.merge_range(dynamic_row2, dynamic_col, dynamic_row2, dynamic_col+5, obj.partner_id.name, demo_data_formate)
            sheet.merge_range(dynamic_row2, dynamic_col+7, dynamic_row2, dynamic_col+13,  obj.partner_id.name, demo_data_formate)
            dynamic_row2 += 1

            sheet.merge_range(dynamic_row2, dynamic_col, dynamic_row2, dynamic_col+5, obj.partner_id.street, demo_data_formate)
            sheet.merge_range(dynamic_row2, dynamic_col+7, dynamic_row2, dynamic_col+13,  obj.partner_id.street, demo_data_formate)
            dynamic_row2 += 1

            sheet.merge_range(dynamic_row2, dynamic_col, dynamic_row2, dynamic_col+5, obj.partner_id.street2, demo_data_formate)
            sheet.merge_range(dynamic_row2, dynamic_col+7, dynamic_row2, dynamic_col+13,  obj.partner_id.city, demo_data_formate)
            dynamic_row2 += 1

            sheet.merge_range(dynamic_row2, dynamic_col, dynamic_row2, dynamic_col+5, obj.partner_id.city, demo_data_formate)
            sheet.merge_range(dynamic_row2, dynamic_col+7, dynamic_row2, dynamic_col+13, obj.partner_id.country_id.name, demo_data_formate)
            dynamic_row2 += 1

            sheet.merge_range(dynamic_row2, dynamic_col, dynamic_row2, dynamic_col+5, obj.partner_id.country_id.name, demo_data_formate)
            sheet.merge_range(dynamic_row2, dynamic_col+7, dynamic_row2, dynamic_col+10, 'Country of Origin of good',  sub_header_small_formate)
            sheet.merge_range(dynamic_row2, dynamic_col+11, dynamic_row2, dynamic_col+13, 'Country of Final Destination',  sub_header_small_formate)
            dynamic_row2 += 1

            sheet.merge_range(dynamic_row2, dynamic_col+7, dynamic_row2, dynamic_col+10, obj.company_id.country_id.name, demo_data_formate)
            sheet.merge_range(dynamic_row2, dynamic_col+11, dynamic_row2, dynamic_col+13, obj.partner_shipping_id.country_id.name, demo_data_formate)
            dynamic_row2 += 1

            sheet.merge_range(dynamic_row2, dynamic_col, dynamic_row2, dynamic_col+2, 'Vessel/Flight No.', sub_header_small_formate)
            sheet.merge_range(dynamic_row2, dynamic_col+3, dynamic_row2, dynamic_col+5, 'Port of Loading', sub_header_small_formate)
            sheet.merge_range(dynamic_row2, dynamic_col+7, dynamic_row2, dynamic_col+13, 'Terms of Delivery and Payment', sub_header_small_formate)
            dynamic_row2 += 1
            

            sheet.merge_range(dynamic_row2, dynamic_col, dynamic_row2, dynamic_col+2, 'data here', demo_data_formate)
            sheet.merge_range(dynamic_row2, dynamic_col+3, dynamic_row2, dynamic_col+5, 'port data', demo_data_formate)
            sheet.merge_range(dynamic_row2, dynamic_col+7, dynamic_row2, dynamic_col+9, 'Terms', sub_header_small_formate)
            sheet.merge_range(dynamic_row2, dynamic_col+10, dynamic_row2, dynamic_col+13, obj.invoice_incoterm_id.name, demo_data_formate)
            dynamic_row2 += 1


            sheet.merge_range(dynamic_row2,  dynamic_col+0, dynamic_row2,  dynamic_col+2, 'Port of Discharge', sub_header_small_formate)
            sheet.merge_range(dynamic_row2,  dynamic_col+3, dynamic_row2,  dynamic_col+5, 'Final Destination', sub_header_small_formate)
            sheet.merge_range(dynamic_row2,  dynamic_col+7, dynamic_row2,  dynamic_col+9, 'PAYMENT', sub_header_small_formate)
            sheet.merge_range(dynamic_row2,  dynamic_col+10, dynamic_row2,  dynamic_col+13, obj.invoice_payment_term_id.name, demo_data_formate)
            dynamic_row2 += 1


            sheet.merge_range(dynamic_row2, dynamic_col+0, dynamic_row2,dynamic_col+ 2, obj.shipping_port_discharge, demo_data_formate)
            sheet.merge_range(dynamic_row2, dynamic_col+3, dynamic_row2, dynamic_col+5, obj.partner_shipping_id.country_id.name, demo_data_formate)
            sheet.merge_range(dynamic_row2, dynamic_col+7, dynamic_row2, dynamic_col+9, 'Net Weight', sub_header_small_formate)
            sheet.merge_range(dynamic_row2, dynamic_col+10, dynamic_row2, dynamic_col+13, 'data', demo_data_formate)
            dynamic_row2 += 1


            sheet.merge_range(dynamic_row2, dynamic_col+0, dynamic_row2, dynamic_col+2, 'Ack. No', sub_header_small_formate)
            sheet.merge_range(dynamic_row2, dynamic_col+3, dynamic_row2, dynamic_col+5, 'Ack. Date',sub_header_small_formate)
            sheet.merge_range(dynamic_row2, dynamic_col+7, dynamic_row2, dynamic_col+9, 'Gross Weight', sub_header_small_formate)
            sheet.merge_range(dynamic_row2, dynamic_col+10, dynamic_row2, dynamic_col+13, 'Data here', demo_data_formate)
            dynamic_row2 += 1

            sheet.merge_range(dynamic_row2, dynamic_col+0, dynamic_row2, dynamic_col+2, 'Data', demo_data_formate)
            sheet.merge_range(dynamic_row2, dynamic_col+3, dynamic_row2, dynamic_col+5, 'Data', demo_data_formate)
            sheet.merge_range(dynamic_row2, dynamic_col+7, dynamic_row2, dynamic_col+9, 'IRN', sub_header_small_formate)
            sheet.merge_range(dynamic_row2, dynamic_col+10, dynamic_row2, dynamic_col+13, 'Data here', demo_data_formate)
            dynamic_row2 += 1

            sheet.merge_range(dynamic_row2, dynamic_col+0, dynamic_row2, dynamic_col+1, 'Carton No.', sub_header_small_formate)
            sheet.merge_range(dynamic_row2, dynamic_col+2, dynamic_row2, dynamic_col+7, 'Description of goods', sub_header_small_formate)
            sheet.write(dynamic_row2, dynamic_col+8, 'Quantity', sub_header_small_formate)
            sheet.merge_range(dynamic_row2, dynamic_col+9, dynamic_row2, dynamic_col+10, 'Rate(US)', sub_header_small_formate)
            sheet.merge_range(dynamic_row2, dynamic_col+11, dynamic_row2, dynamic_col+13, 'Amount(US)', sub_header_small_formate)
            dynamic_row2 += 1
            

            sheet.merge_range(2, 11, 9, 13, '....', demo_data_formate)


# Dynamic functionality

            product_row = 26

            total_quantity = []
            total_rate = []
            total_amount = []
            

            for product in obj.invoice_line_ids.product_id:
                sheet.merge_range(product_row, 0, product_row, 1, 'Data', demo_data_formate)  
                sheet.merge_range(product_row, 2, product_row, 7, product.name, demo_data_formate)
                for q in  obj.invoice_line_ids:
                    if(product.id == q.product_id.id):
                        sheet.write(product_row, 8, q.quantity, demo_data_formate)
                        total_quantity.append(q.quantity)
                        sheet.merge_range(product_row, 9, product_row, 10, q.price_unit, demo_data_formate)
                        sheet.merge_range(product_row, 11, product_row, 13, q.price_subtotal, demo_data_formate)
                                
                product_row += 1
            sumOfQuantity = 0
            for quantity in total_quantity:
                sumOfQuantity += quantity


            sumofRate = 0
            for rate in total_rate:
                sumofRate += rate

            sumofAmount = 0
            for amount in total_amount:
                sumofAmount += amount    



# calculation section

            sheet.merge_range(product_row, 0, product_row, 2, 'Total No. of. Cartons', sub_header_small_formate)
            sheet.merge_range(product_row, 3, product_row, 5, 'Data', demo_data_formate)
            sheet.merge_range(product_row, 6, product_row, 7, 'Total', sub_header_small_formate)
            sheet.write(product_row, 8, sumOfQuantity, demo_data_formate)
            sheet.merge_range(product_row, 9, product_row, 10, sumofRate, demo_data_formate)
            sheet.merge_range(product_row, 11, product_row, 13, sumofAmount, demo_data_formate)

# footer section

            sheet.merge_range(product_row+1, 0, product_row+2, 9, 'Amount chargeable (In words)\n the amount ', demo_data_formate2)
            sheet.merge_range(product_row+3, 0, product_row+5, 9, 'Declaration\n We declare that this invoice shows the actual price of the goods described and that all particulars\n are true and correct', demo_data_formate2)
            sheet.merge_range(product_row+1, 10, product_row+5, 13, 'Signature & Date', demo_data_formate)



