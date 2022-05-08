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
        

        row = 3
        col = 0

        sheet.set_row(0, 45)
        sheet.set_row(1, 8)
        sheet.set_row(2, 20)
        sheet.set_row(10, 4)
        sheet.set_row(24, 3)
        
        for obj in invoice:
            sheet.merge_range(0, 0, 0, 13, 'Invoice', header_formate)
            sheet.merge_range(2, 0, 2, 5, 'Exporter', sub_header_formate)
            # sheet.merge_range(3, 0, 9, 5, 'Data here', demo_data_formate)
            
            # Exporter data here
            sheet.merge_range(3, 0, 3, 5, obj.company_id.name, demo_data_formate)
            sheet.merge_range(4, 0, 4, 5, obj.company_id.street, demo_data_formate)
            sheet.merge_range(5, 0, 5, 5, obj.company_id.street2, demo_data_formate)
            sheet.merge_range(6, 0, 6, 5, obj.company_id.city, demo_data_formate)
            sheet.merge_range(7, 0, 7, 5, obj.company_id.state_id.name, demo_data_formate)
            sheet.merge_range(8, 0, 8, 5, obj.company_id.zip, demo_data_formate)
            sheet.merge_range(9, 0, 9, 5, obj.company_id.country_id.name, demo_data_formate)
            # sheet.write(row, col, obj.company_id.name,)
            # row += 1
            # sheet.write(row, col, obj.company_id.street)
            # row += 1
            sheet.merge_range(2, 7, 2, 10, 'Invoice No. & Date', sub_header_formate)
            # sheet.merge_range(3, 7, 5, 10, 'Date info here', demo_data_formate)
            
            # name and date formate here
            sheet.merge_range(3, 7, 3, 10, obj.name, demo_data_formate)
            sheet.merge_range(4, 7, 4, 10, obj.invoice_date, demo_data_formate)


            sheet.merge_range(6, 7, 6, 10, 'Buyers Order No. & Date', sub_header_formate)
            # sheet.merge_range(7, 7, 9, 10, 'Buyers Order No. & Date here', demo_data_formate)

            sheet.write(7, 7, 'IEC', sub_header_small_formate)
            sheet.write(8, 7, 'PAN', sub_header_small_formate)
            sheet.write(9, 7, 'CIN', sub_header_small_formate)

            sheet.merge_range(7, 8, 7, 10, obj.company_id.iec_no, demo_data_formate)
            sheet.merge_range(8, 8, 8, 10, obj.company_id.pan_no, demo_data_formate)
            sheet.merge_range(9, 8, 9, 10, obj.company_id.cin_no, demo_data_formate)

            sheet.merge_range(2, 11, 9, 13, '....', demo_data_formate)
            sheet.merge_range(11, 0, 11, 5, 'Consignee', sub_header_formate)
            # sheet.merge_range(12, 0, 17, 5, 'Consignee data here', demo_data_formate)
            
            #Consignee data here
            sheet.merge_range(12, 0, 12, 5, obj.partner_id.name, demo_data_formate)
            sheet.merge_range(13, 0, 13, 5, obj.partner_id.street, demo_data_formate)
            sheet.merge_range(14, 0, 14, 5, obj.partner_id.street2, demo_data_formate)
            sheet.merge_range(15, 0, 15, 5, obj.partner_id.city, demo_data_formate)
            sheet.merge_range(16, 0, 16, 5, obj.partner_id.country_id.name, demo_data_formate)

            sheet.merge_range(11, 7, 11, 13, 'Bill to', sub_header_formate)
            # sheet.merge_range(12, 7, 15, 13, 'Bill to data here', demo_data_formate)
            
            #Bill to data here
            sheet.merge_range(12, 7, 12, 13,  obj.partner_id.name, demo_data_formate)
            sheet.merge_range(13, 7, 13, 13,  obj.partner_id.street, demo_data_formate)
            sheet.merge_range(14, 7, 14, 13,  obj.partner_id.city, demo_data_formate)
            sheet.merge_range(15, 7, 15, 13, obj.partner_id.country_id.name, demo_data_formate)


            sheet.merge_range(16, 7, 16, 10, 'Country of Origin of good',  sub_header_small_formate)
            sheet.merge_range(16, 11, 16, 13, 'Country of Final Destination',  sub_header_small_formate)
            sheet.merge_range(17, 7, 17, 10, obj.company_id.country_id.name, demo_data_formate)
            sheet.merge_range(17, 11, 17, 13, obj.partner_shipping_id.country_id.name, demo_data_formate)
            sheet.merge_range(18, 0, 18, 2, 'Vessel/Flight No.', sub_header_small_formate)
            sheet.merge_range(18, 3, 18, 5, 'Port of Loading', sub_header_small_formate)
            sheet.merge_range(19, 0, 19, 2, 'data here', demo_data_formate)
            sheet.merge_range(19, 3, 19, 5, 'port data', demo_data_formate)
            sheet.merge_range(19, 7, 19, 9, 'Terms', sub_header_small_formate)
            sheet.merge_range(19, 10, 19, 13, obj.invoice_incoterm_id.name, demo_data_formate)
            sheet.merge_range(20, 7, 20, 9, 'PAYMENT', sub_header_small_formate)
            sheet.merge_range(20, 10, 20, 13, obj.invoice_payment_term_id.name, demo_data_formate)
            sheet.merge_range(21, 7, 21, 9, 'Net Weight', sub_header_small_formate)
            sheet.merge_range(21, 10, 21, 13, 'data', demo_data_formate)
            sheet.merge_range(22, 7, 22, 9, 'Gross Weight', sub_header_small_formate)
            sheet.merge_range(22, 10, 22, 13, 'Data here', demo_data_formate)
            sheet.merge_range(23, 7, 23, 9, 'IRN', sub_header_small_formate)
            sheet.merge_range(23, 10, 23, 13, 'Data here', demo_data_formate)
            sheet.merge_range(18, 7, 18, 13, 'Terms of Delivery and Payment', sub_header_small_formate)
            sheet.merge_range(20, 0, 20, 2, 'Port of Discharge', sub_header_small_formate)
            sheet.merge_range(20, 3, 20, 5, 'Final Destination', sub_header_small_formate)
            sheet.merge_range(21, 0, 21, 2, obj.shipping_port_discharge, demo_data_formate)
            sheet.merge_range(21, 3, 21, 5, obj.partner_shipping_id.country_id.name, demo_data_formate)

            # sheet.merge_range(21, 3, 21, 5, 'data here', demo_data_formate)
            sheet.merge_range(22, 0, 22, 2, 'Ack. No', sub_header_small_formate)
            sheet.merge_range(22, 3, 22, 5, 'Ack. Date',sub_header_small_formate)
            sheet.merge_range(23, 0, 23, 2, 'Data', demo_data_formate)
            sheet.merge_range(23, 3, 23, 5, 'Data', demo_data_formate)

            #code for product discription need to dynamic customize
            sheet.merge_range(25, 0, 25, 1, 'Carton No.', sub_header_small_formate)
            sheet.merge_range(25, 2, 25, 7, 'Description of goods', sub_header_small_formate)
            sheet.write(25, 8, 'Quantity', sub_header_small_formate)
            sheet.merge_range(25, 9, 25, 10, 'Rate(US)', sub_header_small_formate)
            sheet.merge_range(25, 11, 25, 13, 'Amount(US)', sub_header_small_formate)


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



