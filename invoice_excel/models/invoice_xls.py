# from pkg_resources import safe_extra
# from odoo import models


# class InvoiceFormateXlsx(models.AbstractModel):
#     _name = 'report.invoice_excel.report_invoice_xlsx'
#     _inherit = ['report.report_type_xlsx.abstract']

#     def generate_xlsx_report(self, workbook, data, Sale_order):
#         sheet = workbook.add_worksheet('Invoice')
#         bold = workbook.add_format({'bold': True})
#         format_1 = workbook.add_format({'bold': True, 'align': 'center', 'bg_color': 'yellow'})
#         format_2 = workbook.add_format({'bold': True, 'align': 'center', 'bg_color': 'yellow'})
#         merge_format1 = workbook.add_format({'bold': 1, 'border': 1, 'align': 'center', 'valign': 'vcenter', 'fg_color': 'yellow'})
#         merge_format2 = workbook.add_format({'bold': 1, 'border': 1,})


#         row = 2
#         col = 0

#         sheet.set_column('A:B', 50)
#         sheet.set_row(0, 30)
#         # sheet.set_column('B:B', 40)

#         for obj in Sale_order:
#             # One sheet by partner
#             sheet.merge_range(0, 0, 0, 4, 'Invoice',  merge_format1)
#             # sheet.merge_range('A2:A9', 'Exporter:', merge_format2)

#             # Exporter Information
#             sheet.write(1, 0, 'Exporter:', bold)
#             sheet.write(row, col,  obj.partner_id.name)
#             row += 1
#             sheet.write(row, col,  obj.partner_id.street)
#             row += 1
#             sheet.write(row, col,  obj.partner_id.street2)
#             row += 1
#             sheet.write(row, col,  obj.partner_id.city)
#             row += 1
#             sheet.write(row, col,  obj.partner_id.state_id.name)
#             row += 1
#             sheet.write(row, col,  obj.partner_id.country_id.name)

#             sheet.write(1, 1, 'Invoice No. & Date', bold)







            
