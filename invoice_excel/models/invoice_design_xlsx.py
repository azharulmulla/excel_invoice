from pkg_resources import safe_extra
from odoo import models


class InvoiceFormateXlsx(models.AbstractModel):
    _name = 'report.invoice_excel.report_invoice_xlsx'
    _inherit = ['report.report_type_xlsx.abstract']

    def generate_xlsx_report(self, workbook, data, invoice):
        sheet = workbook.add_worksheet('Invoice')
        bold = workbook.add_format({'bold': True})
        header_formate = workbook.add_format({'bold': 1, 'border': 1, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#c7bfbf',})
        header_formate.set_font_size(17)
        sub_header_formate = workbook.add_format({'bold': 1, 'border': 1, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#c7bfbf',})
        demo_data_formate = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter',})
        sub_header_small_formate = workbook.add_format({'bold': 1, 'border': 1, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#c7bfbf',})
        sub_header_small_formate.set_font_size(10)
        # header_font = workbook.add_format()
        # header_font.set_font_size(40)

        sheet.set_row(0, 45)
        sheet.set_row(1, 8)
        sheet.set_row(2, 20)
        sheet.set_row(10, 4)
        
        for obj in invoice:
            sheet.merge_range(0, 0, 0, 13, 'Invoice', header_formate)
            sheet.merge_range(2, 0, 2, 5, 'Exporter', sub_header_formate)
            sheet.merge_range(3, 0, 9, 5, 'Exporter data here', demo_data_formate)
            sheet.merge_range(2, 7, 2, 10, 'Invoice No. & Date', sub_header_formate)
            sheet.merge_range(3, 7, 5, 10, 'Date info here', demo_data_formate)
            sheet.merge_range(6, 7, 6, 10, 'Buyers Order No. & Date', sub_header_formate)
            sheet.merge_range(7, 7, 9, 10, 'Buyers Order No. & Date here', demo_data_formate)
            sheet.merge_range(2, 11, 9, 13, '....', demo_data_formate)
            sheet.merge_range(11, 0, 11, 5, 'Consignee', sub_header_formate)
            sheet.merge_range(12, 0, 17, 5, 'Consignee data here', demo_data_formate)
            sheet.merge_range(11, 7, 11, 13, 'Bill to', sub_header_formate)
            sheet.merge_range(12, 7, 15, 13, 'Bill to data here', demo_data_formate)
            sheet.merge_range(16, 7, 16, 10, 'Country of Origin of good',  sub_header_small_formate)
            sheet.merge_range(16, 11, 16, 13, 'Country of Final Destination',  sub_header_small_formate)
            sheet.merge_range(17, 7, 17, 10, 'Country here', demo_data_formate)
            sheet.merge_range(17, 11, 17, 13, 'Country here', demo_data_formate)


             
            
            

