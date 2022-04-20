{
    'name': "Account Import/Export",
    'summary': """Import/Export Extension for defining the company export/import details""",
    'author': "SPS",
    'website': "http://www.shorepointsys.com",
    'category': 'Account Extra Tools',
    'version': '0.1',
    'depends': ['account','documents_sales','account_import_export','report_type_xlsx',],
    'data': [
        'views/report_account_invoice_excel.xml',
        # 'views/report_template_invoice_excel.xml',
    ],
}