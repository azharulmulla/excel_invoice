# -*- coding: utf-8 -*-
{
    'name': "Report Type XLSX",
    'summary': """Adding XLSX format in report type""",
    'description': """""",
    'author': "SPS",
    'website': "http://www.shorepointsys.com",
    'category': "Reporting",
    "external_dependencies": {"python": ["xlsxwriter", "xlrd"]},
    'version': '0.1',
    'depends': ["base", "web"],
    'data': [
       "views/webclient_templates.xml",
    ],
    "demo": ["demo/report.xml"],
    "installable": True,
}