

import logging
import re
from io import BytesIO

from odoo import models

_logger = logging.getLogger(__name__)

try:
    import xlsxwriter

    class PatchedXlsxWorkbook(xlsxwriter.Workbook):
        def _check_sheetname(self, sheetname, is_chartsheet=False):
       
            try:
                return super()._check_sheetname(sheetname, is_chartsheet=is_chartsheet)
            except xlsxwriter.exceptions.DuplicateWorksheetName:
                pattern = re.compile(r"~[0-9]{2}$")
                duplicated_secuence = (
                    re.search(pattern, sheetname) and int(sheetname[-2:]) or 0
                )
                # Only up to 100 duplicates
                deduplicated_secuence = "~{:02d}".format(duplicated_secuence + 1)
                if duplicated_secuence > 99:
                    raise xlsxwriter.exceptions.DuplicateWorksheetName
                if duplicated_secuence:
                    sheetname = re.sub(pattern, deduplicated_secuence, sheetname)
                elif len(sheetname) <= 28:
                    sheetname += deduplicated_secuence
                else:
                    sheetname = sheetname[:28] + deduplicated_secuence
            # Refeed the method until we get an unduplicated name
            return self._check_sheetname(sheetname, is_chartsheet=is_chartsheet)

    # "Short string"

    xlsxwriter.Workbook = PatchedXlsxWorkbook

except ImportError:
    _logger.debug("Can not import xlsxwriter`.")


class ReportXlsxAbstract(models.AbstractModel):
    _name = "report.report_type_xlsx.abstract"
    _description = "Abstract XLSX Report"

    def _get_objs_for_report(self, docids, data):
     
        if docids:
            ids = docids
        elif data and "context" in data:
            ids = data["context"].get("active_ids", [])
        else:
            ids = self.env.context.get("active_ids", [])
        return self.env[self.env.context.get("active_model")].browse(ids)

    def create_xlsx_report(self, docids, data):
        objs = self._get_objs_for_report(docids, data)
        file_data = BytesIO()
        workbook = xlsxwriter.Workbook(file_data, self.get_workbook_options())
        self.generate_xlsx_report(workbook, data, objs)
        workbook.close()
        file_data.seek(0)
        return file_data.read(), "xlsx"

    def get_workbook_options(self):
        
        return {}

    def generate_xlsx_report(self, workbook, data, objs):
        raise NotImplementedError()
