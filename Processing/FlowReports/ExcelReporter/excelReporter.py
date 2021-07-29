from Processing.FlowReports.tables_headers.headers import Headers
from Utils.Logger.main_logger import get_logger

log = get_logger("ExcelReport")


class ExcelReport:
    def __init__(self, items_dict, report_name, report):
        self.items_dict = items_dict
        self.cur_write_row = 0
        self.report_name = report_name
        self.report = report

    def write_row(self, ws, m_item_id, values):
        if self.report_name == 'items_history':
            ws.write(self.cur_write_row, 0, m_item_id)
            ws.write(self.cur_write_row, 1, values.get('m_item_order'))
            ws.write(self.cur_write_row, 2, values.get('m_item_date_created'))
            ws.write(self.cur_write_row, 3, values.get('m_item_comment'))
            ws.write(self.cur_write_row, 4, values.get('step_id'))
            ws.write(self.cur_write_row, 5, values.get('date_start'))
            ws.write(self.cur_write_row, 6, values.get('comment'))
            ws.write(self.cur_write_row, 7, values.get('m_item_title'))
            ws.write(self.cur_write_row, 8, values.get('m_item_size'))
            ws.write(self.cur_write_row, 9, values.get('m_item_color'))
            ws.write(self.cur_write_row, 10, values.get('m_item_podklad'))
            ws.write(self.cur_write_row, 11, values.get('m_item_stopped'))
            ws.write(self.cur_write_row, 12, values.get('m_item_model'))
        elif self.report_name == 'stw_modules_colors_local':
            ws.write(self.cur_write_row, 0, m_item_id)
            ws.write(self.cur_write_row, 1, values.get('m_item_lang'))
            ws.write(self.cur_write_row, 2, values.get('m_item_title'))
            ws.write(self.cur_write_row, 3, values.get('m_item_visible'))
        else:
            ws.write(self.cur_write_row, 0, m_item_id)
            ws.write(self.cur_write_row, 1, values.get('m_item_lang'))
            ws.write(self.cur_write_row, 2, values.get('m_item_title'))
        self.cur_write_row += 1
        return ws

    def get_header_values(self):
        if self.report_name == 'items_history':
            return Headers.headers_items_history
        elif self.report_name == 'stw_modules_colors_local':
            return Headers.headers_colors_local
        else:
            return Headers.headers_supporting

    def items_processing(self):
        ws = self.report.add_worksheet(self.report_name)
        ws = self.write_row(ws=ws, m_item_id='m_item_id', values=self.get_header_values())
        log.info(f"Writing report {self.report_name}")
        for m_item_id, values in self.items_dict.items():
            ws = self.write_row(ws=ws, m_item_id=m_item_id, values=values)
        log.info(f"Saving report {self.report_name}")

