from Processing.FlowReports.tables_headers.headers import Headers
from Utils.Logger.main_logger import get_logger

log = get_logger("ExcelReport")


class ExcelReport:
    def __init__(self, items_dict, report_name, report):
        self.items_dict = items_dict
        self.cur_write_row = 2
        self.report_name = report_name
        self.report = report

    def write_row(self, ws, m_item_id, values):
        if self.report_name == "items_history":
            ws.cell(self.cur_write_row, 1).value = m_item_id
            ws.cell(self.cur_write_row, 2).value = values.get("m_item_order")
            ws.cell(self.cur_write_row, 3).value = values.get("m_item_date_created")
            ws.cell(self.cur_write_row, 4).value = values.get("m_item_comment")
            ws.cell(self.cur_write_row, 5).value = values.get("step_id")
            ws.cell(self.cur_write_row, 6).value = values.get("date_start")
            ws.cell(self.cur_write_row, 7).value = values.get("comment")
            ws.cell(self.cur_write_row, 8).value = values.get("m_item_title")
            ws.cell(self.cur_write_row, 9).value = values.get("m_item_size")
            ws.cell(self.cur_write_row, 10).value = values.get("m_item_color")
            ws.cell(self.cur_write_row, 11).value = values.get("m_item_podklad")
            ws.cell(self.cur_write_row, 12).value = values.get("m_item_stopped")
            ws.cell(self.cur_write_row, 13).value = values.get("m_item_model")

            ws.cell(self.cur_write_row, 19).value = values.get("m_item_status")
            ws.cell(self.cur_write_row, 20).value = values.get("m_item_step_status")
            ws.cell(self.cur_write_row, 21).value = values.get("m_item_comment_status")
            ws.cell(self.cur_write_row, 22).value = values.get("m_item_size_status")
            ws.cell(self.cur_write_row, 23).value = values.get("m_item_color_status")
            ws.cell(self.cur_write_row, 24).value = values.get("m_item_podklad_status")
            ws.cell(self.cur_write_row, 25).value = values.get("m_item_model_status")


        elif self.report_name == 'deleted_items_history':
            ws.cell(self.cur_write_row, 1).value = m_item_id
            ws.cell(self.cur_write_row, 2).value = values.get("m_item_order")
            ws.cell(self.cur_write_row, 3).value = values.get("m_item_date_created")
            ws.cell(self.cur_write_row, 4).value = values.get("m_item_comment")
            ws.cell(self.cur_write_row, 5).value = values.get("step_id")
            ws.cell(self.cur_write_row, 6).value = values.get("date_start")
            ws.cell(self.cur_write_row, 7).value = values.get("comment")
            ws.cell(self.cur_write_row, 8).value = values.get("m_item_title")
            ws.cell(self.cur_write_row, 9).value = values.get("m_item_size")
            ws.cell(self.cur_write_row, 10).value = values.get("m_item_color")
            ws.cell(self.cur_write_row, 11).value = values.get("m_item_podklad")
            ws.cell(self.cur_write_row, 12).value = values.get("m_item_stopped")
            ws.cell(self.cur_write_row, 13).value = values.get("m_item_model")
            ws.cell(self.cur_write_row, 14).value = values.get("m_item_delete_date")

        elif self.report_name == "stw_modules_colors_local":
            ws.cell(self.cur_write_row, 1).value = m_item_id
            ws.cell(self.cur_write_row, 2).value = values.get("m_item_lang")
            ws.cell(self.cur_write_row, 3).value = values.get("m_item_title")
            ws.cell(self.cur_write_row, 4).value = values.get("m_item_visible")
        else:
            ws.cell(self.cur_write_row, 1).value = m_item_id
            ws.cell(self.cur_write_row, 2).value = values.get("m_item_lang")
            ws.cell(self.cur_write_row, 3).value = values.get("m_item_title")
        self.cur_write_row += 1
        return ws

    def get_header_values(self):
        if self.report_name == "items_history":
            return Headers.headers_items_history
        elif self.report_name == "stw_modules_colors_local":
            return Headers.headers_colors_local
        else:
            return Headers.headers_supporting

    def items_processing(self):
        if self.report_name == 'items_history':
            worksheet = 'Актуальная выгрузка'
        elif self.report_name == 'deleted_items_history':
            worksheet = 'Удаленные заказы'
        else:
            worksheet = self.report_name
        ws = self.report[worksheet]
        # ws = self.write_row(ws=ws, m_item_id="m_item_id", values=self.get_header_values())
        log.info(f"Writing report {self.report_name}")
        for m_item_id, values in self.items_dict.items():
            if m_item_id != 'm_item_id':
                ws = self.write_row(ws=ws, m_item_id=m_item_id, values=values)
        log.info(f"Saving report {self.report_name}")

