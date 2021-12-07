import datetime

from Processing.FlowReports.ExcelReporter.utils.RowWriter import RowWriter
from Processing.FlowReports.tables_headers.headers import Headers
from Utils.Logger.main_logger import get_logger

log = get_logger("ExcelReport")


class ExcelReport:
    def __init__(self, items_dict, report_name, report, flow_sheet_dict=None, full_items_list=None,
                 deleted_items_dict=None):
        self.deleted_items_dict = deleted_items_dict
        self.full_items_list = full_items_list
        self.items_dict = items_dict
        self.flow_sheet_dict = flow_sheet_dict
        self.cur_write_row = 2
        self.report_name = report_name
        self.report = report

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
        if worksheet == 'Актуальная выгрузка':
            ws.cell(1, 16).value = datetime.datetime.now().strftime("%d.%m.%Y")
            ws.cell(1, 17).value = datetime.datetime.now().strftime("%H:%M")

        log.info(f"Writing report {self.report_name}")
        if worksheet == 'Flow':
            self.cur_write_row = 5

        if worksheet != 'Flow':
            for m_item_id, values in self.items_dict.items():
                if m_item_id != 'm_item_id':
                    log.debug(worksheet)
                    ws = RowWriter.write_row(ws=ws,
                                             m_item_id=m_item_id,
                                             values=values,
                                             report_name=self.report_name,
                                             cur_write_row=self.cur_write_row)
        else:
            for m_item_id in self.full_items_list:
                if m_item_id != 'm_item_id':
                    log.debug(f'Writing values to Flow worksheet: {m_item_id}')
                    log.debug(self.flow_sheet_dict.get(m_item_id))
                    log.debug(self.deleted_items_dict.get(m_item_id))

                    if self.deleted_items_dict.get(m_item_id):
                        m_item_delete_date=self.deleted_items_dict.get(m_item_id).get('m_item_delete_date')
                        ws = RowWriter.write_row(ws=ws,
                                                 m_item_id=m_item_id,
                                                 values=self.deleted_items_dict.get(m_item_id),
                                                 report_name=self.report_name,
                                                 cur_write_row=self.cur_write_row,
                                                 values_flow=self.deleted_items_dict.get(m_item_id),
                                                 deleted=True,
                                                 m_item_delete_date=m_item_delete_date)

                    elif self.items_dict.get(m_item_id):
                        ws = RowWriter.write_row(ws=ws,
                                                 m_item_id=m_item_id,
                                                 values=self.items_dict.get(m_item_id),
                                                 report_name=self.report_name,
                                                 cur_write_row=self.cur_write_row,
                                                 values_flow=self.flow_sheet_dict.get(m_item_id))

        log.info(f"Saving report {self.report_name}")
