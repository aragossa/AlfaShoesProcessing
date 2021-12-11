import datetime

from Processing.FlowReports.ExcelReporter.utils.RowWriter import RowWriter
from Processing.FlowReports.ExcelReporter.utils.formulas import get_excel_formula
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
                    ws = self.write_row(ws=ws,
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
                        ws = self.write_row(ws=ws,
                                                 m_item_id=m_item_id,
                                                 values=self.deleted_items_dict.get(m_item_id),
                                                 report_name=self.report_name,
                                                 cur_write_row=self.cur_write_row,
                                                 values_flow=self.deleted_items_dict.get(m_item_id),
                                                 deleted=True,
                                                 m_item_delete_date=m_item_delete_date)

                    elif self.items_dict.get(m_item_id):
                        ws = self.write_row(ws=ws,
                                                 m_item_id=m_item_id,
                                                 values=self.items_dict.get(m_item_id),
                                                 report_name=self.report_name,
                                                 cur_write_row=self.cur_write_row,
                                                 values_flow=self.flow_sheet_dict.get(m_item_id))

        log.info(f"Saving report {self.report_name}")


    def __detect_changes(self, values):
        log.debug(values)
        if values.get("m_item_status") == 'новое изделие':
            changes = True
        elif values.get("m_item_step_status") != '':
            changes = True
        elif values.get("m_item_comment_status") != '':
            changes = True
        elif values.get("m_item_size_status") != '':
            changes = True
        elif values.get("m_item_color_status") != '':
            changes = True
        elif values.get("m_item_podklad_status") != '':
            changes = True
        elif values.get("m_item_model_status") != '':
            changes = True
        else:
            changes = False
        log.debug(f'detected changes {changes}')
        log.debug('---------------------------')
        return changes

    def __prepare_write_value(self, values, changes, col_num, values_flow, col_name, row_num, m_item_id=None,
                              deleted=False):
        log.debug(values)
        if col_num <= 14:
            if col_name == 'step_id' and deleted:
                return 'удален'
            elif col_name == 'date_start' and deleted:
                return self.deleted_items_dict.get(m_item_id).get('m_item_delete_date')
            else:
                return values.get(col_name)
        elif values.get('m_item_status') == 'новое изделие':
            log.debug('writing formula for new m_item')
            log.debug(get_excel_formula(col_num, row_num))
            return get_excel_formula(col_num, row_num)
        elif changes or col_num >= 30 or col_num <= 43:
            log.debug('writing formula')
            # log.debug(col_num)
            log.debug(get_excel_formula(col_num, row_num))
            return get_excel_formula(col_num, row_num)
        else:
            log.debug('writing data')
            log.debug(col_name)
            log.debug(values_flow.get(col_name))
            log.debug(values_flow.get(col_name))
            # log.debug(col_name)
            # log.debug(values)
            # input('any key')
            return values_flow.get(col_name)

    # def write_row(self, ws, m_item_id, values, values_flow=None, deleted=False):
    def write_row(self, ws, m_item_id, values, report_name, cur_write_row, values_flow=None, deleted=False,
                  m_item_delete_date=None):
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

        elif self.report_name == 'Flow':
            log.debug(deleted)
            if deleted:
                changes = True
            else:
                changes = self.__detect_changes(values)
            log.debug(f'm_item_id {m_item_id}')
            ws.cell(self.cur_write_row, 1).value = m_item_id
            ws.cell(self.cur_write_row, 2).value = self.__prepare_write_value(values=values, values_flow=values_flow,
                                                                              col_num=2,
                                                                              col_name='m_item_order',
                                                                              changes=changes,
                                                                              row_num=self.cur_write_row)
            ws.cell(self.cur_write_row, 3).value = self.__prepare_write_value(values=values, values_flow=values_flow,
                                                                              col_num=3,
                                                                              col_name='m_item_date_created',
                                                                              changes=changes,
                                                                              row_num=self.cur_write_row)
            ws.cell(self.cur_write_row, 4).value = self.__prepare_write_value(values=values, values_flow=values_flow,
                                                                              col_num=4,
                                                                              col_name='m_item_title',
                                                                              changes=changes,
                                                                              row_num=self.cur_write_row)
            ws.cell(self.cur_write_row, 5).value = self.__prepare_write_value(values=values, values_flow=values_flow,
                                                                              col_num=5,
                                                                              col_name='m_item_comment',
                                                                              changes=changes,
                                                                              row_num=self.cur_write_row)
            ws.cell(self.cur_write_row, 6).value = self.__prepare_write_value(values=values, values_flow=values_flow,
                                                                              col_num=6,
                                                                              col_name='step_id',
                                                                              changes=changes,
                                                                              row_num=self.cur_write_row,
                                                                              deleted=deleted)

            ws.cell(self.cur_write_row, 7).value = self.__prepare_write_value(values=values, values_flow=values_flow,
                                                                              col_num=7,
                                                                              col_name='date_start',
                                                                              changes=changes,
                                                                              row_num=self.cur_write_row,
                                                                              m_item_id=m_item_id,
                                                                              deleted=deleted)
            ws.cell(self.cur_write_row, 8).value = self.__prepare_write_value(values=values, values_flow=values_flow,
                                                                              col_num=8,
                                                                              col_name='comment',
                                                                              changes=changes,
                                                                              row_num=self.cur_write_row)
            ws.cell(self.cur_write_row, 9).value = self.__prepare_write_value(values=values, values_flow=values_flow,
                                                                              col_num=9,
                                                                              col_name='m_item_size',
                                                                              changes=changes,
                                                                              row_num=self.cur_write_row)
            ws.cell(self.cur_write_row, 10).value = self.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=10,
                                                                               col_name='m_item_color',
                                                                               changes=changes,
                                                                               row_num=self.cur_write_row)
            ws.cell(self.cur_write_row, 11).value = self.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=11,
                                                                               col_name='m_item_podklad',
                                                                               changes=changes,
                                                                               row_num=self.cur_write_row)
            ws.cell(self.cur_write_row, 12).value = self.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=12,
                                                                               col_name='m_item_stopped',
                                                                               changes=changes,
                                                                               row_num=self.cur_write_row)
            ws.cell(self.cur_write_row, 13).value = self.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=13,
                                                                               col_name='m_item_model',
                                                                               changes=changes,
                                                                               row_num=self.cur_write_row)

            ws.cell(self.cur_write_row, 15).value = self.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=15,
                                                                               col_name='flow_shoe_block',
                                                                               changes=changes,
                                                                               row_num=self.cur_write_row)
            ws.cell(self.cur_write_row, 17).value = self.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=17,
                                                                               col_name='flow_model_article',
                                                                               changes=changes,
                                                                               row_num=self.cur_write_row)
            ws.cell(self.cur_write_row, 20).value = self.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=20,
                                                                               col_name='flow_color_article',
                                                                               changes=changes,
                                                                               row_num=self.cur_write_row)
            ws.cell(self.cur_write_row, 24).value = self.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=24,
                                                                               col_name='flow_size',
                                                                               changes=changes,
                                                                               row_num=self.cur_write_row)
            ws.cell(self.cur_write_row, 26).value = self.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=26,
                                                                               col_name='flow_range_position',
                                                                               changes=changes,
                                                                               row_num=self.cur_write_row)
            ws.cell(self.cur_write_row, 27).value = self.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=27,
                                                                               col_name='flow_sample_numbler',
                                                                               changes=changes,
                                                                               row_num=self.cur_write_row)
            ws.cell(self.cur_write_row, 28).value = self.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=28,
                                                                               col_name='flow_model_article2',
                                                                               changes=changes,
                                                                               row_num=self.cur_write_row)
            ws.cell(self.cur_write_row, 29).value = self.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=29,
                                                                               col_name='flow_model_article2',
                                                                               changes=changes,
                                                                               row_num=self.cur_write_row)
            ws.cell(self.cur_write_row, 30).value = self.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=30,
                                                                               col_name='flow_color_article2',
                                                                               changes=changes,
                                                                               row_num=self.cur_write_row)
            ws.cell(self.cur_write_row, 31).value = self.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=31,
                                                                               col_name='flow_size2',
                                                                               changes=changes,
                                                                               row_num=self.cur_write_row)
            ws.cell(self.cur_write_row, 32).value = self.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=32,
                                                                               col_name='flow_podklad',
                                                                               changes=changes,
                                                                               row_num=self.cur_write_row)
            ws.cell(self.cur_write_row, 33).value = self.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=33,
                                                                               col_name='flow_color_filter',
                                                                               changes=changes,
                                                                               row_num=self.cur_write_row)
            # ws.cell(self.cur_write_row, 34).value = self.__prepare_write_value(values=values, values_flow=values_flow,
            #                                                                    col_num=34,
            #                                                                    col_name='flow_ordered',
            #                                                                    changes=changes,
            #                                                                    row_num=self.cur_write_row)
            ws.cell(self.cur_write_row, 34).value = 1
            ws.cell(self.cur_write_row, 35).value = self.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=35,
                                                                               col_name='flow_canceled',
                                                                               changes=changes,
                                                                               row_num=self.cur_write_row)

            ws.cell(self.cur_write_row, 36).value = self.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=36,
                                                                               col_name='flow_blanc',
                                                                               changes=changes,
                                                                               row_num=self.cur_write_row)
            ws.cell(self.cur_write_row, 37).value = self.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=37,
                                                                               col_name='flow_finish',
                                                                               changes=changes,
                                                                               row_num=self.cur_write_row)
            ws.cell(self.cur_write_row, 38).value = self.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=38,
                                                                               col_name='flow_ready',
                                                                               changes=changes,
                                                                               row_num=self.cur_write_row)

            ws.cell(self.cur_write_row, 39).value = self.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=39,
                                                                               col_name='flow_ready',
                                                                               changes=changes,
                                                                               row_num=self.cur_write_row)
            ws.cell(self.cur_write_row, 40).value = self.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=40,
                                                                               col_name='flow_ready',
                                                                               changes=changes,
                                                                               row_num=self.cur_write_row)
            ws.cell(self.cur_write_row, 41).value = self.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=41,
                                                                               col_name='flow_ready',
                                                                               changes=changes,
                                                                               row_num=self.cur_write_row)
            ws.cell(self.cur_write_row, 42).value = self.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=42,
                                                                               col_name='flow_ready',
                                                                               changes=changes,
                                                                               row_num=self.cur_write_row)
            ws.cell(self.cur_write_row, 43).value = self.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=43,
                                                                               col_name='flow_ready',
                                                                               changes=changes,
                                                                               row_num=self.cur_write_row)

            ws.cell(self.cur_write_row, 46).value = self.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=46,
                                                                               col_name='flow_order_comment',
                                                                               changes=changes,
                                                                               row_num=self.cur_write_row)
            ws.cell(self.cur_write_row, 47).value = self.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=47,
                                                                               col_name='flow_barcode',
                                                                               changes=changes,
                                                                               row_num=self.cur_write_row)
            ws.cell(self.cur_write_row, 48).value = self.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=48,
                                                                               col_name='flow_control_digit',
                                                                               changes=changes,
                                                                               row_num=self.cur_write_row)
            ws.cell(self.cur_write_row, 49).value = self.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=49,
                                                                               col_name='flow_ean_13',
                                                                               changes=changes,
                                                                               row_num=self.cur_write_row)

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


