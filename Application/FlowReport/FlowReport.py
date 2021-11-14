import datetime

import openpyxl

from Processing.FlowReports.ExcelReporter.excelReporter import ExcelReport
from Processing.FlowReports.row_object.rowOject import RowObject
from Application.FlowReport.TableNames.TableNames import TABLES_NAMES
from Utils.DbConnector.dbConnector import DbConnector
from Utils.Logger.main_logger import get_logger

log = get_logger("FlowReport")


class FlowReport:
    def __init__(self):
        self.local_storage_path = 'TempFolder'

    def __parse_row_object(self, table_name, row, items_dict):
        if table_name == "items_history":
            return RowObject(
                m_item_id='' if row[0] is None else row[0],
                table_name=table_name,
                m_item_order='' if row[1] is None else row[1],
                m_item_date_created='' if row[2] is None else row[2],
                m_item_comment='' if row[3] is None else row[3],
                step_id='' if row[4] is None else row[4],
                date_start='' if row[5] is None else row[5],
                comment='' if row[6] is None else row[6],
                m_item_title='' if row[7] is None else row[7],
                m_item_size='' if row[8] is None else row[8],
                m_item_color='' if row[9] is None else row[9],
                m_item_podklad='' if row[10] is None else row[10],
                m_item_stopped='' if row[11] is None else row[11],
                items_list=items_dict,
                m_item_model='' if row[12] is None else row[12],
            )
        elif table_name == "deleted_items_history":
            return RowObject(
                m_item_id='' if row[0] is None else row[0],
                table_name=table_name,
                m_item_order='' if row[1] is None else row[1],
                m_item_date_created='' if row[2] is None else row[2],
                m_item_comment='' if row[3] is None else row[3],
                step_id='' if row[4] is None else row[4],
                date_start='' if row[5] is None else row[5],
                comment='' if row[6] is None else row[6],
                m_item_title='' if row[7] is None else row[7],
                m_item_size='' if row[8] is None else row[8],
                m_item_color='' if row[9] is None else row[9],
                m_item_podklad='' if row[10] is None else row[10],
                m_item_stopped='' if row[11] is None else row[11],
                items_list=items_dict,
                m_item_model='' if row[12] is None else row[12],
                m_item_delete_date='' if row[13] is None else row[13],
            )
        elif table_name == "flow_sheet":
            return RowObject(
                m_item_id='' if row[0] is None else row[0],
                table_name=table_name,
                flow_shoe_block='' if row[14] is None else row[14],
                flow_model_article='' if row[16] is None else row[16],
                flow_color_article='' if row[19] is None else row[19],
                flow_size='' if row[23] is None else row[23],
                flow_range_position='' if row[25] is None else row[25],
                flow_product_article='' if row[26] is None else row[26],
                flow_sample_numbler='' if row[27] is None else row[27],
                flow_model_article2='' if row[28] is None else row[28],
                flow_color_article2='' if row[29] is None else row[29],
                flow_size2='' if row[30] is None else row[30],
                flow_podklad='' if row[31] is None else row[31],
                flow_color_filter='' if row[32] is None else row[32],
                flow_ordered='' if row[33] is None else row[33],
                flow_canceled='' if row[34] is None else row[34],
                flow_blanc='' if row[35] is None else row[35],
                flow_finish='' if row[36] is None else row[36],
                flow_ready='' if row[37] is None else row[37],
                flow_order_comment='' if row[45] is None else row[45],
                flow_barcode='' if row[46] is None else row[46],
                flow_control_digit='' if row[47] is None else row[47],
                flow_ean_13='' if row[48] is None else row[48],
                items_list=items_dict
            )

        elif table_name == "stw_modules_colors_local":
            return RowObject(
                m_item_id='' if row[0] is None else row[0],
                table_name=table_name,
                m_item_lang='' if row[1] is None else row[1],
                m_item_title='' if row[2] is None else row[2],
                m_item_visible='' if row[3] is None else row[3],
                items_list=items_dict,
            )
        else:
            return RowObject(
                m_item_id='' if row[0] is None else row[0],
                table_name=table_name,
                m_item_lang='' if row[1] is None else row[1],
                m_item_title='' if row[2] is None else row[2],
                items_list=items_dict,
            )

    def __read_prev_flow_data(self, report):
        ws = report['Актуальная выгрузка']
        max_row = ws.max_row
        prev_items_dict = {}
        for row in range(1, max_row + 1):
            cur_row = []
            if ws.cell(row=row, column=1).value:
                for col in range(1, 14):
                    cur_row.append(ws.cell(row=row, column=col).value)
                cur_row_object = self.__parse_row_object(table_name='items_history',
                                                         row=cur_row,
                                                         items_dict=prev_items_dict)
            cur_row_object.item_processing()
        return prev_items_dict

    def __read_flow_sheet_data(self, report):
        ws = report['Flow']
        max_row = ws.max_row
        flow_sheet_dict = {}
        for row in range(4, max_row + 1):
            cur_row = []
            if ws.cell(row=row, column=1).value:
                for col in range(1, 50):
                    cur_row.append(ws.cell(row=row, column=col).value)
                log.debug(cur_row)
                cur_row_object = self.__parse_row_object(table_name='flow_sheet',
                                                         row=cur_row,
                                                         items_dict=flow_sheet_dict)
                log.debug(
                    f"""Read m_item_id from Flow sheet: {cur_row_object.m_item_id}""")

                cur_row_object.item_processing()

        return flow_sheet_dict

    def __read_deleted_flow_data(self, report):
        ws = report['Удаленные заказы']
        max_row = ws.max_row
        deleted_items_dict = {}
        for row in range(1, max_row + 1):
            cur_row = []
            if ws.cell(row=row, column=1).value:
                for col in range(1, 15):
                    cur_row.append(ws.cell(row=row, column=col).value)
                cur_row_object = self.__parse_row_object(table_name='deleted_items_history',
                                                         row=cur_row,
                                                         items_dict=deleted_items_dict)
            cur_row_object.item_processing()
        return deleted_items_dict

    def __get_min_m_item_id_prev(self, items_dict):
        return min(items_dict)

    def __compare_attributes(self, values_new, values_prev):

        """ сравнение статусов изготовления изделий (шаг и останов)"""
        if (values_new.get('step_id') != values_prev.get("step_id") or
                values_new.get('m_item_stopped') != values_prev.get("m_item_stopped")):
            values_new.update({'m_item_step_status': 'Изменился статус изделия'})
        else:
            values_new.update({'m_item_step_status': ''})

        """ сравнение комментариев к изделию (comment и m_item_comment)"""

        if (values_new.get('comment') != values_prev.get("comment") or
                values_new.get('m_item_comment') != values_prev.get("m_item_comment")):
            values_new.update({'m_item_comment_status': 'Добавлен комментарий'})
            log.debug(f'result {values_new.get("m_item_comment_status")}')
        else:
            values_new.update({'m_item_comment_status': ''})
            log.debug(f'result {values_new.get("m_item_comment_status")}')

        """ сравнение размера изделия (m_item_size)"""
        if values_new.get('m_item_size') != values_prev.get("m_item_size"):
            values_new.update({'m_item_size_status': 'Изменился размер изделия'})
        else:
            values_new.update({'m_item_size_status': ''})

        """ сравнение цвета изделия (m_item_color)"""
        if values_new.get('m_item_color') != values_prev.get("m_item_color"):
            values_new.update({'m_item_color_status': 'Изменился цвет изделия'})
        else:
            values_new.update({'m_item_color_status': ''})

        """ сравнение подклада изделия (m_item_podklad)"""
        if values_new.get('m_item_podklad') != values_prev.get("m_item_podklad"):
            values_new.update({'m_item_podklad_status': 'Изменился подклад изделия'})
        else:
            values_new.update({'m_item_podklad_status': ''})

        """ сравнение модели изделия (m_item_model)"""
        if values_new.get('m_item_model') != values_prev.get("m_item_model"):
            values_new.update({'m_item_model_status': 'Изменилась модель изделия'})
        else:
            values_new.update({'m_item_model_status': ''})
        return values_new

    def __compare_items_dict(self, items_dict, prev_items_dict, deleted_items_dict):
        for m_item_id, values in items_dict.items():
            if m_item_id in prev_items_dict:
                """ статус изделия - учтено: проверить атрибуты на измненеия с предыдущей выгрузкой """
                values.update({'m_item_status': "учтено"})
                values = self.__compare_attributes(values_new=values, values_prev=prev_items_dict.get(m_item_id))
            else:
                """ статус изделия - новый, добавить на лист Актуальная выгрузка """
                values.update({'m_item_status': "новое изделие",
                               'm_item_step_status': '',
                               'm_item_comment_status': '',
                               'm_item_size_status': '',
                               'm_item_color_status': '',
                               'm_item_podklad_status': '',
                               'm_item_model_status': ''
                               })
            items_dict[m_item_id] = values

        for m_item_id, values in prev_items_dict.items():
            try:
                if int(m_item_id) < int(self.__get_min_m_item_id_prev(items_dict=items_dict)):
                    """ статус изделия - архивный, ничего не делаем """
                    pass
                else:
                    if m_item_id in items_dict:
                        log.debug("passing")
                    else:

                        """ статус изделия - удаленный, если отсутствует в удаленных изделиях, добавить на лист
                            удаленные заказы если есть, то ничего не делаем """
                        if m_item_id not in deleted_items_dict:
                            values.update({'m_item_delete_date': datetime.datetime.now().strftime("%d.%m.%Y")})
                            deleted_items_dict[m_item_id] = values

            except ValueError:
                continue

    def __get_flow_report(self, date_file_prefix, report_file_name):

        log.debug(f"report_path is {report_file_name}")
        report_data = openpyxl.load_workbook(filename=report_file_name, data_only=True)
        flow_sheet_dict = self.__read_flow_sheet_data(report=report_data)

        flow_items_dict = {}
        for elem, val in flow_sheet_dict.items():
            log.debug(f"readed from flow {elem}")
            log.debug(f"readed from flow {val}")
            flow_items_dict[elem] = val

        report_data.close()

        report = openpyxl.load_workbook(filename=report_file_name)
        for table_name in TABLES_NAMES:
            connector = DbConnector(table_name=table_name, db_name="origin")
            items_list = connector.get_items_list()
            items_dict = {}
            log.info(f"Processing query result table {table_name}")
            for row in items_list:
                cur_row_object = self.__parse_row_object(table_name=table_name, row=row, items_dict=items_dict)
                cur_row_object.item_processing()
            if table_name == 'items_history':
                prev_items_dict = self.__read_prev_flow_data(report=report)
                deleted_items_dict = self.__read_deleted_flow_data(report=report)
                self.__compare_items_dict(items_dict=items_dict,
                                          prev_items_dict=prev_items_dict,
                                          deleted_items_dict=deleted_items_dict)
                excel_reporter = ExcelReport(items_dict=deleted_items_dict,
                                             report_name='deleted_items_history',
                                             report=report)
                excel_reporter.items_processing()
                del excel_reporter

                deleted_items_dict = self.__read_deleted_flow_data(report=report)
                excel_reporter = ExcelReport(items_dict=items_dict, report_name=table_name, report=report)
                excel_reporter.items_processing()
                max_key = max(items_dict)
                full_items_list = list(range(max_key + 1))
                full_items_list.pop(0)
                excel_reporter_flow = ExcelReport(items_dict=items_dict,
                                                  report_name='Flow',
                                                  report=report,
                                                  flow_sheet_dict=flow_sheet_dict,
                                                  full_items_list=full_items_list,
                                                  deleted_items_dict=deleted_items_dict)
                excel_reporter_flow.items_processing()

        report.save(filename=report_file_name)
        del items_list
        del excel_reporter_flow
        del deleted_items_dict
        del excel_reporter
        del report

        log.info("Recalculating values")
        # report_data = openpyxl.load_workbook(filename=report_file_name, data_only=True)
        # report.save(filename=report_file_name)
        return report_file_name

    def run_db_report(self, date_file_prefix, report_file_name):
        return self.__get_flow_report(date_file_prefix=date_file_prefix, report_file_name=report_file_name)

    def read_prev_report(self):
        report = openpyxl.load_workbook(filename='TempFolder/20210917_1757_01_Flow_cut.xlsx')
        self.__read_prev_flow_data(report=report)
