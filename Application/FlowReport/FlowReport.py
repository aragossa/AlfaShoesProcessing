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
                m_item_id=row[0],
                table_name=table_name,
                m_item_order=row[1],
                m_item_date_created=row[2],
                m_item_comment=row[3],
                step_id=row[4],
                date_start=row[5],
                comment=row[6],
                m_item_title=row[7],
                m_item_size=row[8],
                m_item_color=row[9],
                m_item_podklad=row[10],
                m_item_stopped=row[11],
                items_list=items_dict,
                m_item_model=row[12],
            )
        elif table_name == "deleted_items_history":
            return RowObject(
                m_item_id=row[0],
                table_name=table_name,
                m_item_order=row[1],
                m_item_date_created=row[2],
                m_item_comment=row[3],
                step_id=row[4],
                date_start=row[5],
                comment=row[6],
                m_item_title=row[7],
                m_item_size=row[8],
                m_item_color=row[9],
                m_item_podklad=row[10],
                m_item_stopped=row[11],
                items_list=items_dict,
                m_item_model=row[12],
                m_item_delete_date=row[13],
            )
        elif table_name == "stw_modules_colors_local":
            return RowObject(
                m_item_id=row[0],
                table_name=table_name,
                m_item_lang=row[1],
                m_item_title=row[2],
                m_item_visible=row[3],
                items_list=items_dict,
            )
        else:
            return RowObject(
                m_item_id=row[0],
                table_name=table_name,
                m_item_lang=row[1],
                m_item_title=row[2],
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
            values_new.update({'m_item_step_status': 'Статус изделия изменился'})
        else:
            values_new.update({'m_item_step_status': ''})

        """ сравнение комментариев к изделию (comment и m_item_comment)"""
        log.debug(f"values_new.get('comment') != values_prev.get('comment') -- {values_new.get('comment') != values_prev.get('comment')}")
        log.debug(f"values_new.get('m_item_comment') != values_prev.get('m_item_comment')) -- {values_new.get('m_item_comment') != values_prev.get('m_item_comment')}")
        if (values_new.get('comment') != values_prev.get("comment") or
                values_new.get('m_item_comment') != values_prev.get("m_item_comment")):
            values_new.update({'m_item_comment_status': 'Добавлен комментарий'})
        else:
            values_new.update({'m_item_comment_status': ''})

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

                        """ статус изделия - удаленный, если отсутствует в удаленных изделиях, добавить на лист удаленные заказы
                            если есть, то ничего не делаем """
                        if m_item_id not in deleted_items_dict:
                            values.update({'m_item_delete_date': datetime.datetime.now().strftime("%d.%m.%Y")})
                            deleted_items_dict[m_item_id] = values

            except ValueError:
                continue


    def __get_flow_report(self, date_file_prefix, report_file_name):

        log.debug(f"report_path is {report_file_name}")

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

            excel_reporter = ExcelReport(items_dict=items_dict, report_name=table_name, report=report)
            excel_reporter.items_processing()

        report.save(filename=report_file_name)
        return report_file_name

    def run_db_report(self, date_file_prefix, report_file_name):
        return self.__get_flow_report(date_file_prefix=date_file_prefix, report_file_name=report_file_name)

    def read_prev_report(self):
        report = openpyxl.load_workbook(filename='TempFolder/20210917_1757_01_Flow_cut.xlsx')
        self.__read_prev_flow_data(report=report)
