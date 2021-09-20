import os

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
        for row in ws.rows:
            for cell in row:
                log.debug(cell.value)








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

            excel_reporter = ExcelReport(items_dict=items_dict, report_name=table_name, report=report)
            excel_reporter.items_processing()
        report.save(filename=report_file_name)
        return report_file_name

    def run_db_report(self, date_file_prefix, report_file_name):
        return self.__get_flow_report(date_file_prefix=date_file_prefix, report_file_name=report_file_name)

    def read_prev_report(self):
        report = openpyxl.load_workbook(filename='TempFolder/20210917_1757_01_Flow.xlsx')
        self.__read_prev_flow_data(report=report)
