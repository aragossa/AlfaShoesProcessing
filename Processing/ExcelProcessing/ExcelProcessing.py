import string

import openpyxl


class ExcelProcessing:
    def __init__(self, file_list):
        self.file_list = file_list

    def __column_2_number_converter(self, col):
        num = 0
        for c in col:
            if c in string.ascii_letters:
                num = num * 26 + (ord(c.upper()) - ord('A')) + 1
        return num

    def __get_report_file(self, report_type):
        for file in self.file_list:
            if report_type in file:
                return file

    def __open_excel_file(self, filepath, data_only=True):
        wb = openpyxl.load_workbook(filename=filepath, data_only=data_only)
        return wb

    def __get_sheet(self, workbook, sheetname):
        return workbook.get_sheet_by_name(sheetname)

    def get_data(self, report_type):
        filename = self.__get_report_file(report_type=report_type)
        wb = self.__open_excel_file(filepath=filename)
        return wb

    def get_cell_values(self, start_col, stop_col, start_row, stop_row):
        start_col_num = self.__column_2_number_converter(start_col)
        stop_col_num = self.__column_2_number_converter(stop_col)





