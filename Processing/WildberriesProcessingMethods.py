import json
import requests

from Utils.Logger.main_logger import get_logger

log = get_logger("ProcessingMethods")


class WildberriesProcessingMethods:

    def __init__(self, workbook, url, report_name, headers):
        self.__url = url
        self.__report_name = report_name
        self.__row_cursor = 1
        self.__headers = headers
        self.workbook = workbook

    def __get_data(self):
        log.info(f"Sending GET request")
        log.info(f"URL: {self.__url}")
        response = requests.get(self.__url).text
        return response

    def __get_json(self):
        return json.loads(self.__get_data())

    def __add_worksheet(self):
        return self.workbook.add_worksheet(self.__report_name)

    def __write_row(self, row, ws):
        col_cursor = 0
        for item in row.values():
            ws.write(self.__row_cursor, col_cursor, item)
            col_cursor += 1

    def __write_headers(self, ws):
        col_cursor = 0
        for item in self.__headers:
            ws.write(0, col_cursor, item)
            col_cursor += 1

    def main_processor(self):
        log.info(f'Processing report {self.__report_name}')
        ws = self.__add_worksheet()
        log.info(f"Adding Excel worksheet {self.__report_name}")
        self.__write_headers(ws=ws)
        data = self.__get_json()
        for row in data:
            self.__write_row(row=row, ws=ws)
            self.__row_cursor += 1
        log.info(f'Report {self.__report_name} finished')
        return ws
