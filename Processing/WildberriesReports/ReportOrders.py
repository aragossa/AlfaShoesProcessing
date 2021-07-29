from datetime import datetime

from Processing.WildberriesReports.Headers.TableHeaders import wildberries_orders_headers_list
from Utils.ConfigReader.ConfigReader import read_config
from Processing.WildberriesProcessingMethods import WildberriesProcessingMethods
from Utils.Logger.main_logger import get_logger

log = get_logger("ReportOrders")

class ReportOrders(WildberriesProcessingMethods):
    def __init__(self, workbook):
        self.report_name = 'Заказы'
        super().__init__(workbook=workbook,
                         url=self.__url_builder(),
                         report_name=self.report_name,
                         headers=self.__get_headers())

    @staticmethod
    def __url_builder():
        app_mode = read_config('application').get('mode')
        if app_mode == 'prod':
            api_key = read_config('wildberries').get('api')
            date_config = read_config('wildberries.ordersreport').get('reportdate')
            time_config = read_config('wildberries.ordersreport').get('reporttime')
            log.info(f"Report datetime {' '.join([date_config, time_config])}")
            date = datetime.strptime(' '.join([date_config, time_config]), '%d.%m.%Y %H:%M')
            date_form = f'{date.year:04d}-{date.month:02d}-{date.day:02d}T{date.hour:02d}%3A00%3A00.000Z'
            url = f'https://suppliers-stats.wildberries.ru/api/v1/supplier/orders?dateFrom={date_form}&key={api_key}'
        else:
            url = 'http://127.0.0.1:5000/wildberries/orders'
        return url

    @staticmethod
    def __get_headers():
        return wildberries_orders_headers_list

