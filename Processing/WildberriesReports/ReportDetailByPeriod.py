from datetime import datetime

from Processing.WildberriesReports.Headers.TableHeaders import wildberries_reportdetailbyperiod_headers_list
from Utils.ConfigReader.ConfigReader import read_config
from Processing.WildberriesProcessingMethods import WildberriesProcessingMethods
from Utils.Logger.main_logger import get_logger

log = get_logger("ReportDetailByPeriod")


class ReportDetailByPeriod(WildberriesProcessingMethods):
    def __init__(self, workbook):
        self.report_name = 'Отчет о продажах по реализации'
        super().__init__(workbook=workbook,
                         url=self.__url_builder(),
                         report_name=self.report_name,
                         headers=self.__get_headers())

    @staticmethod
    def __url_builder():
        app_mode = read_config('application').get('mode')
        if app_mode == 'prod':
            api_key = read_config('wildberries').get('api')
            date_from = read_config('wildberries.reportdetailbyperiod').get('datefrom')
            print(date_from)
            date_to = read_config('wildberries.reportdetailbyperiod').get('dateto')
            print(date_to)
            limit = read_config('wildberries.reportdetailbyperiod').get('limit')
            print(limit)
            rrdid = read_config('wildberries.reportdetailbyperiod').get('rrdid')
            print(rrdid)
            log.info(f"Report dates {' '.join([date_from, date_to])}")
            date_from_date = datetime.strptime(date_from, '%d.%m.%Y')
            date_from_form = f'{date_from_date.year:04d}-{date_from_date.month:02d}-{date_from_date.day:02d}'
            date_to_date = datetime.strptime(date_to, '%d.%m.%Y')
            date_to_form = f'{date_to_date.year:04d}-{date_to_date.month:02d}-{date_to_date.day:02d}'
            url = f"https://suppliers-stats.wildberries.ru/api/v1/supplier/reportDetailByPeriod?dateFrom={date_from_form}&key={api_key}&limit={limit}&rrdid={rrdid}&dateto={date_to_form}"
        else:
            url = 'http://127.0.0.1:5000/wildberries/reportdetailbyperiod'
        return url

    @staticmethod
    def __get_headers():
        return wildberries_reportdetailbyperiod_headers_list
