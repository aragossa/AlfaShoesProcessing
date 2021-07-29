import os

import xlsxwriter

from Processing.WildberriesReports.ReportDetailByPeriod import ReportDetailByPeriod
from Processing.WildberriesReports.ReportIncomes import ReportIncomes
from Processing.WildberriesReports.ReportOrders import ReportOrders
from Processing.WildberriesReports.ReportSales import ReportSales
from Processing.WildberriesReports.ReportStocks import ReportStocks


class WildberriesReport:

    @staticmethod
    def run():
        workbook = xlsxwriter.Workbook(os.path.join('Reports', 'WildberriesReport.xlsx'))
        stocks = ReportStocks(workbook)
        stocks.main_processor()
        sales = ReportSales(workbook)
        sales.main_processor()
        incomes = ReportIncomes(workbook)
        incomes.main_processor()
        orders = ReportOrders(workbook)
        orders.main_processor()
        #report_detail_by_period = ReportDetailByPeriod(workbook)
        #report_detail_by_period.main_processor()
        workbook.close()