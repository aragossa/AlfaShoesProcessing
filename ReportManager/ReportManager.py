import sys

from Updating.UpdatingMethods import UpdatingMethods
from Utils.Logger.main_logger import get_logger

from ReportManager.WildberriesReport import WildberriesReport
from ReportManager.FlowReport.FlowReport import FlowReport

log = get_logger("ReportManager")


class ReportManager:
    def __init__(self, args):
        self.args = args

    def run(self):
        if len(self.args) == 1:
            print(
                """Application usage:
                main.py -<report shortname1> -<report shortname1>
                -w for Wildberries report
                -f for flow report""")
            sys.exit(2)
        else:
            for arg in self.args:
                if '-w' == arg:
                    self.run_wildberries()
                elif '-f' == arg:
                    self.run_flow_update()
                    self.run_flow_export()
                elif arg == 'main.py':
                    log.info('Starting application')
                else:
                    print(
                        """main.py -<report shortname1> -<report shortname1>
                        -w for Wildberries report
                        -f for flow report""")

    def run_wildberries(self):
        log.info('Running Wildberries report')
        WildberriesReport.run()

    def run_flow_export(self):
        log.info('Running Flow.Alfashoes report')
        FlowReport.run()

    def run_flow_update(self):
        UpdatingMethods.run()
        log.info('Running Flow.Alfashoes update')
