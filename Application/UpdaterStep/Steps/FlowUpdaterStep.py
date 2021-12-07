from Application.FlowReport.FlowReport import FlowReport

from Application.UpdaterStep.UpdaterStep import UpdaterStep
from Utils.Logger.main_logger import get_logger

log = get_logger("FlowUpdaterStep")


class FlowUpdaterStep(UpdaterStep):

    @staticmethod
    def run_flow_export(report_file_name):
        log.info("Running Flow.Alfashoes report")
        flow_report = FlowReport()
        flow_report.run_db_report(report_file_name=report_file_name)
