import sys
import datetime

from Application.FlowReport.ReportNames.ReportNames import FLOW_REPORT_NAMES, ACCEPTANCE_REPORT_NAMES, \
    FLOW_RENEW_REPORT_NAMES
from Application.UpdaterStep.Steps.AcceptanceUpdaterStep import AcceptanceUpdaterStep
from Application.UpdaterStep.Steps.FlowUpdaterStep import FlowUpdaterStep
from Application.UpdaterStep.Steps.RenewFlowUpdaterStep import RenewFlowUpdaterStep
from Utils.ConfigReader.ConfigReader import read_config
from Utils.Logger.main_logger import get_logger

log = get_logger("Application")


class ReportManager:
    def __init__(self, args):
        self.args = args
        self.local_storage_path = read_config("local.storage").get("local_storage_path")

    def __choose_updater_type(self, step_num, date_file_prefix):
        if step_num == 'step_1':
            return FlowUpdaterStep(date_file_prefix=date_file_prefix)
        elif step_num == 'step_2':
            return AcceptanceUpdaterStep(date_file_prefix=date_file_prefix)
        elif step_num == 'step_3':
            return RenewFlowUpdaterStep(date_file_prefix=date_file_prefix)
        else:
            pass

    def updater_step(self, step_num, processing_reports):
        date_file_prefix = datetime.datetime.now()
        step_updater = self.__choose_updater_type(step_num=step_num, date_file_prefix=date_file_prefix)
        new_files = {}
        downloaded_files, downloaded_files_clean = step_updater.run_download_root_dir(
            processing_reports=processing_reports)

        step_updater.update_local_files(downloaded_files=downloaded_files,
                                        new_files=new_files)

        if step_num == 'step_1':
            step_updater.run_flow_export(report_file_name=new_files.get('01_Flow'), date_file_prefix=date_file_prefix)
            step_updater.clean_root_dir(files=downloaded_files_clean)
            step_updater.upload_local_files(new_files=new_files)
        elif step_num == 'step_2':
            acceptance_report_archive = step_updater.download_acceptance_archive()
            acceptance_report = step_updater.download_acceptance_template()
            log.debug(f'new_files {new_files}')
            log.debug(f'acceptance_report {acceptance_report}')
            step_updater.run_acceptance_updater(files=new_files,
                                                acceptance_report=acceptance_report,
                                                acceptance_report_archive=acceptance_report_archive)
            # step_updater.clean_root_dir(files=downloaded_files_clean)
            step_updater.upload_local_files(acceptance_report=acceptance_report)
        elif step_num == 'step_3':
            step_updater.run_renew_flow_updater(files=new_files)
            step_updater.clean_root_dir(files=downloaded_files_clean)
            step_updater.upload_local_files(new_files=new_files)

        step_updater.remove_local_files()

    def run(self):
        if len(self.args) == 1:
            print(
                """Application usage:
                main.py -<report shortname1> -<report shortname1>
                -f for flow report""")
            sys.exit(2)
        else:
            for arg in self.args:
                if arg == "-1":
                    self.updater_step(step_num='step_1', processing_reports=FLOW_REPORT_NAMES)
                elif arg == "-2":
                    self.updater_step(step_num='step_2', processing_reports=ACCEPTANCE_REPORT_NAMES)
                elif arg == "-3":
                    self.updater_step(step_num='step_3', processing_reports=FLOW_RENEW_REPORT_NAMES)
                elif arg == "main.py":
                    log.info("Starting application")
                else:
                    log.info("Unknown argument")
