import datetime
import os
import sys

from Application.FlowUpdateFiles.FlowUpdateFiles import FlowUpdateFiles
from Application.FlowUpdateFiles.ReportNames.ReportNames import FLOW_REPORT_NAMES
from Utils.ConfigReader.ConfigReader import read_config
from Utils.Logger.main_logger import get_logger

from Application.FlowReport.FlowReport import FlowReport
from Utils.YandexDisk.YaDiskConnector import YaDiskConnector

log = get_logger("Application")


class ReportManager:
    def __init__(self, args):
        self.args = args
        self.local_storage_path = read_config("local.storage").get("local_storage_path")

    def step_one(self):
        ya = YaDiskConnector()
        step_one_reports = FLOW_REPORT_NAMES
        downloaded_files = self.run_archive_root_dir(yandex_connector=ya,
                                                     processing_reports=step_one_reports)
        date_file_prefix = datetime.datetime.now().strftime("%Y%m%d_%H%M")
        new_files = {}
        self.update_local_files(downloaded_files=downloaded_files,
                                date_file_prefix=date_file_prefix,
                                new_files=new_files)

        self.run_flow_export(yandex_connector=ya,
                             date_file_prefix=date_file_prefix,
                             report_file_name=new_files.get('01_Flow'))

        self.renew_data_in_local_files(new_files=new_files)

        self.upload_local_files(yandex_connector=ya,
                                new_files=new_files)
        # self.remove_local_files()

    def __flow_updater(self):
        flow_report = FlowReport()
        flow_report.read_prev_report()

    def run(self):
        if len(self.args) == 1:
            print(
                """Application usage:
                main.py -<report shortname1> -<report shortname1>
                -f for flow report""")
            sys.exit(2)
        else:
            for arg in self.args:
                if "-1" == arg:
                    self.step_one()
                elif "-t" == arg:
                    self.__flow_updater()
                elif arg == "main.py":
                    log.info("Starting application")
                else:
                    log.info("Unknown argument")

    def run_archive_root_dir(self, yandex_connector, processing_reports):
        log.info("Running Yandex Disk archiving")
        downloaded_file = yandex_connector.archive_root_dir(processing_reports)
        return downloaded_file

    def run_flow_export(self, yandex_connector, date_file_prefix, report_file_name):
        log.info("Running Flow.Alfashoes report")
        flow_report = FlowReport()
        flow_report.run_db_report(date_file_prefix=date_file_prefix, report_file_name=report_file_name)
        # new_files.append(report_file_name)

    def update_filename(self, file, date_file_prefix, new_files, report_name):
        log.debug(f'Updating report {report_name}')
        new_filename = f"{date_file_prefix}_{report_name}.xlsx"
        prepared_filename = f'{self.local_storage_path}/{new_filename}'
        os.rename(file, prepared_filename)
        # new_files.append(prepared_filename)
        new_files.update({report_name: prepared_filename})

    def update_local_files(self, downloaded_files, date_file_prefix, new_files):
        for file in downloaded_files:
            log.debug(f"Working with {file}")

            if '00_Справочник Альфа' in file:
                self.update_filename(file=file,
                                     date_file_prefix=date_file_prefix,
                                     new_files=new_files,
                                     report_name='00_Справочник Альфа')
            elif '00_Справочник Номенклатуры WB' in file:
                self.update_filename(file=file,
                                     date_file_prefix=date_file_prefix,
                                     new_files=new_files,
                                     report_name='00_Справочник Номенклатуры WB')
            elif '01_Flow' in file:
                self.update_filename(file=file,
                                     date_file_prefix=date_file_prefix,
                                     new_files=new_files,
                                     report_name='01_Flow')

    def upload_local_files(self, yandex_connector, new_files):
        for report_type, file in new_files.items():
            clean_filename = file.replace('TempFolder/', '')
            yandex_connector.upload_file(src_path=file,
                                         destination_file_path=f"/Учет Альфа/{clean_filename}")

    def remove_local_files(self):
        local_files = os.listdir(self.local_storage_path)
        for file in local_files:
            filepath = os.path.join(self.local_storage_path, file)
            log.info(f"Removing local file {filepath}")
            os.remove(filepath)

    def renew_data_in_local_files(self, new_files):
        flow_updater = FlowUpdateFiles(new_files=new_files)
        flow_updater.run()
