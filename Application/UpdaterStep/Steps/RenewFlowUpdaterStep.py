import os

import openpyxl

from Application.UpdaterStep.UpdaterStep import UpdaterStep
from Utils.Logger.main_logger import get_logger
from Application.UpdaterStep.Steps.Intervals import Intervals

log = get_logger("RenewFlowUpdaterStep")


class RenewFlowUpdaterStep(UpdaterStep):

    def update_filename(self, file, new_files, report_name):
        log.debug(f'Updating report {report_name}')
        new_filename = f"{self.date_file_prefix}_{report_name}.xlsx"
        prepared_filename = os.path.join(self.local_storage_path, new_filename)
        os.rename(file, prepared_filename)
        new_files.update({report_name: prepared_filename})

    def update_local_files(self, downloaded_files, new_files):
        for file in downloaded_files:
            log.debug(f"Working with {file}")

            if '01_Flow' in file:
                self.update_filename(file=file,
                                     new_files=new_files,
                                     report_name='01_Flow')
            elif '02_Приемка на склад' in file:
                new_files.update({'02_Приемка на склад': file})

    def prepare_renew_flow_values(self, report_file_name, report):
        read_values = Intervals.renew_flow_intervals.get('02_Приемка на склад')
        log.debug(f'opening {report_file_name}')
        data = openpyxl.load_workbook(filename=report_file_name, data_only=True, read_only=True)
        for sheet_name, params in read_values.items():
            log.debug(sheet_name)
            worksheet = data[sheet_name]
            report_sheet = report['Flow']
            for interval in params.get('intervals'):
                excel_data = self.read_interval(worksheet=worksheet,
                                                start_row=interval.get('read').get('start_row'),
                                                stop_row=interval.get('read').get('stop_row'),
                                                start_col=interval.get('read').get('start_col'),
                                                stop_col=interval.get('read').get('stop_col'))
                self.write_interval(worksheet=report_sheet,
                                    start_row=interval.get('write').get('start_row'),
                                    start_col=interval.get('write'),
                                    excel_data=excel_data)
        data.close()

    def run_renew_flow_updater(self, files):
        log.debug(files)
        report = openpyxl.load_workbook(filename=files.get('01_Flow'))
        log.debug(f"opening {files.get('01_Flow')}")
        self.prepare_renew_flow_values(report_file_name=files.get('02_Приемка на склад'), report=report)
        log.info(f"saving report {files.get('01_Flow')}")
        report.save(files.get('01_Flow'))
        report.close()

    def clean_root_dir(self, files):
        for file in files:
            if '01_Flow' in file:
                src_file = file.replace("TempFolder/", "")
                prepared_file_name = src_file.replace(".xlsx", "_архив.xlsx")
                destination_folder_name = prepared_file_name.split('_')[0]
                destination_path = f"/Учет Альфа/Архив учета Альфа/{destination_folder_name}/"
                destination_file_path = f"{destination_path}{prepared_file_name}"
                self.ya.mkdir(destination_path)
                self.ya.copy_file(src_path=f"/Учет Альфа/{src_file}",
                                  destination_file_path=destination_file_path,
                                  overwrite=True)

                self.ya.delete_file(src_path=f"/Учет Альфа/{src_file}")

    def upload_local_files(self, new_files):
        for report_type, file in new_files.items():
            if '01_Flow' in file:
                clean_filename = file.replace('TempFolder/', '').replace('TempFolder\\', '')
                self.ya.upload_file(src_path=file,
                                    destination_file_path=f"/Учет Альфа/{clean_filename}")