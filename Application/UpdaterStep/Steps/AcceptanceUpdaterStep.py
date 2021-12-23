import datetime
import os

import openpyxl

from Application.UpdaterStep.Steps.Intervals import Intervals
from Application.UpdaterStep.UpdaterStep import UpdaterStep
from Utils.Logger.main_logger import get_logger

log = get_logger("AcceptanceUpdaterStep")


class AcceptanceUpdaterStep(UpdaterStep):

    def download_acceptance_template(self):
        acceptance_template_dst_filepath = os.path.join(self.local_storage_path, self.acceptance_template_dst_filename)
        acceptance_template_src_filepath = 'Учет Альфа/Шаблоны учет Альфа/шаблон_02_Приемка на склад.xlsx'
        self.ya.download_file(acceptance_template_src_filepath, acceptance_template_dst_filepath)
        return acceptance_template_dst_filepath

    def prepare_acceptance_values(self, report_name, report_file_name, report):
        log.debug(f'reading {report_name}')
        read_values = Intervals.acceptance_intervals.get(report_name)
        log.debug(read_values)
        log.debug(f'opening {report_file_name}')
        data = openpyxl.load_workbook(filename=report_file_name, data_only=True, read_only=True)
        for sheet_name, params in read_values.items():
            log.debug(sheet_name)
            worksheet = data[sheet_name]
            if sheet_name == 'Flow':
                write_sheetname = 'Flow данные для приемки'
            else:
                write_sheetname = sheet_name
            report_sheet = report[write_sheetname]
            if params.get('cells') is not None:
                for cells in params.get('cells'):
                    value = self.read_cell_value(worksheet=worksheet, col=cells.get('col'), row=cells.get('row'))
                    self.write_cell_value(worksheet=report_sheet, col=cells.get('col'), row=cells.get('row'),
                                          value=value)
            for interval in params.get('intervals'):
                log.debug(f"current_interval")
                log.debug(f"{interval}")
                excel_data = self.read_interval(worksheet=worksheet,
                                                start_row=interval.get("read").get('start_row'),
                                                stop_row=interval.get("read").get('stop_row'),
                                                start_col=interval.get("read").get('start_col'),
                                                stop_col=interval.get("read").get('stop_col'))
                self.write_interval(worksheet=report_sheet,
                                    start_row=interval.get("write").get('start_row'),
                                    start_col=interval.get("write").get('start_col'),
                                    excel_data=excel_data)
        data.close()

    def read_cell_value(self, worksheet, col, row):
        log.debug('reading cell value')
        return worksheet.cell(row=row, column=col).value

    def write_cell_value(self, worksheet, col, row, value):
        log.debug('writing cell value')
        worksheet.cell(row=row, column=col).value = value

    def run_acceptance_updater(self, files, acceptance_report):
        report = openpyxl.load_workbook(filename=acceptance_report)
        for report_name, report_file_name in files.items():
            log.debug(f'opening {acceptance_report}')
            self.prepare_acceptance_values(report_name=report_name, report_file_name=report_file_name, report=report)
            log.info(f'saving report {acceptance_report}')
        report.save(acceptance_report)
        report.close()

    def clean_root_dir(self, files):
        for file in files:
            if '02_Приемка на склад' in file:
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
            if '02_Приемка на склад' in file:
                clean_filename = file.replace('TempFolder/', '').replace('TempFolder\\', '')
                self.ya.upload_file(src_path=file,
                                    destination_file_path=f"/Учет Альфа/{clean_filename}")
