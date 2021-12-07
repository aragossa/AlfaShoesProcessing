import os
import datetime

from Utils.ConfigReader.ConfigReader import read_config
from Utils.Logger.main_logger import get_logger
from Utils.YandexDisk.YaDiskConnector import YaDiskConnector

log = get_logger("TemplateUpdaterStep")


class UpdaterStep:
    def __init__(self):
        self.local_storage_path = read_config("local.storage").get("local_storage_path")
        self.ya = YaDiskConnector()
        self.date_file_prefix = datetime.datetime.now().strftime("%Y%m%d_%H%M")
        self.acceptance_template_dst_filename = f"{self.date_file_prefix}_02_Приемка на склад.xlsx"

    def run_download_root_dir(self, processing_reports):
        log.info("Running Yandex Disk archiving")
        downloaded_file = self.ya.download_root_dir(processing_reports)
        return downloaded_file

    def clean_root_dir(self, files):
        for file in files:
            src_file = file.replace("TempFolder/", "")
            self.ya.delete_file(src_path=f"/Учет Альфа/{src_file}")

    def update_filename(self, file, new_files, report_name):
        log.debug(f'Updating report {report_name}')
        new_filename = f"{self.date_file_prefix}_{report_name}.xlsx"
        prepared_filename = os.path.join(self.local_storage_path, new_filename)
        os.rename(file, prepared_filename)
        new_files.update({report_name: prepared_filename})

    def update_local_files(self, downloaded_files, new_files):
        for file in downloaded_files:
            log.debug(f"Working with {file}")

            if '00_Справочник Альфа' in file:
                self.update_filename(file=file,
                                     new_files=new_files,
                                     report_name='00_Справочник Альфа')
            elif '00_Справочник Номенклатуры WB' in file:
                self.update_filename(file=file,
                                     new_files=new_files,
                                     report_name='00_Справочник Номенклатуры WB')
            elif '01_Flow' in file:
                self.update_filename(file=file,
                                     new_files=new_files,
                                     report_name='01_Flow')
            elif '02_Приемка на склад' in file:
                self.update_filename(file=file,
                                     new_files=new_files,
                                     report_name='02_Приемка на склад')

    def upload_local_files(self, new_files):
        for report_type, file in new_files.items():
            clean_filename = file.replace('TempFolder/', '').replace('TempFolder\\', '')
            self.ya.upload_file(src_path=file,
                                destination_file_path=f"/Учет Альфа/{clean_filename}")

    def remove_local_files(self):
        local_files = os.listdir('TempFolder')
        for file in local_files:
            filepath = os.path.join('TempFolder', file)
            log.info(f"Removing local file {filepath}")
            os.remove(filepath)

    """ Acceptance report """

    def check_null_row(self, row):
        all_none = True
        for elem in row:
            if elem != None:
                all_none = False
        return all_none

    def read_interval(self, worksheet, start_row, stop_row, start_col, stop_col):
        log.debug('reading intervals values')
        max_row = worksheet.max_row
        log.debug(max_row)
        if max_row < stop_row:
            break_row = max_row
        else:
            break_row = stop_row
        excel_data = []
        log.debug(f'max_row: {max_row}, stop_row:{stop_row}')
        cont = True
        for row in worksheet.iter_rows(min_row=start_row, max_row=break_row, min_col=start_col, max_col=stop_col):
            if cont:
                cur_row = []
                for cell in row:
                    cur_row.append(cell.value)
                if self.check_null_row(cur_row):
                    cont = False
                else:
                    excel_data.append(cur_row)
            else:
                log.debug('found empty row')
                break
        return excel_data

    def write_interval(self, worksheet, start_row, start_col, excel_data):
        log.debug('writing intervals values')
        row = start_row
        for rows in excel_data:
            log.debug(rows)
            col = start_col
            for cell_value in rows:
                log.debug(f'writing to {worksheet} ({row}, {col}) value - {cell_value}')
                worksheet.cell(row=row, column=col).value = cell_value
                col += 1
            row += 1
