"""
        for elem in list(yandex_disk.listdir('Учет Альфа/')):
            log.debug(f"-----------------------------------------------------------------")
            log.debug(f"file {elem.file}")
            log.debug(f"size {elem.size}")
            log.debug(f"public_key {elem.public_key}")
            log.debug(f"name {elem.name}")
            log.debug(f"modified {elem.modified}")
            log.debug(f"created {elem.created}")
            log.debug(f"path {elem.path}")
"""
import datetime
import os

import yadisk

from Utils.ConfigReader.ConfigReader import read_config
from Utils.Logger.main_logger import get_logger

log = get_logger("YaDiskConnector")


class YaDiskConnector:
    def __init__(self):
        self.token = read_config("yandex.disk").get("token")
        self.local_storage_path = read_config("local.storage").get("local_storage_path")
        self.connection = yadisk.YaDisk(token=self.token)

    def upload_file(self, src_path, destination_file_path, overwrite=False):
        log.info(f"Uploading file {src_path} to {destination_file_path}")
        self.connection.upload(path_or_file=src_path, dst_path=destination_file_path, overwrite=overwrite)
        log.info("Done")

    def download_file(self, src_path, path_or_file):
        log.info(f"Downloading file {src_path} to {path_or_file}")
        self.connection.download(src_path=src_path, path_or_file=path_or_file)
        log.info("Done")

    def copy_file(self, src_path, destination_file_path, overwrite=False):
        log.info(f"Copying file {src_path} to {destination_file_path}")
        self.connection.copy(src_path=src_path, dst_path=destination_file_path, overwrite=overwrite)
        log.info("Done")

    def move_file(self, src_path, destination_file_path, overwrite=False):
        log.info(f"Moving file {src_path} to {destination_file_path}")
        self.connection.move(src_path=src_path, dst_path=destination_file_path, overwrite=overwrite)
        log.info("Done")

    def mkdir(self, dir_name):
        try:
            log.info(f"Creating folder: {dir_name}")
            self.connection.mkdir(dir_name)
            log.info("Done")
        except yadisk.exceptions.DirectoryExistsError:
            log.debug(f"folder '{dir_name}' already exist")

    def list_dir(self, dir_name):
        return self.connection.listdir({dir_name})

    def find_max_dir(self):
        list_dir = self.connection.listdir("/Учет Альфа/Архив учета Альфа/")
        dir_names = []
        for dir in list_dir:
            if dir.type == "dir":
                try:
                    log.debug(f"{dir.name}")
                    this_datetime = datetime.datetime.strptime(dir.name, "%Y%m%d")
                    dir_names.append(this_datetime)
                except ValueError:
                    log.debug("dir is not archive")
        return max(dir_names).strftime("%Y%m%d")

    def find_files_in_dir(self, dir_name):
        files = []
        for elem in self.list_dir(dir_name=dir_name):
            if elem.type == 'file':
                files.append(elem)
        return files

    def archive_root_dir(self, processing_reports):
        src_path = "/Учет Альфа/"
        files = self.find_files_in_dir(src_path)
        filenames = [elem.name for elem in files]
        destination_folder_name = datetime.datetime.now().strftime('%Y%m%d')

        destination_path = f"/Учет Альфа/Архив учета Альфа/{destination_folder_name}/"
        self.mkdir(destination_path)
        downloaded_files = []
        for file in filenames:
            archive_report = False
            for report_name in processing_reports:
                if report_name in file:
                    archive_report = True
            if archive_report:
                log.debug(f"Processing file {file}")
                prepared_file_name = file.replace(".xlsx", "_архив.xlsx")
                src_file_path = f"{src_path}{file}"
                destination_file_path = f"{destination_path}{prepared_file_name}"

                download_file_path = os.path.join(self.local_storage_path, prepared_file_name)
                self.download_file(src_path=src_file_path,
                                   path_or_file=download_file_path)
                downloaded_files.append(download_file_path)

                self.move_file(src_path=src_file_path,
                               destination_file_path=destination_file_path,
                               overwrite=True)
            else:
                log.debug(f"File {file} skipped")

        return downloaded_files



