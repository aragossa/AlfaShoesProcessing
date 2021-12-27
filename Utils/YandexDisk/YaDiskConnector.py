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

    def delete_file(self, src_path):
        log.info(f"Removing file {src_path}")
        self.connection.remove(path=src_path, permanently=True)
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

    def find_archive_dirs(self):
        list_dir = self.connection.listdir("/Учет Альфа/Архив учета Альфа/")
        dir_names = []
        for dir in list_dir:
            if dir.type == "dir":
                try:
                    log.debug(f"{dir.name}")
                    this_datetime = datetime.datetime.strptime(dir.name, "%Y%m%d")
                    log.debug(this_datetime)
                    dir_names.append(f"/Учет Альфа/Архив учета Альфа/{dir.name}")
                except ValueError:
                    log.debug("dir is not archive")
        return dir_names

    def find_max_report(self, report_name, dir_name):
        list_dir = self.connection.listdir(dir_name)
        file_name = None
        for file in list_dir:
            if (file.type == "file") and (report_name in file.name):
                try:
                    log.debug(f"{file.name}")
                    file_change_date = f"{file.name.split('_')[0]}_{file.name.split('_')[1]}"
                    log.debug(file_change_date)
                    this_datetime = datetime.datetime.strptime(file_change_date, "%Y%m%d_%H%M")
                    if file_name is not None:
                        exist_file_change_date = f"{file_name.name.split('_')[0]}_{file_name.name.split('_')[1]}"
                        log.debug(exist_file_change_date)
                        exist_this_datetime = datetime.datetime.strptime(exist_file_change_date, "%Y%m%d_%H%M")
                        if this_datetime > exist_this_datetime:
                            file_name = file.name
                    else:
                        file_name = file.name
                except ValueError:
                    log.debug("file is not archive")
        return file_name

    def find_max_max(self):
        list_dir = self.connection.listdir("/Учет Альфа/")
        dates = []
        for file in list_dir:
            if file.type == "file":
                try:
                    log.debug(f"{file.name}")
                    this_datetime = datetime.datetime.strptime(file.name.split('_')[0], "%Y%m%d")
                    dates.append(this_datetime)
                except ValueError:
                    log.debug("file is not archive")
        return max(dates).strftime("%Y%m%d")

    def find_files_in_dir(self, dir_name):
        files = []
        for elem in self.list_dir(dir_name=dir_name):
            if elem.type == 'file':
                files.append(elem)
        return files
