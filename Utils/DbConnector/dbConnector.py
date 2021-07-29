import mysql.connector

from Processing.FlowReports.ProcessingQueries.Queries import ProcessingQueries
from Utils.ConfigReader.ConfigReader import read_config
from Utils.Logger.main_logger import get_logger

log = get_logger("DbConnector")


class DbConnector:
    def __init__(self, table_name=None, db_name=None):
        self.db_name = db_name
        self.table_name = table_name
        self.config = self.__get_config()
        self.connection = mysql.connector.connect(**self.config)
        self.cursor = self.connection.cursor(buffered=True)

    def __get_config(self):
        if self.db_name == 'origin':
            return read_config('flow.origin')
        else:
            return read_config('flow.buffer')

    def get_query(self):
        if self.table_name == 'items_history':
            return ProcessingQueries.items_history_query
        elif self.table_name == 'stw_modules_models_local':
            return ProcessingQueries.stw_modules_models_local_query
        elif self.table_name == 'stw_modules_colors_local':
            return ProcessingQueries.stw_modules_colors_local_query
        elif self.table_name == 'stw_modules_podklads_local':
            return ProcessingQueries.stw_modules_podklads_local_query
        elif self.table_name == 'stw_modules_sizes_local':
            return ProcessingQueries.stw_modules_sizes_local_query

    def get_data(self, cnx, cur):
        cur.execute(self.get_query())
        log.info("Execute query")
        items_list = cur.fetchall()
        log.info("Fetch result")
        cur.close()
        cnx.close()
        return items_list

    def get_items_list(self):
        cnx = mysql.connector.connect(**self.config)
        cur = cnx.cursor(buffered=True)
        log.info("Connect to db")
        if self.table_name:
            return self.get_data(cnx=cnx, cur=cur)
        else:
            return None
