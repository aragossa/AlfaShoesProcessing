from datetime import datetime

from Utils.Logger.main_logger import get_logger

log = get_logger("RowObject")


class RowObject:

    def __init__(self,
                 m_item_id,
                 table_name,
                 m_item_order=None,
                 m_item_date_created=None,
                 m_item_comment=None,
                 step_id=None,
                 date_start=None,
                 comment=None,
                 m_item_title=None,
                 m_item_size=None,
                 m_item_color=None,
                 m_item_podklad=None,
                 m_item_stopped=None,
                 items_list=None,
                 m_item_visible=None,
                 m_item_lang=None,
                 m_item_model=None,
                 ):

        self.m_item_id = m_item_id
        self.m_item_order = m_item_order
        self.m_item_date_created = m_item_date_created
        self.m_item_comment = m_item_comment
        self.step_id = step_id
        self.date_start = date_start
        self.comment = comment
        self.m_item_title = m_item_title
        self.m_item_size = m_item_size
        self.m_item_color = m_item_color
        self.m_item_podklad = m_item_podklad
        self.m_item_stopped = m_item_stopped
        self.items_list = items_list
        self.m_item_lang = m_item_lang
        self.m_item_visible = m_item_visible
        self.table_name = table_name
        self.m_item_model = m_item_model

    def check_items_list(self):
        if self.items_list.get(self.m_item_id):

            return True
        else:
            return False

    def convert_to_date_string(self, timestamp):
        return datetime.fromtimestamp(timestamp).strftime("%d.%m.%Y")

    def create_dict_obj(self):
        if self.table_name == 'items_history':
            return {
                'm_item_order': self.m_item_order,
                'm_item_date_created': self.convert_to_date_string(self.m_item_date_created),
                'm_item_comment': self.m_item_comment,
                'step_id': self.step_id,
                'date_start': self.convert_to_date_string(self.date_start),
                'comment': self.comment,
                'm_item_title': self.m_item_title,
                'm_item_size': self.m_item_size,
                'm_item_color': self.m_item_color,
                'm_item_podklad': self.m_item_podklad,
                'm_item_stopped': self.m_item_stopped,
                'm_item_model': self.m_item_model,
            }
        elif self.table_name == 'stw_modules_colors_local':
            log.debug({
                'm_item_title': self.m_item_title,
                'm_item_lang': self.m_item_lang,
                'm_item_visible': self.m_item_visible,
            })
            return {
                'm_item_title': self.m_item_title,
                'm_item_lang': self.m_item_lang,
                'm_item_visible': self.m_item_visible,
            }

        else:
            return {
                'm_item_title': self.m_item_title,
                'm_item_lang': self.m_item_lang,
            }

    def insert_item(self):
        self.items_list[self.m_item_id] = self.create_dict_obj()

    def update_item(self):
        self.items_list[self.m_item_id].update(self.create_dict_obj())

    def compare_step_id(self):
        if self.items_list[self.m_item_id].get('step_id') < self.step_id:
            self.update_item()

    def item_processing(self):
        if self.check_items_list():
            self.compare_step_id()
        else:
            self.insert_item()
        return self.items_list
