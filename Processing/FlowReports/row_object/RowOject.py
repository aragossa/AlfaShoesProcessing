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
                 m_item_delete_date=None,
                 flow_shoe_block=None,
                 flow_model_article=None,
                 flow_color_article=None,
                 flow_size=None,
                 flow_range_position=None,
                 flow_product_article=None,
                 flow_sample_numbler=None,
                 flow_model_article2=None,
                 flow_color_article2=None,
                 flow_size2=None,
                 flow_podklad=None,
                 flow_color_filter=None,
                 flow_ordered=None,
                 flow_canceled=None,
                 flow_blanc=None,
                 flow_finish=None,
                 flow_ready=None,
                 flow_order_comment=None,
                 flow_barcode=None,
                 flow_control_digit=None,
                 flow_ean_13=None
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
        self.m_item_delete_date = m_item_delete_date
        self.flow_shoe_block = flow_shoe_block
        self.flow_model_article = flow_model_article
        self.flow_color_article = flow_color_article
        self.flow_size = flow_size
        self.flow_product_article = flow_product_article
        self.flow_range_position = flow_range_position
        self.flow_sample_numbler = flow_sample_numbler
        self.flow_model_article2 = flow_model_article2
        self.flow_color_article2 = flow_color_article2
        self.flow_size2 = flow_size2
        self.flow_podklad = flow_podklad
        self.flow_color_filter = flow_color_filter
        self.flow_ordered = flow_ordered
        self.flow_canceled = flow_canceled
        self.flow_blanc = flow_blanc
        self.flow_finish = flow_finish
        self.flow_ready = flow_ready
        self.flow_order_comment = flow_order_comment
        self.flow_barcode = flow_barcode
        self.flow_control_digit = flow_control_digit
        self.flow_ean_13 = flow_ean_13

    def check_items_list(self):
        if self.items_list.get(self.m_item_id):

            return True
        else:
            return False

    def convert_to_date_string(self, timestamp):
        try:
            return datetime.fromtimestamp(timestamp).strftime("%d.%m.%Y")
        except TypeError:
            return timestamp

    def create_dict_obj(self):
        if self.table_name == "items_history":
            return {
                "m_item_order": self.m_item_order,
                "m_item_date_created": self.convert_to_date_string(self.m_item_date_created),
                "m_item_comment": self.m_item_comment,
                "step_id": self.step_id,
                "date_start": self.convert_to_date_string(self.date_start),
                "comment": self.comment,
                "m_item_title": self.m_item_title,
                "m_item_size": self.m_item_size,
                "m_item_color": self.m_item_color,
                "m_item_podklad": self.m_item_podklad,
                "m_item_stopped": self.m_item_stopped,
                "m_item_model": self.m_item_model,
            }
        elif self.table_name == "deleted_items_history":
            return {
                "m_item_order": self.m_item_order,
                "m_item_date_created": self.convert_to_date_string(self.m_item_date_created),
                "m_item_comment": self.m_item_comment,
                "step_id": self.step_id,
                "date_start": self.convert_to_date_string(self.date_start),
                "comment": self.comment,
                "m_item_title": self.m_item_title,
                "m_item_size": self.m_item_size,
                "m_item_color": self.m_item_color,
                "m_item_podklad": self.m_item_podklad,
                "m_item_stopped": self.m_item_stopped,
                "m_item_model": self.m_item_model,
                "m_item_delete_date": self.m_item_delete_date,
            }
        elif self.table_name == "stw_modules_colors_local":
            return {
                "m_item_title": self.m_item_title,
                "m_item_lang": self.m_item_lang,
                "m_item_visible": self.m_item_visible,
            }

        elif self.table_name == "flow_sheet":
            return {
                'flow_shoe_block': self.flow_shoe_block,
                'flow_model_article': self.flow_model_article,
                'flow_color_article': self.flow_color_article,
                'flow_size': self.flow_size,
                'flow_product_article': self.flow_product_article,
                'flow_range_position': self.flow_range_position,
                'flow_sample_numbler': self.flow_sample_numbler,
                'flow_model_article2': self.flow_model_article2,
                'flow_color_article2': self.flow_color_article2,
                'flow_size2': self.flow_size2,
                'flow_podklad': self.flow_podklad,
                'flow_color_filter': self.flow_color_filter,
                'flow_ordered': self.flow_ordered,
                'flow_canceled': self.flow_canceled,
                'flow_blanc': self.flow_blanc,
                'flow_finish': self.flow_finish,
                'flow_ready': self.flow_ready,
                'flow_order_comment': self.flow_order_comment,
                'flow_barcode': self.flow_barcode,
                'flow_control_digit': self.flow_control_digit,
                'flow_ean_13': self.flow_ean_13
            }

        elif self.table_name == "flow_sheet_full":
            return {
                'flow_shoe_block': self.flow_shoe_block,
            }

        else:
            return {
                "m_item_title": self.m_item_title,
                "m_item_lang": self.m_item_lang,
            }

    def insert_item(self):
        self.items_list[self.m_item_id] = self.create_dict_obj()

    def update_item(self):
        self.items_list[self.m_item_id].update(self.create_dict_obj())

    def compare_step_id(self):
        log.debug(self.items_list[self.m_item_id])
        log.debug(self.items_list[self.m_item_id].get("step_id"))
        log.debug(self.step_id)
        if self.items_list[self.m_item_id].get("step_id") < self.step_id:
            self.update_item()

    def item_processing(self):
        if self.check_items_list():
            self.compare_step_id()
        else:
            self.insert_item()
        return self.items_list
