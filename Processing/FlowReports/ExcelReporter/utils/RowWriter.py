from Processing.FlowReports.ExcelReporter.utils.formulas import get_excel_formula
from Utils.Logger.main_logger import get_logger

log = get_logger("RowWriter")


class RowWriter:

    @staticmethod
    def __detect_changes(values):
        log.debug(values)
        if values.get("m_item_status") == 'новое изделие':
            changes = True
        elif values.get("m_item_step_status") != '':
            changes = True
        elif values.get("m_item_comment_status") != '':
            changes = True
        elif values.get("m_item_size_status") != '':
            changes = True
        elif values.get("m_item_color_status") != '':
            changes = True
        elif values.get("m_item_podklad_status") != '':
            changes = True
        elif values.get("m_item_model_status") != '':
            changes = True
        else:
            changes = False
        log.debug(f'detected changes {changes}')
        log.debug('---------------------------')
        return changes

    @staticmethod
    def __prepare_write_value(values, changes, col_num, values_flow, col_name, row_num, m_item_delete_date,
                              m_item_id=None, deleted=False):
        log.debug(values)
        if col_num <= 14:
            if col_name == 'step_id' and deleted:
                return 'удален'
            elif col_name == 'date_start' and deleted:
                return m_item_delete_date
            else:
                return values.get(col_name)
        elif values.get('m_item_status') == 'новое изделие':
            log.debug('writing formula for new m_item')
            log.debug(get_excel_formula(col_num, row_num))
            return get_excel_formula(col_num, row_num)
        elif changes or col_num >= 30 or col_num <= 43:
            log.debug('writing formula')
            log.debug(get_excel_formula(col_num, row_num))
            return get_excel_formula(col_num, row_num)
        else:
            log.debug('writing data')
            log.debug(col_name)
            log.debug(values_flow.get(col_name))
            log.debug(values_flow.get(col_name))
            return values_flow.get(col_name)

    @staticmethod
    def write_row(ws, m_item_id, values, report_name, cur_write_row, values_flow=None, deleted=False,
                  m_item_delete_date=None):
        if report_name == "items_history":
            ws.cell(cur_write_row, 1).value = m_item_id
            ws.cell(cur_write_row, 2).value = values.get("m_item_order")
            ws.cell(cur_write_row, 3).value = values.get("m_item_date_created")
            ws.cell(cur_write_row, 4).value = values.get("m_item_comment")
            ws.cell(cur_write_row, 5).value = values.get("step_id")
            ws.cell(cur_write_row, 6).value = values.get("date_start")
            ws.cell(cur_write_row, 7).value = values.get("comment")
            ws.cell(cur_write_row, 8).value = values.get("m_item_title")
            ws.cell(cur_write_row, 9).value = values.get("m_item_size")
            ws.cell(cur_write_row, 10).value = values.get("m_item_color")
            ws.cell(cur_write_row, 11).value = values.get("m_item_podklad")
            ws.cell(cur_write_row, 12).value = values.get("m_item_stopped")
            ws.cell(cur_write_row, 13).value = values.get("m_item_model")

            ws.cell(cur_write_row, 19).value = values.get("m_item_status")
            ws.cell(cur_write_row, 20).value = values.get("m_item_step_status")
            ws.cell(cur_write_row, 21).value = values.get("m_item_comment_status")
            ws.cell(cur_write_row, 22).value = values.get("m_item_size_status")
            ws.cell(cur_write_row, 23).value = values.get("m_item_color_status")
            ws.cell(cur_write_row, 24).value = values.get("m_item_podklad_status")
            ws.cell(cur_write_row, 25).value = values.get("m_item_model_status")

        elif report_name == 'deleted_items_history':
            ws.cell(cur_write_row, 1).value = m_item_id
            ws.cell(cur_write_row, 2).value = values.get("m_item_order")
            ws.cell(cur_write_row, 3).value = values.get("m_item_date_created")
            ws.cell(cur_write_row, 4).value = values.get("m_item_comment")
            ws.cell(cur_write_row, 5).value = values.get("step_id")
            ws.cell(cur_write_row, 6).value = values.get("date_start")
            ws.cell(cur_write_row, 7).value = values.get("comment")
            ws.cell(cur_write_row, 8).value = values.get("m_item_title")
            ws.cell(cur_write_row, 9).value = values.get("m_item_size")
            ws.cell(cur_write_row, 10).value = values.get("m_item_color")
            ws.cell(cur_write_row, 11).value = values.get("m_item_podklad")
            ws.cell(cur_write_row, 12).value = values.get("m_item_stopped")
            ws.cell(cur_write_row, 13).value = values.get("m_item_model")
            ws.cell(cur_write_row, 14).value = values.get("m_item_delete_date")

        elif report_name == 'Flow':
            log.debug(deleted)
            if deleted:
                changes = True
            else:
                changes = RowWriter.__detect_changes(values)
            log.debug(f'm_item_id {m_item_id}')
            ws.cell(cur_write_row, 1).value = m_item_id
            ws.cell(cur_write_row, 2).value = RowWriter.__prepare_write_value(values=values, values_flow=values_flow,
                                                                              col_num=2,
                                                                              col_name='m_item_order',
                                                                              changes=changes,
                                                                              row_num=cur_write_row,
                                                                              m_item_delete_date=m_item_delete_date)
            ws.cell(cur_write_row, 3).value = RowWriter.__prepare_write_value(values=values, values_flow=values_flow,
                                                                              col_num=3,
                                                                              col_name='m_item_date_created',
                                                                              changes=changes,
                                                                              row_num=cur_write_row,
                                                                              m_item_delete_date=m_item_delete_date)
            ws.cell(cur_write_row, 4).value = RowWriter.__prepare_write_value(values=values, values_flow=values_flow,
                                                                              col_num=4,
                                                                              col_name='m_item_title',
                                                                              changes=changes,
                                                                              row_num=cur_write_row,
                                                                              m_item_delete_date=m_item_delete_date)
            ws.cell(cur_write_row, 5).value = RowWriter.__prepare_write_value(values=values, values_flow=values_flow,
                                                                              col_num=5,
                                                                              col_name='m_item_comment',
                                                                              changes=changes,
                                                                              row_num=cur_write_row,
                                                                              m_item_delete_date=m_item_delete_date)
            ws.cell(cur_write_row, 6).value = RowWriter.__prepare_write_value(values=values, values_flow=values_flow,
                                                                              col_num=6,
                                                                              col_name='step_id',
                                                                              changes=changes,
                                                                              row_num=cur_write_row,
                                                                              deleted=deleted,
                                                                              m_item_delete_date=m_item_delete_date)

            ws.cell(cur_write_row, 7).value = RowWriter.__prepare_write_value(values=values, values_flow=values_flow,
                                                                              col_num=7,
                                                                              col_name='date_start',
                                                                              changes=changes,
                                                                              row_num=cur_write_row,
                                                                              m_item_id=m_item_id,
                                                                              deleted=deleted,
                                                                              m_item_delete_date=m_item_delete_date)
            ws.cell(cur_write_row, 8).value = RowWriter.__prepare_write_value(values=values, values_flow=values_flow,
                                                                              col_num=8,
                                                                              col_name='comment',
                                                                              changes=changes,
                                                                              row_num=cur_write_row,
                                                                              m_item_delete_date=m_item_delete_date)
            ws.cell(cur_write_row, 9).value = RowWriter.__prepare_write_value(values=values, values_flow=values_flow,
                                                                              col_num=9,
                                                                              col_name='m_item_size',
                                                                              changes=changes,
                                                                              row_num=cur_write_row,
                                                                              m_item_delete_date=m_item_delete_date)
            ws.cell(cur_write_row, 10).value = RowWriter.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=10,
                                                                               col_name='m_item_color',
                                                                               changes=changes,
                                                                               row_num=cur_write_row,
                                                                               m_item_delete_date=m_item_delete_date)
            ws.cell(cur_write_row, 11).value = RowWriter.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=11,
                                                                               col_name='m_item_podklad',
                                                                               changes=changes,
                                                                               row_num=cur_write_row,
                                                                               m_item_delete_date=m_item_delete_date)
            ws.cell(cur_write_row, 12).value = RowWriter.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=12,
                                                                               col_name='m_item_stopped',
                                                                               changes=changes,
                                                                               row_num=cur_write_row,
                                                                               m_item_delete_date=m_item_delete_date)
            ws.cell(cur_write_row, 13).value = RowWriter.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=13,
                                                                               col_name='m_item_model',
                                                                               changes=changes,
                                                                               row_num=cur_write_row,
                                                                               m_item_delete_date=m_item_delete_date)

            ws.cell(cur_write_row, 15).value = RowWriter.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=15,
                                                                               col_name='flow_shoe_block',
                                                                               changes=changes,
                                                                               row_num=cur_write_row,
                                                                               m_item_delete_date=m_item_delete_date)
            ws.cell(cur_write_row, 17).value = RowWriter.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=17,
                                                                               col_name='flow_model_article',
                                                                               changes=changes,
                                                                               row_num=cur_write_row,
                                                                               m_item_delete_date=m_item_delete_date)
            ws.cell(cur_write_row, 20).value = RowWriter.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=20,
                                                                               col_name='flow_color_article',
                                                                               changes=changes,
                                                                               row_num=cur_write_row,
                                                                               m_item_delete_date=m_item_delete_date)
            ws.cell(cur_write_row, 24).value = RowWriter.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=24,
                                                                               col_name='flow_size',
                                                                               changes=changes,
                                                                               row_num=cur_write_row,
                                                                               m_item_delete_date=m_item_delete_date)
            ws.cell(cur_write_row, 26).value = RowWriter.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=26,
                                                                               col_name='flow_range_position',
                                                                               changes=changes,
                                                                               row_num=cur_write_row,
                                                                               m_item_delete_date=m_item_delete_date)
            ws.cell(cur_write_row, 27).value = RowWriter.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=27,
                                                                               col_name='flow_sample_numbler',
                                                                               changes=changes,
                                                                               row_num=cur_write_row,
                                                                               m_item_delete_date=m_item_delete_date)
            ws.cell(cur_write_row, 28).value = RowWriter.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=28,
                                                                               col_name='flow_model_article2',
                                                                               changes=changes,
                                                                               row_num=cur_write_row,
                                                                               m_item_delete_date=m_item_delete_date)
            ws.cell(cur_write_row, 29).value = RowWriter.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=29,
                                                                               col_name='flow_model_article2',
                                                                               changes=changes,
                                                                               row_num=cur_write_row,
                                                                               m_item_delete_date=m_item_delete_date)
            ws.cell(cur_write_row, 30).value = RowWriter.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=30,
                                                                               col_name='flow_color_article2',
                                                                               changes=changes,
                                                                               row_num=cur_write_row,
                                                                               m_item_delete_date=m_item_delete_date)
            ws.cell(cur_write_row, 31).value = RowWriter.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=31,
                                                                               col_name='flow_size2',
                                                                               changes=changes,
                                                                               row_num=cur_write_row,
                                                                               m_item_delete_date=m_item_delete_date)
            ws.cell(cur_write_row, 32).value = RowWriter.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=32,
                                                                               col_name='flow_podklad',
                                                                               changes=changes,
                                                                               row_num=cur_write_row,
                                                                               m_item_delete_date=m_item_delete_date)
            ws.cell(cur_write_row, 33).value = RowWriter.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=33,
                                                                               col_name='flow_color_filter',
                                                                               changes=changes,
                                                                               row_num=cur_write_row,
                                                                               m_item_delete_date=m_item_delete_date)
            ws.cell(cur_write_row, 34).value = 1
            ws.cell(cur_write_row, 35).value = RowWriter.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=35,
                                                                               col_name='flow_canceled',
                                                                               changes=changes,
                                                                               row_num=cur_write_row,
                                                                               m_item_delete_date=m_item_delete_date)

            ws.cell(cur_write_row, 36).value = RowWriter.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=36,
                                                                               col_name='flow_blanc',
                                                                               changes=changes,
                                                                               row_num=cur_write_row,
                                                                               m_item_delete_date=m_item_delete_date)
            ws.cell(cur_write_row, 37).value = RowWriter.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=37,
                                                                               col_name='flow_finish',
                                                                               changes=changes,
                                                                               row_num=cur_write_row,
                                                                               m_item_delete_date=m_item_delete_date)
            ws.cell(cur_write_row, 38).value = RowWriter.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=38,
                                                                               col_name='flow_ready',
                                                                               changes=changes,
                                                                               row_num=cur_write_row,
                                                                               m_item_delete_date=m_item_delete_date)

            ws.cell(cur_write_row, 39).value = RowWriter.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=39,
                                                                               col_name='flow_ready',
                                                                               changes=changes,
                                                                               row_num=cur_write_row,
                                                                               m_item_delete_date=m_item_delete_date)
            ws.cell(cur_write_row, 40).value = RowWriter.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=40,
                                                                               col_name='flow_ready',
                                                                               changes=changes,
                                                                               row_num=cur_write_row,
                                                                               m_item_delete_date=m_item_delete_date)
            ws.cell(cur_write_row, 41).value = RowWriter.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=41,
                                                                               col_name='flow_ready',
                                                                               changes=changes,
                                                                               row_num=cur_write_row,
                                                                               m_item_delete_date=m_item_delete_date)
            ws.cell(cur_write_row, 42).value = RowWriter.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=42,
                                                                               col_name='flow_ready',
                                                                               changes=changes,
                                                                               row_num=cur_write_row,
                                                                               m_item_delete_date=m_item_delete_date)
            ws.cell(cur_write_row, 43).value = RowWriter.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=43,
                                                                               col_name='flow_ready',
                                                                               changes=changes,
                                                                               row_num=cur_write_row,
                                                                               m_item_delete_date=m_item_delete_date)

            ws.cell(cur_write_row, 46).value = RowWriter.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=46,
                                                                               col_name='flow_order_comment',
                                                                               changes=changes,
                                                                               row_num=cur_write_row,
                                                                               m_item_delete_date=m_item_delete_date)
            ws.cell(cur_write_row, 47).value = RowWriter.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=47,
                                                                               col_name='flow_barcode',
                                                                               changes=changes,
                                                                               row_num=cur_write_row,
                                                                               m_item_delete_date=m_item_delete_date)
            ws.cell(cur_write_row, 48).value = RowWriter.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=48,
                                                                               col_name='flow_control_digit',
                                                                               changes=changes,
                                                                               row_num=cur_write_row,
                                                                               m_item_delete_date=m_item_delete_date)
            ws.cell(cur_write_row, 49).value = RowWriter.__prepare_write_value(values=values, values_flow=values_flow,
                                                                               col_num=49,
                                                                               col_name='flow_ean_13',
                                                                               changes=changes,
                                                                               row_num=cur_write_row,
                                                                               m_item_delete_date=m_item_delete_date)

        elif report_name == "stw_modules_colors_local":
            ws.cell(cur_write_row, 1).value = m_item_id
            ws.cell(cur_write_row, 2).value = values.get("m_item_lang")
            ws.cell(cur_write_row, 3).value = values.get("m_item_title")
            ws.cell(cur_write_row, 4).value = values.get("m_item_visible")
        else:
            ws.cell(cur_write_row, 1).value = m_item_id
            ws.cell(cur_write_row, 2).value = values.get("m_item_lang")
            ws.cell(cur_write_row, 3).value = values.get("m_item_title")
        cur_write_row += 1
        return ws
