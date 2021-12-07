from Processing.FlowReports.row_object.RowOject import RowObject


class RowObjectParser:

    @staticmethod
    def parse(table_name, row, items_dict):
        if table_name == "items_history":
            return RowObject(
                m_item_id='' if row[0] is None else row[0],
                table_name=table_name,
                m_item_order='' if row[1] is None else row[1],
                m_item_date_created='' if row[2] is None else row[2],
                m_item_comment='' if row[3] is None else row[3],
                step_id='' if row[4] is None else row[4],
                date_start='' if row[5] is None else row[5],
                comment='' if row[6] is None else row[6],
                m_item_title='' if row[7] is None else row[7],
                m_item_size='' if row[8] is None else row[8],
                m_item_color='' if row[9] is None else row[9],
                m_item_podklad='' if row[10] is None else row[10],
                m_item_stopped='' if row[11] is None else row[11],
                items_list=items_dict,
                m_item_model='' if row[12] is None else row[12],
            )

        elif table_name == "deleted_items_history":
            return RowObject(
                m_item_id='' if row[0] is None else row[0],
                table_name=table_name,
                m_item_order='' if row[1] is None else row[1],
                m_item_date_created='' if row[2] is None else row[2],
                m_item_comment='' if row[3] is None else row[3],
                step_id='' if row[4] is None else row[4],
                date_start='' if row[5] is None else row[5],
                comment='' if row[6] is None else row[6],
                m_item_title='' if row[7] is None else row[7],
                m_item_size='' if row[8] is None else row[8],
                m_item_color='' if row[9] is None else row[9],
                m_item_podklad='' if row[10] is None else row[10],
                m_item_stopped='' if row[11] is None else row[11],
                items_list=items_dict,
                m_item_model='' if row[12] is None else row[12],
                m_item_delete_date='' if row[13] is None else row[13],
            )

        elif table_name == "stw_modules_colors_local":
            return RowObject(
                m_item_id='' if row[0] is None else row[0],
                table_name=table_name,
                m_item_lang='' if row[1] is None else row[1],
                m_item_title='' if row[2] is None else row[2],
                m_item_visible='' if row[3] is None else row[3],
                items_list=items_dict,
            )

        elif table_name == "flow_sheet":
            return RowObject(
                m_item_id='' if row[0] is None else row[0],
                table_name=table_name,
                flow_shoe_block='' if row[14] is None else row[14],
                flow_model_article='' if row[16] is None else row[16],
                flow_color_article='' if row[19] is None else row[19],
                flow_size='' if row[23] is None else row[23],
                flow_range_position='' if row[25] is None else row[25],
                flow_product_article='' if row[26] is None else row[26],
                flow_sample_numbler='' if row[27] is None else row[27],
                flow_model_article2='' if row[28] is None else row[28],
                flow_color_article2='' if row[29] is None else row[29],
                flow_size2='' if row[30] is None else row[30],
                flow_podklad='' if row[31] is None else row[31],
                flow_color_filter='' if row[32] is None else row[32],
                flow_ordered='' if row[33] is None else row[33],
                flow_canceled='' if row[34] is None else row[34],
                flow_blanc='' if row[35] is None else row[35],
                flow_finish='' if row[36] is None else row[36],
                flow_ready='' if row[37] is None else row[37],
                flow_order_comment='' if row[45] is None else row[45],
                flow_barcode='' if row[46] is None else row[46],
                flow_control_digit='' if row[47] is None else row[47],
                flow_ean_13='' if row[48] is None else row[48],
                items_list=items_dict
            )

        elif table_name == "flow_sheet_full":
            return RowObject(
                m_item_id='' if row[0] is None else row[0],
                table_name=table_name,
                flow_sheet_full='' if row[14] is None else row[14],
                flow_model_article='' if row[16] is None else row[16],
                flow_color_article='' if row[19] is None else row[19],
                flow_size='' if row[23] is None else row[23],
                flow_range_position='' if row[25] is None else row[25],
                flow_product_article='' if row[26] is None else row[26],
                flow_sample_numbler='' if row[27] is None else row[27],
                flow_model_article2='' if row[28] is None else row[28],
                flow_color_article2='' if row[29] is None else row[29],
                flow_size2='' if row[30] is None else row[30],
                flow_podklad='' if row[31] is None else row[31],
                flow_color_filter='' if row[32] is None else row[32],
                flow_ordered='' if row[33] is None else row[33],
                flow_canceled='' if row[34] is None else row[34],
                flow_blanc='' if row[35] is None else row[35],
                flow_finish='' if row[36] is None else row[36],
                flow_ready='' if row[37] is None else row[37],
                flow_order_comment='' if row[45] is None else row[45],
                flow_barcode='' if row[46] is None else row[46],
                flow_control_digit='' if row[47] is None else row[47],
                flow_ean_13='' if row[48] is None else row[48],
                items_list=items_dict
            )

        else:
            return RowObject(
                m_item_id='' if row[0] is None else row[0],
                table_name=table_name,
                m_item_lang='' if row[1] is None else row[1],
                m_item_title='' if row[2] is None else row[2],
                items_list=items_dict,
            )