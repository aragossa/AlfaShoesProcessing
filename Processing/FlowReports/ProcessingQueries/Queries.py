class ProcessingQueries:
    items_history_query = """SELECT smi.m_item_id, smi.m_item_order, smi.m_item_date_created, smol.m_item_comment,
                                    ith.step_id, ith.date_start, ith.comment, smil.m_item_title, smi.m_item_size,
                                    smi.m_item_color, smi.m_item_podklad, smi.m_item_stopped, mo.m_item_model 
                             FROM items_history ith
                             INNER JOIN stw_modules_items smi ON ith.item_id = smi.m_item_id
                             INNER JOIN stw_modules_orders mo ON smi.m_item_order = mo.m_item_id
                             INNER JOIN stw_modules_orders_local smol ON mo.m_item_id = smol.m_item_id 
                             INNER JOIN stw_modules_items_local smil ON ith.item_id = smil.m_item_id
                             ORDER BY (smi.m_item_id)"""

    stw_modules_models_local_query = """SELECT m_item_id, m_item_lang, m_item_title FROM stw_modules_models_local"""

    stw_modules_colors_local_query = """SELECT smcl.m_item_id, smcl.m_item_lang, smcl.m_item_title, smc.m_item_visible
                                        FROM stw_modules_colors_local smcl 
                                        INNER JOIN stw_modules_colors smc ON smcl.m_item_id = smc.m_item_id"""

    stw_modules_podklads_local_query = """SELECT m_item_id, m_item_lang, m_item_title FROM stw_modules_podklads_local"""

    stw_modules_sizes_local_query = """SELECT m_item_id, m_item_lang, m_item_title FROM stw_modules_sizes_local"""
