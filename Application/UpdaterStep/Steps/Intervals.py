class Intervals:
    renew_flow_intervals = {
        '02_Приемка на склад': {
            'СДАНО данные для Flow': {
                'cells': None,
                'intervals': [
                    {
                        'start_row': 3,
                        'stop_row': 50000,
                        'start_col': 1,
                        'stop_col': 19
                    }
                ]
            }
        }
    }
    acceptance_intervals = {
        '01_Flow': {
            'Flow': {
                'cells': [
                    {
                        'col': 51,
                        'row': 3,
                    },
                    {
                        'col': 52,
                        'row': 3,
                    }
                ],
                'intervals': [
                    {
                        'start_row': 5,
                        'stop_row': 50000,
                        'start_col': 1,
                        'stop_col': 49
                    }
                ]
            }
        },
        '02_Приемка на склад': {
            'Приемка рюкзаков': {
                'cells': None,
                'intervals': [
                    {
                        'start_row': 3,
                        'stop_row': 1000,
                        'start_col': 1,
                        'stop_col': 1,
                    },
                    {
                        'start_row': 3,
                        'stop_row': 1000,
                        'start_col': 3,
                        'stop_col': 4
                    },
                    {
                        'start_row': 3,
                        'stop_row': 1000,
                        'start_col': 7,
                        'stop_col': 10
                    },
                    {
                        'start_row': 3,
                        'stop_row': 1000,
                        'start_col': 15,
                        'stop_col': 15
                    }
                ]
            }
        },
        '00_Справочник Номенклатуры WB': {
            'Номенклатуры ВБ': {
                'cells': [
                    {
                        'col': 17,
                        'row': 1,
                    }
                ],
                'intervals': [
                    {
                        'start_row': 2,
                        'stop_row': 5008,
                        'start_col': 1,
                        'stop_col': 15
                    }
                ]
            }
        },
        '00_Справочник Альфа': {
            'Справочник кож': {
                'cells': [
                    {
                        'col': 10,
                        'row': 1,
                    }
                ],
                'intervals': [
                    {
                        'start_row': 2,
                        'stop_row': 508,
                        'start_col': 1,
                        'stop_col': 15
                    }
                ]
            },
            'Справочник модели обуви': {
                'cells': [
                    {
                        'col': 14,
                        'row': 1,
                    }
                ],
                'intervals': [
                    {
                        'start_row': 2,
                        'stop_row': 508,
                        'start_col': 1,
                        'stop_col': 12
                    }
                ]
            },
            'Справочник аксессуары': {
                'cells': [
                    {
                        'col': 14,
                        'row': 1,
                    }
                ],
                'intervals': [
                    {
                        'start_row': 2,
                        'stop_row': 508,
                        'start_col': 1,
                        'stop_col': 12
                    }
                ]
            },
            'Цены ВБ': {
                'cells': [
                    {
                        'col': 20,
                        'row': 1,
                    }
                ],
                'intervals': [
                    {
                        'start_row': 2,
                        'stop_row': 999,
                        'start_col': 1,
                        'stop_col': 18
                    }
                ]
            }
        }
    }
