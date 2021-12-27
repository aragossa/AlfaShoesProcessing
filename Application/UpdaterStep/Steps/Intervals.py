class Intervals:
    renew_flow_intervals = {
        '02_Приемка на склад': {
            'СДАНО данные для Flow': {
                'cells': None,
                'intervals': [
                    {
                        'read':
                            {
                                'start_row': 3,
                                'stop_row': 50000,
                                'start_col': 1,
                                'stop_col': 19
                            },
                        'write':
                            {
                                'start_row': 3,
                                'start_col': 1
                            }
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
                        'read': {
                            'col': 51,
                            'row': 3,
                        },
                        'write': {
                            'col': 51,
                            'row': 2,
                        }
                    },
                    {
                        'read': {
                            'col': 52,
                            'row': 3,
                        },
                        'write': {
                            'col': 52,
                            'row': 2,
                        }
                    }
                ],
                'intervals': [
                    {
                        'read':
                            {
                                'start_row': 5,
                                'stop_row': 50000,
                                'start_col': 1,
                                'stop_col': 49
                            },
                        'write':
                            {
                                'start_row': 4,
                                'start_col': 1
                            }
                    }
                ]
            }
        },
        '02_Приемка на склад': {
            'Приемка рюкзаков': {
                'cells': None,
                'intervals': [

                    {
                        'read':
                            {
                                'start_row': 3,
                                'stop_row': 1000,
                                'start_col': 1,
                                'stop_col': 1
                            },
                        'write':
                            {
                                'start_row': 3,
                                'start_col': 1
                            }
                    },

                    {
                        'read':
                            {
                                'start_row': 3,
                                'stop_row': 1000,
                                'start_col': 3,
                                'stop_col': 4
                            },
                        'write':
                            {
                                'start_row': 3,
                                'start_col': 3
                            }
                    },

                    {
                        'read':
                            {
                                'start_row': 3,
                                'stop_row': 1000,
                                'start_col': 7,
                                'stop_col': 10
                            },
                        'write':
                            {
                                'start_row': 3,
                                'start_col': 7
                            }
                    },
                    {
                        'read':
                            {
                                'start_row': 3,
                                'stop_row': 1000,
                                'start_col': 15,
                                'stop_col': 15
                            },
                        'write':
                            {
                                'start_row': 3,
                                'start_col': 15
                            }
                    }
                ]
            }
        },
        '00_Справочник Номенклатуры WB': {
            'Номенклатуры ВБ': {
                'cells': [
                    {
                        'read': {
                            'col': 16,
                            'row': 1,
                        },
                        'write': {
                            'col': 16,
                            'row': 2,
                        }
                    }

                ],
                'intervals': [
                    {
                        'read':
                            {
                                'start_row': 2,
                                'stop_row': 5000,
                                'start_col': 1,
                                'stop_col': 15
                            },
                        'write':
                            {
                                'start_row': 3,
                                'start_col': 2
                            }
                    }

                ]
            }
        },
        '00_Справочник Альфа': {
            'Справочник кож': {
                'cells': [
                    {
                        'read': {
                            'col': 10,
                            'row': 1,
                        },
                        'write': {
                            'col': 10,
                            'row': 2,
                        }
                    }
                ],
                'intervals': [
                    {
                        'read':
                            {
                                'start_row': 2,
                                'stop_row': 500,
                                'start_col': 1,
                                'stop_col': 15
                            },
                        'write':
                            {
                                'start_row': 3,
                                'start_col': 1
                            }
                    }

                ]
            },
            'Справочник модели обуви': {
                'cells': [
                    {
                        'read': {
                            'col': 14,
                            'row': 1,
                        },
                        'write': {
                            'col': 14,
                            'row': 2,
                        }
                    }
                ],
                'intervals': [
                    {
                        'read':
                            {
                                'start_row': 2,
                                'stop_row': 508,
                                'start_col': 1,
                                'stop_col': 12
                            },
                        'write':
                            {
                                'start_row': 3,
                                'start_col': 1
                            }
                    }

                ]
            },
            'Справочник аксессуары': {
                'cells': [
                    {
                        'read': {
                            'col': 14,
                            'row': 1,
                        },
                        'write': {
                            'col': 14,
                            'row': 2,
                        }
                    }
                ],
                'intervals': [
                    {
                        'read':
                            {
                                'start_row': 2,
                                'stop_row': 508,
                                'start_col': 1,
                                'stop_col': 12
                            },
                        'write':
                            {
                                'start_row': 3,
                                'start_col': 1
                            }
                    }

                ]
            },
            'Цены ВБ': {
                'cells': [
                    {
                        'read': {
                            'col': 20,
                            'row': 1,
                        },
                        'write': {
                            'col': 21,
                            'row': 2,
                        }
                    }
                ],
                'intervals': [
                    {
                        'read':
                            {
                                'start_row': 2,
                                'stop_row': 999,
                                'start_col': 1,
                                'stop_col': 18
                            },
                        'write':
                            {
                                'start_row': 3,
                                'start_col': 2
                            }
                    }
                ]
            }
        }
    }
