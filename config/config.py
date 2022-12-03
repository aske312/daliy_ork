# -*- coding: utf-8 -*-

# VK_BOT_API
TOKEN_VK = 'TOKEN'
GROUP_VK = 'ID_GROUP'

# TELEGRAM_API
TOKEN_TG = 'TOKEN'

# schedule_param
USER = 'USER@EMAIL'
CREDENTIALS_FILE = 'JSON_FILE'
SHEETS_FILE = 'SHEETS_FILE'
CELL_LIST = ['B', 'Z', 'AX', 'BV', 'CT', 'DR', 'EP',
             'FN', 'GL', 'HJ', 'IH', 'JF', 'KD', 'LB', 'LZ',
             'MX', 'NV', 'OT', 'PR', 'QP', 'RN', 'SL', 'TJ',
             'UH', 'VF', 'WD', 'XB', 'XZ', 'YX', 'ZV', 'AAT']

# balance_night
DAY_TIME = {'end_night': 5, 'start_night': 22}

# event_by_schedule
TRAINING_RUN = {
    'time_start': 0,
    'type': 'run',
    'allotted_time': '2h',
    'stop': True,
}
MEETING_FRD = {
    'time_start': 0,
    'type': 'run',
    'allotted_time': '4h',
    'stop': None,
}
MEETING_JOB = {
    'time_start': 0,
    'type': 'run',
    'allotted_time': '2h',
    'stopper': True,
}
WORK_JOB = {
    'time_start': 0,
    'type': 'run',
    'allotted_time': '4h',
    'stop': None,
}
WORK_HOB = {}
WORK_STD = {}
