import datetime
from typing import Final

import pandas as pd
import requests
from pandas import DataFrame

URL: Final = "https://www.mos.ru/api/stats/v1/frontend/json/evp"
META_KEY: Final = '_meta'
DATA_KEY: Final = 'items'
PAGE_COUNT_KEY: Final = 'pageCount'
SHEET_NAME: Final = 'Visit Report'

OK: Final = 'OK'
ERROR: Final = 'ERROR'


def parse():
    date_today = datetime.datetime.today().strftime('%d%m%Y')

    try:
        parsing_process(date_today)
    except Exception:
        log_process(ERROR, date_today)


def parsing_process(date_today: str):
    page = 1
    process_is_live = True

    payload = {'sort': '-visits_CURRENT_MONTH', 'per-page': 50}
    data = pd.read_json("{}")

    while process_is_live:
        payload.update({'page': page})
        response = requests.get(URL, params=payload)

        if response and response.ok:
            json = response.json()
            page_count = json[META_KEY][PAGE_COUNT_KEY]
            data = pd.concat([data, pd.DataFrame(json[DATA_KEY])], ignore_index=True)

            if page < page_count:
                page += 1
            else:
                process_is_live = False
        else:
            log_process('ERROR', date_today)
            process_is_live = False

    if data.shape[0] > 0:
        write_to_exel(data, date_today)
        log_process('OK', date_today)


def write_to_exel(data: DataFrame, date_today: str):
    file_name = f'reports/visit_report_{date_today}.xlsx'

    data.to_excel(file_name, sheet_name=SHEET_NAME)


def log_process(result, date_today):
    file = open('logs/logs.txt', 'a')
    file.write(f'{date_today}: {result}\n')
    file.close()
