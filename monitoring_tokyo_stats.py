#coding: utf-8

import io, re, datetime
import requests
from openpyxl import load_workbook
import pandas as pd
import pandas_bokeh
pandas_bokeh.output_notebook()


class MonitoringTokyoStats:
    url = 'https://www.fukushihoken.metro.tokyo.lg.jp/iryo/kansen/kensa/kensuu.files/syousaisenryakukensa.xlsx'
    first_row = 5
    column_interval = 'A'
    column_n_test = 'C'
    column_n_positive = 'D'

    def show(self):
        org = pd.get_option('plotting.backend')
        pd.set_option('plotting.backend', 'pandas_bokeh')
        df = self.read()
        df[['陽性率']].plot(
            figsize=(800, 300),
            legend='top_left',
        )
        df[['検査実施件数', '陽性件数']].plot(
            figsize=(800, 300),
            legend='top_left',
            logy=True,
        )
        pd.set_option('plotting.backend', org)

    def read(self):
        r = requests.get(self.url)
        assert r.status_code == 200
        wb = load_workbook(io.BytesIO(r.content))
        sheet = wb['週報詳細']

        row = self.first_row
        interval = sheet[f'{self.column_interval}{row}'].value
        data = []
        while interval != '累計' and interval is not None:
            m = re.match(r'(\d+)年\d+月第\d週[\n][(（](\d+)/(\d+)～(\d+)/(\d+)[)）]', interval)
            y_start = int(m.group(1))
            m_start = int(m.group(2))
            d_start = int(m.group(3))
            m_end = int(m.group(4))
            d_end = int(m.group(5))
            y_end = y_start + 1 if (m_start == 12 and m_end == 1) else y_start
            date_start = datetime.date(y_start, m_start, d_start)
            date_end = datetime.date(y_end, m_end, d_end)
            n_test = sheet[f'{self.column_n_test}{row}'].value
            n_positive = sheet[f'{self.column_n_positive}{row}'].value
            data.append([date_end, n_test, n_positive, n_positive / n_test])
            row += 1
            interval = sheet[f'{self.column_interval}{row}'].value
        df = pd.DataFrame(data,
            columns=['week', '検査実施件数', '陽性件数', '陽性率']).set_index('week')
        return df

