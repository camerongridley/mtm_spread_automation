import openpyxl
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

class ExcelAutomator(object):
    def __init__(self, filename) -> None:
        self.filename = filename
        self.wb = load_workbook(filename)

    def get_sheet(self, sheet_name):
        ws = self.wb[sheet_name]
        return ws
    
    def updated_dai(self, ws, df:pd.DataFrame):
        d = df.to_dict(orient='index')

        for cell, data in d.items():
            ws[cell]=data['value']

    def move_to_last_week(self, ws, df:pd.DataFrame):
        ws.move_range()

    def save(self, filename):
        self.wb.save(filename=filename)


if __name__ == "__main__":
    cols = ['date', 'metric', 'value', 'cell']
    data = [['1-1-21', 'Uses Impression Rate (%)', 1.0, 'D33'], 
            ['1-1-21', 'Fill rate, Ad only (%)', 2.3, 'D34'], 
            ['1-1-21', 'Fill rate, Ad+Promo (%)', 3.6, 'D35'], 
            ['1-1-21', 'Completed Impression Rate (%)', 4.9, 'D36']
        ]

    df = pd.DataFrame(columns=cols, data=data)
    df.index = df['cell']
    df.drop('cell', axis=1)
    filename = 'TeamMetrics.xlsx'

    exa = ExcelAutomator(filename)
    ws = exa.get_sheet('SlingTV')
    exa.updated_dai(ws, df)
    exa.save(filename)


'''
,
            ['1-8-21', 'Uses Impression Rate (%)', 1.1, 'D33'], 
            ['1-8-21', 'Fill rate, Ad only (%)', 2.4, 'D34'], 
            ['1-8-21', 'Fill rate, Ad+Promo (%)', 3.7, 'D35'], 
            ['1-8-21', 'Completed Impression Rate (%)', 5.0, 'D36'],
            ['1-15-21', 'Uses Impression Rate (%)', 0.9, 'D33'], 
            ['1-15-21', 'Fill rate, Ad only (%)', 2.2, 'D34'], 
            ['1-15-21', 'Fill rate, Ad+Promo (%)', 3.5, 'D35'], 
            ['1-15-21', 'Completed Impression Rate (%)', 4.8, 'D36'],
            ['1-22-21', 'Uses Impression Rate (%)', 1.0, 'D33'], 
            ['1-22-21', 'Fill rate, Ad only (%)', 2.3, 'D34'], 
            ['1-22-21', 'Fill rate, Ad+Promo (%)', 3.6, 'D35'], 
            ['1-22-21', 'Completed Impression Rate (%)', 4.9, 'D36'],
            ['1-30-21', 'Uses Impression Rate (%)', 1.3, 'D33'], 
            ['1-30-21', 'Fill rate, Ad only (%)', 2.6, 'D34'], 
            ['1-30-21', 'Fill rate, Ad+Promo (%)', 3.9, 'D35'], 
            ['1-30-21', 'Completed Impression Rate (%)', 5.2, 'D36'],
'''