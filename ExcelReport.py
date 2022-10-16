from readline import read_history_file
import pandas as pd
import numpy as np
import io
import matplotlib.pyplot as plt
from matplotlib.gridspec import GridSpec
import seaborn as sns
from datetime import date
import os

import openpyxl
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Font, Border, Side, PatternFill, Alignment
from openpyxl.formatting.rule import CellIsRule
from openpyxl.drawing.image import Image
from openpyxl.utils.dataframe import dataframe_to_rows


class ExcelReport():

    def __init__(self):
        self.cur_sheet = None
        self.cur_content = None
        self.ws = None


    def create(self, sheets_data:dict, path:str):
        if not isinstance(sheets_data, dict):
            raise ValueError('"data" parameter should be dictionary')

        folder = os.path.split(path)[0]
        if not os.path.exists(folder):
            os.makedirs(folder)

        wb = openpyxl.Workbook()
        for sheet, contents in sheets_data.items():
            if not isinstance(sheet, str):
                raise ValueError('"data" parameter should contains string as key')
            else:
                self.cur_sheet = sheet
            if not hasattr(contents, '__iter__'):
                raise ValueError('"data" parameter should contains iteriable type as value')


            self.ws = wb.create_sheet(sheet)
            for content in contents:
                self.cur_content = content

                data = self.cur_content.get('data')
                asTable = self.cur_content.get('asTable')

                if isinstance(data, pd.DataFrame):
                    if data.shape[0] == 0:
                        pass

                    self.write_pandas(data)
                    if asTable:
                        self.convertto_Table(data)

        wb.save(path)
    
    def write_pandas(self, data):
        startrow = self.cur_content.get('startrow')
        startcol = self.cur_content.get('startcol')
        rows = dataframe_to_rows(data, index=False, header=True)

        for rowidx, row in enumerate(rows):
            for colidx, value in enumerate(row):
                wrow, wcol = (rowidx + startrow + 1), (colidx + startcol + 1)
                self.ws.cell(row=wrow, column=wcol).value = value
    
    def convertto_Table(self, df):
        nrows, ncols = df.shape
        cell_range = f'A1:{self.get_ColumnLetter(ncols)}{nrows+1}'
        tab = Table(displayName="prediction", ref=cell_range)
        style = TableStyleInfo(name="TableStyleLight1", showFirstColumn=False,
                            showLastColumn=False, showRowStripes=True, showColumnStripes=False)
        tab.tableStyleInfo = style
        self.ws.add_table(tab)

        # change columns width
        for num in range(1, ncols+1):
            ind = self.get_ColumnLetter(num)
            self.ws.column_dimensions[ind].width = 18
        return
    
    def get_ColumnLetter(self, col_num):
        letter1, letter2 = col_num // 26, col_num % 26
        
        if letter2 == 0:
            letter1, letter2 = -1, 26
        
        if letter1 > 0:
            ColumnLetter = f"{chr(ord('@') + letter1)}{chr(ord('@') + letter2)}"
        else:
            ColumnLetter = chr(ord('@') + letter2)
        return ColumnLetter
