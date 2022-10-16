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


class ExcelReport():

    def __init__(self):
        pass

    def create(self, sheets_data:dict, path:str):
        if not isinstance(sheets_data, dict):
            raise ValueError('"data" parameter should be dictionary')

        folder = os.path.split(path)[0]
        if not os.path.exists(folder):
            os.makedirs(folder)

        with pd.ExcelWriter(path, engine='openpyxl', mode='w') as writer:           
            for sheet, contents in sheets_data.items():
                if not isinstance(sheet, str):
                    raise ValueError('"data" parameter should contains string as key')
                if not hasattr(contents, '__iter__'):
                    raise ValueError('"data" parameter should contains iteriable type as value')

                for content in contents:
                    data = content.get('data')
                    startrow = content.get('startrow')
                    startcol = content.get('startcol')

                    if isinstance(data, pd.DataFrame):
                        if data.shape[0] == 0:
                            pass
                        data.to_excel(writer, sheet_name=sheet, startrow=startrow, startcol=startcol, index=False)

            writer.save()
