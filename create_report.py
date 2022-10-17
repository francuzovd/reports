from array import array
import pandas as pd
import numpy as np
from ExcelReport import ExcelReport
import seaborn as sns
import matplotlib.pyplot as plt
from matplotlib.gridspec import GridSpec
import io


def mape(fact:np.array, pred:np.array):
    """ Calculate Mean Absolute Percantage Error (MAPE) by rows

    Args:
        fact (np.array): fact data
        pred (np.array): prediction data

    Returns:
        np.array: array of mape
    """
    return np.abs(fact - pred)/ fact

def mape_report(df_plot:pd.DataFrame, x:str, y:str, hue:str, compare:list, palette:sns.palettes, font:dict):
    """_summary_

    Args:
        df_plot (pd.DataFrame): Data to visualize
        x (str): name of value column
        y (str): name of column for y-axis in 1st graph
        hue (str): name of column to split data by group
        compare (list): list of 2 names of column to conpare values of 2d graph
        palette (sns.palettes): seaborn pallete
        font (dict): fictionary of font style

    Returns:
        Figure: Figure of report
    """
    fig = plt.figure(figsize=(16,8))
    gs = GridSpec(2, 2, figure=fig, wspace=0.1, width_ratios=[2, 3], height_ratios=[1,1], left=0.05, right=0.93, top=0.9, bottom=0.1)
    sns.set_palette(palette)
    sns.set_style('whitegrid')
    ax = []

    ax_ = fig.add_subplot(gs[0, :])
    ax.append(ax_)

    sns.boxplot(data=df_plot, y=y, x=x, hue=hue, ax=ax_, palette=palette[1:])
    # ax_.set_title()
    ax_.set_ylabel(f'{y.upper()}', fontdict=font['label'])
    # ax_.legend(loc='upper right', fontsize=9)
    ax_.set_title(f'MAPE distribution by {hue}', fontdict=font['title'])

    ax_ = fig.add_subplot(gs[1, 0])
    ax.append(ax_)
    sns.barplot(data=df_plot, x=compare[0], y=hue, estimator=np.sum, ci=None, ax=ax_, color=palette[1])
    sns.barplot(data=df_plot, x=compare[1], y=hue, estimator=np.sum, ci=None, ax=ax_, color=palette[2])
    ax_.set_title('Fact vs Pred', fontdict=font['title'])

    ax_ = fig.add_subplot(gs[1, 1])
    ax.append(ax_)
    sns.violinplot(data=df_plot, x=hue, y=y, hue='hasPromo', split=True)
    ax_.set_title('Promo vs NotPromo', fontdict=font['title'])
    # ax_.legend(loc='upper right', fontsize=9)

    for ax_ in ax:
        ax_.set_xlabel('')
        ax_.set_ylabel('')
        ax_.tick_params(axis='x', colors=palette[0])
        ax_.tick_params(axis='y', colors=palette[1])
        ax_.legend(loc='upper right', fontsize=9)

    return fig

def last_Nmonth(dates:array, nmonth:int):
    """Function to calculate interval between last date and N months back from it.

    Args:
        dates (array): np.array with dates
        nmonth (int): number of month

    Returns:
        (tuple): (start date of interval, last date of interval)
    """
    last_date = dates.max()
    first_date = last_date - pd.Timedelta(30*nmonth - 10, unit='D')
    first_date - pd.Timedelta(first_date.day-1, unit='D')
    return (first_date, last_date)

xrep = ExcelReport()
outpath = './output/report.xlsx'
# import data
sales = pd.read_csv('./data/sales.csv', sep=';')
sales['date'] = sales['date'].astype('datetime64')
# calculate interval from last 3 month
last_3months = last_Nmonth(sales.date, 3)

# create first Frame to import
df_sheet1 = sales.head(15)

# create second Frame to import
df_report = sales.loc[sales['date'].between(*last_3months), ]
df_report['mape'] = mape(df_report.weight_fact, df_report.weight_predict)
df_report['accuracy'] = 1 - df_report['mape']
df_sheet2 = pd.pivot_table(data=df_report, index=['client','region'], columns=['prod_name', 'year_month'], values=['accuracy'], aggfunc=np.median)

# create graph to import
palette = sns.color_palette('Spectral')
font = {
    'title': {
        'family': 'serif',
        'fontsize': 14,
        'color': palette[1],
    },
    'label': {
        'family': 'serif',
        'fontsize': 12,
        'color': palette[1],
    },
    'suptitle': {
        'family': 'serif',
        'size': 16,
        'weight': 'bold'
    },
}

client, product = 'client_1', 'prod_name_14'
df_plot = df_report.loc[
    (sales.client == client) & (sales.prod_name == product)
] 
graph = mape_report(df_plot, x='year_month', y='mape', hue='region', compare=('weight_fact', 'weight_predict'), palette=palette, font=font)
graph.suptitle(f'Prediction result for "{product}" in "{client}"', fontproperties=font['suptitle'], color=palette[1])

# create congig dictionary to create Excel Report
xlxContent = {
    'data' :[
        {'data':df_sheet1, 'startrow': 0, 'startcol': 0, 'importIndex': False, 'asTable': True}
    ],
    'mape' :[
        {'data':df_sheet2, 'startrow': 6, 'startcol': 3, 'importIndex': True, 'formatTable': True, 'title_name': 'Accuracy of prediction, 2021'},
        {'data':graph, 'startrow': 6 + df_sheet2.shape[0] + len(df_sheet2.columns.names) + 4, 'startcol': 3},
    ]
}

# run
xrep.create(xlxContent, path=outpath)
