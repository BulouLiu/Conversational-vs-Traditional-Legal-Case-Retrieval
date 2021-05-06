import warnings
warnings.filterwarnings('ignore')
import numpy as np
from scipy import stats
from itertools import combinations
stats.chisqprob = lambda chisq, df: stats.chi2.sf(chisq, df)
import pandas as pd


def cal_Scheirer_Ray_Hare(data, name):
    data['name'] = data[name]
    data['rank'] = data.name.sort_values().rank(numeric_only=float)
    rows = data.groupby(['Expertise'], as_index=False).agg({'rank': ['count', 'mean', 'var']}).rename(
        columns={'rank': 'row'})

    rows.columns = ['_'.join(col) for col in rows.columns]
    rows.columns = rows.columns.str.replace(r'_$', "")
    rows['row_mean_rows'] = rows.row_mean.mean()
    rows['sqdev'] = (rows.row_mean - rows.row_mean_rows) ** 2

    cols = data.groupby(['Method'], as_index=False).agg({'rank': ['count', 'mean', 'var']}).rename(
        columns={'rank': 'col'})
    cols.columns = ['_'.join(col) for col in cols.columns]
    cols.columns = cols.columns.str.replace(r'_$', "")
    cols['col_mean_cols'] = cols.col_mean.mean()
    cols['sqdev'] = (cols.col_mean - cols.col_mean_cols) ** 2

    data_sum = data.groupby(['Expertise', 'Method'], as_index=False).agg({'rank': ['count', 'mean', 'var']})
    data_sum.columns = ['_'.join(col) for col in data_sum.columns]
    data_sum.columns = data_sum.columns.str.replace(r'_$', "")

    nobs_row = rows.row_count.mean()
    nobs_total = rows.row_count.sum()
    nobs_col = cols.col_count.mean()

    Columns_SS = cols.sqdev.sum() * nobs_col
    Rows_SS = rows.sqdev.sum() * nobs_row
    Within_SS = data_sum.rank_var.sum() * (data_sum.rank_count.min() - 1)
    MS = data['rank'].var()
    TOTAL_SS = MS * (nobs_total - 1)
    Inter_SS = TOTAL_SS - Within_SS - Rows_SS - Columns_SS

    H_rows = Rows_SS / MS
    H_cols = Columns_SS / MS
    H_int = Inter_SS / MS

    df_rows = len(rows) - 1
    df_cols = len(cols) - 1
    df_int = df_rows * df_cols
    df_total = len(data) - 1
    df_within = df_total - df_int - df_cols - df_rows

    p_rows = round(1 - stats.chi2.cdf(H_rows, df_rows), 4)
    p_cols = round(1 - stats.chi2.cdf(H_cols, df_cols), 4)
    p_inter = round(1 - stats.chi2.cdf(H_int, df_int), 4)

    return [p_rows, p_cols, p_inter]


names = ["Cases", "Queries", "Time", "Self-reported", "Satisfaction", "Success"]
for name in names:
    print(name)
    data = pd.read_csv("./Expertise/" + name + ".csv")
    print(cal_Scheirer_Ray_Hare(data, name))
