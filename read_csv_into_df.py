import pandas as pd
import pathlib
import re
import numpy as np

def read_csv_export(export_filename: str): # todo
    df_ = pd.read_csv(export_filename)
    target_columns = ['Financial Instrument', 'Position', 'Last', 'Underlying Price']
    df2 = pd.concat([df_[key] for key in target_columns], axis=1)

    # Filter alphabetical chars and nan from 'Last'
    df2 = df2.replace({'Position': ','}, '', regex=True)
    df2 = df2.replace({'Last':'[A-Za-z]'}, '', regex=True)
    df2 = df2.replace({'Last': np.nan}, '0', regex=True)
    df2 = df2.replace({'Underlying Price': '[A-Za-z]'}, '', regex=True)
    df2 = df2.replace({'Underlying Price': ''}, '0', regex=True)
    df2 = df2.replace({'Underlying Price': np.nan}, '', regex=True) # todo - super weird tws bug where sometimes this returns for stocks...
    df2['Position'] = df2['Position'].astype(float)
    df2['Last'] = df2['Last'].where(lambda x: x != '', 0).astype(float)
    df2 = df2.replace({'Last': 0}, '')
    df2['Underlying Price'] = df2['Underlying Price'].where(lambda x: x != '', 0).astype(float)
    df2 = df2.replace({'Underlying Price': 0}, '')
    # Add "Option Flag" and "Option Strike" column to df
    option_strike = pd.Series(np.zeros((df2.shape[0]), dtype='1>U'), name='Option Strike')
    netliq_cont = pd.Series(np.zeros((df2.shape[0]), dtype='1>U'), name='Netliq Contribution')
    df2 = pd.concat([df2, option_strike,netliq_cont], axis=1)

    def func_(x):
        string_list = re.findall('\s\d{1,4}\s|\s\d{1,4}\.\d{1,4}\s', x)
        if len(string_list) == 1:
            return float(string_list[0])
        else:
            return 0

    df2['Option Strike'] = df2['Financial Instrument'].map(func_)
    return df2


if __name__ == '__main__':
    # filename = pathlib.Path(
    #     "C:/Users/Anton/PycharmProjects/Portfolio_rebalancing_tool/input_files_for_scripts/portfolio_export_for_testing.csv")
    filename = pathlib.Path(
        "C:/Users/Anton/PycharmProjects/Portfolio_rebalancing_tool/input_files_for_scripts/portfolio_export_for_testing_easy.csv")
    df = read_csv_export(filename)
