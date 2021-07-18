import pathlib
import re

import pandas as pd
import numpy as np


def filter_data(df_):

    df_ = df_.replace({'Position': ','}, '', regex=True)
    df_ = df_.replace({'Last': '[A-Za-z]'}, '', regex=True)
    df_ = df_.replace({'Last': np.nan}, '0', regex=True)
    df_ = df_.replace({'Underlying Price': '[A-Za-z]'}, '', regex=True)
    df_ = df_.replace({'Underlying Price': ''}, '0', regex=True)
    df_ = df_.replace({'Underlying Price': np.nan}, '', regex=True)
    df_['Position'] = df_['Position'].astype(float)
    df_['Last'] = df_['Last'].where(lambda x: x != '', 0).astype(float)
    df_ = df_.replace({'Last': 0}, '')
    df_['Underlying Price'] = df_['Underlying Price'].where(lambda x: x != '', 0).astype(float)
    df_ = df_.replace({'Underlying Price': 0}, '')

    def find_option_strike(x):
        string_list = re.findall('\s\d{1,4}\s|\s\d{1,4}\.\d{1,4}\s', x)
        if len(string_list) == 1:
            return float(string_list[0])
        else:
            return 0
    df_['Option Strike'] = df_['Financial Instrument'].map(find_option_strike)

    def find_option_type(x):
        string_list = re.findall('\s[Cc][Aa][Ll][Ll]|\s[Pp][Uu][Tt]|\s[Ww][Aa][Rr]', x)
        if len(string_list) >= 1:
            return str(string_list[0][1:])
        else:
            return 0
    df_['Option Type'] = df_['Financial Instrument'].map(find_option_type)
    return df_


def rearrange_columns(df_):  # todo this func does two things...
    option_type = pd.Series(np.zeros((df_.shape[0]), dtype='1>U'), name='Option Type')
    option_strike = pd.Series(np.zeros((df_.shape[0]), dtype='1>U'), name='Option Strike')
    market_value = pd.Series(np.zeros((df_.shape[0]), dtype='1>U'), name='MarketValue')
    netliq_cont = pd.Series(np.zeros((df_.shape[0]), dtype='1>U'), name='Netliq Contribution')
    percent_of_netliq = pd.Series(np.zeros((df_.shape[0]), dtype='1>U'), name='% of NetLiq')
    df_ = pd.concat([df_, option_type, option_strike, market_value, netliq_cont, percent_of_netliq], axis=1)
    target_columns = ['Financial Instrument', 'Position', 'Last', 'Underlying Price', 'Option Type', 'Option Strike', 'MarketValue',
                      'Netliq Contribution', '% of NetLiq']

    return pd.concat([df_[key] for key in target_columns], axis=1)


def read_csv_export(export_filename: str):  # todo refactor into a nice pipe
                                            # todo naming is kinda aweful
    df_ = pd.read_csv(export_filename)
    df2 = rearrange_columns(df_)
    df2 = filter_data(df2)
    return df2


if __name__ == '__main__':
    # filename = pathlib.Path(
    #     "C:/Users/Anton/PycharmProjects/Portfolio_rebalancing_tool/input_files_for_scripts/portfolio_export_for_testing.csv")
    filename = pathlib.Path(
        "C:/Users/Anton/PycharmProjects/Portfolio_rebalancing_tool/input_files_for_scripts/portfolio_export_for_testing_v2.csv")
    df = read_csv_export(filename)

