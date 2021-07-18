import pathlib
import re
from typing import Union

import pandas as pd
import numpy as np


def add_json_inputs(df_, json_filename):  # todo - an elegant way to add to the df if there's data to add

    json_df = pd.read_json(json_filename)

    for i in range(json_df.shape[0]):
        for j in range(df_.shape[0]):
            if json_df['Financial Instrument'][i] == df_['Financial Instrument'][j]:
                for strat in json_df.columns[1:]:
                    df_[f'{strat}'][j] = json_df[f'{strat}'][i]  # todo this is improper pandas
                break

    return df_


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

    def func_(x):
        string_list = re.findall('\s\d{1,4}\s|\s\d{1,4}\.\d{1,4}\s', x)
        if len(string_list) == 1:
            return float(string_list[0])
        else:
            return 0
    df_['Option Strike'] = df_['Financial Instrument'].map(func_)
    return df_


def rearrange_columns(df_, json_filename=None): # todo this func does two things...
    option_strike = pd.Series(np.zeros((df_.shape[0]), dtype='1>U'), name='Option Strike')
    market_value = pd.Series(np.zeros((df_.shape[0]), dtype='1>U'), name='MarketValue')
    netliq_cont = pd.Series(np.zeros((df_.shape[0]), dtype='1>U'), name='Netliq Contribution')
    percent_of_netliq = pd.Series(np.zeros((df_.shape[0]), dtype='1>U'), name='% of NetLiq')
    df_ = pd.concat([df_, option_strike, market_value, netliq_cont, percent_of_netliq], axis=1)
    target_columns = ['Financial Instrument', 'Position', 'Last', 'Underlying Price', 'Option Strike', 'MarketValue',
                      'Netliq Contribution', '% of NetLiq']
    if json_filename is not None:
        names_ = pd.read_json(json_filename).columns[1:]
        strats = [pd.Series(np.zeros((df_.shape[0]), dtype='1>U'), name=name_) for name_ in names_]
        df_ = pd.concat([df_] + strats, axis=1)
        return pd.concat([df_[key] for key in target_columns + [name_ for name_ in names_]], axis=1)

    return pd.concat([df_[key] for key in target_columns], axis=1)


def read_csv_export(export_filename: str, add_json: Union[None, str] = None):  # todo refactor into a nice pipe
                                                                               # todo naming is kinda aweful
    df_ = pd.read_csv(export_filename)
    df2 = rearrange_columns(df_, add_json)
    df2 = filter_data(df2)
    if add_json is not None:
        return add_json_inputs(df2, add_json)
    return df2


if __name__ == '__main__':
    # filename = pathlib.Path(
    #     "C:/Users/Anton/PycharmProjects/Portfolio_rebalancing_tool/input_files_for_scripts/portfolio_export_for_testing.csv")
    filename = pathlib.Path(
        "C:/Users/Anton/PycharmProjects/Portfolio_rebalancing_tool/input_files_for_scripts/portfolio_export_for_testing_v2.csv")
    df = read_csv_export(filename)

