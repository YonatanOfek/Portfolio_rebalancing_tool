import pandas as pd
import pathlib
import numpy as np

def read_csv_export(export_filename: str):
    df_ = pd.read_csv(export_filename)
    target_columns = ['Financial Instrument', 'Position', 'Last']
    df2 = pd.concat([df_[key] for key in target_columns], axis=1)

    # clean data  todo Clean position data-get rid of ',' 'C' and convert to numeric.
    #             todo Clean Last data - get rid of alphabeticals and convert to numeric.
    # df2['Position'] = pd.to_numeric(df2['Position'])
    #
    # regex = '[a-z]'
    # for i in df2['Last']:
    #     re.sub(regex,i,'')
    #
    # df2['Last'] = pd.to_numeric(df2['Last'])



    df2 = df2.replace({'Last':'[A-Za-z]'},'',regex=True)
    df2 = df2.replace({'Last': np.nan}, '0', regex=True)

    return df2


if __name__ == '__main__':
    # filename = pathlib.Path(
    #     "C:/Users/Anton/PycharmProjects/Portfolio_rebalancing_tool/input_files_for_scripts/portfolio_export_for_testing.csv")
    filename = pathlib.Path(
        "C:/Users/Anton/PycharmProjects/Portfolio_rebalancing_tool/input_files_for_scripts/portfolio_export_for_testing_easy.csv")
    df = read_csv_export(filename)
