import pandas as pd
import pathlib


def read_csv_export(export_filename: str):
    df_ = pd.read_csv(export_filename)
    target_columns = ['Financial Instrument', 'Position', 'Last']
    df2 = pd.concat([df_[key] for key in target_columns], axis=1)
    return df2


if __name__ == '__main__':
    filename = pathlib.Path(
        "C:/Users/Anton/PycharmProjects/Portfolio_rebalancing_tool/portfolio_export_for_testing.csv")
    df = read_csv_export(filename)
