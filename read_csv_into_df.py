import pandas as pd
import pathlib
import matplotlib


filename = pathlib.Path("C:/Users/Anton/PycharmProjects/Portfolio_rebalancing_tool/portfolio_export_for_testing.csv")
df = pd.read_csv(filename)


# create new df from selected columns ( keep only instrument name, position size, last value)


target_columns = ['Financial Instrument', 'Size', 'Last']
df2 = pd.concat([df[key] for key in target_columns], axis=1)


# add cash data using wizard TODO
# USD_CASH_POSITION = 5000
# ILS_CASH_POSITION = 49999

# add strat relationship columns, and cash data using wizard TODO

# strat_list = ["UD", "Sven", "Andrew", "Else"]

