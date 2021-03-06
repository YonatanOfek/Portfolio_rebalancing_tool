import pathlib
from typing import List, Union

import xlsxwriter
import pandas as pd

from read_csv_into_df import read_csv_export


class CurrentPortfolioWorksheet(xlsxwriter.worksheet.Worksheet):

    def load_portfolio_export_data(self, portfolio_export_data, list_of_strats):
        self.portfolio_export_data = portfolio_export_data
        self.list_of_strats = list_of_strats

    @property
    def pos_table_name(self):
        return 'Position_list'

    @property
    def strat_table_name(self):
        return 'Strat_distribution'

    @property
    def netliq_formula(self):
        return f"{self.usdcash_cell_loc} + SUM({self.pos_table_name}[[#Data],[Netliq Contribution]])"

    @property
    def market_value_formula_stock(self):
        return f'[[#This Row],[Position]]*[[#This Row],[Last]]'

    @property
    def otm_formula_put(self):
        return f'MAX(0,[[#This Row],[Underlying Price]]-[[#This Row],[Option Strike]])'

    @property
    def maxing_formula_short_put(self):
        return f'Max(0.2* [[#This Row],[Underlying Price]] - {self.otm_formula_put},[[#This Row],[Option Strike]] * 0.1)'

    @property
    def short_put_market_value_formula(self):
        return f'[[#This Row],[Position]]*(-100)*([[#This Row],[Last]]+{self.maxing_formula_short_put})'

    @property
    def put_market_value_formula(self):
        return f'0'

    @property
    def otm_formula_call(self):
        return f'MAX(0,[[#This Row],[Option Strike]]- [[#This Row],[Underlying Price]])'

    @property
    def maxing_formula_short_call(self):
        return f'Max(0.2* [[#This Row],[Underlying Price]] - {self.otm_formula_call},[[#This Row],[Underlying Price]] * 0.1)'

    @property
    def short_call_market_value_formula(self):
        return f'[[#This Row],[Position]]*(-100)*([[#This Row],[Last]] + {self.maxing_formula_short_call})' # This uses the formula for naked calls because it is more stringent. Also - it is used by IBKR for PMCCs

    @property
    def call_and_warrant_market_value_formula(self):
        return f'{self.long_option_netliq_contribution_formula}'

    @property
    def polymorphic_long_short_call_option_market_value_formula(self):
        return f'If([[#This Row],[Position]] > 0,{self.call_and_warrant_market_value_formula},{self.short_call_market_value_formula})'

    @property
    def polymorphic_long_short_put_option_market_value_formula(self):
        return f'If([[#This Row],[Position]] > 0,{self.put_market_value_formula},{self.short_put_market_value_formula})'

    @property
    def polymorphic_call_put_option_market_value_formula(self):
        return f'If([[#This Row],[Option Type]] ="PUT",{self.polymorphic_long_short_put_option_market_value_formula},' \
               f'{self.polymorphic_long_short_call_option_market_value_formula})'

    @property
    def polymorphic_market_value_formula(self):
        return f'If([[#This Row],[Option Strike]] > 0,{self.polymorphic_call_put_option_market_value_formula},{self.market_value_formula_stock})'

    @property
    def short_option_netliq_contribution_formula(self):
        return '(-100)*[[#This Row],[Last]]*[[#This Row],[Position]]'

    @property
    def long_option_netliq_contribution_formula(self):
        return 'If([[#This Row],[Option Type]]="WAR",[[#This Row],[Last]]*[[#This Row],[Position]],(100)*[[#This Row],[Last]]*[[#This Row],[Position]])'

    @property
    def polymorphic_long_short_option_netliq_contribution_formula(self):
        return f'If([[#This Row],[Position]] > 0,{self.long_option_netliq_contribution_formula},{self.short_option_netliq_contribution_formula})'

    @property
    def stock_netliq_contribution_formula(self):
        return '[[#This Row],[MarketValue]]'

    @property
    def polymorphic_netliq_contribution_formula(self):
        return f'If([[#This Row],[Option Strike]] > 0,{self.polymorphic_long_short_option_netliq_contribution_formula},' \
               f'{self.stock_netliq_contribution_formula})'

    @property
    def percent_netliq_formula(self):
        return f'([[#This Row],[MarketValue]])/{self.netliq_cell_loc}'

    def weighted_exposure_formula(self, column_header):
        return f'([[#This Row],[MarketValue]])*([[#This Row],[{column_header}]])'

    def strat_summing_formula(self, strat_name):
        return f'SUM({self.pos_table_name}[[#Data],[{strat_name}]])'

    @property
    def redhead_strat_summing_formula(self):
        return f'SUM({self.pos_table_name}[[#Data],[Redhead]])'

    @property
    def workhorse_strat_summing_formula(self):
        return f'SUM({self.pos_table_name}[[#Data],[Workhorse]])'

    @property
    def safe_strat_summing_formula(self):
        return f'SUM({self.pos_table_name}[[#Data],[Safe]])'

    @property
    def top_left_corner_of_pos_table(self):
        return [2, 1]  # B3 todo add xl to row here too..

    @property
    def pos_list_table_range(self):
        return [self.top_left_corner_of_pos_table[0], self.top_left_corner_of_pos_table[1],
                self.portfolio_export_data.shape[0] + self.top_left_corner_of_pos_table[0],
                self.portfolio_export_data.shape[1] + 2 * len(self.list_of_strats)]

    @property
    def top_left_corner_of_strats_table(self):
        return [2, self.pos_list_table_range[3] + 2]  # todo add xl to row here too..

    @property
    def strat_list_table_range(self):
        return [self.top_left_corner_of_strats_table[0], self.top_left_corner_of_strats_table[1],
                self.top_left_corner_of_strats_table[0] + len(self.list_of_strats) + 1,
                self.top_left_corner_of_strats_table[1] + 1]

    @property
    def data_columns(self):
        return ['Financial Instrument', 'Position', 'Last', 'Underlying Price', 'Option Type', 'Option Strike']  # todo get these headers programmatically after switching data from a df.values into a df

    @property
    def pos_table_headers_list(self):
        data_columns_headers = self.data_columns
        calculated_columns_headers = ['MarketValue', 'Netliq Contribution', '% of NetLiq']
        strats_list_percent_columns_headers = [strat + ' %' for strat in self.list_of_strats]
        return data_columns_headers + calculated_columns_headers + strats_list_percent_columns_headers + self.list_of_strats

    @property
    def pos_table_formulas_list(self):
        data_columns_formulas = [''] * len(self.data_columns)

        calculated_columns_formulas = [f'={self.polymorphic_market_value_formula}',
                                       f'={self.polymorphic_netliq_contribution_formula}',
                                       f'={self.percent_netliq_formula}']

        strats_list_percent_columns_formulas = [''] * len(self.list_of_strats)
        strats_list_columns_formulas = [f'={self.weighted_exposure_formula(strat_header)}' for strat_header in
                                        [strat + ' %' for strat in self.list_of_strats]]
        return data_columns_formulas + calculated_columns_formulas + strats_list_percent_columns_formulas + strats_list_columns_formulas

    @property # should this be a property though?
    def pos_list_table_columns(self):
        return [{'header': header, 'formula': formula} for header, formula in zip(self.pos_table_headers_list, self.pos_table_formulas_list)]

    @property
    def netliq_cell_loc(self):
        return '$B$1'

    @property
    def usdcash_cell_loc(self):
        return '$B$2'

    @property
    def warning_cell_loc(self):
        return [0, self.top_left_corner_of_strats_table[1] + 5]

    def add_warnings_and_misc_cells(self):

        # Add netliq, cash
        self.write('A1', 'Net Liquidity:') # todo incorporate xl_row_col_to_str methods to make this proper...
        self.write_formula(self.netliq_cell_loc, f'={self.netliq_formula}')

        self.write('A2', 'USD Cash Position:')

        # Add warning
        self.set_column(first_col=self.warning_cell_loc[1], last_col=self.warning_cell_loc[1], width=115)  # todo incorporate xl_row_col_to_str methods to make this proper...
        self.write(self.warning_cell_loc[0], self.warning_cell_loc[1],
                   'Note for OPTION POSITIONS - Margin requirement can change at any time, so position sizing values SHOULD BE REVIEWED MANUALLY')

    def add_positions_list_table(self):

        self.add_table(self.pos_list_table_range[0], self.pos_list_table_range[1], self.pos_list_table_range[2],
                       self.pos_list_table_range[3], {'name': f'{self.pos_table_name}',
                                                      'data': self.portfolio_export_data,
                                                      'columns': self.pos_list_table_columns})

    def add_strategies_table(self):

        self.add_table(self.strat_list_table_range[0], self.strat_list_table_range[1], self.strat_list_table_range[2],
                       self.strat_list_table_range[3], {'name': self.strat_table_name, 'columns': [{'header': 'Strategy'},
                                                                                        {'header': 'Portfolio Weight'}
                                                                                        ]})

        for i, j in zip(range(1, len(self.list_of_strats) + 2), self.list_of_strats + ['Cash']):
            self.write(self.top_left_corner_of_strats_table[0] + i, self.top_left_corner_of_strats_table[1], j)
        for i, j in zip(range(1, len(self.list_of_strats) + 2), [f'={self.strat_summing_formula(strat)}' for strat in self.list_of_strats] + [f'={self.usdcash_cell_loc}']):
            self.write_formula(self.top_left_corner_of_strats_table[0] + i, self.top_left_corner_of_strats_table[1] + 1, j)

    def add_strat_dist_chart(self):
        # plot a pie chart
        strategy_destribution_chart = wb.add_chart({'type': 'pie'})
        strategy_destribution_chart.set_title({'name': 'Portfolio Strategies Distribution Chart'})
        strategy_destribution_chart.add_series({
            'categories': f'={self.strat_table_name}[Strategy]',
            'values': f'={self.strat_table_name}[Portfolio Weight]',
            'data_labels': {'percentage': True},
        })
        self.insert_chart(self.top_left_corner_of_strats_table[0], self.top_left_corner_of_strats_table[1] + 5, strategy_destribution_chart)

    def add_json_inputs(self, json_filename):  # todo - an elegant way to add to the df if there's data to add

        json_df = pd.read_json(json_filename)

        for i in range(json_df.shape[0]):
            for j in range(self.pos_list_table_range[2] - self.pos_list_table_range[0]):
                if json_df['Financial Instrument'][i] == self.portfolio_export_data[j][0]:
                    for strat in json_df.columns[1:]:
                        self.write_number(j + self.top_left_corner_of_pos_table[0] + 1, self.pos_table_headers_list.index(strat)
                                          + self.top_left_corner_of_pos_table[1], json_df[f'{strat}'][i])
                    break

class PortfolioBalanceWorkbook(xlsxwriter.Workbook):
    """
    This object creates the Portfolio Balance Workbook. It calls methods in the CurrentPortfolioWorkSheet object which it
    generates during its run.
    """

    # def __init__(self, portfolio_export_data, filename=None, options=None):
    #     super().__init__(filename, options)
    #     self.curr_port_ws = self.add_worksheet('Current Portfolio', worksheet_class=CurrentPortfolioWorksheet) todo

    @property
    def input_is_required_format(self):
        return self.add_format({'bg_color': '#FFC7CE'})

    def add_formatting(self):
        self.curr_port_ws.conditional_format(self.curr_port_ws.usdcash_cell_loc, {'type': 'blanks',
                                                                                  'format': self.input_is_required_format})
        self.curr_port_ws.conditional_format(self.curr_port_ws.warning_cell_loc[0], self.curr_port_ws.warning_cell_loc[1],
                                             self.curr_port_ws.warning_cell_loc[0],
                                             self.curr_port_ws.warning_cell_loc[1],
                                             {'type':'text', 'criteria': 'containing', 'value':'Note',
                                                    'format':wb.add_format({'bold': True, 'bg_color': 'black',
                                                                            'font_color': 'red'})})
        range_for_table = self.curr_port_ws.tables[0]['range']
        self.curr_port_ws.conditional_format(range_for_table,
                                             {'type': 'blanks', 'format': self.input_is_required_format})

    def create_curr_port_worksh(self, portfolio_export_data, list_of_strats: List[str], json_filename: Union[None, str] = None):
        self.curr_port_ws = self.add_worksheet('Current Portfolio', worksheet_class=CurrentPortfolioWorksheet)
        self.curr_port_ws.load_portfolio_export_data(portfolio_export_data, list_of_strats)
        self.curr_port_ws.add_warnings_and_misc_cells()
        self.curr_port_ws.add_positions_list_table()
        if json_filename is not None:
            self.curr_port_ws.add_json_inputs(json_filename)
        self.curr_port_ws.add_strategies_table()
        self.curr_port_ws.add_strat_dist_chart()
        self.add_formatting()
        self.close()


if __name__ == '__main__':
    data_filename = pathlib.Path(
        "C:/Users/Anton/PycharmProjects/Portfolio_rebalancing_tool/input_files_for_scripts/E19072021_clean.csv")
    strats_list = ['Andrew', 'Sven', 'UD', 'Yonatan', 'Oz']
    strats_list1 = ['Redhead', 'Workhorse', 'Safe']
    strats_list2 = ['Crypto', 'MJ', 'Materials and Defensives', 'Researched Growth Companies', 'Momentum Hype']
    json_filename = 'C:/Users/Anton/PycharmProjects/Portfolio_rebalancing_tool/jsons/save_weights_E18072021.json'
    data_ = read_csv_export(data_filename).values
    workbook_filename = 'C:/Users/Anton/PycharmProjects/Portfolio_rebalancing_tool/outputs/E19072021_clean_filled.xlsx'
    wb = PortfolioBalanceWorkbook(workbook_filename, {'nan_inf_to_errors': True})
    wb.create_curr_port_worksh(data_, strats_list2, json_filename)
