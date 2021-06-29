import xlsxwriter
from xlsxwriter.worksheet import Worksheet
from read_csv_into_df import read_csv_export
import pathlib
import numpy as np


class CurrentPortfolioWorksheet(Worksheet):

    def __init__(self, portfolio_export_data):
        super().__init__()
        self.portfolio_export_data = portfolio_export_data

    @property
    def netliq_cell_loc(self):
        return '$B$1'
    
    @property
    def usdcash_cell_loc(self):
        return '$B$2'
    
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
    def stock_market_value_formula(self):
        return f'[[#This Row],[Position]]*[[#This Row],[Last]]'
    
    @property
    def otm_formula(self):
        return f'MAX(0,[[#This Row],[Underlying Price]]-[[#This Row],[Option Strike]])'
    
    @property
    def maxing_formula(self):
        return f'Max(0.2* [[#This Row],[Underlying Price]] - {self.otm_formula},[[#This Row],[Option Strike]] * 0.1)'
    
    @property
    def short_put_market_value_formula(self):
        return f'[[#This Row],[Position]]*(-100)*([[#This Row],[Last]]+{self.maxing_formula})'
    
    @property
    def put_market_value_formula(self):
        return 0  # todo v3
    
    @property
    def short_call_market_value_formula(self):
        return 0  # todo v3
    
    @property
    def call_market_value_formula(self):
        return 0  # todo v3
    
    @property
    def polymorphic_market_value_formula(self):
        return f'If([[#This Row],[Option Strike]] > 0,{self.short_put_market_value_formula},{self.stock_market_value_formula})'

    @property
    def short_option_netliq_contribution_formula(self):
        return '(-100)*[[#This Row],[Last]]*[[#This Row],[Position]]'
    
    @property
    def stock_netliq_contribution_formula(self):
        return '[[#This Row],[Market Value]]'
    
    @property
    def polymorphic_netliq_contribution_formula(self):
        return f'If([[#This Row],[Option Strike]] > 0,{self.short_option_netliq_contribution_formula},{self.stock_netliq_contribution_formula})'

    @property
    def percent_netliq_formula(self):
        return f'([[#This Row],[Market Value]])/{self.netliq_cell_loc}'
    
    @property
    def weighted_exposure_formula_redhead(self):
        return f'([[#This Row],[Market Value]])*([[#This Row],[Redhead %]])'
    
    @property
    def weighted_exposure_formula_workhorse(self):
        return f'([[#This Row], [Market Value]])*([[#This Row], [Workhorse %]])'
    
    @property
    def weighted_exposure_formula_safe(self):
        return f'([[#This Row], [Market Value]])*([[#This Row],[Safe %]])'

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
    def topleft_corner_of_t1(self):
        return [2, 1]  # B3 todo add xl to row here too..

    @property
    def topleft_corner_of_t2(self):
        return [2, self.portfolio_export_data.shape[1] + 10] # todo add xl to row here too..

    @property
    def t1_range(self):
        return [self.topleft_corner_of_t1[0], self.topleft_corner_of_t1[1],
                self.portfolio_export_data.shape[0] + self.topleft_corner_of_t1[0], self.portfolio_export_data.shape[1] + 8]

    @property
    def t2_range(self):
        return [self.topleft_corner_of_t2[0], self.topleft_corner_of_t2[1], self.topleft_corner_of_t2[0] + 3,
                self.topleft_corner_of_t2[1] + 1]

    def add_warnings_and_misc_cells(self):

        # Add netliq, cash
        self.write('A1', 'Net Liquidity:') # todo incorporate xl_row_col_to_str methods to make this proper...
        self.write_formula(self.netliq_cell_loc, f'={self.netliq_formula}')

        self.write('A2', 'USD Cash Position:')

        # Add warning
        self.set_column('V:V', width=115)  #data.shape[1] + 15, data.shape[1] + 15 todo incorporate xl_row_col_to_str methods to make this proper...
        self.write('V1',                     # 0, data.shape[1] + 15, todo incorporate xl_row_col_to_str methods to make this proper...
                   'Note for OPTION POSITIONS - Margin requirement can change at any time, so position sizing values SHOULD BE REVIEWED MANUALLY')

        return

    def add_positions_list_table(self):

        self.add_table(self.t1_range[0], self.t1_range[1], self.t1_range[2], self.t1_range[3], {'name': f'{self.pos_table_name}',
                                                                          'data': self.portfolio_export_data,
                                                                          'columns': [
                                                                              {'header': 'Financial Instrument'}, # todo get these headers programmatically after switching data from a df.values into a df
                                                                              {'header': 'Position'},
                                                                              {'header': 'Last'},
                                                                              {'header': 'Underlying Price'},
                                                                              {'header': 'Option Strike'},
                                                                              {'header': 'Market Value',
                                                                               'formula': f'={self.polymorphic_market_value_formula}'},
                                                                              {'header': 'Netliq Contribution',
                                                                               'formula': f'={self.polymorphic_netliq_contribution_formula}'},
                                                                              {'header': '% of Net Liq',
                                                                               'formula': f'={self.percent_netliq_formula}'},
                                                                              {'header': 'Redhead %'},
                                                                              {'header': 'Workhorse %'},
                                                                              {'header': 'Safe %'},
                                                                              {'header': 'Redhead',
                                                                               'formula': f'={self.weighted_exposure_formula_redhead}'}, # todo add some polymorphism
                                                                              {'header': 'Workhorse',
                                                                               'formula': f'={self.weighted_exposure_formula_workhorse}'},
                                                                              {'header': 'Safe',
                                                                               'formula': f'={self.weighted_exposure_formula_safe}'}
                                                                              ]})

        return # todo returns

    def add_strategies_table(self, strategy_names):
        # data2 = np.zeros([3, 2], '<U1') todo not need this right?

        self.add_table(self.t1_range[0], self.t1_range[1], self.t1_range[2], self.t1_range[3], {'name': self.strat_table_name,
                                                                            # 'data': data2 todo not need this right?
                                                                            'columns': [{'header': 'Strategy'},
                                                                                        {'header': 'Portfolio Weight'}
                                                                                        ]})
        for i, j in zip([1, 2, 3], ['Redhead', 'Workhorse', 'Safe']):
            self.write(self.topleft_corner_of_t2[0] + i, self.topleft_corner_of_t2[1], j)
        for i, j in zip([1, 2, 3], [f'={self.redhead_strat_summing_formula}', f'={self.workhorse_strat_summing_formula}',
                                    f'={self.safe_strat_summing_formula}']):
            self.write_formula(self.topleft_corner_of_t2[0] + i, self.topleft_corner_of_t2[1] + 1, j)

        return

    def add_strat_dist_chart(self):
        # plot a pie chart
        strategy_destribution_chart = wb.add_chart({'type': 'pie'})
        strategy_destribution_chart.set_title({'name': 'Portfolio Strategies Distribution Chart'})
        strategy_destribution_chart.add_series({
            'categories': f'={self.strat_table_name}[Strategy]',
            'values': f'={self.strat_table_name}[Portfolio Weight]',
            'data_labels': {'percentage': True},
        })
        self.insert_chart(self.topleft_corner_of_t2[0], self.topleft_corner_of_t2[1] + 5, strategy_destribution_chart)

        return # todo returns


class PortfolioBalanceWorkbook(xlsxwriter.Workbook):
    """
    This object creates the Portfolio Balance Workbook. It calls methods in the CurrentPortfolioWorkSheet object which it
    generates during its run.
    """

    def __init__(self, portfolio_export_data, filename=None, options=None):
        super().__init__(filename, options)
        self.curr_port_ws = self.add_worksheet('Current Portfolio', worksheet_class=CurrentPortfolioWorksheet, data=portfolio_export_data)

    @property
    def input_is_required_format(self):
        return self.add_format({'bg_color': '#FFC7CE'})

    def add_formatting(self):
        self.curr_port_ws.conditional_format(self.curr_port_ws.usdcash_cell_loc, {'type': 'blanks',
                                                                                  'format': self.input_is_required_format})
        self.curr_port_ws.write('V1', wb.add_format({'bold': True, 'bg_color': 'black', 'font_color': 'red'}))
        self.curr_port_ws.conditional_format(self.curr_port_ws.pos_list_table_range,
                                             {'type': 'blanks', 'format': self.input_is_required_format})
    def create_curr_port_worksh(self, portfolio_export_data):
        self.curr_port_ws.add_warnings_and_misc_cells()
        self.curr_port_ws.add_positions_list_table()
        self.curr_port_ws.add_strategies_table()
        self.curr_port_ws.add_strat_dist_chart()
        self.add_formatting()

        return # todo returns

if __name__ == '__main__':
    data_filename = pathlib.Path(
        "C:/Users/Anton/PycharmProjects/Portfolio_rebalancing_tool/input_files_for_scripts/portfolio_export_for_testing_v2.csv")
    data_ = read_csv_export(data_filename).values
    workbook_filename = 'C:/Users/Anton/PycharmProjects/Portfolio_rebalancing_tool/outputs/Portfolio_mgmt_testing.xlsx'
    wb = PortfolioBalanceWorkbook(workbook_filename, {'nan_inf_to_errors': True})
    wb.create_curr_port_worksh(data_)
    wb.close()


