import xlsxwriter
from xlsxwriter.worksheet import Worksheet
from read_csv_into_df import read_csv_export
import pathlib
import numpy as np


class CurrentPortfolioWorksheet(Worksheet):

    def add_formatting(self):
        # add formatting
        pleasefillin_format = wb.add_format({'bg_color': '#FFC7CE'})

        cond_pleasefillin_format = wb.add_format()
        # add cash data using wizard  - - cond. formatting RED
        ws.conditional_format('B2', {'type': 'blanks',
                                     'format': pleasefillin_format})
        # add strat relationship columns using wizard - cond. formatting RED.
        ws.conditional_format(pos_list_table_range, {'type': 'blanks', 'format': pleasefillin_format})

        return

    def add_warnings_and_misc_cells(self):
        netliq_cell_loc = '$B$1'
        usdcash_cell_loc = '$B$2'
        # Add notes and netliq, cash

        ws.write('A1', 'Net Liquidity:')

        ws.write_formula(netliq_cell_loc,
                         f'={usdcash_cell_loc} + SUM({t1_name}[[#Data],[Netliq Contribution]])')  # todo this is dumb because doesn't consider premia given from put options rather margin load...

        ws.write('A2', 'USD Cash Position:')
        ws.write('B2', '')
        ws.set_column(data.shape[1] + 15, data.shape[1] + 15, width=115)
        ws.write(0, data.shape[1] + 15,
                 'Note for OPTION POSITIONS - Margin requirement can change at any time, so position sizing values SHOULD BE REVIEWED MANUALLY',
                 wb.add_format({'bold': True, 'bg_color': 'black', 'font_color': 'red'}))

        return

    def add_positions_list_table(self, portfolio_export_data):
        # initate a table for one position and generate all the formulas

        t1_name = 'Position_list'

        stock_market_value_component = f'[[#This Row],[Position]]*[[#This Row],[Last]]'
        # do something for puts and short assets
        # Put Price + Maximum ((20% 2* Underlying Price - Out of the Money Amount),(10% * Strike Price))
        otm_component = f'MAX(0,[[#This Row],[Underlying Price]]-[[#This Row],[Option Strike]])'
        maxing_component = f'Max(0.2* [[#This Row],[Underlying Price]] - {otm_component},[[#This Row],[Option Strike]] * 0.1)'
        short_put_market_value_component = f'[[#This Row],[Position]]*(-100)*([[#This Row],[Last]]+{maxing_component})'
        put_market_value_formula = 0  # todo v3
        short_call_market_value_formula = 0  # todo v3
        call_market_value_formula = 0  # todo v3
        polymorphic_market_value_formula = f'=If([[#This Row],[Option Strike]] > 0,{short_put_market_value_component},{stock_market_value_component})'

        short_option_netliq_contribution_component = '(-100)*[[#This Row],[Last]]*[[#This Row],[Position]]'
        stock_netliq_contribution_component = '[[#This Row],[Market Value]]'
        polymorphic_netliq_contribution_formula = f'=If([[#This Row],[Option Strike]] > 0,{short_option_netliq_contribution_component},{stock_netliq_contribution_component})'

        percent_netliq_formula = f'=([[#This Row],[Market Value]])/{netliq_cell_loc}'
        weighted_exposure_formula_redhead = f'=([[#This Row],[Market Value]])*([[#This Row],[Redhead %]])'
        weighted_exposure_formula_workhorse = f'=([[#This Row], [Market Value]])*([[#This Row], [Workhorse %]])'
        weighted_exposure_formula_safe = f'=([[#This Row], [Market Value]])*([[#This Row],[Safe %]])'

        # add table
        topleft_corner_of_t1 = [2, 1]  # B1
        t1_range = [topleft_corner_of_t1[0], topleft_corner_of_t1[1], data.shape[0] + topleft_corner_of_t1[0],
                    data.shape[1] + 8]

        ws.add_table(t1_range[0], t1_range[1], t1_range[2], t1_range[3], {'name': f'{t1_name}',
                                                                          'data': data,
                                                                          'columns': [
                                                                              {'header': 'Financial Instrument'},
                                                                              {'header': 'Position'},
                                                                              {'header': 'Last'},
                                                                              {'header': 'Underlying Price'},
                                                                              {'header': 'Option Strike'},
                                                                              {'header': 'Market Value',
                                                                               'formula': polymorphic_market_value_formula},
                                                                              {'header': 'Netliq Contribution',
                                                                               'formula': polymorphic_netliq_contribution_formula},
                                                                              {'header': '% of Net Liq',
                                                                               'formula': percent_netliq_formula},
                                                                              {'header': 'Redhead %',
                                                                               'format': cond_pleasefillin_format},
                                                                              {'header': 'Workhorse %',
                                                                               'format': cond_pleasefillin_format},
                                                                              {'header': 'Safe %',
                                                                               'format': cond_pleasefillin_format},
                                                                              {'header': 'Redhead',
                                                                               'formula': weighted_exposure_formula_redhead},
                                                                              {'header': 'Workhorse',
                                                                               'formula': weighted_exposure_formula_workhorse},
                                                                              {'header': 'Safe',
                                                                               'formula': weighted_exposure_formula_safe}
                                                                              ]})

        pos_list_table_range = ws.tables[0]['range']

        return

    def add_strategies_table(self, strategy_names):
        # initialize t2
        t2_name = 'Strat_distribution'
        topleft_corner_of_t2 = [2, data.shape[1] + 10]
        t1_range = [topleft_corner_of_t2[0], topleft_corner_of_t2[1], topleft_corner_of_t2[0] + 3,
                    topleft_corner_of_t2[1] + 1]
        data2 = np.zeros([3, 2], '<U1')

        # calculate strat distribution in table

        ws.add_table(t1_range[0], t1_range[1], t1_range[2], t1_range[3], {'name': t2_name,
                                                                          'data': data2,
                                                                          'columns': [{'header': 'Strategy'},
                                                                                      {'header': 'Portfolio Weight'}
                                                                                      ]})
        for i, j in zip([1, 2, 3], ['Redhead', 'Workhorse', 'Safe']):
            ws.write(topleft_corner_of_t2[0] + i, topleft_corner_of_t2[1], j)
        for i, j in zip([1, 2, 3], [f'=SUM({t1_name}[[#Data],[Redhead]])', f'=SUM({t1_name}[[#Data],[Workhorse]])',
                                    f'=SUM({t1_name}[[#Data],[Safe]])']):
            ws.write_formula(topleft_corner_of_t2[0] + i, topleft_corner_of_t2[1] + 1, j)

        return

    def add_strat_dist_chart(self):
        # plot a pie chart
        strategy_destribution_chart = wb.add_chart({'type': 'pie'})
        strategy_destribution_chart.set_title({'name': 'Portfolio Strategies Distribution Chart'})
        strategy_destribution_chart.add_series({
            'categories': f'={t2_name}[Strategy]',
            'values': f'={t2_name}[Portfolio Weight]',
            'data_labels': {'percentage': True},
        })
        ws.insert_chart(topleft_corner_of_t2[0], topleft_corner_of_t2[1] + 5, strategy_destribution_chart)

        return

if __name__ == '__main__':
        filename_ = 'C:/Users/Anton/PycharmProjects/Portfolio_rebalancing_tool/outputs/Portfolio_mgmt_testing.xlsx'
        wb = xlsxwriter.Workbook(filename_, {'nan_inf_to_errors': True})
        data = read_csv_export(pathlib.Path("C:/Users/Anton/PycharmProjects/Portfolio_rebalancing_tool/input_files_for_scripts/portfolio_export_for_testing_v2.csv")).values  # todo why n = 59?
        ws = wb.add_worksheet('Current Portfolio', worksheet_class=CurrentPortfolioWorksheet)
        ws.add_positions_list_table(data)
        wb.close()


