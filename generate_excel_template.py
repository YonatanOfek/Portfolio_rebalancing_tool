import xlsxwriter
from read_csv_into_df import read_csv_export
import pathlib
# Create a workbook and add a worksheet.

wb = xlsxwriter.Workbook('C:/Users/Anton/PycharmProjects/Portfolio_rebalancing_tool/outputs/Portfolio_mgmt_testing.xlsx', {'nan_inf_to_errors': True})
ws = wb.add_worksheet('Data')



# STRAT_LIST = ["Safe", "Workhorse", "Redhead"]
# initate a table for one position and generate all the formulas


# initiate headers
# headers = ['Financial Instrument', 'Position', 'Last', 'Market Value', '% of Net Liq'] + STRAT_LIST

ws.write('A1', 'Net Liquidity:')
netliq_cell_loc = '$B$1'
ws.write_formula(netliq_cell_loc, '=1000000') # todo
ws.write('A2', 'USD Cash Position:')
ws.write('B2', '')

data = read_csv_export(pathlib.Path("C:/Users/Anton/PycharmProjects/Portfolio_rebalancing_tool/input_files_for_scripts/portfolio_export_for_testing_easy.csv")).values
t1_name = 'Position_list'
market_value_formula = f'=[[#This Row],[Position]]*[[#This Row],[Last]]'
percent_netliq_formula = f'=([[#This Row],[Market Value]])/{netliq_cell_loc}'
weighted_exposure_formula_redhead = f'=([[#This Row],[Market Value]])*([[#This Row],[Redhead %]])'
weighted_exposure_formula_workhorse = f'=([[#This Row], [Market Value]])*([[#This Row], [Workhorse %]])'
weighted_exposure_formula_safe = f'=([[#This Row], [Market Value]])*([[#This Row],[Safe %]])'


ws.add_table('B3:L9', {'name': f'{t1_name}',
                        'data': data,
                        'columns': [{'header': 'Financial Instrument'},
                                    {'header': 'Position'},
                                    {'header': 'Last'},
                                    {'header': 'Market Value', 'formula': market_value_formula},
                                    {'header': '% of Net Liq', 'formula': percent_netliq_formula},
                                    {'header': 'Redhead %'},
                                    {'header': 'Workhorse %'},
                                    {'header': 'Safe %'}, 
                                    {'header': 'Redhead', 'formula': weighted_exposure_formula_redhead},
                                    {'header': 'Workhorse', 'formula': weighted_exposure_formula_workhorse},
                                    {'header': 'Safe', 'formula': weighted_exposure_formula_safe}
                                    ]})

pleasefillin_format = wb.add_format({'bg_color':   '#FFC7CE'})
# add cash data using wizard  - - cond. formatting RED
ws.conditional_format('B3', {'type':   'blanks',
                                       'format': pleasefillin_format})
# add strat relationship columns using wizard - cond. formatting RED.
# ws.conditional_format(f'{t1_name}[[Safe %]]', {'type':   'blanks', 'format': pleasefillin_format})
# ws.conditional_format(f'{t1_name}[[Workhorse %]]', {'type':   'blanks', 'format': pleasefillin_format})
# ws.conditional_format(f'{t1_name}[[Redhead %]]', {'type':   'blanks', 'format': pleasefillin_format})


# do something for puts and short assets



# calculate strat distribution in table
t2_name = 'Strat_distribution'
strat_name = f'{t2_name}[@[Strategy]]'
strat_distribution_formula = f'=SUM({t1_name}[{strat_name}])' #todo - best way to choose for row?


ws.add_table('B13:G17', {'name': t2_name,
                         'data': data,
                         'columns': [{'header': 'Strategy'},
                                    {'header': 'Portfolio Weight', 'formula': strat_distribution_formula}
                                    ]})


# plot a pie chart
strategy_destribution_chart = wb.add_chart({'type': 'pie'})

strategy_destribution_chart.add_series({
    'values':      f'={t2_name}[Portfolio Weight]', #todo
    'data_labels': {'percentage': True},
})

# Save file on desktop
wb.close()

