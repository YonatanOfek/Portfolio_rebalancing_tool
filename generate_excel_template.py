import xlsxwriter
from read_csv_into_df import read_csv_export
import pathlib
# Create a workbook and add a worksheet.

wb = xlsxwriter.Workbook('Portfolio_mgmt.xlsx')
ws = wb.add_worksheet('Data')



# STRAT_LIST = ["Safe", "Workhorse", "Redhead"]
# initate a table for one position and generate all the formulas


# initiate headers
# headers = ['Financial Instrument', 'Size', 'Last', 'Market Value', '% of Net Liq'] + STRAT_LIST

ws.write('A1', 'Net Liquidity:')
netliq_cell_loc = '$B$1'
ws.write_formula(netliq_cell_loc, '=') # todo
ws.write('A2', 'USD Cash Position:')
ws.write('B2', '')

data = read_csv_export(pathlib.Path(
        "C:/Users/Anton/PycharmProjects/Portfolio_rebalancing_tool/portfolio_export_for_testing.csv")).values
t1_name = 'Position_list'
market_value_formula = f'={t1_name}[@[Size]*[Last]]'
percent_netliq_formula = f'=({t1_name}[@[Market Value]])/{netliq_cell_loc}'
weighted_exposure_formula_redhead = f'=({t1_name}[@[Market Value]])*({t1_name}[@[Redhead %]])'
weighted_exposure_formula_workhorse = f'=({t1_name}[@[Market Value]])*({t1_name}[@[Workhorse %]])'
weighted_exposure_formula_safe = f'=({t1_name}[@[Market Value]])*({t1_name}[@[Safe %]])'



ws.add_table('B3:G11', {'name': f'{t1_name}',
                        'data': data,
                        'columns': [{'header': 'Financial Instrument'},
                                    {'header': 'Size'},
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
ws.conditional_format(f'{t1_name}[[Safe %]]', {'type':   'blanks', 'format': pleasefillin_format})
ws.conditional_format(f'{t1_name}[[Workhorse %]]', {'type':   'blanks', 'format': pleasefillin_format})
ws.conditional_format(f'{t1_name}[[Redhead %]]', {'type':   'blanks', 'format': pleasefillin_format})


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

chart.add_series({
    'values':      '=Table2[]', #todo
    'data_labels': {'percentage': True},
})
ws.Range('A1').FormulaR1C1 = 'X'
ws.Range('B1').FormulaR1C1 = 'Y'
ws.Range('A2').FormulaR1C1 = 1
ws.Range('A3').FormulaR1C1 = 2
ws.Range('A4').FormulaR1C1 = 3
ws.Range('B2').FormulaR1C1 = 4
ws.Range('B3').FormulaR1C1 = 5
ws.Range('B4').FormulaR1C1 = 6
#ws.Range('A1:B4').Select()
ch = ws.Shapes.AddChart().Select()
xl.ActiveChart.ChartType = c.xlXYScatterLines
xl.ActiveChart.SetSourceData(Source=ws.Range("A1:B4"))


# Save file on desktop
wb.close()

