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

market_value_formula = f'=Table1[@[Size]*[Last]]'
percent_netliq_formula = f'=(Table1[@[Market Value]])/{netliq_cell_loc}'


ws.add_table('B3:G11', {'data': data,
                        'columns': [{'header': 'Financial Instrument'},
                                    {'header': 'Size'},
                                    {'header': 'Last'},
                                    {'header': 'Market Value',
                                     'formula': market_value_formula},
                                    {'header': '% of Net Liq',
                                     'formula': percent_netliq_formula},
                                    {'header': 'Safe %'},
                                    {'header': 'Workhorse %'},
                                    {'header': 'Redhead %'}
                                    ]})

pleasefillin_format = wb.add_format({'bg_color':   '#FFC7CE'})
# add cash data using wizard  - - cond. formatting RED
ws.conditional_format('B3', {'type':   'blanks',
                                       'format': pleasefillin_format})
# add strat relationship columns using wizard - cond. formatting RED.
ws.conditional_format('Table1[[Safe %]]', {'type':   'blanks',
                                                     'format': pleasefillin_format})
ws.conditional_format('Table1[[Workhorse %]]', {'type':   'blanks',
                                                          'format': pleasefillin_format})
ws.conditional_format('Table1[[Redhead %]]', {'type':   'blanks',
                                                        'format': pleasefillin_format})


# do something for puts and short assets



# calculate strat distribution in table

ws.add_table() #todo

strat_mkt_values = {strat_name:df2[strat_name] * df2['Market Value'] for strat_name in strat_list}


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

