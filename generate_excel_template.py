
# generate formuleic excels, print into them only correct columns: instrument name, strat relationships, position size, last value
from win32com.client import Dispatch

xl = Dispatch("Excel.application")


for n in range(1, 10):
    worksheet.write(n, 0, random.randint(1,150000))
    worksheet.write(n, 1, random.randint(1, 100))
    cell1 = xl_rowcol_to_cell(n,0)
    cell2 = xl_rowcol_to_cell(n,1)
    worksheet.write_formula(n, 2, f'=({cell1}*{cell2})/100')
workbook.close()


# calculate strat distribution and plot a pie chart

strat_mkt_values = {strat_name:df2[strat_name] * df2['Market Value'] for strat_name in strat_list}
matplotlib.something

# net liq, %of liq for position, etc.

df2['Market Value'] = df2['Size'] * df2['Last']
net_liq = sum(df2['Market Value'])
df2['% of Liq'] = df2['Market Value'] / net_liq

# do something for puts and short assets


# print(df)