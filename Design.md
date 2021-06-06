



### Present t0 portfolio state, as positions and also transposed to sub-portfolios


### Generate a Delta portfolio state



### Sum the Delta portfolio and t0 portfolio to arrive at a t1 portfolio (which can than be validated)




###

### User Stories:
# 1) Re-balance a portfolio of sub-portfolios (which in-turn contain individual positions).
#   a) Review a portfolio of sub-portfolio
#   b) Simulate a set of change to the portfolio, and review the resulting portfolio
#   c) Extract a portfolio from TWS
#       - Portfolio positions have notes with their sub-portfolio tags, maybe also with their strategy relations
#       - Also export from IBI
#       - Also export from Leumi
#   d) Define positions, edit their parameters and give tags (e.g. "PSTH warrants are notionally 70% strategy A and 30% strategy B. Their current price is 4")
#   e) Handle short positions intelligently (especially short puts)
#   f) Output to and read from an excel format (with equations?)
#   g) Display tax and commission implications, margin implications
#   h) Generate delta portfolio's orders automatically?
# 2) Review risk parameters and historical correlations for a portfolio of sub-portfolios. Todo
#
# 3) Estimate kelly performance according to arbitrarily assigned/historical parameters Todo

### Objects and design:

# There's a python backend and an Excel frontend. Mainly 1 sheet in Excel, and a python backend which fills that sheet with data.
# Pbackend also sums two excel sheets.
# In the future the python backend should write the excel sheet.


