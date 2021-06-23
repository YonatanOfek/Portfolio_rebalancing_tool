import xlsxwriter

# Create a workbook and add a worksheet.

wb = xlsxwriter.Workbook('testing_1.xlsx')
ws = wb.add_worksheet('Data1')


ws.add_table('B3:G7', {
                              'columns': [{'header': 'Product'},
                                          {'header': 'Quarter 1'},
                                          {'header': 'Quarter 2'},
                                          {'header': 'Quarter 3'},
                                          {'header': 'Quarter 4'},
                                          {'header': 'Year',
                                           },
                                          ]})

wb.close()