import xlrd

book = xlrd.open_workbook("kodbandar.xls")
sh = book.sheet_by_index(0)

