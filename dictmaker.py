import xlrd

book = xlrd.open_workbook("kodbandar.xls")
sh = book.sheet_by_index(1)

d = {}

myList = []

#for i in range(0,16):
#	myList.insert(0,sh.cell_value(i,1))
#for i in range(0,449):
#	if myList.count(sh.cell_value(i,1)) != 1:
#		print("jumlah bandar {0} {1}".format(sh.cell_value(i,1),myList.count(sh.cell_value(i,1))))

for i in range(0,16):
#	if d[sh.cell_value(i,1)] != None:	
#		print("{0} already have values".format(sh.cell_value(i,1)))
	d[sh.cell_value(i,1)] = str(sh.cell_value(i,0))
print(d)