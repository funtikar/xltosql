import xlrd
import pymysql.cursors

#setting up the excel file connections
book = xlrd.open_workbook("OUT PT - PEADS THERAPY.xls")
sh = book.sheet_by_index(8)
#####################################################


#setting up the sql connection
#connection = pymysql.connect(host='10.72.101.36',
#                             user='carakerja',
#                             password='ot101',
#                             db='carakerja_ot',
#                             charset='utf8mb4',
#                             cursorclass=pymysql.cursors.SSCursor)
#####################################################


		
thedash = set('-')
def testdash(s):
		return set(thedash).issubset(s)
							 
							 
def giveICwithnoh(xx,yy):
	print(sh.cell_value(xx,yy))
	if type(sh.cell_value(xx,yy)) == float:
		return str(sh.cell_value(xx,yy))
	if type(sh.cell_value(xx,yy)) == str:
		if testdash(sh.cell_value(xx,yy)):
			#print('yay')
			#print(sh.cell_value(xx,yy))
			sss = sh.cell_value(xx,yy).replace('-','')
			return sss
		
print(giveICwithnoh(46,4))		
		
#print(giveICwithnoh(47,4))
		
def givenamefromdb(rx,cx):	
	try:
		with connection.cursor() as cursor:
			sql = "SELECT `Nama` FROM `demoot101` WHERE `No_Pengenalan_Diri`='{0}'".format(giveICwithnoh(rx,cx))
			cursor.execute(sql)
			result = cursor.fetchone()
			if result != None:
				return result[0]
	finally:
		connection.close()

def getdemofromxl(rx,cx):
	varr = []
	varr.append(giveICwithnoh(rx,4))
	varr.append(sh.cell_value(rx,5))
	varr.append(sh.cell_value(rx,6))
	varr.append(sh.cell_value(rx,8))
	
	
	



#mykad(rx,4),nama(rx,5),alamat(rx,6),poskod(rx,8),kod_negeri(rx,9),kod_bandar(rx,7),telefon(rx,10),telefonbimbit(rx,10),tarikhlahir(rx,11),umurhari,umurbulan,umurtahun,kod_jantina(rx,13),kod_warganegara(rx,32),kod_pekerjaan(rx,33),kot_etnik(rx,16),pekerjaanlain
	
#def masukdemobaru(rx,cx):
	#if givenamefromdb(45,4)!=None:
	#	sql = "INSERT INTO 'demoot101' VALUES ('

#print("The number of worksheets is {0}".format(book.nsheets))
#print("Worksheet name(s): {0}".format(book.sheet_names()))

#print("{0} {1} {2}".format(sh.name, sh.nrows, sh.ncols))
#print("Cell D30 is {0}".format(sh.cell_value(rowx=18, colx=4)))




#print(giveICno(18,4))
#for rx in range(sh.nrows):
#    print(sh.row(rx))