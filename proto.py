import xlrd
import pymysql.cursors

#setting up the excel file connections
book = xlrd.open_workbook("sample.xls")
sh = book.sheet_by_index(0)
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

def getkodnegeri(rx,cx):
	str1 = "johor"
	str2 = "kedah"
	str3 = "kelantan"
	str4 = "melaka"
	str5 = "sembilan"
	str6 = "pahang"
	str7 = "pinang"
	str8 = "perak"
	str9 = "perlis"
	str10 = "selangor"
	str11 = "terengganu"
	str12 = "sabah"
	str13 = "sarawak"
	str14 = "lumpur"
	str15 = "labuan"
	str16 = "putrajaya"
	str17 = "luar negara"
	ck = sh.cell_value(rx,cx).lower()
	if ck.find(str1) != -1:
		return "01"
	elif ck.find(str2) != -1:
		return "02"
	elif ck.find(str3) != -1:
		return "03"
	elif ck.find(str4) != -1:
		return "04"
	elif ck.find(str5) != -1:
		return "05"
	elif ck.find(str6) != -1:
		return "06"	
	elif ck.find(str6) != -1:
		return "06"
	elif ck.find(str7) != -1:
		return "07"
	elif ck.find(str8) != -1:
		return "08"
	elif ck.find(str9) != -1:
		return "09"
	elif ck.find(str10) != -1:
		return "10"
	elif ck.find(str11) != -1:
		return "11"
	elif ck.find(str12) != -1:
		return "12"
	elif ck.find(str13) != -1:
		return "13"
	elif ck.find(str14) != -1:
		return "14"
	elif ck.find(str15) != -1:
		return "15"
	elif ck.find(str16) != -1:
		return "16"
	elif ck.find(str17) != -1:
		return "98"
	else:
		return "tiada"
		
print(getkodnegeri(60,9))
		
def getdemofromxl(rx,cx):
	varr = []
	varr.append(giveICwithnoh(rx,4))
	varr.append(sh.cell_value(rx,5))#nama
	varr.append(sh.cell_value(rx,6))#alamat
	varr.append(sh.cell_value(rx,8))#poskod
	
	
	
	



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