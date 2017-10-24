import xlrd
import pymysql.cursors
import sys
import datetime

#setting up the excel file connections
book = xlrd.open_workbook("M:\\Unit Perubatan Carakerja\\BUKU DAFTAR HARIAN SEMUA DISIPLIN 2017\\OUT PT - PEADS THERAPY.xls")
sh = book.sheet_by_index(6)
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
	#print(sh.cell_value(xx,yy))
	if type(sh.cell_value(xx,yy)) == float:
		return str(sh.cell_value(xx,yy))
	if type(sh.cell_value(xx,yy)) == str:
		if testdash(sh.cell_value(xx,yy)):
			#print('yay')
			#print(sh.cell_value(xx,yy))
			sss = sh.cell_value(xx,yy).replace('-','')
			return sss
		
#print(giveICwithnoh(46,4))		
		
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
		
#print(getkodnegeri(60,9))

def getkodbandar(rx,cx):
	ck = sh.cell_value(rx,cx).lower()
	delist = ['Labuan', 'Oya', 'Sibuti', 'Pakan', 'Tebedu', 'Sibu', 'Tebakang', 'Tatau', 'Sundar', 'Sri Aman', 'Spaoh', 'Song', 'Simunjan', 'Serian', 'Sebuyau', 'Sebauh', 'Sarikei', 'Saratok', 'Roban', 'Pusa', 'Pekenu', 'Niah', 'Nanga Medamit', 'Mukah', 'Miri', 'Matu', 'Lundu', 'Lubok antu', 'Long Lama', 'Lingga', 'Limbang', 'Lawas', 'Kuching', 'Kota Samarahan', 'Kapit', 'Kanowit', 'Kabong', 'Julau', 'Engkilili', 'Debak', 'Daro', 'Dalat', 'Bintulu', 'Bintagor', 'Betong', 'Belawi', 'Belaga', 'Bekenu', 'Bau', 'Baram', 'Balingan', 'Asajaya', 'K. Kinabatangan', 'Tuaran', 'Tenom', 'Tawau', 'Tanjung Aru', 'Tamparuli', 'Tambunan', 'Sipitang', 'Semporna', 'Sandakan', 'Ranau', 'Penampang', 'Papar', 'Nabawan', 'Menumbuk', 'Membakut', 'Likas', 'Lamag', 'Lahad Datu', 'Kunak', 'Kudat', 'Kuala Penyu', 'kota Marudu', 'kota Kinabalu', 'Kota Belud', 'Keningau', 'Inanam', 'Bongawan', 'Beluran', 'Beaufort', 'Cukai', 'Kampong Raja', 'Al Muktafi B S', 'Kertih', 'Paka', 'Marang', 'K Terengganu', 'Kuala Besut', 'Kuala Berang', 'Kijal', 'Kemasek', 'Kamaman', 'Permaisuri', 'Jertih', 'Dungun', 'Bukit Besi', 'Besut', 'Ajil', 'Ulu Yam Baru', 'B Baru Sungai Buluh', 'Cyberjaya', 'Putrajaya', 'Bdr. Baru Bangi', 'Hulu Bernam', 'Upm Serdang', 'Ukm Bangi', 'Telok P. Garang', 'Tanjong Sepat', 'Tanjong Karang', 'Sungai Pelek', 'Sungai Buloh', 'Sungai Besar', 'Ayer TAwar', 'Subang Jaya', 'Shah Alam', 'Sri Kembangan', 'Sepang', 'Semenyih', 'Sekincan', 'Sabak Bernam', 'Rawang', 'Rasa', 'Pulau Lumut', 'Pulau Ketam', 'Pulau Carey', 'Puchong', 'Petaling Jaya', 'Pelabohan Klang', 'Kuala Selangor', 'Kuala Lumpur', 'Kubu Bharu', 'Kerling', 'Klang', 'Kapar', 'Kajang', 'Jeram', 'Jenjarom', 'Hulu langat', 'Dengkil', 'Bukit Rotan', 'Beranang', 'Batu Caves', 'Batu Arang', 'Batu 9 Cheras', 'Batang Kali', 'Btg. Berjuntai', 'Banting', 'Ampang', 'Padang Besar', 'Kuala Perlis', 'Kangar', 'Kaki Bukit', 'Arau', 'Seri Manjung', 'Kuala Kurau', 'Kuala Dipang', 'Trolak', 'Behrang Stesen', 'Siput', 'Ulu Bernam', 'Ulu Kinta', 'Teronoh', 'Trong', 'Temoh', 'Telok Intan', 'Tapah Road', 'Tapah', 'Tanjong Tualang', 'Tg Rambutan', 'Tg Piandang', 'Tanjung Malim', 'Taiping', 'Sungkai', 'Sungai Sumun', 'Sungai Siput', 'Slim River', 'Sitiawan', 'Simpang', 'Semanggol', 'Selekoh', 'Selama', 'Sauk', 'Pusing', 'Pengkalan Hulu', 'Parit Buntar', 'Parit', 'Pantai Remis', 'Pangkor', 'Padang Rengas', 'Menglembu', 'Matang', 'Manong', 'Mambang Diawan', 'Malim Mawar', 'Lumut', 'Lenggong', 'Langkap', 'Lahat', 'Kuala Sepetang', 'Kuala Kangsar', 'Kamunting', 'Kepayang', 'Kampong Gajah', 'Kampar', 'Ipoh', 'Klian Intan', 'Hutan Melintang', 'Gopeng', 'Gerik', 'Enggor', 'Chenderong Balai', 'Chenderiang', 'Chemor', 'Changkat Keruing', 'Bota', 'Bidor', 'Beruas', 'Batu Kurau', 'Batu Gajah', 'Bagan Serai', 'Bagan Datoh', 'Ayer Tawar', 'Sungai Bakap', 'Usm P/Pinang', 'Tasek Gelugor', 'Tanjong Bungah', 'Sungai Jawi', 'Simpang Ampat', 'Pulau Pinang', 'Permatang Pauh', 'Perai', 'Penang Hill', 'Penaga', 'Nibong Tebal', 'Georgetown', 'Gelugor', 'Butterworth', 'Bukit Mertajam', 'Bayan Lepas', 'Batu Ferringhi', 'Balik Pulau', 'Ayer Itam', 'Cameron', 'Rompin', 'Tun Razak', 'Triang', 'Temerloh', 'Tanah Rata', 'Sungai Ruan', 'Sungai Lembing', 'Ringlet', 'Raub', 'Pekan', 'Padang Tengku', 'Muadzam Shah', 'Mentakab', 'Mengkarak', 'Maran', 'Lurah Bilut', 'Lancang', 'Kuantan', 'Kuala Rompin', 'Kuala Lipis', 'Kemayang', 'Karak', 'Jerantut', 'Genting', 'Gambang', 'Dong', 'Chenor', 'Bukit Fraser', 'Beserah', 'Bentong', 'Benta', 'Bandar Jengka', 'B Baru Jempol', 'Juasseh', 'Baru Serting', 'Jelebu', 'Peng. Kempas', 'Teluk Kemang', 'Lukut', 'Chuah', 'Bukit Pelandok', 'Titi', 'Tanjong Ipoh', 'Sungai Gadut', 'Simpang Pertang', 'Silau', 'Si Rusa', 'Sri Menanti', 'Seremban', 'Seliau', 'Rembau', 'Rantau', 'Port Dickson', 'Pedas', 'Pasir Panjang', 'Nilai', 'Mantin', 'Lubok China', 'Linggi', 'Lenggeng', 'Labu', 'Kuala Pilah', 'Kuala Kelawang', 'Kota', 'Johol', 'Gemencheh', 'Durian Tipus', 'Batu Kikir', 'Bahau', 'Asahan2', 'Tanjong Keling', 'Tampin', 'Sungai Udang', 'Sungai Rambai', 'Merlimau', 'Melaka', 'Masjid Tanah', 'Lubok china', 'Sungai Baru', 'Jasin', 'Durian Tunggal', 'Bemban', 'Batang Melaka', 'Asahan', 'Alor Gajah', 'Wakaf Baru', 'Tumpat', 'Temangan', 'Tanah Merah', 'Rantau Panjang', 'Pulai Chondong', 'Pasir Puteh', 'Pasir Mas', 'Melor', 'Machang', 'Kuala Krai', 'Kota Bharu', 'Ketereh', 'K Desa Pahlawan', 'Jeli', 'Gua Musang', 'Dabong', 'Cherang Ruku', 'Bachok', 'Ayer Lanas', 'Sg. Bakap', 'Changloon', 'Yan', 'Sungai Petani', 'Sik', 'Serdang', 'Pokok Sena', 'Pendang', 'Padang Serai', 'Merbok', 'Lunas', 'Langkawi', 'Langgar', 'Kupang', 'Kulim', 'Kuala Nerang', 'Kuala Ketil', 'Kuala Kedah', 'Kota Kuala Muda', 'Kodiang', 'Kepala Batas', 'Karangan', 'Jitra', 'Gurun', 'Bedong', 'Bandar Baharu', 'Baling', 'Alor Setar', 'Renggam', 'Gemas Baru', 'Sungai Balang', 'Lenga', 'Bukit Kepong', 'Sagil', 'Yong Peng', 'Ulu Tiram', 'Tangkak', 'Sungai Mati', 'Simpang Renggam', 'Senggarang', 'Senai', 'Semerah', 'Skudai', 'Segamat', 'Rengit', 'Pontian', 'Pengerang', 'Pekan Nenas', 'Pasir Gudang', 'Parit Sulong', 'Parit Raja', 'Parit Jawa', 'Panchor', 'Paloh', 'Pagoh', 'Muar', 'Mersing', 'Masai', 'Layang', 'Labis', 'Kulai', 'Kukup', 'Kota Tinggi', 'Kg Kngan T Dr', 'Kluang', 'Kahang', 'Johor Bahru', 'Jementah', 'Andak', 'Gerisek', 'Gemas 1', 'Patah', 'Endau', 'Chaah', 'Bukit Pasir', 'Gambir', 'Benut', 'Bekok', 'Pahat', 'Anam', 'Bakri', 'Ayer Hitam', 'Ayer Baloi', 'Asahan Johor']
	dedict = {'Asahan Johor': '0101', 'Ayer Baloi': '0102', 'Ayer Hitam': '0103', 'Bakri': '0104', 'Anam': '0105', 'Pahat': '0106', 'Bekok': '0107', 'Benut': '0108', 'Gambir': '0109', 'Bukit Pasir': '0110', 'Chaah': '0111', 'Endau': '0112', 'Patah': '0113', 'Gemas 1': '0114', 'Gerisek': '0115', 'Andak': '0116', 'Jementah': '0117', 'Johor Bahru': '0118', 'Kahang': '0119', 'Kluang': '0120', 'Kg Kngan T Dr': '0121', 'Kota Tinggi': '0122', 'Kukup': '0123', 'Kulai': '0124', 'Labis': '0125', 'Layang': '0126', 'Masai': '0127', 'Mersing': '0128', 'Muar': '0129', 'Pagoh': '0130', 'Paloh': '0131', 'Panchor': '0132', 'Parit Jawa': '0133', 'Parit Raja': '0134', 'Parit Sulong': '0135', 'Pasir Gudang': '0136', 'Pekan Nenas': '0137', 'Pengerang': '0138', 'Pontian': '0139', 'Rengit': '0140', 'Segamat': '0141', 'Skudai': '0142', 'Semerah': '0143', 'Senai': '0144', 'Senggarang': '0145', 'Simpang Renggam': '0146', 'Sungai Mati': '0147', 'Tangkak': '0148', 'Ulu Tiram': '0149', 'Yong Peng': '0150', 'Sagil': '0151', 'Bukit Kepong': '0157', 'Lenga': '0177', 'Sungai Balang': '0191', 'Gemas Baru': '0192', 'Renggam': '0193', 'Alor Setar': '0201', 'Baling': '0202', 'Bandar Baharu': '0203', 'Bedong': '0204', 'Gurun': '0205', 'Jitra': '0206', 'Karangan': '0207', 'Kepala Batas': '0208', 'Kodiang': '0209', 'Kota Kuala Muda': '0210', 'Kuala Kedah': '0211', 'Kuala Ketil': '0212', 'Kuala Nerang': '0213', 'Kulim': '0214', 'Kupang': '0215', 'Langgar': '0216', 'Langkawi': '0217', 'Lunas': '0218', 'Merbok': '0219', 'Padang Serai': '0220', 'Pendang': '0221', 'Pokok Sena': '0222', 'Serdang': '0223', 'Sik': '0224', 'Sungai Petani': '0225', 'Yan': '0226', 'Changloon': '0234', 'Sg. Bakap': '0281', 'Ayer Lanas': '0301', 'Bachok': '0302', 'Cherang Ruku': '0303', 'Dabong': '0304', 'Gua Musang': '0305', 'Jeli': '0306', 'K Desa Pahlawan': '0307', 'Ketereh': '0308', 'Kota Bharu': '0309', 'Kuala Krai': '0310', 'Machang': '0311', 'Melor': '0312', 'Pasir Mas': '0313', 'Pasir Puteh': '0314', 'Pulai Chondong': '0315', 'Rantau Panjang': '0316', 'Tanah Merah': '0317', 'Temangan': '0318', 'Tumpat': '0319', 'Wakaf Baru': '0320', 'Alor Gajah': '0401', 'Asahan': '0402', 'Batang Melaka': '0403', 'Bemban': '0404', 'Durian Tunggal': '0405', 'Jasin': '0407', 'Sungai Baru': '0408', 'Lubok china': '0409', 'Masjid Tanah': '0410', 'Melaka': '0411', 'Merlimau': '0412', 'Sungai Rambai': '0413', 'Sungai Udang': '0414', 'Tampin': '0415', 'Tanjong Keling': '0416', 'Asahan2': '0501', 'Bahau': '0502', 'Batu Kikir': '0504', 'Durian Tipus': '0505', 'Gemencheh': '0507', 'Johol': '0508', 'Kota': '0509', 'Kuala Kelawang': '0510', 'Kuala Pilah': '0511', 'Labu': '0512', 'Lenggeng': '0513', 'Linggi': '0514', 'Lubok China': '0515', 'Mantin': '0516', 'Nilai': '0517', 'Pasir Panjang': '0518', 'Pedas': '0519', 'Port Dickson': '0520', 'Rantau': '0521', 'Rembau': '0522', 'Seliau': '0524', 'Seremban': '0525', 'Sri Menanti': '0526', 'Si Rusa': '0527', 'Silau': '0528', 'Simpang Pertang': '0529', 'Sungai Gadut': '0530', 'Tanjong Ipoh': '0532', 'Titi': '0533', 'Bukit Pelandok': '0534', 'Chuah': '0535', 'Lukut': '0536', 'Teluk Kemang': '0537', 'Peng. Kempas': '0538', 'Jelebu': '0539', 'Baru Serting': '0540', 'Juasseh': '0541', 'B Baru Jempol': '0542', 'Bandar Jengka': '0601', 'Benta': '0602', 'Bentong': '0603', 'Beserah': '0604', 'Bukit Fraser': '0605', 'Chenor': '0606', 'Dong': '0607', 'Gambang': '0608', 'Genting': '0609', 'Jerantut': '0610', 'Karak': '0611', 'Kemayang': '0612', 'Kuala Lipis': '0613', 'Kuala Rompin': '0614', 'Kuantan': '0615', 'Lancang': '0616', 'Lurah Bilut': '0617', 'Maran': '0618', 'Mengkarak': '0619', 'Mentakab': '0620', 'Muadzam Shah': '0622', 'Padang Tengku': '0623', 'Pekan': '0624', 'Raub': '0625', 'Ringlet': '0626', 'Sungai Lembing': '0628', 'Sungai Ruan': '0629', 'Tanah Rata': '0630', 'Temerloh': '0631', 'Triang': '0632', 'Tun Razak': '0633', 'Rompin': '0634', 'Cameron': '0635', 'Ayer Itam': '0701', 'Balik Pulau': '0702', 'Batu Ferringhi': '0703', 'Bayan Lepas': '0704', 'Bukit Mertajam': '0705', 'Butterworth': '0706', 'Gelugor': '0707', 'Georgetown': '0708', 'Nibong Tebal': '0710', 'Penaga': '0711', 'Penang Hill': '0712', 'Perai': '0713', 'Permatang Pauh': '0714', 'Pulau Pinang': '0715', 'Simpang Ampat': '0716', 'Sungai Jawi': '0717', 'Tanjong Bungah': '0718', 'Tasek Gelugor': '0719', 'Usm P/Pinang': '0720', 'Sungai Bakap': '0721', 'Ayer Tawar': '0801', 'Bagan Datoh': '0802', 'Bagan Serai': '0803', 'Batu Gajah': '0804', 'Batu Kurau': '0805', 'Beruas': '0806', 'Bidor': '0807', 'Bota': '0808', 'Changkat Keruing': '0809', 'Chemor': '0810', 'Chenderiang': '0811', 'Chenderong Balai': '0812', 'Enggor': '0813', 'Gerik': '0814', 'Gopeng': '0815', 'Hutan Melintang': '0816', 'Klian Intan': '0817', 'Ipoh': '0818', 'Kampar': '0819', 'Kampong Gajah': '0820', 'Kepayang': '0821', 'Kamunting': '0822', 'Kuala Kangsar': '0823', 'Kuala Sepetang': '0824', 'Lahat': '0825', 'Langkap': '0826', 'Lenggong': '0827', 'Lumut': '0828', 'Malim Mawar': '0829', 'Mambang Diawan': '0830', 'Manong': '0831', 'Matang': '0832', 'Menglembu': '0833', 'Padang Rengas': '0834', 'Pangkor': '0835', 'Pantai Remis': '0836', 'Parit': '0837', 'Parit Buntar': '0838', 'Pengkalan Hulu': '0839', 'Pusing': '0840', 'Sauk': '0841', 'Selama': '0842', 'Selekoh': '0843', 'Semanggol': '0844', 'Simpang': '0845', 'Sitiawan': '0847', 'Slim River': '0848', 'Sungai Siput': '0849', 'Sungai Sumun': '0850', 'Sungkai': '0851', 'Taiping': '0852', 'Tanjung Malim': '0853', 'Tg Piandang': '0854', 'Tg Rambutan': '0855', 'Tanjong Tualang': '0856', 'Tapah': '0857', 'Tapah Road': '0858', 'Telok Intan': '0859', 'Temoh': '0860', 'Trong': '0861', 'Teronoh': '0862', 'Ulu Kinta': '0863', 'Ulu Bernam': '0864', 'Siput': '0865', 'Behrang Stesen': '0866', 'Trolak': '0867', 'Kuala Dipang': '0880', 'Kuala Kurau': '0881', 'Seri Manjung': '0882', 'Arau': '0901', 'Kaki Bukit': '0902', 'Kangar': '0903', 'Kuala Perlis': '0904', 'Padang Besar': '0905', 'Ampang': '1001', 'Banting': '1002', 'Btg. Berjuntai': '1003', 'Batang Kali': '1004', 'Batu 9 Cheras': '1005', 'Batu Arang': '1006', 'Batu Caves': '1007', 'Beranang': '1008', 'Bukit Rotan': '1009', 'Dengkil': '1010', 'Hulu langat': '1011', 'Jenjarom': '1012', 'Jeram': '1013', 'Kajang': '1014', 'Kapar': '1015', 'Klang': '1016', 'Kerling': '1017', 'Kubu Bharu': '1018', 'Kuala Lumpur': '1019', 'Kuala Selangor': '1020', 'Pelabohan Klang': '1023', 'Petaling Jaya': '1024', 'Puchong': '1025', 'Pulau Carey': '1026', 'Pulau Ketam': '1027', 'Pulau Lumut': '1028', 'Rasa': '1030', 'Rawang': '1031', 'Sabak Bernam': '1032', 'Sekincan': '1033', 'Semenyih': '1034', 'Sepang': '1035', 'Sri Kembangan': '1037', 'Shah Alam': '1038', 'Subang Jaya': '1039', 'Ayer TAwar': '1040', 'Sungai Besar': '1041', 'Sungai Buloh': '1042', 'Sungai Pelek': '1043', 'Tanjong Karang': '1044', 'Tanjong Sepat': '1045', 'Telok P. Garang': '1046', 'Ukm Bangi': '1047', 'Upm Serdang': '1048', 'Hulu Bernam': '1049', 'Bdr. Baru Bangi': '1050', 'Putrajaya': '1051', 'Cyberjaya': '1052', 'B Baru Sungai Buluh': '1087', 'Ulu Yam Baru': '1088', 'Ajil': '1101', 'Besut': '1102', 'Bukit Besi': '1103', 'Dungun': '1104', 'Jertih': '1105', 'Permaisuri': '1106', 'Kamaman': '1107', 'Kemasek': '1108', 'Kijal': '1109', 'Kuala Berang': '1110', 'Kuala Besut': '1111', 'K Terengganu': '1112', 'Marang': '1113', 'Paka': '1114', 'Kertih': '1115', 'Al Muktafi B S': '1116', 'Kampong Raja': '1117', 'Cukai': '1118', 'Beaufort': '1201', 'Beluran': '1202', 'Bongawan': '1203', 'Inanam': '1204', 'Keningau': '1205', 'Kota Belud': '1206', 'kota Kinabalu': '1207', 'kota Marudu': '1208', 'Kuala Penyu': '1209', 'Kudat': '1210', 'Kunak': '1211', 'Lahad Datu': '1212', 'Lamag': '1213', 'Likas': '1214', 'Membakut': '1215', 'Menumbuk': '1216', 'Nabawan': '1217', 'Papar': '1218', 'Penampang': '1219', 'Ranau': '1220', 'Sandakan': '1221', 'Semporna': '1222', 'Sipitang': '1223', 'Tambunan': '1224', 'Tamparuli': '1225', 'Tanjung Aru': '1226', 'Tawau': '1227', 'Tenom': '1228', 'Tuaran': '1229', 'K. Kinabatangan': '1230', 'Asajaya': '1301', 'Balingan': '1302', 'Baram': '1303', 'Bau': '1304', 'Bekenu': '1305', 'Belaga': '1306', 'Belawi': '1307', 'Betong': '1308', 'Bintagor': '1309', 'Bintulu': '1310', 'Dalat': '1311', 'Daro': '1312', 'Debak': '1313', 'Engkilili': '1314', 'Julau': '1315', 'Kabong': '1316', 'Kanowit': '1317', 'Kapit': '1318', 'Kota Samarahan': '1319', 'Kuching': '1320', 'Lawas': '1321', 'Limbang': '1322', 'Lingga': '1323', 'Long Lama': '1324', 'Lubok antu': '1325', 'Lundu': '1326', 'Matu': '1327', 'Miri': '1328', 'Mukah': '1329', 'Nanga Medamit': '1330', 'Niah': '1331', 'Pekenu': '1332', 'Pusa': '1333', 'Roban': '1334', 'Saratok': '1335', 'Sarikei': '1336', 'Sebauh': '1337', 'Sebuyau': '1338', 'Serian': '1339', 'Simunjan': '1340', 'Song': '1341', 'Spaoh': '1342', 'Sri Aman': '1343', 'Sundar': '1344', 'Tatau': '1345', 'Tebakang': '1346', 'Sibu': '1347', 'Tebedu': '1349', 'Pakan': '1378', 'Sibuti': '1379', 'Oya': '1380', 'Labuan': '1501'}
	#print(dedict)
	for i in range(0,434):
		if ck.find(delist[i].lower()) != -1:
			return(dedict[delist[i]])
	
#print(getkodbandar(59,7))
def getdemofromxl(rx,cx):
	varr = []
	varr.append(giveICwithnoh(rx,4))
	varr.append(sh.cell_value(rx,5))#nama
	varr.append(sh.cell_value(rx,6))#alamat
	vposkod = str(sh.cell_value(rx,8))
	varr.append(vposkod[:-2])#poskod
	varr.append(getkodnegeri(rx,9))#kodnegeri
	varr.append(getkodbandar(rx,7))#kodbandar
	varr.append(sh.cell_value(rx,10))#telefon
	varr.append(sh.cell_value(rx,10))#tbimbit
	vartar = sh.cell_value(rx,11)
	if type(vartar) == str:
		varr.append(sh.cell_value(rx,11))#tlahir
	elif type(vartar) == float:
		vartar = datetime.datetime(*xlrd.xldate_as_tuple(vartar, book.datemode)) 
		vartar = str(vartar)
		vartar = vartar.split("-")
		#print(vartar[2][])
		newstuff = vartar[2][0:2] + "/" +vartar[1] + "/" + vartar[0]#
		varr.append(newstuff)
	vtlahir = sh.cell_value(rx,11)
	if type(vtlahir) == str:
		xv = vtlahir.split("/")
		varr.append(xv[0])#tarikh
		divo = {1:'Januari',2:'Februari',3:'Mac',4:'April',5:'Mei',6:'Jun',7:'Julai',8:'Ogos',9:'September',10:'Oktober',11:'November',12:'Disember'}
		print(xv[0])
		varr.append(divo[int(xv[1])])#bulan
		varr.append(xv[2])
	elif type(vtlahir) == float:
		vtlahir = datetime.datetime(*xlrd.xldate_as_tuple(vtlahir, book.datemode)) 
		vtlahir = str(vtlahir)
		vtlahira = vtlahir.split("-")
		#print(vtlahira[2][])
		varr.append(vtlahira[2][0:2])#tarikh
		varr.append(vtlahira[1])#bulan
		varr.append(vtlahira[0])#tahun
	#print(vtlahira)

	if sh.cell_value(rx,13) == "Lelaki ":#jantinakod
		varr.append("L")
	elif sh.cell_value(rx,13) == "Perempuan ":
		varr.append("P")
	varr.append("14")#kodpekerjaan
	kodetnikdata = {'Melayu': '01', 'Cina ': '02', 'India': '03', 'Orang Asli Semenanjung': '04', 'Bajau': '05', 'Dusun': '06', 'Kadazan': '07', 'Murut': '08', 'Bumiputra Sabah Lain': '10', 'Melanau': '11', 'Kedayan': '12', 'Iban ': '13', 'Bidayuh': '14', 'Bumiputra Sarawak Lain': '15', 'Lain-lain': '16', 'Bukan Warganegara': '17'}
	vkodetnik = sh.cell_value(rx,16)
	print(vkodetnik)
	varr.append(kodetnikdata[vkodetnik])
	print(varr)
	

for i in range(11,98):
	getdemofromxl(i,4)
	
#getdemofromxl(int(sys.argv[1]),int(sys.argv[2]))
#print(sh.cell_value(int(sys.argv[1]),int(sys.argv[2])))

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