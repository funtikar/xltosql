from ftplib import FTP
import sys

#domain name or server ip:
ftp = FTP('ftp.epizy.com')
ftp.login(user='epiz_20939056', passwd = '')
ftp.cwd('/htdocs')

def placeFile():

    filename = sys.argv[1]
    ftp.storbinary('STOR '+filename, open(filename, 'rb'))
    ftp.quit()

def grabFile():

    filename = sys.argv[2]

    localfile = open(filename, 'wb')
    ftp.retrbinary('RETR ' + filename, localfile.write, 1024)

    ftp.quit()
    localfile.close()


if sys.argv[1] != "-git":
	placeFile()
elif sys.argv[1] == "-git":
	grabFile()

