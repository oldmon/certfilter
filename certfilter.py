#!/usr/bin/env python3
import sys,ssl,socket,urllib.request,validators
from openpyxl import load_workbook
from OpenSSL import crypto
#please `pip install pyopenssl openpyxl validators` before running

#Global variables as settings
GCIS_URL='http://data.gcis.nat.gov.tw/od/data/api/6BBA2268-1367-4B42-9CCA-BC17499EBE8C'
evoidfile='evoid'
timeout=4
socket.setdefaulttimeout(timeout)

#Print Usage
if len(sys.argv)<2:
	print('Usage: certfilter.py inputfile.xlsx')
	sys.exit(1)

#prevent empty extraction of cert element
def xstr(s):
	if s is None:
		return ''
	return str(s)

#simeple way to check if host accept TLS connection
def isSiteAlive(hostname):
	try:
		urllib.request.urlopen('https://'+hostname)
		return True
	except urllib.error.URLError as e:
		print(e.reason)
		return False

#prepare set for EVSSL oids
with open(evoidfile)as f:
	evoid=set(f.read().splitlines())

f.close()

#worksheet preparation
input_file=sys.argv[1]
in_wb=load_workbook(filename=input_file)
in_ws=in_wb[in_wb.sheetnames[0]]

#worker loop
for row in range(1,in_ws.max_row+1):
	sub_c=sub_o=sub_ou=sub_cn=iss_c=iss_o=iss_cn=iss_ou=nota=notb=None
	host=in_ws['A'+str(row)].value
	weight=in_ws['B'+str(row)].value
	if validators.url('https://'+xstr(host)) and isSiteAlive(xstr(host)):
		getcert=ssl.get_server_certificate((host,443))
		workcert=crypto.load_certificate(crypto.FILETYPE_PEM,getcert)
		sub_c=workcert.get_subject().C
		sub_o=workcert.get_subject().O
		sub_ou=workcert.get_subject().OU
		sub_cn=workcert.get_subject().CN
		iss_c=workcert.get_issuer().C
		iss_o=workcert.get_issuer().O
		iss_ou=workcert.get_issuer().OU
		iss_cn=workcert.get_issuer().CN
		notb=workcert.get_notBefore().decode('utf-8')
		nota=workcert.get_notAfter().decode('utf-8')
		print(str(row)+'\t'+xstr(host)+'\t'+xstr(sub_c)+'\t'+xstr(sub_o)+'\t'+xstr(sub_ou)+'\t'+xstr(sub_cn)+'\t'+xstr(iss_c)+'\t'+xstr(iss_o)+'\t'+xstr(iss_ou)+'\t'+xstr(iss_cn)+'\t'+xstr(notb)+'\t'+xstr(nota))
