#!/usr/bin/env python3
import sys,ssl,traceback
from openpyxl import load_workbook
from OpenSSL import SSL,crypto
#please `pip install pyopenssl openpyxl` before running

#Global variables as settings
GCIS_URL='http://data.gcis.nat.gov.tw/od/data/api/6BBA2268-1367-4B42-9CCA-BC17499EBE8C'
evoidfile='evoid'

#Print Usage
if len(sys.argv)<2:
	print('Usage: certfilter.py inputfile.xlsx')
	sys.exit(1)

#prevent empty extraction of cert element
def xstr(s):
	if s is None:
		return ''
	return str(s)

#prepare set for EVSSL oids
with open(evoidfile)as f:
	evoid=set(f.read().splitlines())
f.close()

#worksheet preparation
input_file=sys.argv[1]
in_wb=load_workbook(filename=input_file)
in_ws=in_wb[in_wb.sheetnames[0]]

#worker loop
for row in range(2,in_ws.max_row+1):
	host=in_ws['A'+str(row)].value
	weight=in_ws['B'+str(row)].value
	try:
		getcert=ssl.get_server_certificate((host,443))
	except(ssl.SSLError,TimeoutError,ConnectionRefusedError):
		sub_c=sub_o=sub_ou=sub_cn=iss_c=iss_o=iss_cn=iss_ou=nota=notb=''
		exc_type,exc_value,exc_traceback=sys.exc_info()
		lines=traceback.format_exception(exc_type,exc_value,exc_traceback)
		print(''.join('!! '+line for line in lines))
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
	print(str(row)+'\t'+host+'\t'+xstr(sub_c)+'\t'+xstr(sub_o)+'\t'+xstr(sub_ou)+'\t'+sub_cn+'\t'+iss_c+'\t'+iss_o+'\t'+xstr(iss_ou)+'\t'+iss_cn+'\t'+str(notb)+'\t'+str(nota))
