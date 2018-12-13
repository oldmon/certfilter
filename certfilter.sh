#!/bin/bash
input=$1
output=$2
counter=0
evfile='evoid'
while read line;do
	evoid[$counter]=$line
	let counter++
done<$evfile

#for item in ${evoid[*]};do
#	echo $item
#done

counter=0
while read line;do
	certtemp=`tempfile`
	let counter++
	host=`echo $line|awk -F"," '{print $1}'`
	power=`echo $line|awk -F"," '{print $2}'`
	echo "#$counter Try $host ..."
	curl -k -m 3 https://$host > /dev/null 2>&1
	if [ $? -eq 0 ];then
		echo |openssl s_client -quiet -connect $host:443 2>/dev/null
		if [ $? -eq 0 ];then
			echo |openssl s_client -connect $host:443 2>/dev/null|openssl x509 -out $certtemp
			subject=`openssl x509 -noout -in $certtemp -subject|sed 's/subject=//'`
			issuer=`openssl x509 -noout -in $certtemp -issuer|sed 's/issuer=//'`
			startdate=`openssl x509 -noout -in $certtemp -startdate|sed 's/notBefore=//'`
			enddate=`openssl x509 -noout -in $certtemp -enddate|sed 's/notAfter=//'`
			openssl x509 -noout -in $certtemp -text|grep -f evoid >/dev/null
			if [ $? -eq 0 ];then
				EV='Y'
			else
				EV='N'
			fi
			echo "$host,$subject,$issuer,$startdate,$enddate,$EV,$power"
		fi
	fi
done<$input

