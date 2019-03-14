#Certificate Filter Helper Tool#
=============
Creating virtual environment by using python 3 venv suggested.

Please `pip install openpyxl requests validators` for required libraries before running this tool.

Usage: certfilter.py inputfile.xlsx outputfile.xlsx

process.log in same directory help monitoring

Update evoid file while it changed, may from [https://en.wikipedia.org/wiki/Extended_Validation_Certificate](https://en.wikipedia.org/wiki/Extended_Validation_Certificate) or somewhere else applicable

Input file example: xlsx e.g.
|      host      |  rank  |
| -------------- | ------ |
| www.google.com |  65535 |
| www.hinet.net  |  32768 |
| www.twgate.net |  16384 |
