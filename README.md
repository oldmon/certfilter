#Certificate Filter Helper Tool#
=============
Using python 3 venv suggested
Please `pip install openpyxl requests validators` before running this tool.

Usage: certfilter.py inputfile.xlsx outputfile.xlsx
and produce process.log for monitoring
Update evoid file while it changed, may from https://en.wikipedia.org/wiki/Extended_Validation_Certificate or somewhere else applicable

Input file: xlsx e.g.
|      host      |  rank  |
| -------------- | ------ |
| www.google.com |  65535 |
| www.hinet.net  |  32768 |
