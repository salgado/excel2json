# Read excel and transform to json - CSI_Pais_Recife
# Create by alex salgado on July 9, 2015
import xlrd
from collections import OrderedDict
import simplejson as json
import sys

#pega o tipo de relatorio e o arquivo excel
if len(sys.argv) != 3:
	sys.exit("usage: CSI_JSON [report_type] [excel_file]")

rep_type = sys.argv[1]
file_excel = sys.argv[2]

# name the file aregarding to first parameter
if rep_type in 'pais':
	rep_type = 'out_CSI_Pais__Recife__results'
elif rep_type in 'adultos':
	rep_type = 'out_CSI_Adultos_em_Geral__Recife__results'
elif rep_type in '13_17':
	rep_type = 'out_CSI_13_17_anos__Recife__results'
elif rep_type in '8_12':
	rep_type = 'out_CSI_8_12_anos__Recife__results'

# Open the workbook
wb = xlrd.open_workbook( file_excel )

# Get the second sheet named DatabaseAgo6
sh = wb.sheet_by_index(1)

# List to hold CSI_Pais_Recife
lst_CSI_Pais_Recife = []

# Iterate to recover the data
header = sh.row_values(4)

for rownum in range(5, sh.nrows):
	CSI_Pais_Recife = OrderedDict()
	row_values = sh.row_values(rownum)
	if row_values[0] in rep_type:
		# extract each cell value for this row
		for col in range(0,70):
			CSI_Pais_Recife[header[col]] = row_values[col]

		lst_CSI_Pais_Recife.append(CSI_Pais_Recife)


# Serialize the list of dicts to json
j = json.dumps(lst_CSI_Pais_Recife)

# write to file
with open(rep_type[4::] + '.json', 'w') as f:
	f.write(j)
