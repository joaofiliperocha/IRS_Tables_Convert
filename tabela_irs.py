import xlrd
from collections import OrderedDict
import simplejson as json

wb = xlrd.open_workbook('tabelas_ret_irs_2017.xlsx')
sh = wb.sheet_by_index(0)
tabela = []
tab = OrderedDict()
for rownum in range(0, sh.nrows):  
   row_values = sh.row_values(rownum)
   firstColVals = row_values[1].split(' ')
   if len(firstColVals) == 11 and firstColVals[0] == 'TABELAS':      
      #if firstColVals[8] not in tab:
      #tab['local'] = firstColVals[8] 
      localRowNum = rownum
   if rownum - (localRowNum + 4) == 0:
      escaloes = OrderedDict()
      categoria = row_values[1]     
   if row_values[1] == 'Até' or   row_values[1] == 'Superior a' :
       dependentes = OrderedDict();       
       dependentes["0"] = row_values[3]
       dependentes["1"] = row_values[4]
       dependentes["2"] = row_values[5]
       dependentes["3"] = row_values[6]
       dependentes["4"] = row_values[7]
       dependentes["5"] = row_values[8]
       escaloes[ row_values[2]]  = dependentes
       tab[categoria] = escaloes
    
print(tab)
tabela.append(tab)

j= json.dumps(tabela)

with open('data.json', 'w') as f:
   f.write(j)