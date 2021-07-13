import openpyxl


workbook = openpyxl.load_workbook('SCC-RoadsAndParking-Variances.xlsx')
workbook2 = openpyxl.Workbook()

sheet1 = workbook.active
sheet2 = workbook2.active
_max=200000
ic=3
lines=[]
files=[]
defs_ln=[]
lines_pattern=['','']
ult_cache_nm=[]
caches_call=[]
lista = []

ic2=2

sheet2['A1']='Question'
sheet2['B1']='Answer'
sheet2['C1']='Action'
sheet2['D1']='Sign'
utt=[]

while ic < _max:
  #print 'Process utterences:',str(ic)
  #print workbook.get_sheet_names()
  #w==============
  ask1=sheet1['B'+str(ic)].value
  sign=sheet1['C'+str(ic)].value
  response=sheet1['G'+str(ic)].value
  action = sheet1['H' + str(ic)].value
  ask2 = sheet1['F' + str(ic)].value
  if ask1== None:
        break
  if ask1.replace(' ','')=='':
        break
  if ask1 in utt:
    ask1=ask2
  else:
   utt.append(ask1)
  sheet2['A' + str(ic2)].value = ask1
  sheet2['B' + str(ic2)].value = response
  sheet2['C' + str(ic2)].value = action
  sheet2['D' + str(ic2)].value = sign
  #
  ic += 1
  ic2 += 1

workbook2.save('SCC-RoadsAndParking-VariancesDONE.xlsx')