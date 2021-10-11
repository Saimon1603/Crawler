import openpyxl


workbook = openpyxl.load_workbook('Clientes.xlsx') # Nome da planilha a ser verificada
workbook2 = openpyxl.Workbook() # Ativa uma planilha em branco

sheet1 = workbook.active
sheet2 = workbook2.active
_max=200000
ic=2 #Onde o scraping vai comçar a pegar os dados no caso 2° Linha
lines=[]
files=[]
defs_ln=[]
lines_pattern=['','']
ult_cache_nm=[]
caches_call=[]
lista = []

ic2=2

#Nomes a serem colocados na planilha
sheet2['A1']='Nome'
#sheet2['B1']='Data de Nascimento'
#sheet2['C1']='CPF'
utt=[]

while ic < _max:
  #print 'Process utterences:',str(ic)
  #print workbook.get_sheet_names()
  #w==============
  #campos a serem coletados
  nome=sheet1['A'+str(ic)].value
 # data=sheet1['B'+str(ic)].value
 # cpf=sheet1['C'+str(ic)].value
  if nome== None:
        break
  if nome.replace(' ','')=='':
        break
  if nome in utt:
    nome=nome2
  else:
   utt.append(nome)
  #Onde vai ser colocado os dados e seus respectivos campos
  sheet2['A' + str(ic2)].value = nome
#  sheet2['B' + str(ic2)].value = data
#  sheet2['C' + str(ic2)].value = cpf
  #
  ic += 1
  ic2 += 1

workbook2.save('Planilha nova.xlsx') #salva a nova planilha com os campos desejados
