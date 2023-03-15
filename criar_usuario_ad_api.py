from tkinter import Y
import requests
import pandas as pd
from pandas.io.json import json_normalize 
import json
from datetime import date
import openpyxl

from pyad import *
import pandas as pd
from datetime import date
import pyad

from unidecode import unidecode 

URL = "https://cloud.plataforma.senac.br/segserver/api/logins"


myobj = {
"email": "integracaosig@ma.senac.br",
"senha": "Senac1234",
"tipoLogin": "01"
}
response = requests.post(URL, json = myobj)
response_data = response.json()
token = response_data["access_token"]

##########################
token2 = 'Bearer '
token2 += token
##########################

URL2= "https://cloud.plataforma.senac.br/integracaoserver/api/matricula/obter-alunos/"


headers = {
"Authorization": token2
}

myobj = {
    "ApenasTumasEmProcesso": "true",
    "UnidadesOperativasIds": ["115","116", "117","118"] 
}

datareq = "ApenasTumasEmProcesso: true, UnidadesOperativasIds: 115, UnidadesOperativasIds: 116, UnidadesOperativasIds: 117, UnidadesOperativasIds: 118" 

print(myobj)

response = requests.get(URL2, data = myobj, headers = headers )
pd.read_json = response.json()
df = pd.read_json['data']


x=pd.DataFrame(df)
x['data'] = date.today()


####EXPORTANDO PARA EXCEL
x.to_excel('C:/python/bi.xlsx')

file = pd.read_excel('C:/python/bi.xlsx')
file_norepeat = file.drop_duplicates(subset=['cpf'])

result = []

def atualizaListaAD():
    q = pyad.adquery.ADQuery()
    q.execute_query(
    attributes = ["sAMAccountName"],
    where_clause = "objectClass = '*'",
    base_dn = "DC=senaclab, DC=local"
    )
    for row in q.get_results():
        result.append(row["sAMAccountName"])
    return result
resultAtualizado = []
resultAtualizado = atualizaListaAD()

for i in file_norepeat.index:
    cpfExcel = str(file_norepeat['cpf'][i]).rjust(11, '0')
    nomeExcel = str(file_norepeat['nome'][i]).upper()
    senha = nomeExcel.lower().strip().split(' ')[0] + cpfExcel[:5]
    print(cpfExcel)
    print(nomeExcel)
    print(unidecode(senha))
    #resultAtualizado = []
    #resultAtualizado = atualizaListaAD()
    if cpfExcel in resultAtualizado:
        print('Ja existe')
    else:
        ou = pyad.adcontainer.ADContainer.from_dn('ou=Alunos,dc=senaclab,dc=local')
        new_user = pyad.aduser.ADUser.create(cpfExcel, ou, password=unidecode(senha), optional_attributes={"displayName": nomeExcel, "givenName": nomeExcel }) 
        #resultAtualizado = []
        #resultAtualizado = atualizaListaAD()




