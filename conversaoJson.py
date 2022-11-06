import openpyxl as op
import json
import os

# Carrega o arquivo de planilha
planilha = op.load_workbook('cobranca.xlsx')
aba = planilha.active
t = aba.max_row
a = 2
clientes = {}
while a < t+1:

    nome = aba.cell(row=a, column=1).value
    vecimento = aba.cell(row=a, column=2).value
    vecimento = vecimento.strftime('%d/%m/%Y')
    telefone = aba.cell(row=a, column=3).value
    celular = aba.cell(row=a, column=4).value
    y = {
        'nome': nome,
        'vencimento': vecimento,
        'telefone': telefone,
        'celular': celular
    }
    clientes[a-2] = y
    a += 1

if os.path.exists('cobranca.json'):
    os.remove('cobranca.json')
clientes = json.dumps(clientes, indent=4)
with open('cobranca.json', 'w') as f:
    f.write(clientes)
