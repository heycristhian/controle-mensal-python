import json
import os

import pandas as pd
from dateutil.relativedelta import relativedelta

file_name = '01.xlsx'
file_path = f'./data/{file_name}'


def read_excel():
    vendas = pd.read_excel(file_path, sheet_name='Plan1')
    gastos_variaveis = pd.read_excel(file_path, sheet_name='Plan2')
    gastos_fixos = pd.read_excel(file_path, sheet_name='Plan3')

    return vendas, gastos_variaveis, gastos_fixos


def add_months(key, expiration_date, row, count, months, index):
    if key not in months.keys():
        months[key] = []

    months[key].append({
        'VENDAS': row['VENDAS'],
        'VALORES': row['VALORES'] / count,
        'PARCELAS':  f'{index}/ {count}',
        'PAGAMENTO': row['PAGAMENTO'],
        'DATA VENCIMENTO': expiration_date.strftime('%d/%m/%Y'),
    })


def generate_excel(months):
    # Verificar se o arquivo já existe
    if os.path.isfile(file_path):
        # Carregar o arquivo Excel existente
        with pd.ExcelFile(file_path) as xls:
            # Criar um Pandas Excel writer usando o XlsxWriter como motor
            with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
                # Copiar as planilhas existentes para o novo arquivo
                for sheet_name in xls.sheet_names:
                    if sheet_name not in months:  # Excluir as planilhas existentes que não estão em months
                        df = xls.parse(sheet_name)
                        df.to_excel(writer, sheet_name=sheet_name, index=False)

                # Adicionar as novas planilhas
                for key in months.keys():
                    df = pd.DataFrame(months[key])
                    df.to_excel(writer, sheet_name=key, index=False)
    else:
        # Se o arquivo não existir, criar um novo
        with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
            for key in months.keys():
                df = pd.DataFrame(months[key])
                df.to_excel(writer, sheet_name=key, index=False)


vendas, gastos_variaveis, gastos_fixos = read_excel()
vendas['ano_mes'] = vendas['DATA ENTREGA'].dt.strftime('%Y-%m')

print(vendas.head())
print()
print(gastos_variaveis.head())
print()
print(gastos_fixos.head())

months = {}

for index, row in vendas.iterrows():
    actual_date = row['DATA ENTREGA']

    days = list(map(int, str(row['DIAS']).split('/')))
    count = len(days)

    for index in range(len(days)):
        expiration_date = row['DATA ENTREGA'] + relativedelta(days=days[index])
        key = expiration_date.strftime('%Y-%m')

        add_months(key, expiration_date, row, count, months, index + 1)

months = dict(sorted(months.items()))

generate_excel(months)

print()
print(json.dumps(months, indent=2))

print()
print(months.keys())
