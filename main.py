import json
import math
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
        'VALORES': math.ceil((row['VALORES'] / count) * 100) / 100,
        'PARCELAS': f'{index}/{count}',
        'PAGAMENTO': row['PAGAMENTO'],
        'DATA VENCIMENTO': expiration_date.strftime('%d/%m/%Y'),
    })


def generate_main_sheets(writer):
    vendas.to_excel(writer, sheet_name='Vendas', index=False)
    gastos_variaveis.to_excel(writer, sheet_name='Gastos variaveis', index=False)
    gastos_fixos.to_excel(writer, sheet_name='Gastos fixos', index=False)


def change_column_size(writer, key, df, column_size):
    for col_num in range(len(df.columns)):
        writer.sheets[key].set_column(col_num, col_num, column_size)


def copy_exists_plan_to_new_file(xls, writer, column_size):
    for sheet_name in xls.sheet_names:
        if sheet_name not in months:  # Excluir as planilhas existentes que não estão em months
            df = xls.parse(sheet_name)
            df.to_excel(writer, sheet_name=sheet_name, index=False)

            change_column_size(writer, sheet_name, df, column_size)


def add_new_plans(writer, column_size):
    for key in months.keys():
        df = pd.DataFrame(months[key])
        df.to_excel(writer, sheet_name=key, index=False)

        change_column_size(writer, key, df, column_size)


def file_exists(file_path):
    return os.path.isfile(file_path)


def generate_excel(months, vendas, gastos_variaveis, gastos_fixos):
    # Tamanho específico que você deseja para a coluna
    column_size = 20

    # months['vendas'] = vendas.to_dict(orient='records')

    file_path = './data/02.xlsx'

    if file_exists(file_path):
        with pd.ExcelFile(file_path) as xls:
            with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
                generate_main_sheets(writer)
                copy_exists_plan_to_new_file(xls, writer, column_size)
                add_new_plans(writer, column_size)
    else:
        with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
            generate_main_sheets(writer)
            add_new_plans(writer, column_size)


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

generate_excel(months, vendas, gastos_variaveis, gastos_fixos)

print()
print(json.dumps(months, indent=2))

print()
print(months.keys())
