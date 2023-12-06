import math
import os

import pandas as pd
from dateutil.relativedelta import relativedelta

file_name = '01.xlsx'
path = './content/drive/MyDrive/CALCULO_MENSAL'
file_path = f'{path}/{file_name}'


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


def insert_by_position_dict(dictionary, index, df, key):
    items = list(dictionary.items())
    items.insert(index, (key, df.to_dict(orient='records')))

    return dict(items)


def generate_main_sheets(months, vendas):
    vendas['DATA ENTREGA'] = vendas['DATA ENTREGA'].dt.strftime('%d/%m/%Y')
    vendas = vendas.drop('ano_mes', axis=1)

    return insert_by_position_dict(months, 0, vendas, 'VENDAS')


def change_column_size(writer, key, df, column_size):
    for col_num in range(len(df.columns)):
        writer.sheets[key].set_column(col_num, col_num, column_size)


def copy_exists_plan_to_new_file(xls, writer, column_size, months):
    for sheet_name in xls.sheet_names:
        if sheet_name.lower() not in map(str.lower, months):  # Excluir as planilhas existentes que não estão em months
            df = xls.parse(sheet_name)
            df.to_excel(writer, sheet_name=sheet_name, index=False)

            change_column_size(writer, sheet_name, df, column_size)


def add_new_plans(writer, column_size, months):
    for key in months.keys():
        df = pd.DataFrame(months[key])
        df.to_excel(writer, sheet_name=key, index=False)

        change_column_size(writer, key, df, column_size)


def file_exists(file_path):
    return os.path.isfile(file_path)


def generate_excel(months):
    column_size = 20

    months = generate_main_sheets(months, vendas)

    file_path = f'{path}/02.xlsx'

    if file_exists(file_path):
        with pd.ExcelFile(file_path) as xls:
            with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
                copy_exists_plan_to_new_file(xls, writer, column_size, months)
                add_new_plans(writer, column_size, months)
    else:
        with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
            add_new_plans(writer, column_size, months)


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

# GERA TOTAL DE CADA MES
for key in months.keys():
    total = sum(item['VALORES'] for item in months[key])

    months[key].append({'VENDAS': '', 'TOTAL': ''})
    months[key].append({'VENDAS': 'TOTAL DO MES', 'VALORES': total})

months = dict(sorted(months.items()))

generate_excel(months)
