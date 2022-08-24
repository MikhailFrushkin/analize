import re

import pandas as pd
import csv
import datetime

smena_1 = ['825078', '825116', '825053', '825054', '825055']
smena_2 = ['825065', '825066', '825070', '825079', '825052']


def read(num):
    colums_list = ['Название документа', 'Тип документа', 'Зона выдачи заказа',
                   'Время создания', 'Время завершения', 'Исполнитель']
    operation_list = ['Подбор', 'Отгрузка', 'Внутрискладское перемещение']
    users_list = []
    works_dict = {
        'Подбор': [0, 0],
        'Отгрузка': [0, 0],
        'Внутрискладское перемещение': [0, 0],
        'ПСТ с зала': [0, 0],
        'Приемка': 0

    }
    users_works = []
    users_works1 = []
    users_works2 = []
    excel_data_df = pd.read_excel('pst.xlsx', sheet_name='Лист1', usecols=colums_list)
    excel_data_df.to_csv('pst.csv')

    with open('pst.csv', newline='', encoding='utf-8') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            if row['Исполнитель'] not in users_list and row['Исполнитель'] not in ['717863', '825003', '825101',
                                                                                   '825098',
                                                                                   '825092']:
                users_list.append(row['Исполнитель'])
        users_list = sorted(users_list)
    for user in users_list:
        with open('pst.csv', newline='', encoding='utf-8') as csvfile:
            reader = csv.DictReader(csvfile)
            for row in reader:
                if 'шт' in row['Название документа']:
                    if row['Исполнитель'] == str(user) and row['Тип документа'] in operation_list:
                        works_dict['ПСТ с зала'][0] += 1

                        pattern = ".*,00"
                        match = re.search(pattern, row['Название документа'])
                        result = match[0].split('-')

                        works_dict['ПСТ с зала'][1] += int(result[-1].replace(',00', '').replace(' ', ''))
                elif row['Тип документа'] == 'Приемка' and row['Исполнитель'] == str(user):
                    works_dict['Приемка'] += 1
                else:
                    if row['Исполнитель'] == str(user) and row['Тип документа'] in operation_list:
                        works_dict[row['Тип документа']][0] += 1

                        date_time_str = row['Время создания']
                        date_time_obj = datetime.datetime.strptime(date_time_str, '%Y-%m-%d %H:%M:%S')
                        date_time_str2 = row['Время завершения']
                        date_time_obj2 = datetime.datetime.strptime(date_time_str2, '%Y-%m-%d %H:%M:%S')

                        if date_time_obj.date() == date_time_obj2.date():
                            time = date_time_obj2 - date_time_obj
                            try:
                                works_dict[row['Тип документа']][1] += time.total_seconds()
                            except ZeroDivisionError as ex:
                                works_dict[row['Тип документа']][1] = 0

        users_works.append((user, works_dict))
        user = user[:6]
        if user in smena_1:
            users_works1.append((user, works_dict))
        elif user in smena_2:
            users_works2.append((user, works_dict))

        works_dict = {
            'Подбор': [0, 0],
            'Отгрузка': [0, 0],
            'Внутрискладское перемещение': [0, 0],
            'ПСТ с зала': [0, 0],
            'Приемка': 0
        }
    if num == 0:
        qwe(users_works)
        data = users_works
    elif num == 1:
        qwe(users_works1)
        data = users_works1
    elif num == 2:
        qwe(users_works2)
        data = users_works2
    else:
        data = 0
    return data, num


def qwe(iter):
    for i in iter:
        for key, value in i[1].items():
            try:
                if key != 'ПСТ с зала' and key != 'Приемка':
                    value[1] = round((value[1] / value[0] / 60), 2)
            except ZeroDivisionError as ex:
                value[1] = 0
    return iter


def multiple_replace(target_str, replace_values):
    for i, j in replace_values.items():
        target_str = target_str.replace(i, j)
    return target_str


def save_csv(data):
    print(data)
    with open('result{}.csv'.format(data[1]), 'w', encoding='utf-8') as file:
        file.write('Табельный номер,Тип документа,Кол-во,Ср.время сборки(мин),'
                   'Операция (Учитывается тому кто выдал на выдаче) ,Кол-во,Ср.время сборки(мин),'
                   'Операция,Кол-во,Ср.время работы с переносом(мин),Операция,Кол-во,Кол-во шт.,'
                   'Операция,Кол-во (учитывается кто закрыл приемку)\n')
        for i in data[0]:
            replace_values = {'(': '', ')': '', '{': '', '}': '', '[': '', ']': '', ':': ',', "'": ''}
            replace_values2 = {'Внутрискладскоеперемещение': 'Перенос(склад)', 'ПСТсзала': 'ПСТ(зал)'}
            line = multiple_replace(str(i), replace_values)
            for item in line[6:]:
                try:
                    if int(item) > 0:
                        file.write('{}\n'.format(multiple_replace(line.replace(' ', ''), replace_values2)))
                        break
                except Exception as ex:
                    pass


def save_exsel(data):
    with open('result{}.csv'.format(data[1]), 'w', encoding='utf-8') as file:
        file.write('Табельный номер,Тип документа,Кол-во,Ср.время сборки(мин),'
                   'Операция (Учитывается тому кто выдал на выдаче) ,Кол-во,Ср.время сборки(мин),'
                   'Операция,Кол-во,Ср.время работы с переносом(мин),Операция,Кол-во,Кол-во шт.,'
                   'Операция,Кол-во (учитывается кто закрыл приемку)\n')
        for i in data[0]:
            replace_values = {'(': '', ')': '', '{': '', '}': '', '[': '', ']': '', ':': ',', "'": ''}
            replace_values2 = {'Внутрискладскоеперемещение': 'Перенос(склад)', 'ПСТсзала': 'ПСТ(зал)'}
            line = multiple_replace(str(i), replace_values)
            for item in line[6:]:
                try:
                    if int(item) > 0:
                        file.write('{}\n'.format(multiple_replace(line.replace(' ', ''), replace_values2)))
                        break
                except Exception as ex:
                    pass
    df = pd.read_csv('result{}.csv'.format(data[1]), encoding='utf-8')
    df.to_excel('output{}.xlsx'.format(data[1]), 'Sheet1', index=False)


def main():
    save_exsel(read(0))
    save_exsel(read(1))
    save_exsel(read(2))


if __name__ == '__main__':
    main()


