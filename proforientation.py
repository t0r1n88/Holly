import pandas as pd
import openpyxl
import math
pd.set_option('display.max_columns', None)
def check_digit(cell):
    """
    Функция для подсчета данных
    :param cell: содержимое ячейки
    :return: число в зависимости от типа
    """
    if type(cell) == str:
        return 0
    # Проверяем на пустую ячейку
    if math.isnan(cell):
        return 0
    if type(cell) == int:
        return cell
    if type(cell) == float:
        return int(cell)
    else:
        return 0


# Создаем список порядковых номеров столбцов которые нужно удалить
drop_col = [i for i in range(8,159,3)]
# Получаем сырой датафрейм
temp_df = pd.read_excel('data/Календарь профориентационных мероприятий 2022.xlsx',skiprows=1)

# Удаляем столбцы с ссылками
df = temp_df.drop(temp_df.columns[drop_col],axis=1)
df.to_excel('check_col.xlsx',index=False)


# Очищаем данные в колонках наименование и место проведения
df['Наименование ПОО'] = df['Наименование ПОО'].apply(lambda x: x.strip() if str(x) !='nan' else x)
df['Место проведения профориентационного мероприятия'] = df['Место проведения профориентационного мероприятия'].apply(lambda x: x.strip() if str(x) !='nan' else x)
# Создаем список ПОО для создания ключей верхнего уровня
temp_lst_poo = df['Наименование ПОО'].unique()
# Очищаем от nan
lst_poo = [x for x in temp_lst_poo if str(x) !='nan']


# Создаем словарь который будет содержать данные по количеству прошедших профробы
data_dct = {poo:{} for poo in lst_poo}

# Итерируемся по таблице чтобы получить места проведения проб
for row in df.itertuples():
    if str(row[1]) and str(row[2]) != 'nan':
        data_dct[row[1]][row[2]] = 0
length = df.shape[1]
# Проводим подсчет по местам проведения проб
for row in df.itertuples():
    if str(row[2]) == 'nan':
        continue
    for i in range(7,length,2):
        data_dct[row[1]][row[2]] += check_digit(row[i])

print(data_dct)

# Подсчитываем  сумму для каждого ПОО и общую сумму

total = 0

for poo,places in data_dct.items():
    total_poo = 0
    for place,value in places.items():
        total_poo += value
    data_dct[poo][f'Итого {poo}'] = total_poo
    print(total_poo)
    total += total_poo

print(data_dct)
print(total)
# Обрабатываем словарь для перевода в читаемую таблицу
first_step_df = pd.DataFrame.from_dict(data_dct,orient='index')
first_step_df.to_excel('temp.xlsx')
# развертываем колонки в строки(индексы)
stack_df = first_step_df.stack()
# превращаем в выходной датафрейм
out_df = stack_df.to_frame().reset_index()
# Переименовываем колонки
out_df.rename(columns={'level_0':'Наименование ПОО','level_1':'Место проведения',0:'Количество школьников посетивших профробы'},inplace=True)
# Добавляем результирующую строку
itog_row = {'Наименование ПОО':'Итого по Республике Бурятия','Место проведения':'','Количество школьников посетивших профробы':total}
out_df = out_df.append(itog_row,ignore_index=True)
print(out_df)
out_df.to_excel('Базовый отчет по профориентации.xlsx',index=False)
#
# # df.to_excel('temp.xlsx',index=False)

