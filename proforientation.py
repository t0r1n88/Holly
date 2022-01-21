import pandas as pd
import openpyxl
import math

pd.set_option('display.max_columns', None)


def processing_note(cell:str,poo,place,date):
    """
    Функция для обработки ячейки с примечаниями в которых находится информация по пробам в формате
    Название пробы,Школа,класс-количество школьников; таких строк может быть несколько в зависимости от количества
    проведенных проб
    :param cell: ячейка с примечаниями
    :param poo: название поо
    :param place место проведения
    :param date: дата проведения мероприятия
    :return: Обновляет значения в словаре
    """
    try:
        # Сплитим по ;
        temp_probs = cell.split(';')
        # Сплитим по запятой
        for prob in temp_probs:
            lst_probs = prob.split(',')
            # Обрабатываем имя пробы чтобы увеличить единообразие
            name_prob = lst_probs[0].strip().capitalize()
            school = lst_probs[1]
            # Проводим еще один сплит по -, чтобы получить количество школьников в классе посетивших пробу
            lst_class = lst_probs[2].split('-')
            # Получаем класс и количество
            school_class = lst_class[0]
            quantity_class = int(lst_class[1])

            # Обновляем значени в словаре, если такая проба уже есть, если нет то создаем такой ключ с начальным значением 0
            if name_prob in probs_dct:
                probs_dct[name_prob] += quantity_class
            else:
                probs_dct[name_prob] = quantity_class



    except:
        with open(f'ERRORS.txt', 'a', encoding='utf-8') as f:
            f.write(f'ПОО {poo} в строке {place} ячейка {cell} номер колонки {date} не обработана!!!\n')


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
drop_col = [i for i in range(8, 159, 3)]
# Получаем сырой датафрейм
temp_df = pd.read_excel('data/Календарь профориентационных мероприятий 2022.xlsx', skiprows=1)

# Удаляем столбцы с ссылками
df = temp_df.drop(temp_df.columns[drop_col], axis=1)
df.to_excel('check_col.xlsx', index=False)

# Очищаем данные в колонках наименование и место проведения
df['Наименование ПОО'] = df['Наименование ПОО'].apply(lambda x: x.strip() if str(x) != 'nan' else x)
df['Место проведения профориентационного мероприятия'] = df['Место проведения профориентационного мероприятия'].apply(
    lambda x: x.strip() if str(x) != 'nan' else x)
# Создаем список ПОО для создания ключей верхнего уровня
temp_lst_poo = df['Наименование ПОО'].unique()
# Очищаем от nan
lst_poo = [x for x in temp_lst_poo if str(x) != 'nan']

# Создаем словарь который будет содержать данные по количеству прошедших профроб для каждого ПОО
data_dct = {poo: {} for poo in lst_poo}

# Создаем словарь для подсчета количества школьников посетивших каждую профпробу в общем
probs_dct = dict()
# Итерируемся по таблице чтобы получить места проведения проб
for row in df.itertuples():
    if str(row[1]) and str(row[2]) != 'nan':
        data_dct[row[1]][row[2]] = 0
length = df.shape[1]
# Проводим подсчет по местам проведения проб
for row in df.itertuples():
    if str(row[2]) == 'nan':
        continue
    for i in range(7, length, 2):
        data_dct[row[1]][row[2]] += check_digit(row[i])
    # Обрабатываем ячейку с примечаниями, где записаны данные по профпробам
    for i in range(8, length + 1, 2):
        processing_note(row[i],row[1],row[2],i)

print(data_dct)

# Подсчитываем  сумму для каждого ПОО и общую сумму

total = 0

for poo, places in data_dct.items():
    total_poo = 0
    for place, value in places.items():
        total_poo += value
    data_dct[poo][f'Итого {poo}'] = total_poo
    print(total_poo)
    total += total_poo

print(probs_dct)
# Обрабатываем словарь для перевода в читаемую таблицу
first_step_df = pd.DataFrame.from_dict(data_dct, orient='index')
# развертываем колонки в строки(индексы)
stack_df = first_step_df.stack()
# превращаем в выходной датафрейм
out_df = stack_df.to_frame().reset_index()
# Переименовываем колонки
out_df.rename(columns={'level_0': 'Наименование ПОО', 'level_1': 'Место проведения',
                       0: 'Количество школьников посетивших профробы'}, inplace=True)
# Добавляем результирующую строку
itog_row = {'Наименование ПОО': 'Итого по Республике Бурятия', 'Место проведения': '',
            'Количество школьников посетивших профробы': total}
out_df = out_df.append(itog_row, ignore_index=True)
out_df.to_excel('Базовый отчет по профориентации.xlsx', index=False)

#Выводим список проб и общее количество посетивших их
prob_df = pd.DataFrame.from_dict(probs_dct,orient='index')
# развертываем колонки в строки(индексы)
stack_df = prob_df.stack()
# превращаем в выходной датафрейм
prob_out_df = stack_df.to_frame().reset_index()
# Переименовываем колонки
prob_out_df.rename(columns={'level_0': 'Наименование пробы',
                       0: 'Количество школьников посетивших пробу'},inplace=True)
# Удаляем колонку
prob_out_df.drop(['level_1'],inplace=True,axis=1)
# Сортируем по убыванию
sorted_prob_out_df = prob_out_df.sort_values(by='Количество школьников посетивших пробу',ascending=False)
# Сохраняем базовый отчет
sorted_prob_out_df.to_excel('Базовый отчет по профпробам.xlsx',index=False)
