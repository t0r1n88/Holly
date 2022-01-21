import pandas as pd
import openpyxl
import math

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
drop_col = [i for i in range(8,153,3)]

# Получаем сырой датафрейм
temp_df = pd.read_excel('data/Календарь профориентационных мероприятий 2022.xlsx',skiprows=1)
# Удаляем столбцы с ссылками
df = temp_df.drop(temp_df.columns[drop_col],axis=1)

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
print(data_dct)
length = df.shape[1]
# Проводим подсчет по местам проведения проб
for row in df.itertuples():
    print(row)
    if str(row[2]) == 'nan':
        continue
    for i in range(7,length,2):

        data_dct[row[1]][row[2]] += check_digit(row[i])

print(data_dct)

# Подсчитываем общую сумму

for poo,places in data_dct.items():
    total = 0
    for place,value in places.items():
        total += value
    data_dct[poo]['Итого'] = total

print(data_dct)

df.to_excel('temp.xlsx',index=False)

