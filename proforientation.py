import pandas as pd
import openpyxl

# Создаем список порядковых номеров столбцов которые нужно удалить
drop_col = [i for i in range(8,153,3)]

# Получаем сырой датафрейм
temp_df = pd.read_excel('data/Календарь профориентационных мероприятий 2022.xlsx',skiprows=1)
print(temp_df.shape)
# Удаляем столбцы с ссылками
temp_df.drop(temp_df.columns[drop_col],axis=1,inplace=True)




temp_df.to_excel('temp.xlsx',index=False)

