import cx_Oracle
import pandas as pd
import text 
# Установите соответствующие параметры подключения к базе данных Oracle
user = text.user
password = text.password
dsn = text.dsn

# Подключение к базе данных Oracle
connection = cx_Oracle.connect(user=user, password=password, dsn=dsn, encoding="utf-8", nencoding="utf-8")

# Путь к файлу Excel
excel_file_path = 'nurtelecom_binding_2023-06-22.csv'

# Чтение данных из файла Excel с помощью библиотеки pandas
excel_data = pd.read_csv(excel_file_path)

# Преобразование данных в формат, совместимый с базой данных Oracle
excel_data = excel_data.where((pd.notnull(excel_data)), None)

# Создание курсора для выполнения SQL-запросов
cursor = connection.cursor()

# Имя таблицы, в которую вы хотите загрузить данные
table_name = 'binding_aiii'

# Очистка таблицы перед загрузкой новых данных (опционально)
cursor.execute(f'TRUNCATE TABLE {table_name}')

# Подготовка SQL-запроса для вставки данных в таблицу
sql_query = f'INSERT INTO {table_name} VALUES ({", ".join([":{}".format(i) for i in range(1, len(excel_data.columns) + 1)])})'

# Загрузка данных в базу данных Oracle
cursor.executemany(sql_query, excel_data.values.tolist())

# Фиксация изменений и закрытие соединения с базой данных
connection.commit()
connection.close()



