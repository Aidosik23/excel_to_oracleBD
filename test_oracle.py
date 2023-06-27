import cx_Oracle
import text

# Установка соединения
connection = cx_Oracle.connect(user=text.user, password=text.password, dsn=text.dsn, encoding="utf-8", nencoding="utf-8")
try:
    # Создание объекта-курсора
    cursor = connection.cursor()

    # Выполнение SQL-запроса для вставки записи
    insert_query = "INSERT INTO binding_ai (ID, Date_b, msisdn, Subscription_date, Activation_date) VALUES (:id, TO_DATE(:date_b, 'YYYY-MM-DD'), :msisdn, :subscription_date, :activation_date)"
    data = [
        {"id": 1, "date_b": "2023-03-26", "msisdn": "996500000014", "subscription_date": 730000003, "activation_date": 123456789},
        {"id": 2, "date_b": "2023-04-01", "msisdn": "996500000015", "subscription_date": 730000004, "activation_date": 987654321},
        # Add more dictionaries for additional rows
    ]
    
    # Вставка записей
    for row in data:
        cursor.execute(insert_query, row)
        print(row)

    # Подтверждение транзакции
    connection.commit()

finally:
    # Закрытие курсора и соединения
    cursor.close()
    connection.close()

