# import os

# folder_path = 'D:\Aidar\Python\outlook_info'  # Укажите путь к папке, в которой нужно удалить файлы

# # Получаем список файлов в папке
# file_list = os.listdir(folder_path)

# # Проходимся по каждому файлу и удаляем его
# for file_name in file_list:
#     file_path = os.path.join(folder_path, file_name)
#     if os.path.isfile(file_path):
#         os.remove(file_path)
import os
import win32com.client
from datetime import datetime, timedelta
import tarfile

# Удаление файлов в текущей директории
folder_path = 'D:\Aidar\Python\outlook_info'   # Получаем текущую директорию
file_list = os.listdir(folder_path)
for file_name in file_list:
    file_path = os.path.join(folder_path, file_name)
    if os.path.isfile(file_path):
        os.remove(file_path)

print("Файлы в текущей директории удалены")

# Продолжаем выполнение вашего кода
print("Идет процесс:")

outlook = win32com.client.Dispatch('outlook.application')
mapi = outlook.GetNamespace("MAPI")

# Выводим названия учетных записей Outlook
for account in mapi.Accounts:
    print(account.DeliveryStore.DisplayName)  # Название учетной записи Outlook

# Получаем доступ к папке "Входящие" (Inbox)
inbox = mapi.GetDefaultFolder(6)  # Папка "Входящие"

# Получаем доступ к папке внутри "Входящие"
inbox = inbox.Folders['VN']  # Папка внутри "Входящие" (замените "your folder" на имя нужной папки)

# Получаем все сообщения в выбранной папке
messages = inbox.Items 

# Устанавливаем временной диапазон для ограничения поиска писем (последние 24 часа)
received_dt = datetime.now() - timedelta(days=7) #Сюда вводим значения дней по умолчанию 24 часа (1 день)
# received_dt = datetime.now() - timedelta(hours=24) #сюда вводим значение часов 
received_dt = received_dt.strftime('%m/%d/%Y %H:%M %p')

# Задаем отправителей и темы писем для фильтрации
email_senders = ['Василий Наумов', 'fraud@nurtelecom.kg']  # Список отправителей
email_subjects = ['FW: Subscriptions 06.26.23', 'AVAILABLE MSISDN с привязкой к BAN']  # Список тем писем

# Применяем ограничения к набору сообщений
messages = messages.Restrict("[ReceivedTime] >= '"+received_dt+"'")

# Задаем путь для сохранения вложений в текущую директорию
outputDir = r'D:\Aidar\Python\outlook_info' #если хотим в текущую директорию то оставляем outputDir = os.getcwd()

try:
    for message in list(messages):
        # Проверяем соответствие отправителя, темы письма и даты получения
        if message.sender.Name in email_senders or message.SenderEmailAddress in email_senders:
            if message.subject in email_subjects:
                try:
                    # Сохраняем все вложения письма
                    s = message.sender
                    for attachment in message.Attachments:
                        filename = attachment.FileName
                        filepath = os.path.join(outputDir, filename)
                        
                        # Сохраняем вложение
                        attachment.SaveASFile(filepath)
                        print(f"Вложение {filename} от {s} сохранено")
                        
                        # Проверяем, является ли сохраненный файл архивом
                        if filename.endswith(".tar.gz") or filename.endswith(".tar"):
                            tar = tarfile.open(filepath, "r:*")
                            tar.extractall(outputDir)
                            tar.close()
                            print(f"Извлечен архив {filename} от {s}")
                except Exception as e:
                    print("Ошибка при сохранении или разархивации вложения: " + str(e))
except Exception as e:
    print("Ошибка при обработке писем: " + str(e))
