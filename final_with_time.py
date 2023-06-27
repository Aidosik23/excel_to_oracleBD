import win32com.client
import os
from datetime import datetime, timedelta
import tarfile
import schedule
import time

print("Идет процесс:")
# Создаем экземпляр приложения Outlook
outlook = win32com.client.Dispatch('outlook.application')

# Получаем доступ к объекту MAPI
mapi = outlook.GetNamespace("MAPI")

# Выводим названия учетных записей Outlook
for account in mapi.Accounts:
    print(account.DeliveryStore.DisplayName)  # Название учетной записи Outlook

def process_emails():
    # Получаем доступ к папке "Входящие" (Inbox)
    inbox = mapi.GetDefaultFolder(6)  # Папка "Входящие"

    # Получаем доступ к папке внутри "Входящие"
    inbox = inbox.Folders['VN']  # Папка внутри "Входящие" (замените "your folder" на имя нужной папки)

    # Получаем все сообщения в выбранной папке
    messages = inbox.Items

    # Устанавливаем временной диапазон для ограничения поиска писем (последние 24 часа)
    received_dt = datetime.now() - timedelta(days=7)  # Здесь установите желаемый временной диапазон
    received_dt = received_dt.strftime('%m/%d/%Y %H:%M %p')

    # Задаем отправителей и темы писем для фильтрации
    email_senders = ['Василий Наумов', 'fraud@nurtelecom.kg']  # Список отправителей
    email_subjects = ['FW: Subscriptions 06.26.23', 'AVAILABLE MSISDN с привязкой к BAN']  # Список тем писем

    # Применяем ограничения к набору сообщений
    messages = messages.Restrict("[ReceivedTime] >= '"+received_dt+"'")

    # Задаем путь для сохранения вложений в текущую директорию
    outputDir = r'D:\Aidar\Python\outlook_info'  # Если хотим в текущую директорию, оставляем outputDir = os.getcwd()

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
                        print("Процесс закончился")        
                    except Exception as e:
                        print("Ошибка при сохранении или разархивации вложения: " + str(e))                     
    except Exception as e:
        print("Ошибка при обработке писем: " + str(e))


schedule.every().day.at("16:45").do(process_emails)  # Здесь установите желаемое время запуска

# Бесконечный цикл для выполнения заданий по расписанию
while True:
    schedule.run_pending()
    time.sleep(1)
