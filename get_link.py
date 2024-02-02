try:
    from exchangelib import Credentials, Account, DELEGATE, Configuration
except ImportError:
    print("Библиотека 'exchangelib' не найдена. Попытка установки...")
    try:
        import importlib
        importlib.import_module('exchangelib')
        from exchangelib import Credentials, Account, DELEGATE, Configuration
        print("Библиотека успешно установлена.")
    except ImportError:
        print("Ошибка: Не удалось установить библиотеку 'exchangelib'. Установите её вручную с помощью 'pip install exchangelib'")
        exit()

import re
from exchangelib import Credentials, Account, DELEGATE, Configuration

# Чтение логина и пароля из файла
with open('outlook.txt', 'r') as file:
    lines = file.readlines()
    if len(lines) >= 2:
        email_address = lines[0].strip()
        password = lines[1].strip()
    else:
        print("Ошибка: Файл outlook.txt должен содержать логин в первой и пароль во второй строках.")
        exit()

# Инициализация учетных данных и конфигурации аккаунта
credentials = Credentials(email_address, password)
config = Configuration(server='outlook.office365.com', credentials=credentials)

# Создание аккаунта с ручной конфигурацией
account = Account(primary_smtp_address=email_address, config=config, access_type=DELEGATE)

# Получение всех писем во входящих
inbox = account.inbox.all().order_by('-datetime_received')

# Перебираем все письма
for latest_email in inbox:
    print(f"Subject: {latest_email.subject}")
    print(f"From: {latest_email.sender.email_address}")
    print(f"Received: {latest_email.datetime_received}")
    
    # Извлечение ссылки с использованием регулярного выражения
    body_to_search = latest_email.html_body if hasattr(latest_email, 'html_body') and latest_email.html_body else latest_email.text_body
    if body_to_search:
        match = re.search(r'Confirm your email <([^>]+)>', body_to_search)
        if match:
            confirmation_link = match.group(1)
            print("")
            print("Confirm your email:")
            print("")
            print(f"{confirmation_link}")
            print("")
            print("Файл link.txt со ссылкой успешно создан!")
            print("Вы можете удалить все письма с помощью 'python clean_all.py'")
            print("")
            # Запись ссылки в файл
            with open('link.txt', 'w') as file:
                file.write(confirmation_link)
            
            break  # Если ссылка найдена, завершаем цикл
    else:
        print("Тело письма не содержит текста или HTML.")

# Проверяем, была ли найдена ссылка
if 'confirmation_link' not in locals():
    print("В письмах в папке Inbox ссылка не найдена.")
