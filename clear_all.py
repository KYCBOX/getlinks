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

from exchangelib import Credentials, Account, DELEGATE, Configuration

# Чтение логина и пароля из файла
with open('outlook.txt', 'r') as file:
    lines = file.readlines()
    if len(lines) >= 2:
        email_address = lines[0].strip()
        password = lines[1].strip()
    else:
        print("Ошибка: Файл outlook.txt должен содержать логин и пароль в первой и второй строке.")
        exit()

# Инициализация учетных данных и конфигурации аккаунта
credentials = Credentials(email_address, password)
config = Configuration(server='outlook.office365.com', credentials=credentials)

# Создание аккаунта с ручной конфигурацией
account = Account(primary_smtp_address=email_address, config=config, access_type=DELEGATE)

# Получение всех папок в ящике
all_folders = account.root.tree()

# Перебираем все папки
for folder in all_folders:
    print(f"Deleting emails in folder: {folder.name}")

    # Получение всех писем в текущей папке
    emails_to_delete = folder.all().filter(categories__ne='CATEGORY_TO_EXCLUDE')  # Можно добавить фильтры по вашему выбору
    
    # Удаление писем
    for email in emails_to_delete:
        email.delete()

print("Все письма успешно удалены.")
