import os
import json
import time
import datetime
import sys
import pandas as pd
from tkinter import messagebox  
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains

# ==============================================================================
# БЛОК 1: НАСТРОЙКА ПУТЕЙ И ОКРУЖЕНИЯ С УЧЕТОМ СБОРКИ В EXE (PyInstaller)
# ==============================================================================

# Проверяем, запущена ли программа как скомпилированный исполняемый файл (.exe)
if getattr(sys, 'frozen', False):
    # Если приложение «заморожено» с помощью PyInstaller, корнем является папка с main.exe (dist/main)
    PROJECT_ROOT = os.path.dirname(sys.executable)
else:
    # Если скрипт запускается в режиме разработки как обычный файл .py
    PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Конструируем абсолютный путь к целевой рабочей директории "NavAdd"
NavAdd_DIR = os.path.join(PROJECT_ROOT, "NavAdd")
# Абсолютный путь к файлу конфигурации (хранит логины, пароли и параметры мероприятия)
CONFIG_PATH = os.path.join(NavAdd_DIR, "config.json")

# Автоматически создаем папку NavAdd в dist/main, если её ещё нет на диске.
# Это гарантирует, что программа не упадет с ошибкой отсутствия директории при записи логов.
os.makedirs(NavAdd_DIR, exist_ok=True)

# Формируем уникальную временную метку для текущей сессии работы робота
current_time_str = datetime.datetime.now().strftime("%d.%m.%Y_%H-%M-%S")
# Файл для построчной фиксации шагов выполнения и визуального прогресса
steps_filename = os.path.join(NavAdd_DIR, f"шаги_{current_time_str}.txt")
# Файл итогового лога (ФИО — статус обработки)
log_filename = os.path.join(NavAdd_DIR, f"логи_{current_time_str}.txt")


# ==============================================================================
# БЛОК 2: ФУНКЦИЯ БЕЗОПАСНОГО СОХРАНЕНИЯ ДАННЫХ EXCEL
# ==============================================================================
def safe_save_excel(dataframe, filename):
    """
    Пытается перезаписать файл Excel по указанному пути.
    В случае, если файл заблокирован операционной системой (например, открыт пользователем),
    функция не прерывает работу аварийно, а уходит в цикл ожидания, запрашивая действие у пользователя.
    """
    while True:
        try:
            # Попытка стандартного сохранения через pandas
            dataframe.to_excel(filename, index=False)
            break # Если сохранение прошло успешно, выходим из бесконечного цикла
        except PermissionError:
            # Исключение возникает, если файл открыт в Excel или другой программе
            short_name = os.path.basename(filename)
            print(f"\n[ВНИМАНИЕ !!!] Робот не может сохранить файл '{short_name}', так как он ОТКРЫТ в Excel.")
            # Показываем модальное окно Tkinter с предупреждением
            messagebox.showwarning(
                "Файл заблокирован",
                f"Робот не может сохранить изменения в файл '{short_name}', так как он открыт в Excel.\n\n"
                f"Пожалуйста, ЗАКРОЙТЕ этот файл в Excel и нажмите ОК для повторной попытки."
            )
        except Exception as save_err:
            # Обработка непредвиденных критических ошибок файловой системы
            print(f"\n[ОШИБКА СОХРАНЕНИЯ EXCEL]: {save_err}")
            break


# ==============================================================================
# БЛОК 3: УПРАВЛЕНИЕ КОНФИГУРАЦИОННЫМ ФАЙЛОМ JSON
# ==============================================================================
def load_config():
    """Загружает параметры авторизации и мероприятия из файла JSON (если он существует)."""
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_config(**kwargs):
    """Обновляет или создает файл конфигурации, сохраняя переданные именованные аргументы."""
    config = load_config()
    config.update(kwargs) # Объединяем существующие настройки с новыми данными
    
    # Дополнительная превентивная проверка существования папки
    if not os.path.exists(NavAdd_DIR):
        os.makedirs(NavAdd_DIR, exist_ok=True)
        
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        # Пишем JSON с сохранением кириллицы (ensure_ascii=False) и красивыми отступами
        json.dump(config, f, ensure_ascii=False, indent=4)


# ==============================================================================
# БЛОК 4: ОСНОВНАЯ БИЗНЕС-ЛОГИКА АВТОМАТИЗАЦИИ СЕЛЕНИУМА
# ==============================================================================
def run_automation(data, success_callback, error_callback, register_driver_cb):
    """
    Основной воркер автоматизации.
    Принимает:
        data - словарь с логином, паролем, названием и временем мероприятия.
        success_callback - функция для уведомления интерфейса об успешном завершении.
        error_callback - функция для вывода ошибок в UI.
        register_driver_cb - коллбэк для передачи объекта драйвера в главное окно (для ручного закрытия).
    """
    driver = None
    try:
        print("\n[ИНФО] Начинаем работу программы 'Добавление'...")
        
        # --- ОПРЕДЕЛЕНИЕ АБСОЛЮТНОГО ПУТИ К ТАБЛИЦЕ ДАННЫХ LIST.XLSX ---
        # Сначала ищем файл в корневом каталоге проекта (рядом с исполняемым файлом)
        excel_path = os.path.join(PROJECT_ROOT, 'list.xlsx') 
        if not os.path.exists(excel_path):
            # Если там нет, проверяем внутреннюю рабочую подпапку скрипта
            excel_path = os.path.join(NavAdd_DIR, 'list.xlsx') 

        # Если файл так и не найден, прерываем выполнение и генерируем понятное исключение
        if not os.path.exists(excel_path):
            raise FileNotFoundError(
                f"Файл 'list.xlsx' не найден!\n"
                f"Пожалуйста, положите его в корневую папку проекта:\n-> {PROJECT_ROOT}\n"
                f"или в папку первого скрипта:\n-> {NavAdd_DIR}"
            )
            
        print(f"[ИНФО] Успешно найден файл данных по пути: {excel_path}")
        # Читаем Excel-таблицу в оперативную память как DataFrame
        df = pd.read_excel(excel_path)
        
        total_rows = len(df)
        print(f"[ИНФО] Всего строк для обработки в Excel: {total_rows}")

        # Распаковываем управляющие параметры из словаря данных UI
        LOGIN = data['login']
        PASSWORD = data['password']
        EVENT_TEXT = data['event_text']
        EVENT_DATE = data['event_date']
        EVENT_TIME = data['event_time']

        # --- СТАНДАРТИЗАЦИЯ И ИНИЦИАЛИЗАЦИЯ СЛУЖЕБНЫХ СТОЛБЦОВ В ТАБЛИЦЕ ---
        # Если столбца "Обработан" нет, пытаемся переименовать 4-й столбец (индекс 3) или создаем с нуля
        if 'Обработан' not in df.columns:
            if df.shape[1] >= 4:
                df.rename(columns={df.columns[3]: 'Обработан'}, inplace=True)
            else:
                df['Обработан'] = None

        # То же самое проделываем для столбца "Статус" (5-й по счету столбец, индекс 4)
        if 'Статус' not in df.columns:
            if df.shape[1] >= 5:
                df.rename(columns={df.columns[4]: 'Статус'}, inplace=True)
            else:
                df['Статус'] = None

        # Принудительно приводим типы к object, чтобы избежать проблем с записью строковых значений
        df['Обработан'] = df['Обработан'].astype(object)
        df['Статус'] = df['Статус'].astype(object)

        # Динамически определяем индексы/имена колонок, ориентируясь на ключевые слова или порядковый номер
        col_fio = 'ФИО' if 'ФИО' in df.columns else df.columns[0]
        col_date = 'Дата' if 'Дата' in df.columns else (df.columns[1] if len(df.columns) > 1 else df.columns[0])
        col_muni = 'Муниципалитет' if 'Муниципалитет' in df.columns else (df.columns[2] if len(df.columns) > 2 else df.columns[0])

        # --- ЗАПУСК БРАУЗЕРА И ИНИЦИАЛИЗАЦИЯ ВЕБ-ДРАЙВЕРА ---
        print("[ИНФО] Запуск браузера Chrome...")
        driver = webdriver.Chrome()
        driver.maximize_window()
        
        # Регистрируем ссылку на драйвер в главном потоке приложения
        register_driver_cb(driver)
        # Осуществляем переход на целевую страницу админ-панели Навигатора
        driver.get('https://админка02.навигатор.дети/admin/#activity_order') 

        # --- СЕГМЕНТ АВТОМАТИЧЕСКОЙ АВТОРИЗАЦИИ НА САЙТЕ ---
        print("\n[ИНФО] Начинаем автоматическую авторизацию...")
        try:
            # Ожидаем появление поля ввода логина (селектор учитывает как старые id, так и новые placeholder-атрибуты)
            login_input = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="textfield-1598-inputEl"] | //input[@placeholder="E-mail"]'))
            )
            login_input.clear()
            login_input.send_keys(LOGIN)
            print("[УСПЕШНО] Шаг А1: Поле ввода логина заполнено")

            # Ожидаем и заполняем поле пароля
            password_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="textfield-1599-inputEl"] | //input[@placeholder="Пароль"]'))
            )
            password_input.clear()
            password_input.send_keys(PASSWORD)
            print("[УСПЕШНО] Шаг А2: Поле ввода пароля заполнено")

            # Ищем и кликаем кнопку входа
            login_btn = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="button-1605"] | //button[contains(., "ВОЙТИ")] | //*[text()="ВОЙТИ"]'))
            )
            login_btn.click()
            print("[УСПЕШНО] Шаг А3: Нажата кнопка 'ВОЙТИ'")
            
            print("\n[ИНФО] Ожидаем загрузки основного рабочего кабинета...")
            time.sleep(5) # Техническая пауза для рендеринга интерфейса ExtJS
            
            # Контрольная точка: просим пользователя подтвердить, что интерфейс успешно прогрузился
            user_ready = messagebox.askokcancel(
                "Подтверждение авторизации",
                "Убедитесь, что вы успешно зашли и находитесь в разделе 'Заявки на мероприятия'.\n\n"
                "Нажмите 'ОК', чтобы запустить автоматическую обработку списка детей."
            )
            if not user_ready:
                raise InterruptedError("Выполнение отменено пользователем после авторизации.")

        except InterruptedError as ie:
            raise ie
        except Exception as auth_err:
            # Сценарий на случай, если автоматический ввод сломался (например, изменился интерфейс авторизации)
            print(f"\n[КРИТИЧЕСКАЯ ОШИБКА АВТОРИЗАЦИИ] Не удалось войти автоматически: {auth_err}")
            user_manual = messagebox.askokcancel(
                "Требуется ручной вход",
                "Не удалось авторизоваться автоматически.\nПожалуйста, введите данные на сайте вручную.\n\n"
                "Как только зайдете в нужный раздел заявок, нажмите 'ОК' здесь для продолжения."
            )
            if not user_manual:
                raise InterruptedError("Выполнение отменено пользователем.")

        # ==============================================================================
        # БЛОК 5: ОСНОВНОЙ ИТЕРАЦИОННЫЙ ЦИКЛ ОБРАБОТКИ СТРОК ТАБЛИЦЫ EXCEL
        # ==============================================================================
        for index, row in df.iterrows():
            current_row_num = index + 1
            
            # --- РАСЧЕТ И ИНДИКАЦИЯ ПРОГРЕСС-БАРА В КОНСОЛИ И ЛОГАХ ---
            percent = int((current_row_num / total_rows) * 100)
            bar_length = 20
            filled_length = int(bar_length * current_row_num // total_rows)
            bar_str = '█' * filled_length + '░' * (bar_length - filled_length)
            progress_status = f"[{bar_str}] {percent}% ({current_row_num}/{total_rows})"

            # Пишем текущий прогресс в текстовый файл шагов
            try:
                with open(steps_filename, "a", encoding="utf-8") as sf:
                    sf.write(f"[{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Прогресс: {progress_status}\n")
            except Exception:
                pass # Защита от случайных сбоев записи диска (например, занят лог-файл)

            # Получаем очищенное ФИО из текущей строки
            current_child_fio = str(row[col_fio]).strip() if not pd.isna(row[col_fio]) else "Пустая строка"
            
            # Проверяем, обрабатывалась ли данная строка в прошлых сессиях
            val_by_index_3 = str(row.iloc[3]).strip() if len(row) > 3 else ""
            val_by_name = str(row.get('Обработан', '')).strip()

            # Если стоит маркер "да", то просто пропускаем итерацию, экономя время
            if val_by_name.lower() == "да" or val_by_index_3.lower() == "да":
                print(f"--> [ПРОПУСК] {progress_status} Строка №{current_row_num} ({current_child_fio}) пропущена, так как обнаружен флаг 'да'.")
                continue

            # Валидация на пустое имя в ячейке
            first_cell = row[col_fio]
            if pd.isna(first_cell) or str(first_cell).strip() == "" or str(first_cell).strip().lower() == "nan":
                print(f"\n==================================================")
                print(f" Прогресс: {progress_status}")
                print(f" Строка №{current_row_num}: ПРОПУЩЕНА (Обнаружено пустое значение в ФИО)")
                print(f"==================================================")
                
                # Помечаем в DataFrame как обработанный пропуск, чтобы не спотыкаться при перезапуске
                df.at[index, 'Обработан'] = "да"
                df.at[index, 'Статус'] = "пропущен"
                safe_save_excel(df, excel_path)
                
                with open(log_filename, "a", encoding="utf-8") as log_file:
                    log_file.write(f"Строка №{current_row_num} - в списке пустое значение\n")
                continue  

            fio_child = str(first_cell).strip()
            log_status = "пропущен" 
            excel_status = "пропущен" 
            
            # --- ПАРСИНГ И КОРРЕКТИРОВКА ФОРМАТА ДАТЫ РОЖДЕНИЯ ---
            birth_date_val = row[col_date]
            if pd.isna(birth_date_val) or str(birth_date_val).strip() == "" or str(birth_date_val).strip().lower() == "nan":
                birth_date_str = "не указано"
                clean_excel_dob = ""
            else:
                try:
                    # Приводим к единому стандарту "ДД.ММ.ГГГГ" через pandas datetime
                    birth_date_str = pd.to_datetime(birth_date_val, dayfirst=True).strftime("%d.%m.%Y")
                except:
                    # Если формат нестандартный и парсинг упал — оставляем "как есть" в виде строки
                    birth_date_str = str(birth_date_val).strip()
                clean_excel_dob = birth_date_str.strip()
                    
            # Получение данных о муниципалитете
            municipality = "не указано"
            muni_val = row[col_muni]
            if not pd.isna(muni_val) and str(muni_val).strip() != "" and str(muni_val).strip().lower() != "nan":
                municipality = str(muni_val).strip()
            
            print(f"\n==================================================")
            print(f" Прогресс: {progress_status}")
            print(f" Начинаем обработку ребенка №{current_row_num}: {fio_child}")
            print(f" Дата рождения по Excel: {birth_date_str}")
            print(f" Муниципалитет по Excel: {municipality}")
            print(f"==================================================")
            
            try:
                # --- ВЗАИМОДЕЙСТВИЕ С ИНТЕРФЕЙСОМ СОЗДАНИЯ ЗАЯВКИ ---
                create_btn = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@data-qtip="Создать"]'))
                )
                create_btn.click()
                print("[УСПЕШНО] Шаг 1.1: Нажата кнопка 'Создать'")
                
                # Ждем появления модальной формы создания элемента
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '//div[contains(@id, "activity_order-create-form")]'))
                )
                time.sleep(1) # Стабилизационная пауза

                # Находим комбобокс выбора мероприятия
                event_input = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '//div[contains(@id, "activity_order-create-form")]//input[contains(@id, "activitycombo")]'))
                )
                
                # Используем JavaScript injection для мгновенной вставки текста (обход медленного посимвольного send_keys)
                driver.execute_script("arguments[0].value = arguments[1];", event_input, EVENT_TEXT)
                driver.execute_script("arguments[0].dispatchEvent(new Event('input', { bubbles: true }));", event_input)
                time.sleep(0.5)
                
                # Переводим фокус на поле и имитируем нажатие Space+Backspace для триггера внутренних событий JavaScript сайта
                driver.execute_script("arguments[0].focus(); arguments[0].click();", event_input)
                actions = ActionChains(driver)
                actions.send_keys(Keys.SPACE).send_keys(Keys.BACKSPACE).perform()
                
                time.sleep(1.5)  # Ожидаем фильтрацию выпадающего списка на сайте
                actions.send_keys(Keys.ENTER).perform() # Выбираем подставившийся вариант
                print("[УСПЕШНО] Шаг 2.1: Название мероприятия зафиксировано")

                # --- ЗАПОЛНЕНИЕ ПОЛЯ "ДАТА МЕРОПРИЯТИЯ" ---
                date_input = driver.find_element(By.XPATH, '//div[contains(@id, "activity_order-create-form")]//input[contains(@id, "datefield")]')
                driver.execute_script("arguments[0].focus(); arguments[0].click();", date_input)
                time.sleep(0.3)
                
                # Очищаем поле через Ctrl+A -> Delete и вводим дату
                actions = ActionChains(driver)
                actions.key_down(Keys.CONTROL).send_keys("a").key_up(Keys.CONTROL).perform()
                actions.send_keys(Keys.DELETE).perform()
                actions.send_keys(EVENT_DATE).perform()
                time.sleep(0.3)
                actions.send_keys(Keys.ENTER).perform()
                print(f"[УСПЕШНО] Шаг 2.2: Дата ({EVENT_DATE}) заполнена")

                # --- ЗАПОЛНЕНИЕ ПОЛЯ "ВРЕМЯ МЕРОПРИЯТИЯ" ---
                time_input = driver.find_element(By.XPATH, '//div[contains(@id, "activity_order-create-form")]//input[contains(@id, "timefield")]')
                driver.execute_script("arguments[0].focus(); arguments[0].click();", time_input)
                time.sleep(0.3)
                
                actions = ActionChains(driver)
                actions.key_down(Keys.CONTROL).send_keys("a").key_up(Keys.CONTROL).perform()
                actions.send_keys(Keys.DELETE).perform()
                actions.send_keys(EVENT_TIME).perform()
                actions.send_keys(Keys.ENTER).perform()
                print(f"[УСПЕШНО] Шаг 2.3: Время '{EVENT_TIME}' заполнено")

                # --- ПОИСК И ВЫБОР РЕБЕНКА (КРИТИЧЕСКИЙ СЕГМЕНТ) ---
                search_input = driver.find_element(By.XPATH, '//div[contains(@id, "activity_order-create-form")]//input[contains(@id, "kidcombo")]')
                driver.execute_script("arguments[0].value = arguments[1];", search_input, fio_child)
                driver.execute_script("arguments[0].dispatchEvent(new Event('input', { bubbles: true }));", search_input)
                time.sleep(0.5)
                
                driver.execute_script("arguments[0].focus(); arguments[0].click();", search_input)
                actions = ActionChains(driver)
                actions.send_keys(Keys.SPACE).send_keys(Keys.BACKSPACE).perform()
                print(f"[ИНФО] Шаг 3.1: ФИО '{fio_child}' отправлено в поиск. Анализируем варианты...")
                time.sleep(1.5)

                try:
                    # Если даты рождения в таблице нет, робот не сможет сопоставить двойников автоматически
                    if birth_date_str == "не указано":
                        print("[ИНФО] Шаг 4.1: Дата рождения не указана в Excel. Автоматический выбор отключен.")
                        raise Exception("Дата рождения не указана")

                    match_index_to_try = 0
                    while True:
                        # Если это не первая попытка подбора кандидата, сбрасываем и вводим заново для обновления ExtJS-выпадашки
                        if match_index_to_try > 0:
                            print(f"[ИНФО] У текущего ребенка нет статуса 'Подтвержден'. Очищаем поле, вводим ФИО заново и проверяем кандидата №{match_index_to_try + 1}...")
                            search_input = driver.find_element(By.XPATH, '//div[contains(@id, "activity_order-create-form")]//input[contains(@id, "kidcombo")]')
                            driver.execute_script("arguments[0].value = '';", search_input)
                            driver.execute_script("arguments[0].dispatchEvent(new Event('input', { bubbles: true }));", search_input)
                            time.sleep(0.5)
                            
                            driver.execute_script("arguments[0].value = arguments[1];", search_input, fio_child)
                            driver.execute_script("arguments[0].dispatchEvent(new Event('input', { bubbles: true }));", search_input)
                            time.sleep(0.5)
                            
                            driver.execute_script("arguments[0].focus(); arguments[0].click();", search_input)
                            actions = ActionChains(driver)
                            actions.send_keys(Keys.SPACE).send_keys(Keys.BACKSPACE).perform()
                            time.sleep(1.5)

                        # Ищем все элементы списка совпадений в выпадающем окне (picker)
                        picker_items_xpath = "//*[contains(@id, 'kidcombo') and contains(@id, 'picker')]//li[contains(@class, 'x-boundlist-item')]"
                        items = driver.find_elements(By.XPATH, picker_items_xpath)
                        
                        matching_items = []
                        # Очищаем строки от неразрывных пробелов (\xa0) для точного сравнения строк
                        clean_excel_fio = fio_child.lower().replace('\xa0', ' ').strip()
                        
                        fio_matches_count = 0
                        for item in items:
                            item_text = item.text.lower().replace('\xa0', ' ').strip()
                            # Проверяем вхождение имени
                            if clean_excel_fio in item_text or item_text in clean_excel_fio:
                                fio_matches_count += 1
                                
                                dob_matches = True
                                if clean_excel_dob:
                                    # Проверяем совпадение даты рождения в тексте элемента подбора
                                    dob_matches = (clean_excel_dob in item.text)
                                
                                if dob_matches:
                                    matching_items.append(item)
                        
                        # Если индекс перебора превысил количество найденных совпадений — кандидатов больше нет
                        if match_index_to_try >= len(matching_items):
                            print("[ИНФО] Все подходящие кандидаты из списка проверены, подтвержденный ребенок не найден.")
                            raise Exception("Не найдено подтвержденных совпадений среди всех вариантов ФИО + Дата рождения")
                        
                        # Выбираем кандидата по текущему индексу проверки
                        target_element = matching_items[match_index_to_try]
                        driver.execute_script("arguments[0].scrollIntoView(false);", target_element)
                        time.sleep(0.2)
                        
                        try:
                            target_element.click()
                        except Exception:
                            driver.execute_script("arguments[0].click();", target_element)
                        
                        print(f"[ИНФО] Шаг 4.1: Выбран кандидат №{match_index_to_try + 1}. Опрашиваем страницу на статус 'Подтвержден'...")
                        confirmed_found = False
                        
                        # Цикл ожидания подгрузки статуса верификации («Подтвержден») для выбранного ребенка
                        for _ in range(15):
                            try:
                                form_el = driver.find_element(By.XPATH, '//div[contains(@id, "activity_order-create-form")]')
                                form_html_text = form_el.get_attribute('textContent').lower()
                                direct_elements = driver.find_elements(By.XPATH, '//div[contains(@id, "activity_order-create-form")]//*[contains(text(), "Подтвержден") or contains(text(), "Подтверждён")]')
                                
                                if "подтвержден" in form_html_text or "подтверждён" in form_html_text or direct_elements:
                                    confirmed_found = True
                                    break
                            except:
                                pass
                            time.sleep(0.3)
                        
                        if confirmed_found:
                            # Дифференцируем статус для логов, если обнаружены полные тезки
                            if fio_matches_count > 1:
                                log_status = "успешно добавлен среди тёзок"
                                excel_status = "успешно добавлен среди тёзок"
                            else:
                                log_status = "успешно добавлен"
                                excel_status = "успешно добавлен"
                            print(f"[УСПЕШНО] Шаг 4.2: Статус подтверждения подтвержден сайтом.")
                            time.sleep(0.5)
                            break # Выходим из цикла подбора кандидатов, так как нужный ребенок найден и выбран
                        else:
                            print(f"[ОТКЛОНЕНО] У кандидата №{match_index_to_try + 1} отсутствует обязательный статус 'Подтвержден'!")
                            match_index_to_try += 1 # Переходим к проверке следующего тезки
                            
                except Exception:
                    # --- БЛОК ОПЕРАТИВНОГО ИНТЕРФЕЙСНОГО ВМЕШАТЕЛЬСТВА ОПЕРАТОРА ---
                    print(f"\n[ВНИМАНИЕ !!!] Требуется ВМЕШАТЕЛЬСТВО для ребенка '{fio_child}'!")
                    
                    user_skip_choice = messagebox.askyesno(
                        "Требуется вмешательство",
                        f"Робот не смог однозначно выбрать ребенка:\n👉 {fio_child}\nДР: {birth_date_str}\nМуниципалитет: {municipality}\n\n"
                        f"Вы хотите ПРОПУСТИТЬ этого ребенка?\n\n"
                        f"Нажмите 'ДА' — чтобы пропустить и пойти дальше.\n"
                        f"Нажмите 'НЕТ' — если вы выбрали нужного ребенка мышкой на сайте САМИ и хотите продолжить сохранение."
                    )
                    
                    if user_skip_choice:  
                        # Оператор решил пропустить запись
                        print(f"[ИНФО] Пропускаем ребенка {fio_child} по команде пользователя.")
                        log_status = "proпущен"
                        excel_status = "пропущен"
                        
                        df.at[index, 'Обработан'] = "да"
                        df.at[index, 'Статус'] = excel_status
                        safe_save_excel(df, excel_path)
                        
                        with open(log_filename, "a", encoding="utf-8") as log_file:
                            log_file.write(f"{fio_child} - {log_status}\n")
                        try:
                            # Пытаемся корректно закрыть висящую форму на сайте
                            cancel_btn = driver.find_element(By.XPATH, "//div[contains(@id, 'activity_order-create-form')]//*[text()='Закрыть']")
                            cancel_btn.click()
                            time.sleep(1)
                        except Exception:
                            pass
                        continue # Переход к следующему ребенку из DataFrame
                    else:  
                        # Оператор обработал карточку на сайте вручную и разрешил продолжить
                        log_status = "выбран вручную"
                        excel_status = "выбран вручную"
                        
                        time.sleep(0.5) 
                        active_form = driver.find_elements(By.XPATH, '//div[contains(@id, "activity_order-create-form")]')
                        
                        # Проверяем, может оператор сам уже нажал "Сохранить" на форме
                        if not active_form:
                            print(f"[УСПЕШНО] Обнаружено, что форма уже была сохранена вручную.")
                            df.at[index, 'Обработан'] = "да"
                            df.at[index, 'Статус'] = excel_status
                            safe_save_excel(df, excel_path)
                            
                            with open(log_filename, "a", encoding="utf-8") as log_file:
                                log_file.write(f"{fio_child} - {log_status}\n")
                            time.sleep(1.0)
                            continue 

                # --- ШАГ 5: ФИКСАЦИЯ И ИЗМЕНЕНИЕ ДАННЫХ В БАЗЕ САЙТА ---
                save_btn = driver.find_element(By.XPATH, "//div[contains(@id, 'activity_order-create-form')]//*[text()='Сохранить']")
                save_btn.click()
                print("[УСПЕШНО] Шаг 5: Кнопка 'Сохранить' нажата")

                # Синхронизируем статус строки в оперативной памяти и дампим в Excel
                df.at[index, 'Обработан'] = "да"
                df.at[index, 'Статус'] = excel_status
                safe_save_excel(df, excel_path)

                print(f"\n[ГОТОВО] Ребенок {fio_child} успешно обработан.")
                with open(log_filename, "a", encoding="utf-8") as log_file:
                    log_file.write(f"{fio_child} - {log_status}\n")
                    
                time.sleep(1.5) # Пауза перед открытием новой формы создания

            except Exception as e:
                # --- ЛОКАЛЬНЫЙ ОБРАБОТЧИК ОШИБОК ДЛЯ ТЕКУЩЕЙ СТРОКИ ---
                print(f"\n[КРИТИЧЕСКАЯ ОШИБКА] На каком-то этапе произошел сбой: {e}")
                log_status = "пропущен"
                excel_status = "пропущен"
                
                df.at[index, 'Обработан'] = "да"
                df.at[index, 'Статус'] = excel_status
                safe_save_excel(df, excel_path)
                
                with open(log_filename, "a", encoding="utf-8") as log_file:
                    log_file.write(f"{fio_child} - {log_status}\n")
                
                try:
                    # Закрываем аварийную форму, чтобы не ломать логику для последующих записей
                    cancel_btn = driver.find_element(By.XPATH, "//div[contains(@id, 'activity_order-create-form')]//*[text()='Закрыть']")
                    cancel_btn.click()
                    time.sleep(1)
                except Exception:
                    pass

        # --- ЗАВЕРШЕНИЕ ВСЕГО ЦИКЛА ОБРАБОТКИ ---
        print(f"\nУра! Весь список успешно обработан.")
        print(f"Результаты сохранены в лог: {log_filename}")
        success_callback("Программа 'Добавление' успешно обработала весь список детей!")

    except InterruptedError:
        # Исключение генерируется, если пользователь сознательно отменил выполнение в диалоговом окне
        print("[ИНФО] Выполнение прервано пользователем.")
        error_callback("Операция отменена.")
    except Exception as global_err:
        # Глобальный перехват непредвиденных исключений (например, пропал интернет, закрыли браузер)
        print(f"\n[КРИТИЧЕСКАЯ ОШИБКА] {global_err}")
        error_callback(str(global_err))
    finally:
        # Блок гарантирует закрытие инстанса Chrome и освобождение системных ресурсов в любом сценарии завершения
        if driver:
            try:
                driver.quit()
            except Exception:
                pass
            print("[ИНФО] Браузер закрыт.")