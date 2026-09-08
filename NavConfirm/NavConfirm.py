import os
import json
import time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Константы программы
CONFIG_FILE = "config.json"
BASE_URL = "https://xn--02-6kcatyook.xn--80aafey1amqq.xn--d1acj3b/admin/#activity_order"

# --- НАСТРОЙКА ДВОЙНОГО ЛОГИРОВАНИЯ ---
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
LOG_STEPS_FILE = f"LogSteps_{timestamp}.txt"
LOG_KIDS_FILE = f"LogKids_{timestamp}.txt"

def log_step(message):
    """Общий технический лог шагов."""
    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    full_message = f"[{current_time}] {message}"
    print(full_message)
    with open(LOG_STEPS_FILE, "a", encoding="utf-8") as f:
        f.write(full_message + "\n")

def log_kid_success(kid_name):
    """Специализированный лог для ФИО успешно подтвержденных детей."""
    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    entry = f"{kid_name} - успешно подтверждено"
    with open(LOG_KIDS_FILE, "a", encoding="utf-8") as f:
        f.write(f"[{current_time}] {entry}\n")
    log_step(f"🔥 РЕЗУЛЬТАТ: {entry}")

def confirm_action(action_description):
    """Ручное подтверждение ключевых точек."""
    log_step(f"ОЖИДАНИЕ ПОДТВЕРЖДЕНИЯ: {action_description}")
    print(f"\n Нажмите ENTER в консоли, чтобы разрешить: {action_description}...")
    input()
    log_step(f"ДЕЙСТВИЕ РАЗРЕШЕНО: {action_description}")

# --- ШАГ 1: Конфигурация ---
log_step("Старт скрипта. Проверка конфигурационного файла.")
if not os.path.exists(CONFIG_FILE):
    log_step("Файл конфигурации не найден. Запуск интерактивной настройки.")
    print("=" * 60)
    print("  ОБНАРУЖЕН ПЕРВЫЙ ЗАПУСК. НАСТРОЙКА ВСЕХ ПАРАМЕТРОВ  ")
    print("=" * 60)
    
    login = input("1. Введите ваш E-mail (логин): ").strip()
    password = input("2. Введите ваш пароль: ").strip()
    
    print("\n--- Настройка фильтров по умолчанию ---")
    event = input("3. Название мероприятия: ").strip()
    municipality = input("4. Муниципалитет ребёнка (например, 'Караидельский р-н'): ").strip()
    status = input("5. Статус заявок (например, 'Новая'): ").strip()
    
    config_data = {
        "login": login,
        "password": password,
        "event": event if event else "Опрос по результатам деятельности Мобильного технопарка МТ№2",
        "municipality": municipality if municipality else "Караидельский р-н",
        "status": status if status else "Новая"
    }
    
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config_data, f, ensure_ascii=False, indent=4)
        
    log_step(f"Настройки сохранены в '{CONFIG_FILE}'.")

with open(CONFIG_FILE, "r", encoding="utf-8") as f:
    config = json.load(f)
log_step("Конфигурация успешно загружена.")


# --- ШАГ 2: Инициализация браузера ---
log_step("Запуск Chrome...")
driver = webdriver.Chrome()
driver.maximize_window()
log_step(f"Переход на страницу: {BASE_URL}")
driver.get(BASE_URL)

wait = WebDriverWait(driver, 10)


# --- ШАГ 3: Авторизация ---
try:
    log_step("Ввод учетных данных...")
    login_field = wait.until(EC.presence_of_element_located((By.開 if False else By.XPATH, "//input[@placeholder='E-mail']")))
    login_field.clear()
    login_field.send_keys(config["login"])
    
    password_field = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Пароль']")))
    password_field.clear()
    password_field.send_keys(config["password"])
    
    log_step("Клик по кнопке 'ВОЙТИ'...")
    login_button = wait.until(EC.element_to_be_clickable((
        By.XPATH, "//*[contains(text(), 'ВОЙТИ')]/ancestor::button | //*[contains(text(), 'ВОЙТИ')]/ancestor::a | //button[contains(., 'ВОЙТИ')] | //a[contains(., 'ВОЙТИ')]"
    )))
    login_button.click()
    
    time.sleep(4) 
    log_step("Принудительный переход в раздел заявок...")
    driver.get(BASE_URL)
    time.sleep(5)
    log_step("Авторизация пройдена.")
except Exception as auth_error:
    log_step(f"[ОШИБКА] Сбой авторизации: {auth_error}")
    driver.quit()
    exit()


# --- ШАГ 4: Выставление фильтров (РУЧНОЙ ПАУЗЕР НА СТАРТЕ) ---
print("=" * 60)
print("  Разверните окно браузера на весь экран, чтобы были видны поля  ")
print("  Мероприятия, Муниципалитет ребёнка, Статус  ")
print("=" * 60)
confirm_action("Заполнить фильтры таблицы (Мероприятие, Муниципалитет ребёнка, Статус)")

try:
    # 1. Мероприятие
    log_step(f"Фильтр 'Мероприятие' -> '{config['event']}'")
    filter_event = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[contains(@placeholder, 'Мероприятие') or contains(@data-qtip, 'Мероприятие')]")))
    filter_event.click()
    filter_event.clear()
    filter_event.send_keys(config["event"])
    time.sleep(1.5)

    event_dropdown_item = wait.until(EC.element_to_be_clickable((
        By.XPATH, f"//div[contains(@class, 'x-boundlist')]//*[contains(text(), '{config['event']}')] | //li[contains(text(), '{config['event']}')]"
    )))
    event_dropdown_item.click()
    time.sleep(1)

    # 2. Муниципалитет ребёнка
    #log_step(f"Фильтр 'Муниципалитет ребёнка' -> '{config['municipality']}'")
    #filter_municipality = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Муниципалитет ребёнка...' or contains(@placeholder, 'Муниципалитет ребёнка')]")))
    #filter_municipality.click()
    #filter_municipality.clear()
    #filter_municipality.send_keys(config["municipality"])
    #time.sleep(1.5)

    #municipality_dropdown_item = wait.until(EC.element_to_be_clickable((
    #    By.XPATH, f"//div[contains(@class, 'x-boundlist')]//*[contains(text(), '{config['municipality']}')] | //li[contains(text(), '{config['municipality']}')]"
    #)))
    #municipality_dropdown_item.click()
    #time.sleep(1)

    # 3. Статус
    log_step(f"Фильтр 'Статус' -> '{config['status']}'")
    filter_status = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[contains(@placeholder, 'Статус') or @name='status_id-inputEl']")))
    filter_status.click()
    time.sleep(1)

    status_dropdown_item = wait.until(EC.element_to_be_clickable((
        By.XPATH, f"//div[contains(@class, 'x-boundlist')]//*[contains(text(), '{config['status']}')] | //li[contains(text(), '{config['status']}')]"
    )))
    status_dropdown_item.click()

    time.sleep(3)
    log_step("Фильтры успешно выставлены.")
except Exception as filter_error:
    log_step(f"[ПРЕДУПРЕЖДЕНИЕ] Ошибка фильтрации: {filter_error}")


# --- ШАГ 5: Бесконечный цикл обработки ---
is_first_iteration = True

while True:
    try:
        rows = driver.find_elements(By.XPATH, "//tr[contains(@class, 'x-grid-row')]")
        
        if not rows:
            log_step("Все заявки в списке успешно обработаны!")
            break
            
        first_row = rows[0]
        
        # РУЧНОЙ ПАУЗЕР ТОЛЬКО ПЕРЕД САМОЙ ПЕРВОЙ ЗАЯВКОЙ
        if is_first_iteration:
            confirm_action("Запустить автоматический конвейер обработки первой и последующих заявок")
            is_first_iteration = False
        
        # 1. Открываем карточку
        log_step("Открытие карточки заявки (двойной клик)...")
        actions = ActionChains(driver)
        actions.double_click(first_row).perform()
        
        # ЧТЕНИЕ ФИО РЕБЕНКА (С защитой от "Загружаю...")
        log_step("Считывание ФИО ребенка из карточки...")
        try:
            # Находим сам элемент на странице
            kid_element = wait.until(EC.presence_of_element_located((
                By.XPATH, '//*[@id="kid-button-1463-btnInnerEl"] | //*[contains(@id, "kid-button-") and contains(@id, "-btnInnerEl")]'
            )))
            
            # Умное ожидание: ждем, пока текст перестанет быть пустым и в нем ИСЧЕЗНЕТ слово "Загружаю"
            WebDriverWait(driver, 5).until(
                lambda d: kid_element.text.strip() != "" and "Загружаю" not in kid_element.text
            )
            
            current_kid_name = kid_element.text.strip()
            log_step(f"Успешно считано ФИО: {current_kid_name}")
        except Exception as e_kid:
            log_step(f"[ПРЕДУПРЕЖДЕНИЕ] Не удалось дождаться загрузки ФИО ребенка: {e_kid}")
            current_kid_name = "Неизвестный ребенок (ошибка загрузки ФИО)"
        
        # 2. Первое подтверждение (Синяя кнопка)
        log_step("Нажатие синей кнопки 'Подтвердить'...")
        confirm_btn_1 = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Подтвердить']/ancestor::a")))
        confirm_btn_1.click()
        
        # 3. Второе подтверждение (Зеленая кнопка)
        log_step("Подтверждение в модальном окне (зеленая кнопка)...")
        confirm_btn_2 = wait.until(EC.element_to_be_clickable((
            By.XPATH, "//div[contains(@class, 'x-window')]//div[contains(text(), 'Подтверждение заявки')]/ancestor::div[contains(@class, 'x-window')]//span[text()='Подтвердить']/ancestor::a"
        )))
        confirm_btn_2.click()
        
        time.sleep(1.5)
        
        # 4. Закрытие карточки (Красная кнопка)
        log_step("Закрытие карточки (красная кнопка)...")
        close_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Закрыть']/ancestor::a")))
        close_btn.click()
        
        # Запись в лог детей делаем ТОЛЬКО после успешного закрытия карточки
        log_kid_success(current_kid_name)
        
        time.sleep(2)
        
    except Exception as loop_error:
        log_step(f"[КОНЕЦ ЦИКЛА] Выход из процесса обработки. Причина: {loop_error}")
        break

log_step("Скрипт завершил работу.")
driver.quit()
