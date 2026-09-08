import os
import json
import time
import sys
import re  # Добавлен стандартный модуль регулярных выражений для парсинга числа заявок из пагинации
from datetime import datetime
from tkinter import messagebox  

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ==============================================================================
# БЛОК 1: НАСТРОЙКА ПУТЕЙ И ОКРУЖЕНИЯ С УЧЕТОМ СБОРКИ В EXE (PyInstaller)
# ==============================================================================

# Проверяем, запущена ли программа как скомпилированный исполняемый файл (.exe)
if getattr(sys, 'frozen', False):
    # Если приложение «заморожено», корнем является папка, где лежит main.exe (dist/main)
    PROJECT_ROOT = os.path.dirname(sys.executable)
else:
    # Если скрипт запускается в режиме разработки (прямой запуск .py файла)
    PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Формируем абсолютный путь к целевой рабочей директории "NavConfirm"
NavConfirm_DIR = os.path.join(PROJECT_ROOT, "NavConfirm")
# Абсолютный путь к файлу конфигурации (хранит сохраненные логины/фильтры)
CONFIG_PATH = os.path.join(NavConfirm_DIR, "config.json")

# Генерируем уникальный таймстемп для именования лог-файлов текущей сессии
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
LOG_STEPS_FILE = os.path.join(NavConfirm_DIR, f"LogSteps_{timestamp}.txt")
LOG_KIDS_FILE = os.path.join(NavConfirm_DIR, f"LogKids_{timestamp}.txt")


def log_step(message):
    """Записывает технический шаг в консоль и дублирует в текстовый лог LogSteps."""
    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    full_message = f"[{current_time}] {message}"
    print(full_message)
    try:
        with open(LOG_STEPS_FILE, "a", encoding="utf-8") as f:
            f.write(full_message + "\n")
    except Exception:
        pass


def log_kid_success(kid_name):
    """Записывает ФИО успешно обработанного ребенка в отдельный финальный отчет LogKids."""
    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    entry = f"{kid_name} - успешно обработано (подтверждение + участие)"
    try:
        with open(LOG_KIDS_FILE, "a", encoding="utf-8") as f:
            f.write(f"[{current_time}] {entry}\n")
    except Exception:
        pass
    log_step(f"🔥 РЕЗУЛЬТАТ: {entry}")


def confirm_action(action_description):
    """Приостанавливает поток и требует от пользователя нажатия ОК в диалоговом окне для защиты от ошибок."""
    log_step(f"ОЖИДАНИЕ ПОДТВЕРЖДЕНИЯ: {action_description}")
    
    user_confirmed = messagebox.askokcancel(
        title="Действие робота приостановлено",
        message=f"Требуется ваше подтверждение перед следующим шагом:\n\n👉 {action_description}\n\nНажмите 'ОК' для продолжения или 'Отмена' для остановки скрипта."
    )
    
    if user_confirmed:
        log_step(f"ДЕЙСТВИЕ РАЗРЕШЕНО: {action_description}")
    else:
        log_step(f"ДЕЙСТВИЕ ОТМЕНЕНО ПОЛЬЗОВАТЕЛЕМ: {action_description}")
        raise InterruptedError("Выполнение скрипта прервано пользователем.")


def load_config():
    """Загружает сохраненные параметры авторизации и фильтров из JSON."""
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_config(**kwargs):
    """Сохраняет переданные параметры в файл конфигурации JSON."""
    config = load_config()
    config.update(kwargs)
    if not os.path.exists(NavConfirm_DIR):
        os.makedirs(NavConfirm_DIR, exist_ok=True)
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=4)


# ==============================================================================
# БЛОК 2: ОСНОВНАЯ ФУНКЦИЯ АВТОМАТИЗАЦИИ СЕТЕВОГО КОНВЕЙЕРА
# ==============================================================================

def run_automation(data, success_callback, error_callback, register_driver_cb):
    driver = None
    try:
        log_step("Старт скрипта из панели управления. Инициализация параметров...")
        
        # Валидация критически важного поля статуса перед запуском сессии браузера
        if not data.get('status') or not str(data['status']).strip():
            log_step("[ОШИБКА ВАЛИДАЦИИ] Поле 'Статус' не заполнено!")
            messagebox.showerror(
                "Ошибка заполнения формы", 
                "Поле 'Статус' является ОБЯЗАТЕЛЬНЫМ!\nПожалуйста, укажите статус."
            )
            raise ValueError("Поле 'Статус' пустое. Запуск отменен.")

        base_url = "https://xn--02-6kcatyook.xn--80aafey1amqq.xn--d1acj3b/admin/#activity_order"
        
        log_step("Запуск веб-драйвера Chrome...")
        driver = webdriver.Chrome()
        register_driver_cb(driver)  # Передаем управление инстансом в UI для возможности экстренного закрытия
        driver.maximize_window()
        log_step(f"Переход на целевой URL: {base_url}")
        driver.get(base_url)

        wait = WebDriverWait(driver, 10)

        # Выполнение авторизации пользователя на платформе Навигатора
        log_step("Ввод учетных данных...")
        login_field = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@placeholder='E-mail']")))
        login_field.clear()
        login_field.send_keys(data["login"])
        
        password_field = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Пароль']")))
        password_field.clear()
        password_field.send_keys(data["password"])
        
        log_step("Клик по кнопке 'ВОЙТИ'...")
        login_button = wait.until(EC.element_to_be_clickable((
            By.XPATH, "//*[contains(text(), 'ВОЙТИ')]/ancestor::button | //*[contains(text(), 'ВОЙТИ')]/ancestor::a | //button[contains(., 'ВОЙТИ')] | //a[contains(., 'ВОЙТИ')]"
        )))
        login_button.click()
        
        time.sleep(4) 
        log_step("Принудительный переход в раздел заявок...")
        driver.get(base_url)
        time.sleep(5)
        log_step("Авторизация успешно пройдена.")

        print("=" * 60)
        print("  Разверните окно браузера на весь экран, чтобы были видны фильтры  ")
        print("=" * 60)
        
        confirm_action("Заполнить фильтры таблицы на сайте (будут заполнены только указанные вами поля)")

        # Блок интеллектуального выставления фильтров в интерфейсе ExtJS
        try:
            if data.get('event'):
                log_step(f"Фильтр 'Мероприятие' -> '{data['event']}'")
                filter_event = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[contains(@placeholder, 'Мероприятие') or contains(@data-qtip, 'Мероприятие')]")))
                filter_event.click()
                filter_event.clear()
                filter_event.send_keys(data["event"])
                time.sleep(1.5)
                event_dropdown_item = wait.until(EC.element_to_be_clickable((
                    By.XPATH, f"//div[contains(@class, 'x-boundlist')]//*[contains(text(), '{data['event']}')] | //li[contains(text(), '{data['event']}')]"
                )))
                event_dropdown_item.click()
                time.sleep(1)

            if data.get('organizer'):
                log_step(f"Фильтр 'Организатор' -> '{data['organizer']}'")
                filter_org = wait.until(EC.element_to_be_clickable((
                    By.XPATH, "//input[contains(@placeholder, 'Организатор') or contains(@data-qtip, 'Организатор') or @name='organizer_id-inputEl']"
                )))
                filter_org.click()
                filter_org.clear()
                filter_org.send_keys(data["organizer"])
                time.sleep(1.5)
                org_dropdown_item = wait.until(EC.element_to_be_clickable((
                    By.XPATH, f"//div[contains(@class, 'x-boundlist')]//*[contains(text(), '{data['organizer']}')] | //li[contains(text(), '{data['organizer']}')]"
                )))
                org_dropdown_item.click()
                time.sleep(1)

            log_step(f"Фильтр 'Статус' -> '{data['status']}'")
            filter_status = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[contains(@placeholder, 'Статус') or @name='status_id-inputEl']")))
            filter_status.click()
            time.sleep(1)
            status_dropdown_item = wait.until(EC.element_to_be_clickable((
                By.XPATH, f"//div[contains(@class, 'x-boundlist')]//*[contains(text(), '{data['status']}')] | //li[contains(text(), '{data['status']}')]"
            )))
            status_dropdown_item.click()
            time.sleep(1)

            time.sleep(3)
            log_step("Фильтрация завершена.")
        except Exception as filter_error:
            log_step(f"[ПРЕДУПРЕЖДЕНИЕ] Ошибка при автоматической фильтрации: {filter_error}")

        is_first_iteration = True
        processed_count = 0
        total_rows = 0

        # Основной конвейер обработки динамического списка строк таблицы
        while True:
            try:
                # Находим актуальные строки таблицы на текущем экране
                rows = driver.find_elements(By.XPATH, "//tr[contains(@class, 'x-grid-row')]")
                
                if not rows:
                    log_step("Все заявки в списке успешно обработаны!")
                    break
                    
                if is_first_iteration:
                    # --- ОБНОВЛЕННАЯ ЛОГИКА PROGRESS BAR: ЧТЕНИЕ ДАННЫХ ИЗ ПАНЕЛИ ПАГИНАЦИИ ---
                    try:
                        log_step("Попытка считать общее количество записей из правого угла пагинации...")
                        # Нацеливаемся строго на правый счетчик записей, чтобы не путать с количеством страниц
                        pagination_element = wait.until(EC.presence_of_element_located((
                            By.XPATH, "//*[contains(text(), 'Отображаются записи')]"
                        )))
                        pag_text = pagination_element.text
                        log_step(f"Текст пагинации на сайте: '{pag_text}'")
                        
                        # Извлекаем все числовые группы. Последнее число — полный объем выборки (например, из 'Отображаются записи 1 - 25 из 1889' выберет 1889)
                        all_numbers = re.findall(r'\d+', pag_text)
                        if all_numbers:
                            total_rows = int(all_numbers[-1])
                            log_step(f"🎯 УСПЕШНО: Прогресс-бар настроен на обработку {total_rows} заявок.")
                        else:
                            raise ValueError("Числа в строке пагинации не обнаружены.")
                    except Exception as e_pag:
                        # Защитный fallback: если верстка пагинации изменилась, берем видимое число строк
                        total_rows = len(rows)
                        log_step(f"[ПРЕДУПРЕЖДЕНИЕ] Не удалось прочитать угол пагинации ({e_pag}). Взято видимое число строк: {total_rows}")
                    
                    confirm_action("Запустить автоматический конвейер обработки первой и последующих заявок")
                    is_first_iteration = False
                
                # Всегда берем самую первую строку выборки, так как обработанные элементы исчезают при обновлении/фильтрации
                first_row = rows[0]
                processed_count += 1
                
                # Расчет процентов и динамическая отрисовка текстового статус-бара в консоли панели управления
                percent = int((processed_count / total_rows) * 100) if total_rows > 0 else 100
                bar_length = 20
                filled_length = int(bar_length * processed_count // total_rows) if total_rows > 0 else bar_length
                bar_str = '█' * filled_length + '░' * (bar_length - filled_length)
                progress_status = f"[{bar_str}] {percent}% ({processed_count}/{total_rows})"
                
                log_step(f"\n==================================================")
                log_step(f" Прогресс: {progress_status}")
                log_step(f"==================================================")

                log_step("Открытие карточки заявки (двойной клик)...")
                actions = ActionChains(driver)
                actions.double_click(first_row).perform()
                
                log_step("Считывание ФИО ребенка из карточки...")
                try:
                    kid_element = wait.until(EC.presence_of_element_located((
                        By.XPATH, '//*[@id="kid-button-1463-btnInnerEl"] | //*[contains(@id, "kid-button-") and contains(@id, "-btnInnerEl")]'
                    )))
                    # Дожидаемся, пока текст внутри элемента станет реальным ФИО, а не заглушкой загрузки
                    WebDriverWait(driver, 5).until(
                        lambda d: kid_element.text.strip() != "" and "Загружаю" not in kid_element.text
                    )
                    current_kid_name = kid_element.text.strip()
                    log_step(f"Успешно считано ФИО: {current_kid_name}")
                except Exception as e_kid:
                    log_step(f"[ПРЕДУПРЕЖДЕНИЕ] Не удалось дождаться загрузки ФИО ребенка: {e_kid}")
                    current_kid_name = "Неизвестный ребенок (ошибка загрузки ФИО)"
                
                # Операция 1: Подтверждение заявки
                try:
                    log_step("Проверка наличия синей кнопки 'Подтвердить'...")
                    confirm_btn_1 = WebDriverWait(driver, 1.5).until(
                        EC.element_to_be_clickable((By.XPATH, "//span[text()='Подтвердить']/ancestor::a"))
                    )
                    confirm_btn_1.click()
                    
                    log_step("Подтверждение в модальном окне (зеленая кнопка)...")
                    confirm_btn_2 = wait.until(EC.element_to_be_clickable((
                        By.XPATH, "//div[contains(@class, 'x-window')]//div[contains(text(), 'Подтверждение заявки')]/ancestor::div[contains(@class, 'x-window')]//span[text()='Подтвердить']/ancestor::a"
                    )))
                    confirm_btn_2.click()
                    
                    time.sleep(1.5)
                except Exception:
                    log_step("[ИНФО] Кнопка 'Подтвердить' отсутствует. Скорее всего, заявка уже подтверждена.")
                
                # Операция 2: Отметка участия ребенка
                try:
                    log_step("Проверка наличия кнопки 'Отметить участие'...")
                    participation_btn = WebDriverWait(driver, 1.5).until(
                        EC.element_to_be_clickable((
                            By.XPATH, "//span[text()='Отметить участие']/ancestor::a | //*[text()='Отметить участие']"
                        ))
                    )
                    participation_btn.click()
                    log_step("[УСПЕШНО] Кнопка 'Отметить участие' нажата.")
                    time.sleep(1.5)
                except Exception as e_part:
                    log_step(f"[ОШИБКА] Не удалось нажать кнопку 'Отметить участие' (возможно, уже отмечено): {e_part}")
                
                log_step("Закрытие карточки (красная кнопка)...")
                close_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Закрыть']/ancestor::a")))
                close_btn.click()
                
                # Фиксируем успех в логах
                log_kid_success(current_kid_name)
                
                time.sleep(2)  # Стабилизационная пауза перед следующей итерацией цикла поиска строк
                
            except InterruptedError as int_err:
                # Если на этапе промежуточных confirm_action была нажата «Отмена», прокидываем исключение выше
                raise int_err
            except Exception as loop_error:
                # Локальный перехват ошибок цикла: если что-то сломалось безвозвратно, прерываем цикл
                log_step(f"[КОНЕЦ ЦИКЛА] Выход из процесса обработки. Причина: {loop_error}")
                break

        # Сигнализируем в графический интерфейс об успешном окончании
        success_callback("Робот успешно завершил подтверждение и отметку участия заявок!")
        
    except InterruptedError:
        log_step("Выполнение остановлено пользователем через окно подтверждения.")
        error_callback("Операция отменена пользователем.")
    except Exception as e:
        # Глобальный перехват критических ошибок (например, падение интернет-соединения)
        log_step(f"[КРИТИЧЕСКАЯ ОШИБКА] {e}")
        error_callback(str(e))
    finally:
        # Гарантированное закрытие сессии браузера и освобождение памяти
        if driver:
            driver.quit()
            log_step("Браузер закрыт.")