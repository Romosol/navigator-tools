import sys
import os
import threading
import customtkinter as ctk
from tkinter import messagebox

# Импорт модулей автоматизации для соответствующих вкладок
import NavAdd.automation as NavAdd
import NavConfirm.automation as NavConfirm

# ==============================================================================
# ГЛОБАЛЬНЫЕ НАСТРОЙКИ ИНТЕРФЕЙСА
# ==============================================================================
ctk.set_appearance_mode("System")  # Автоматическая тема (светлая/темная) в зависимости от ОС
ctk.set_default_color_theme("blue")   # Синяя цветовая схема для элементов управления


# ==============================================================================
# КЛАСС ПЕРЕНАПРАВЛЕНИЯ КОНСОЛЬНОГО ВЫВОДА
# ==============================================================================
class TextRedirector:
    """
    Класс для перехвата стандартных потоков вывода (stdout/stderr) 
    и их дублирования в графическое текстовое поле (консоль интерфейса).
    """
    def __init__(self, textbox, original_stream):
        self.textbox = textbox                  # Ссылка на виджет CTkTextbox
        self.original_stream = original_stream  # Сохраненный системный поток (например, sys.stdout)

    def write(self, message):
        # Вывод сообщения в стандартную системную консоль
        if self.original_stream:
            self.original_stream.write(message)
        # Безопасное добавление текста в графический виджет
        try:
            self.textbox.after(0, self.append_text, message)
        except Exception:
            pass

    def append_text(self, message):
        """Метод вставки текста в виджет консоли с автопрокруткой вниз."""
        try:
            self.textbox.configure(state="normal")    # Разрешаем редактирование текста
            self.textbox.insert("end", message)       # Вставляем текст в самый конец
            self.textbox.see("end")                  # Прокручиваем лог до конца
            self.textbox.configure(state="disabled")  # Блокируем от ручного ввода пользователя
        except Exception:
            pass

    def flush(self):
        # Поддержка стандартного метода сброса буфера потока
        if self.original_stream:
            self.original_stream.flush()


# ==============================================================================
# ОСНОВНОЕ ПРИЛОЖЕНИЕ ГРАФИЧЕСКОГО ИНТЕРФЕЙСА (GUI)
# ==============================================================================
class AutomationApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Настройки главного окна
        self.title("Панель автоматизации")
        self.geometry("650x850")  
        self.resizable(True, True)
        self.minsize(580, 750)

        # Переменные для контроля запущенных браузеров и экстренной остановки
        self.active_driver = None  # Ссылка на текущий запущенный инстанс WebDriver Selenium
        self.is_stopping = False   # Флаг, указывающий на процесс принудительного закрытия

        # Настройка адаптивной сетки (Grid) главного окна
        self.grid_rowconfigure(0, weight=3)     # Область вкладок (занимает больше всего места)
        self.grid_rowconfigure(1, weight=0)     # Кнопка экстренной остановки
        self.grid_rowconfigure(2, weight=0)     # Заголовок консоли
        self.grid_rowconfigure(3, weight=1)     # Текстовое поле консоли
        self.grid_columnconfigure(0, weight=1)  # Растягивание по ширине

        # ----------------------------------------------------------------------
        # КОМПОНЕНТ 1: Панель вкладок (Tabview)
        # ----------------------------------------------------------------------
        self.tabview = ctk.CTkTabview(self)
        self.tabview.grid(row=0, column=0, padx=20, pady=(10, 5), sticky="nsew")
        self.tabview.add("Навигатор (Добавление)")
        self.tabview.add("Навигатор (Подтверждение)")
        self.tabview.add("Программа 3")

        # ----------------------------------------------------------------------
        # КОМПОНЕНТ 2: Кнопка экстренной остановки
        # ----------------------------------------------------------------------
        self.btn_stop = ctk.CTkButton(
            self, text="🛑 ЭКСТРЕННАЯ ОСТАНОВКА РОБОТА",
            command=self.trigger_emergency_stop,
            font=("Arial", 14, "bold"),
            fg_color="#D32F2F", hover_color="#B71C1C",
            height=45
        )
        self.btn_stop.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        self.btn_stop.configure(state="disabled")  # Изначально выключена, пока робот не запущен

        # ----------------------------------------------------------------------
        # КОМПОНЕНТ 3: Заголовок секции логирования
        # ----------------------------------------------------------------------
        self.console_label = ctk.CTkLabel(self, text="Консоль вывода / Логи роботов:", font=("Arial", 12, "bold"))
        self.console_label.grid(row=2, column=0, padx=25, pady=(5, 0), sticky="w")

        # ----------------------------------------------------------------------
        # КОМПОНЕНТ 4: Текстовое поле вывода (Консоль)
        # ----------------------------------------------------------------------
        self.console_box = ctk.CTkTextbox(self, font=("Courier New", 12), fg_color=("#F0F0F0", "#1D1E22"))
        self.console_box.grid(row=3, column=0, padx=20, pady=(0, 20), sticky="nsew")
        self.console_box.configure(state="disabled")

        # Перенаправление стандартных потоков Python в текстовое поле UI
        sys.stdout = TextRedirector(self.console_box, sys.stdout)
        sys.stderr = TextRedirector(self.console_box, sys.stderr)

        # Инициализация интерфейса вкладок и исправление горячих клавиш
        self.setup_tab_navigator()
        self.setup_tab_confirm()
        self.apply_native_hotkeys_fix()

    # ==============================================================================
    # ВСПОМОГАТЕЛЬНЫЕ МЕТОДЫ ИНТЕРФЕЙСА
    # ==============================================================================
    def create_field(self, parent, placeholder):
        """Универсальный метод для генерации полей ввода с плейсхолдером."""
        show_char = "*" if "Пароль" in placeholder else None
        entry = ctk.CTkEntry(parent, placeholder_text=placeholder, width=400, height=35, show=show_char)
        entry.pack(pady=6)
        return entry

    def open_external_file(self, file_path):
        """Кроссплатформенный метод открытия локальных файлов (инструкций, таблиц, логов)."""
        if not os.path.exists(file_path):
            messagebox.showwarning(
                "Файл не найден", 
                f"Файл '{file_path}' ещё не создан на диске.\n"
                f"Пожалуйста, создайте его в соответствующей папке скрипта вручную."
            )
            return
        try:
            if sys.platform == "win32":
                os.startfile(file_path)
            elif sys.platform == "darwin":
                import subprocess
                subprocess.call(["open", file_path])
            else:
                import subprocess
                subprocess.call(["xdg-open", file_path])
        except Exception as e:
            messagebox.showerror("Ошибка открытия", f"Не удалось открыть файл:\n{e}")

    def register_active_driver(self, driver):
        """Колбэк для регистрации текущего WebDriver Selenium с целью управления им из GUI."""
        if self.is_stopping:
            try:
                driver.quit()
            except Exception:
                pass
            self.active_driver = None
        else:
            self.active_driver = driver

    def clear_console(self):
        """Очистка окна консоли перед новым запуском."""
        self.console_box.configure(state="normal")
        self.console_box.delete("1.0", "end")
        self.console_box.configure(state="disabled")

    # ==============================================================================
    # УПРАВЛЕНИЕ ЭКСТРЕННОЙ ОСТАНОВКОЙ
    # ==============================================================================
    def trigger_emergency_stop(self):
        """Обработчик нажатия кнопки экстренной остановки робота."""
        if self.is_stopping:
            return
        
        self.is_stopping = True
        self.btn_stop.configure(state="disabled", text="Остановка процессов...")
        print("\n🛑 [СИСТЕМА] Получена команда экстренной остановки! Убиваю процессы Chrome...")

        # Асинхронное принудительное закрытие сессии браузера
        def force_close_worker():
            if self.active_driver:
                try:
                    self.active_driver.quit()
                except Exception:
                    pass
                self.active_driver = None
            
            # Возвращаем UI в исходное состояние через главный поток
            self.after(0, self.reset_ui_after_stop)

        threading.Thread(target=force_close_worker, daemon=True).start()

    def reset_ui_after_stop(self):
        """Сброс состояния кнопок GUI после экстренной остановки."""
        self.btn_start1.configure(state="normal", text="Сохранить и запустить Добавление")
        self.btn_start2.configure(state="normal", text="Сохранить и запустить Подтверждение")
        self.btn_stop.configure(state="disabled", text="🛑 ЭКСТРЕННАЯ ОСТАНОВКА РОБОТА")
        messagebox.showwarning("Экстренная остановка", "Робот был принудительно остановлен пользователем!")

    # ==============================================================================
    # ВКЛАДКА 1: ДОБАВЛЕНИЕ (МЕРОПРИЯТИЯ)
    # ==============================================================================
    def setup_tab_navigator(self):
        """Инициализация и отрисовка элементов на вкладке 'Добавление'."""
        tab = self.tabview.tab("Навигатор (Добавление)") 
        lbl = ctk.CTkLabel(tab, text="Параметры мероприятия (Добавление)", font=("Arial", 15, "bold"))
        lbl.pack(pady=10)

        # Создание полей ввода параметров
        self.ent_login1 = self.create_field(tab, "Логин (Email):")
        self.ent_password1 = self.create_field(tab, "Пароль:")
        self.ent_event1 = self.create_field(tab, "Название мероприятия (event_text):")
        self.ent_date1 = self.create_field(tab, "Дата (ДД.ММ.ГГГГ):")
        self.ent_time1 = self.create_field(tab, "Время (ЧЧ:ММ):")

        # Фрейм для размещения кнопок управления файлами
        self.frame_files1 = ctk.CTkFrame(tab, fg_color="transparent")
        self.frame_files1.pack(pady=10)

        self.btn_readme1 = ctk.CTkButton(
            self.frame_files1, text="📖 Инструкция",
            command=lambda: self.open_external_file(os.path.join("NavAdd", "readme.txt")),
            font=("Arial", 12, "bold"), fg_color="#1E88E5", hover_color="#1565C0", width=140, height=35
        )
        self.btn_readme1.pack(side="left", padx=5)

        self.btn_open_list = ctk.CTkButton(
            self.frame_files1, text="📊 Открыть list.xlsx",
            command=lambda: self.open_external_file(os.path.join("NavAdd", "list.xlsx")),
            font=("Arial", 12, "bold"), fg_color="#2E7D32", hover_color="#1B5E20", width=140, height=35
        )
        self.btn_open_list.pack(side="left", padx=5)

        self.btn_open_config1 = ctk.CTkButton(
            self.frame_files1, text="⚙️ Открыть config.json",
            command=lambda: self.open_external_file(os.path.join("NavAdd", "config.json")),
            font=("Arial", 12), fg_color="#616161", hover_color="#424242", width=140, height=35
        )
        self.btn_open_config1.pack(side="left", padx=5)

        self.btn_start1 = ctk.CTkButton(
            tab, text="Сохранить и запустить Добавление",
            command=self.start_NavAdd, font=("Arial", 13, "bold"), height=40, width=340
        )
        self.btn_start1.pack(pady=5)

        # Автозаполнение полей сохраненными ранее значениями из конфигурации
        config = NavAdd.load_config()
        if config:
            self.ent_login1.insert(0, config.get("login", ""))
            self.ent_password1.insert(0, config.get("password", ""))
            self.ent_event1.insert(0, config.get("event_text", ""))
            self.ent_date1.insert(0, config.get("event_date", ""))
            self.ent_time1.insert(0, config.get("event_time", ""))

    def start_NavAdd(self):
        """Сбор данных и запуск процесса автоматизации 'Добавление' в отдельном потоке."""
        data = {
            "login": self.ent_login1.get().strip(), "password": self.ent_password1.get().strip(),
            "event_text": self.ent_event1.get().strip(), "event_date": self.ent_date1.get().strip(), "event_time": self.ent_time1.get().strip()
        }
        if not all(data.values()):
            messagebox.showwarning("Внимание", "Пожалуйста, заполните все поля на вкладке Добавления!")
            return

        self.clear_console()
        NavAdd.save_config(**data) # Сохраняем введенные данные в json конфигурации
        
        self.is_stopping = False
        self.btn_start1.configure(state="disabled", text="Выполняется программа 1...")
        self.btn_stop.configure(state="normal")

        # Запуск фонового рабочего потока, чтобы интерфейс приложения не зависал
        threading.Thread(
            target=NavAdd.run_automation,
            args=(data, 
                  lambda msg: self.success_handler(self.btn_start1, "Сохранить и запустить Добавление", msg),
                  lambda err: self.error_handler(self.btn_start1, "Сохранить и запустить Добавление", err),
                  self.register_active_driver),
            daemon=True
        ).start()


    # ==============================================================================
    # ВКЛАДКА 2: ПОДТВЕРЖДЕНИЕ (ЗАЯВОК)
    # ==============================================================================
    def setup_tab_confirm(self):
        """Инициализация и отрисовка элементов на вкладке 'Подтверждение'."""
        tab = self.tabview.tab("Навигатор (Подтверждение)")
        lbl = ctk.CTkLabel(tab, text="Параметры подтверждения заявок", font=("Arial", 15, "bold"))
        lbl.pack(pady=10)

        # Создание полей ввода параметров
        self.ent_login2 = self.create_field(tab, "Логин (Email):")
        self.ent_password2 = self.create_field(tab, "Пароль:")
        self.ent_event2 = self.create_field(tab, "Название мероприятия (event):")
        self.ent_organizer2 = self.create_field(tab, "Организатор (organizer):")
        self.ent_status2 = self.create_field(tab, "Статус заявок (status):")

        # Фрейм для размещения кнопок управления файлами
        self.frame_files2 = ctk.CTkFrame(tab, fg_color="transparent")
        self.frame_files2.pack(pady=10)

        self.btn_readme2 = ctk.CTkButton(
            self.frame_files2, text="📖 Инструкция",
            command=lambda: self.open_external_file(os.path.join("NavConfirm", "readme.txt")),
            font=("Arial", 12, "bold"), fg_color="#1E88E5", hover_color="#1565C0", width=160, height=35
        )
        self.btn_readme2.pack(side="left", padx=6)

        self.btn_open_config2 = ctk.CTkButton(
            self.frame_files2, text="⚙️ Открыть config.json",
            command=lambda: self.open_external_file(os.path.join("NavConfirm", "config.json")),
            font=("Arial", 12), fg_color="#616161", hover_color="#424242", width=160, height=35
        )
        self.btn_open_config2.pack(side="left", padx=6)

        self.btn_start2 = ctk.CTkButton(
            tab, text="Сохранить и запустить Подтверждение",
            command=self.start_NavConfirm, font=("Arial", 13, "bold"), height=40, width=340
        )
        self.btn_start2.pack(pady=5)

        # Восстановление ранее заполненных данных из конфигурации
        config = NavConfirm.load_config()
        if config:
            self.ent_login2.insert(0, config.get("login", ""))
            self.ent_password2.insert(0, config.get("password", ""))
            self.ent_event2.insert(0, config.get("event", ""))
            self.ent_organizer2.insert(0, config.get("organizer", ""))
            self.ent_status2.insert(0, config.get("status", ""))

    def start_NavConfirm(self):
        """Сбор данных и запуск процесса автоматизации 'Подтверждение' в отдельном потоке."""
        data = {
            "login": self.ent_login2.get().strip(), 
            "password": self.ent_password2.get().strip(),
            "event": self.ent_event2.get().strip(), 
            "organizer": self.ent_organizer2.get().strip(), 
            "status": self.ent_status2.get().strip()
        }
        
        # Валидация обязательных полей авторизации
        if not data["login"] or not data["password"]:
            messagebox.showwarning("Внимание", "Пожалуйста, обязательно заполните поля Логин и Пароль!")
            return

        self.clear_console()
        NavConfirm.save_config(**data)
        
        self.is_stopping = False
        self.btn_start2.configure(state="disabled", text="Выполняется программа 2...")
        self.btn_stop.configure(state="normal")

        # Запуск фонового рабочего потока
        threading.Thread(
            target=NavConfirm.run_automation,
            args=(data, 
                  lambda msg: self.success_handler(self.btn_start2, "Сохранить и запустить Подтверждение", msg),
                  lambda err: self.error_handler(self.btn_start2, "Сохранить и запустить Подтверждение", err),
                  self.register_active_driver),
            daemon=True
        ).start()


    # ==============================================================================
    # ХЕНДЛЕРЫ ЗАВЕРШЕНИЯ ПОТОКОВ РОБОТОВ
    # ==============================================================================
    def success_handler(self, button, original_text, message):
        """Обработчик успешного завершения выполнения скрипта робота."""
        button.configure(state="normal", text=original_text)
        self.btn_stop.configure(state="disabled")
        self.active_driver = None
        if self.is_stopping:
            self.is_stopping = False
            return
        messagebox.showinfo("Успех", message)

    def error_handler(self, button, original_text, error_message):
        """Обработчик ошибок, возникших в процессе выполнения скрипта робота."""
        button.configure(state="normal", text=original_text)
        self.btn_stop.configure(state="disabled")
        self.active_driver = None
        
        if self.is_stopping:
            self.is_stopping = False
            return
            
        messagebox.showerror("Ошибка выполнения", f"Скрипт завершился с ошибкой:\n{error_message}")


    # ==============================================================================
    # БЛОК ФИКСА ГОРЯЧИХ КЛАВИШ (ОБХОД СТАНДАРТНОЙ ОШИБКИ TKINTER С КИРИЛЛИЦЕЙ)
    # ==============================================================================
    def apply_native_hotkeys_fix(self):
        """Навешивание кастомных биндингов для корректной работы Ctrl+A / Ctrl+C / Ctrl+V / Ctrl+X."""
        all_entries = [
            self.ent_login1, self.ent_password1, self.ent_event1, self.ent_date1, self.ent_time1,
            self.ent_login2, self.ent_password2, self.ent_event2, self.ent_organizer2, self.ent_status2
        ]
        for ctk_entry in all_entries:
            ctk_entry._entry.bind("<Control-KeyPress>", self.handle_entry_hotkeys)
        self.console_box._textbox.bind("<Control-KeyPress>", self.handle_textbox_hotkeys)

    def handle_entry_hotkeys(self, event):
        """Обработка сочетаний клавиш внутри однострочных полей ввода."""
        key = event.keysym.lower()
        if event.keycode == 65 or key in ['a', 'cyrillic_ef', 'ф']:
            event.widget.select_range(0, "end")
            event.widget.icursor("end")
            return "break"
        elif event.keycode == 86 or key in ['v', 'cyrillic_em', 'м']:
            event.widget.event_generate("<<Paste>>")
            return "break"
        elif event.keycode == 67 or key in ['c', 'cyrillic_es', 'с']:
            event.widget.event_generate("<<Copy>>")
            return "break"
        elif event.keycode == 88 or key in ['x', 'cyrillic_che', 'ч']:
            event.widget.event_generate("<<Cut>>")
            return "break"

    def handle_textbox_hotkeys(self, event):
        """Обработка сочетаний клавиш внутри многострочного текстового поля (консоли)."""
        key = event.keysym.lower()
        if event.keycode == 65 or key in ['a', 'cyrillic_ef', 'ф']:
            event.widget.tag_add("sel", "1.0", "end")
            return "break"
        elif event.keycode == 67 or key in ['c', 'cyrillic_es', 'с']:
            event.widget.event_generate("<<Copy>>")
            return "break"


# ==============================================================================
# ТОЧКА ВХОДА В ПРИЛОЖЕНИЕ
# ==============================================================================
if __name__ == "__main__":
    app = AutomationApp()
    app.mainloop()