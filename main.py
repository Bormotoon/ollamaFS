import argparse  # Для обработки аргументов командной строки
import concurrent.futures  # Для многопоточной обработки файлов
import gettext  # Для поддержки многоязычности
import hashlib  # Для вычисления хэшей файлов
import json  # Для работы с JSON-данными
import locale  # Для настройки локализации
import os  # Для взаимодействия с файловой системой
import shutil  # Для операций с файлами и папками
import threading  # Для запуска сортировки в отдельном потоке
import time  # Для добавления временных меток в лог
import tkinter as tk  # Базовый модуль для создания GUI
import zipfile  # Для создания резервных копий в ZIP-формате
from tkinter import filedialog, ttk, messagebox, simpledialog  # Компоненты Tkinter для интерфейса

import PyPDF2  # Для работы с PDF-файлами
import docx  # Для работы с DOCX-файлами
import requests  # Для взаимодействия с API Ollama
from dropbox import Dropbox  # Для интеграции с Dropbox
from dropbox.exceptions import ApiError, AuthError  # Обработка ошибок Dropbox
from dropbox.files import WriteMode  # Режим записи файлов в Dropbox
from google.oauth2 import service_account  # Для аутентификации в Google Drive
from googleapiclient.discovery import build  # Для работы с Google Drive API
from googleapiclient.http import MediaIoBaseDownload  # Для скачивания файлов из Google Drive
from langdetect import detect  # Для определения языка текста
from odf import text, teletype  # Для работы с OpenDocument форматами
from tkinterdnd2 import *  # Поддержка Drag-and-Drop в Tkinter

locale.setlocale(locale.LC_ALL, '')  # Настройка локали для корректного отображения дат и текста


def setup_localization(lang="en"):
    """
    Настраивает локализацию приложения для поддержки нескольких языков.

    Аргументы:
        lang (str): Код языка (по умолчанию "en" для английского).

    Возвращает:
        function: Функция перевода текста (_).
    """
    languages = {'en': 'en_US', 'ru': 'ru_RU'}  # Словарь поддерживаемых языков
    loc = languages.get(lang, 'en_US')  # Выбор локали по коду языка, по умолчанию английский
    translation = gettext.translation('sorter', localedir='locale', languages=[loc], fallback=True)
    # Загрузка переводов из папки 'locale' с запасным вариантом
    translation.install()  # Установка функции перевода в глобальное пространство
    return translation.gettext  # Возвращаем функцию для перевода текста


_ = setup_localization("en")  # Инициализация функции перевода с английским языком по умолчанию


class DocumentSorter:
    """
    Класс для сортировки документов с использованием модели Ollama.

    Атрибуты:
        root (TkinterDnD.Tk): Корневое окно приложения с поддержкой Drag-and-Drop.
        ollama_url (str): URL для подключения к API Ollama.
        model (str): Текущая модель Ollama для классификации.
        available_models (list): Список доступных моделей Ollama.
        category_list (list): Список категорий для сортировки.
        cache (dict): Кэш для хранения результатов анализа файлов.
        language (str): Текущий язык интерфейса.
        google_drive_service (Any): Сервис Google Drive API или None.
        dropbox_client (Dropbox): Клиент Dropbox API или None.
    """

    def __init__(self, root):
        """
        Инициализирует экземпляр класса DocumentSorter.

        Аргументы:
            root (TkinterDnD.Tk): Корневое окно приложения.
        """
        self.root = root  # Сохранение корневого окна
        self.root.title(_("Document Sorter with Ollama"))  # Установка заголовка окна
        self.root.geometry("900x700")  # Установка размеров окна
        self.root.resizable(True, True)  # Разрешение изменения размеров окна

        self.ollama_url = "http://localhost:11434/api"  # URL для API Ollama
        self.model = "deepseek-coder"  # Модель по умолчанию
        self.available_models = []  # Список моделей, доступных в Ollama
        self.category_list = []  # Список категорий для сортировки
        self.cache = {}  # Кэш для ускорения повторной обработки
        self.language = "en"  # Язык интерфейса по умолчанию
        self.google_drive_service = None  # Сервис Google Drive (пока не подключён)
        self.dropbox_client = None  # Клиент Dropbox (пока не подключён)

        self.setup_ui()  # Настройка пользовательского интерфейса
        self.check_ollama_status()  # Проверка состояния Ollama

        # Настройка Drag-and-Drop
        self.root.drop_target_register(DND_FILES)  # Регистрация окна для принятия файлов
        self.root.dnd_bind('<<Drop>>', self.handle_drop)  # Привязка обработчика события перетаскивания

    def setup_ui(self):
        """
        Настраивает пользовательский интерфейс приложения.
        """
        # Создание главного меню
        menubar = tk.Menu(self.root)  # Инициализация меню
        self.root.config(menu=menubar)  # Установка меню в окно
        lang_menu = tk.Menu(menubar, tearoff=0)  # Подменю для выбора языка
        menubar.add_cascade(label=_("Language"), menu=lang_menu)  # Добавление пункта "Язык"
        lang_menu.add_command(label="English", command=lambda: self.change_language("en"))  # Английский
        lang_menu.add_command(label="Русский", command=lambda: self.change_language("ru"))  # Русский

        # Главный фрейм для размещения элементов интерфейса
        main_frame = ttk.Frame(self.root, padding="10")  # Создание фрейма с отступами
        main_frame.pack(fill=tk.BOTH, expand=True)  # Размещение фрейма с заполнением пространства

        # Фрейм статуса Ollama
        status_frame = ttk.Frame(main_frame)  # Создание фрейма для статуса
        status_frame.pack(fill=tk.X, pady=5)  # Размещение фрейма
        ttk.Label(status_frame, text=_("Ollama Status:")).pack(side=tk.LEFT, padx=5)  # Надпись "Статус Ollama"
        self.status_label = ttk.Label(status_frame, text=_("Checking..."), foreground="orange")  # Метка статуса
        self.status_label.pack(side=tk.LEFT, padx=5)  # Размещение метки

        # Фрейм выбора модели
        model_frame = ttk.Frame(main_frame)  # Создание фрейма для выбора модели
        model_frame.pack(fill=tk.X, pady=5)  # Размещение фрейма
        ttk.Label(model_frame, text=_("Select Model:")).pack(side=tk.LEFT, padx=5)  # Надпись "Выберите модель"
        self.model_combobox = ttk.Combobox(model_frame, state="readonly")  # Выпадающий список моделей
        self.model_combobox.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)  # Размещение списка
        self.model_combobox.bind("<<ComboboxSelected>>", self.on_model_selected)  # Привязка обработчика выбора
        refresh_button = ttk.Button(model_frame, text=_("Refresh Models"),
                                    command=self.fetch_models)  # Кнопка обновления
        refresh_button.pack(side=tk.RIGHT, padx=5)  # Размещение кнопки

        # Фрейм для облачных сервисов
        cloud_frame = ttk.LabelFrame(main_frame, text=_("Cloud Storage"), padding="10")  # Фрейм для облачных хранилищ
        cloud_frame.pack(fill=tk.X, pady=10)  # Размещение фрейма
        ttk.Button(cloud_frame, text=_("Connect Google Drive"), command=self.connect_google_drive).pack(side=tk.LEFT,
                                                                                                        padx=5)  # Кнопка Google Drive
        ttk.Button(cloud_frame, text=_("Connect Dropbox"), command=self.connect_dropbox).pack(side=tk.LEFT,
                                                                                              padx=5)  # Кнопка Dropbox

        # Фрейм выбора каталогов
        dir_frame = ttk.LabelFrame(main_frame, text=_("Directory Selection"), padding="10")  # Фрейм для выбора папок
        dir_frame.pack(fill=tk.X, pady=10)  # Размещение фрейма
        ttk.Label(dir_frame, text=_("Source Directory:")).grid(row=0, column=0, sticky=tk.W,
                                                               pady=5)  # Надпись "Исходная папка"
        self.source_dir_var = tk.StringVar()  # Переменная для хранения пути к исходной папке
        ttk.Entry(dir_frame, textvariable=self.source_dir_var, width=50).grid(row=0, column=1, padx=5, pady=5,
                                                                              sticky=tk.EW)  # Поле ввода пути
        ttk.Button(dir_frame, text=_("Browse"), command=self.browse_source_dir).grid(row=0, column=2, padx=5,
                                                                                     pady=5)  # Кнопка выбора папки
        ttk.Label(dir_frame, text=_("Destination Directory:")).grid(row=1, column=0, sticky=tk.W,
                                                                    pady=5)  # Надпись "Целевая папка"
        self.dest_dir_var = tk.StringVar()  # Переменная для хранения пути к целевой папке
        ttk.Entry(dir_frame, textvariable=self.dest_dir_var, width=50).grid(row=1, column=1, padx=5, pady=5,
                                                                            sticky=tk.EW)  # Поле ввода пути
        ttk.Button(dir_frame, text=_("Browse"), command=self.browse_dest_dir).grid(row=1, column=2, padx=5,
                                                                                   pady=5)  # Кнопка выбора папки
        dir_frame.columnconfigure(1, weight=1)  # Настройка растяжения столбца

        # Фрейм настроек категорий
        category_frame = ttk.LabelFrame(main_frame, text=_("Categories"), padding="10")  # Фрейм для категорий
        category_frame.pack(fill=tk.BOTH, expand=True, pady=10)  # Размещение фрейма
        self.category_tree = ttk.Treeview(category_frame, height=5)  # Дерево категорий
        self.category_tree.pack(fill=tk.BOTH, expand=True, pady=5)  # Размещение дерева

        # Подфрейм для автоматической сортировки и глубины подкатегорий
        auto_frame = ttk.Frame(category_frame)  # Создание подфрейма
        auto_frame.pack(fill=tk.X, pady=5)  # Размещение подфрейма
        self.auto_sort_var = tk.BooleanVar(
            value=True)  # Переменная для автоматической сортировки (по умолчанию включена)
        self.auto_sort_check = ttk.Checkbutton(auto_frame, text=_("Automatic Sorting"), variable=self.auto_sort_var,
                                               command=self.toggle_auto_sort)  # Чекбокс автоматической сортировки
        self.auto_sort_check.pack(side=tk.LEFT, padx=5)  # Размещение чекбокса
        ttk.Label(auto_frame, text=_("Max Subcategory Depth:")).pack(side=tk.LEFT,
                                                                     padx=5)  # Надпись "Максимальная глубина подкатегорий"
        self.max_depth_var = tk.StringVar(value="3")  # Переменная для максимальной глубины (по умолчанию 3)
        self.max_depth_entry = ttk.Entry(auto_frame, textvariable=self.max_depth_var, width=5)  # Поле ввода глубины
        self.max_depth_entry.pack(side=tk.LEFT, padx=5)  # Размещение поля ввода

        # Подфрейм кнопок управления категориями
        category_buttons_frame = ttk.Frame(category_frame)  # Создание подфрейма для кнопок
        category_buttons_frame.pack(fill=tk.X, pady=5)  # Размещение подфрейма
        self.add_category_btn = ttk.Button(category_buttons_frame, text=_("Add Category"), command=self.add_category,
                                           state=tk.DISABLED)  # Кнопка добавления категории
        self.add_category_btn.pack(side=tk.LEFT, padx=5)  # Размещение кнопки
        self.add_subcategory_btn = ttk.Button(category_buttons_frame, text=_("Add Subcategory"),
                                              command=self.add_subcategory,
                                              state=tk.DISABLED)  # Кнопка добавления подкатегории
        self.add_subcategory_btn.pack(side=tk.LEFT, padx=5)  # Размещение кнопки
        self.remove_category_btn = ttk.Button(category_buttons_frame, text=_("Remove"), command=self.remove_category,
                                              state=tk.DISABLED)  # Кнопка удаления категории
        self.remove_category_btn.pack(side=tk.LEFT, padx=5)  # Размещение кнопки

        # Фрейм настроек удаления дубликатов
        dedupe_frame = ttk.LabelFrame(main_frame, text=_("Duplicate Removal Options"),
                                      padding="10")  # Фрейм для дубликатов
        dedupe_frame.pack(fill=tk.X, pady=10)  # Размещение фрейма
        self.dedupe_mode = tk.StringVar(value="none")  # Переменная для режима удаления дубликатов
        ttk.Radiobutton(dedupe_frame, text=_("No Deduplication"), value="none", variable=self.dedupe_mode).pack(
            side=tk.LEFT, padx=5)  # Радиокнопка "Без удаления"
        ttk.Radiobutton(dedupe_frame, text=_("Normal (Exact Matches)"), value="normal", variable=self.dedupe_mode).pack(
            side=tk.LEFT, padx=5)  # Радиокнопка "Обычный режим"
        ttk.Radiobutton(dedupe_frame, text=_("Hardcore (Similar Files)"), value="hardcore",
                        variable=self.dedupe_mode).pack(side=tk.LEFT, padx=5)  # Радиокнопка "Жёсткий режим"

        # Фрейм лога
        log_frame = ttk.LabelFrame(main_frame, text=_("Log"), padding="10")  # Фрейм для лога
        log_frame.pack(fill=tk.BOTH, expand=True, pady=10)  # Размещение фрейма
        self.log_text = tk.Text(log_frame, height=10, state=tk.DISABLED)  # Текстовое поле для лога
        self.log_text.pack(fill=tk.BOTH, expand=True, pady=5)  # Размещение поля
        ttk.Button(log_frame, text=_("Export Log"), command=self.export_log).pack(side=tk.BOTTOM,
                                                                                  pady=5)  # Кнопка экспорта лога

        # Прогресс-бар
        self.progress_var = tk.DoubleVar()  # Переменная для отслеживания прогресса
        self.progress_bar = ttk.Progressbar(main_frame, variable=self.progress_var, maximum=100)  # Прогресс-бар
        self.progress_bar.pack(fill=tk.X, pady=10)  # Размещение прогресс-бара

        # Фрейм кнопок управления
        button_frame = ttk.Frame(main_frame)  # Создание фрейма для кнопок
        button_frame.pack(fill=tk.X, pady=10)  # Размещение фрейма
        self.sort_button = ttk.Button(button_frame, text=_("Start Sorting"),
                                      command=self.start_sorting)  # Кнопка запуска сортировки
        self.sort_button.pack(side=tk.RIGHT, padx=5)  # Размещение кнопки
        self.backup_button = ttk.Button(button_frame, text=_("Create Backup"),
                                        command=self.create_backup)  # Кнопка создания резервной копии
        self.backup_button.pack(side=tk.RIGHT, padx=5)  # Размещение кнопки

        self.is_processing = False  # Флаг выполнения процесса сортировки
        self.cancel_requested = False  # Флаг запроса отмены сортировки

    def toggle_auto_sort(self):
        """
        Включает или отключает автоматическую сортировку и управление кнопками категорий.
        """
        if self.auto_sort_var.get():  # Если автоматическая сортировка включена
            self.add_category_btn.config(state=tk.DISABLED)  # Отключение кнопки добавления категории
            self.add_subcategory_btn.config(state=tk.DISABLED)  # Отключение кнопки добавления подкатегории
            self.remove_category_btn.config(state=tk.DISABLED)  # Отключение кнопки удаления категории
            self.category_tree.delete(*self.category_tree.get_children())  # Удаление всех элементов из дерева категорий
            self.category_list.clear()  # Очистка списка категорий
        else:  # Если автоматическая сортировка отключена
            self.add_category_btn.config(state=tk.NORMAL)  # Включение кнопки добавления категории
            self.add_subcategory_btn.config(state=tk.NORMAL)  # Включение кнопки добавления подкатегории
            self.remove_category_btn.config(state=tk.NORMAL)  # Включение кнопки удаления категории

    def change_language(self, lang):
        """
        Меняет язык интерфейса приложения.

        Аргументы:
            lang (str): Код языка ("en" для английского, "ru" для русского).
        """
        global _  # Доступ к глобальной функции перевода
        self.language = lang  # Установка нового языка
        _ = setup_localization(lang)  # Обновление функции перевода
        self.root.title(_("Document Sorter with Ollama"))  # Обновление заголовка окна
        self.setup_ui()  # Перерисовка интерфейса с новым языком

    def connect_google_drive(self):
        """
        Подключает приложение к Google Drive через API.
        """
        try:
            creds = service_account.Credentials.from_service_account_file("credentials.json", scopes=[
                "https://www.googleapis.com/auth/drive"])
            # Загрузка учётных данных из файла credentials.json
            self.google_drive_service = build('drive', 'v3', credentials=creds)  # Инициализация сервиса Google Drive
            self.log_message(_("Connected to Google Drive"))  # Сообщение об успешном подключении
        except Exception as e:
            self.log_message(_(f"Google Drive connection error: {str(e)}"))  # Сообщение об ошибке подключения

    def connect_dropbox(self):
        """
        Подключает приложение к Dropbox через API.
        """
        token = simpledialog.askstring(_("Dropbox"), _("Enter Dropbox Access Token:"))  # Запрос токена доступа
        if token:  # Если токен введён
            try:
                self.dropbox_client = Dropbox(token)  # Инициализация клиента Dropbox
                self.log_message(_("Connected to Dropbox"))  # Сообщение об успешном подключении
            except AuthError as e:
                self.log_message(_(f"Dropbox authentication error: {str(e)}"))  # Сообщение об ошибке аутентификации

    def handle_drop(self, event):
        """
        Обрабатывает событие Drag-and-Drop для выбора исходной папки.

        Аргументы:
            event: Событие перетаскивания с данными о сброшенном объекте.
        """
        dropped = event.data  # Получение данных о сброшенном объекте
        if os.path.isdir(dropped):  # Проверка, является ли объект папкой
            self.source_dir_var.set(dropped)  # Установка пути в переменную исходной папки
            self.log_message(_(f"Dropped directory: {dropped}"))  # Логирование события

    def check_ollama_status(self):
        """
        Проверяет статус подключения к Ollama API и обновляет метку статуса.
        """
        try:
            response = requests.get(f"{self.ollama_url}/version")  # Запрос версии API Ollama
            if response.status_code == 200:  # Если запрос успешен
                self.status_label.config(text=_("Connected"), foreground="green")  # Установка статуса "Подключено"
                self.fetch_models()  # Обновление списка моделей
            else:
                self.status_label.config(text=_("Error: API not responding"),
                                         foreground="red")  # Установка статуса ошибки
        except requests.exceptions.ConnectionError:
            self.status_label.config(text=_("Disconnected (Is Ollama running?)"),
                                     foreground="red")  # Статус "Отключено"
            self.root.after(5000, self.check_ollama_status)  # Повторная проверка через 5 секунд

    def fetch_models(self):
        """
        Получает список доступных моделей от Ollama и обновляет выпадающий список.
        """
        try:
            response = requests.get(f"{self.ollama_url}/tags")  # Запрос списка моделей
            if response.status_code == 200:  # Если запрос успешен
                models_data = response.json()  # Парсинг ответа в JSON
                self.available_models = [model["name"] for model in
                                         models_data.get("models", [])]  # Извлечение имён моделей
                self.model_combobox["values"] = self.available_models  # Обновление списка в интерфейсе
                if self.model in self.available_models:  # Если текущая модель доступна
                    self.model_combobox.set(self.model)  # Установка текущей модели
                elif self.available_models:  # Если есть хотя бы одна модель
                    self.model_combobox.set(self.available_models[0])  # Установка первой доступной модели
                    self.model = self.available_models[0]  # Обновление текущей модели
        except requests.exceptions.ConnectionError:
            self.log_message(_("Cannot connect to Ollama"))  # Сообщение об ошибке подключения

    def on_model_selected(self, event):
        """
        Обрабатывает выбор модели из выпадающего списка.

        Аргументы:
            event: Событие выбора элемента в Combobox.
        """
        self.model = self.model_combobox.get()  # Получение выбранной модели
        self.log_message(_(f"Selected model: {self.model}"))  # Логирование выбора модели

    def browse_source_dir(self):
        """
        Открывает диалог для выбора исходной папки.
        """
        dir_path = filedialog.askdirectory(title=_("Select Source Directory"))  # Открытие диалога выбора папки
        if dir_path:  # Если папка выбрана
            self.source_dir_var.set(dir_path)  # Установка пути в переменную

    def browse_dest_dir(self):
        """
        Открывает диалог для выбора целевой папки.
        """
        dir_path = filedialog.askdirectory(title=_("Select Destination Directory"))  # Открытие диалога выбора папки
        if dir_path:  # Если папка выбрана
            self.dest_dir_var.set(dir_path)  # Установка пути в переменную

    def add_category(self):
        """
        Добавляет новую категорию в список категорий.
        """
        category = simpledialog.askstring(_("Add Category"), _("Enter category name:"))  # Запрос имени категории
        if category and category not in self.category_list:  # Если имя введено и уникально
            self.category_tree.insert("", tk.END, text=category)  # Добавление категории в дерево
            self.category_list.append(category)  # Добавление категории в список

    def add_subcategory(self):
        """
        Добавляет подкатегорию к выбранной категории.
        """
        selected = self.category_tree.selection()  # Получение выбранного элемента
        if not selected:  # Если ничего не выбрано
            messagebox.showwarning(_("Warning"), _("Select a category first"))  # Предупреждение
            return
        parent = self.category_tree.item(selected[0])["text"]  # Получение имени родительской категории
        subcategory = simpledialog.askstring(_("Add Subcategory"),
                                             _(f"Enter subcategory for {parent}:"))  # Запрос имени подкатегории
        if subcategory:  # Если имя введено
            self.category_tree.insert(selected[0], tk.END, text=subcategory)  # Добавление подкатегории в дерево
            self.category_list.append(f"{parent}/{subcategory}")  # Добавление подкатегории в список

    def remove_category(self):
        """
        Удаляет выбранную категорию или подкатегорию из списка.
        """
        selected = self.category_tree.selection()  # Получение выбранного элемента
        if selected:  # Если элемент выбран
            item = self.category_tree.item(selected[0])["text"]  # Получение имени элемента
            parent = self.category_tree.parent(selected[0])  # Получение родителя элемента
            if parent:  # Если есть родитель (это подкатегория)
                item = f"{self.category_tree.item(parent)['text']}/{item}"  # Формирование полного пути подкатегории
            self.category_tree.delete(selected[0])  # Удаление элемента из дерева
            self.category_list.remove(item)  # Удаление элемента из списка

    def log_message(self, message):
        """
        Добавляет сообщение в лог с временной меткой.

        Аргументы:
            message (str): Текст сообщения для добавления в лог.
        """
        self.log_text.config(state=tk.NORMAL)  # Включение редактирования текстового поля
        self.log_text.insert(tk.END, f"{time.strftime('%H:%M:%S')} - {message}\n")  # Добавление сообщения с временем
        self.log_text.see(tk.END)  # Автоматическая прокрутка вниз
        self.log_text.config(state=tk.DISABLED)  # Отключение редактирования

    def export_log(self):
        """
        Экспортирует содержимое лога в текстовый файл.
        """
        log_content = self.log_text.get("1.0", tk.END)  # Получение всего текста из лога
        file_path = filedialog.asksaveasfilename(defaultextension=".txt",
                                                 filetypes=[("Text files", "*.txt")])  # Запрос пути для сохранения
        if file_path:  # Если путь выбран
            with open(file_path, "w", encoding="utf-8") as file:  # Открытие файла для записи
                file.write(log_content)  # Запись лога в файл
            self.log_message(_(f"Log saved to {file_path}"))  # Логирование успешного сохранения

    def create_backup(self):
        """
        Создаёт резервную копию исходной папки в ZIP-архиве.
        """
        source_dir = self.source_dir_var.get()  # Получение пути к исходной папке
        if not source_dir or not os.path.isdir(source_dir):  # Проверка корректности пути
            messagebox.showerror(_("Error"), _("Please select a valid source directory"))  # Ошибка, если путь неверный
            return
        backup_path = filedialog.asksaveasfilename(defaultextension=".zip", filetypes=[("ZIP files", "*.zip")],
                                                   title=_("Save Backup As"))  # Запрос пути для архива
        if backup_path:  # Если путь выбран
            try:
                with zipfile.ZipFile(backup_path, 'w', zipfile.ZIP_DEFLATED) as zipf:  # Создание ZIP-архива
                    for root, _, files in os.walk(source_dir):  # Обход всех файлов в папке
                        for file in files:
                            file_path = os.path.join(root, file)  # Формирование полного пути к файлу
                            arcname = os.path.relpath(file_path, source_dir)  # Относительный путь для архива
                            zipf.write(file_path, arcname)  # Добавление файла в архив
                self.log_message(_(f"Backup created at {backup_path}"))  # Логирование успешного создания
            except Exception as e:
                self.log_message(_(f"Backup error: {str(e)}"))  # Логирование ошибки

    def start_sorting(self):
        """
        Запускает процесс сортировки файлов в отдельном потоке.
        """
        if self.is_processing:  # Если сортировка уже идёт
            return
        source_dir = self.source_dir_var.get()  # Получение исходной папки
        dest_dir = self.dest_dir_var.get()  # Получение целевой папки
        if not source_dir or not os.path.isdir(source_dir):  # Проверка исходной папки
            messagebox.showerror(_("Error"), _("Please select a valid source directory"))  # Ошибка, если путь неверный
            return
        if not dest_dir or not os.path.isdir(dest_dir):  # Проверка целевой папки
            messagebox.showerror(_("Error"),
                                 _("Please select a valid destination directory"))  # Ошибка, если путь неверный
            return
        if not self.auto_sort_var.get() and not self.category_list:  # Проверка наличия категорий
            messagebox.showerror(_("Error"),
                                 _("Please define at least one category or enable automatic sorting"))  # Ошибка при отсутствии категорий
            return
        try:
            self.max_depth = int(self.max_depth_var.get())  # Получение максимальной глубины подкатегорий
            if self.max_depth < 1:  # Проверка корректности значения
                raise ValueError("Depth must be at least 1")  # Ошибка при некорректной глубине
        except ValueError:
            messagebox.showerror(_("Error"), _("Invalid subcategory depth"))  # Ошибка ввода глубины
            return
        self.sort_button.config(state=tk.DISABLED)  # Отключение кнопки сортировки
        self.backup_button.config(state=tk.DISABLED)  # Отключение кнопки резервного копирования
        self.is_processing = True  # Установка флага выполнения
        self.cancel_requested = False  # Сброс флага отмены
        thread = threading.Thread(target=self.sort_documents,
                                  args=(source_dir, dest_dir))  # Создание потока для сортировки
        thread.daemon = True  # Установка потока как фонового
        thread.start()  # Запуск потока

    def get_file_hash(self, file_path):
        """
        Вычисляет MD5-хэш файла для поиска дубликатов.

        Аргументы:
            file_path (str): Путь к файлу.

        Возвращает:
            str: Хэш файла в виде строки.
        """
        hasher = hashlib.md5()  # Инициализация объекта для вычисления MD5
        with open(file_path, 'rb') as f:  # Открытие файла в бинарном режиме
            hasher.update(f.read())  # Обновление хэша содержимым файла
        return hasher.hexdigest()  # Возвращение хэша в шестнадцатеричном формате

    def find_and_remove_duplicates(self, files, mode="normal"):
        """
        Находит и удаляет дубликаты файлов в зависимости от режима.

        Аргументы:
            files (list): Список путей к файлам.
            mode (str): Режим удаления дубликатов ("normal" — точное совпадение, "hardcore" — по имени и размеру).

        Возвращает:
            list: Список уникальных файлов после удаления дубликатов.
        """
        if mode == "none":  # Если режим "без удаления"
            return files  # Возвращаем исходный список файлов
        file_info = {}  # Словарь для хранения информации о файлах
        for file_path in files:  # Обход всех файлов
            file_hash = self.get_file_hash(file_path)  # Вычисление хэша файла
            file_size = os.path.getsize(file_path)  # Получение размера файла
            mod_time = os.path.getmtime(file_path)  # Получение времени изменения
            file_name = os.path.basename(file_path)  # Получение имени файла
            file_info[file_path] = {"hash": file_hash, "size": file_size, "mod_time": mod_time,
                                    "name": file_name}  # Сохранение информации
        duplicates = {}  # Словарь для группировки дубликатов
        if mode == "normal":  # Обычный режим (точное совпадение по хэшу)
            for path, info in file_info.items():
                key = info["hash"]  # Ключ — хэш файла
                if key not in duplicates:
                    duplicates[key] = []  # Инициализация списка для хэша
                duplicates[key].append(path)  # Добавление пути в группу
        elif mode == "hardcore":  # Жёсткий режим (совпадение по имени и размеру)
            for path, info in file_info.items():
                key = (info["name"], info["size"])  # Ключ — имя и размер
                if key not in duplicates:
                    duplicates[key] = []  # Инициализация списка для ключа
                duplicates[key].append(path)  # Добавление пути в группу
        unique_files = []  # Список уникальных файлов
        for group in duplicates.values():  # Обход групп дубликатов
            if len(group) > 1:  # Если в группе больше одного файла
                sorted_group = sorted(group, key=lambda x: file_info[x]["mod_time"],
                                      reverse=True)  # Сортировка по времени (новые первыми)
                keep_file = sorted_group[0]  # Оставляем самый новый файл
                unique_files.append(keep_file)  # Добавление файла в список уникальных
                for duplicate in sorted_group[1:]:  # Удаление остальных дубликатов
                    os.remove(duplicate)  # Удаление файла
                    self.log_message(_(f"Removed duplicate: {os.path.basename(duplicate)}"))  # Логирование удаления
            else:
                unique_files.append(group[0])  # Добавление единственного файла в группу
        return unique_files  # Возвращение списка уникальных файлов

    def generate_auto_categories(self, files):
        """
        Генерирует категории автоматически с помощью Ollama на основе анализа файлов.

        Аргументы:
            files (list): Список путей к файлам для анализа.
        """
        self.category_list.clear()  # Очистка текущего списка категорий
        self.category_tree.delete(*self.category_tree.get_children())  # Очистка дерева категорий
        file_info_list = [{"filename": os.path.basename(f), "extension": os.path.splitext(f)[1].lower(),
                           "size_bytes": os.path.getsize(f)} for f in files]  # Формирование списка информации о файлах
        prompt = f"""
        {_('Analyze the following files and suggest a hierarchical category structure with a maximum depth of')} {self.max_depth}.
        {_('Return a JSON object with categories and subcategories based on file names, extensions, and sizes.')}

        Files: {json.dumps(file_info_list, indent=2)}

        {_('Example output:')}
        {{
            "Documents": {{
                "Work": {{}},
                "Personal": {{}}
            }},
            "Images": {{}},
            "PDF": {{}}
        }}
        """  # Формирование запроса для Ollama
        try:
            response = requests.post(f"{self.ollama_url}/generate",
                                     json={"model": self.model, "prompt": prompt, "stream": False})  # Отправка запроса
            if response.status_code == 200:  # Если запрос успешен
                categories = json.loads(response.json().get("response", "{}"))  # Парсинг ответа в JSON
                self._build_category_tree(categories)  # Построение дерева категорий
                self.log_message(_("Automatic categories generated by Ollama"))  # Логирование успешной генерации
            else:
                self.log_message(_(f"Failed to generate categories: {response.status_code}"))  # Логирование ошибки
                self.category_list = ["Default"]  # Установка категории по умолчанию
                self.category_tree.insert("", tk.END, text="Default")  # Добавление категории в дерево
        except Exception as e:
            self.log_message(_(f"Error generating categories: {str(e)}"))  # Логирование ошибки
            self.category_list = ["Default"]  # Установка категории по умолчанию
            self.category_tree.insert("", tk.END, text="Default")  # Добавление категории в дерево

    def _build_category_tree(self, categories, parent=""):
        """
        Рекурсивно строит дерево категорий из структуры JSON.

        Аргументы:
            categories (dict): Словарь категорий и подкатегорий от Ollama.
            parent (str): Идентификатор родительской категории в дереве (по умолчанию пустой).
        """
        for cat, subcats in categories.items():  # Обход категорий и их подкатегорий
            full_cat = f"{parent}/{cat}" if parent else cat  # Формирование полного пути категории
            self.category_list.append(full_cat)  # Добавление категории в список
            cat_id = self.category_tree.insert(parent if parent else "", tk.END,
                                               text=cat)  # Добавление категории в дерево
            if isinstance(subcats, dict) and subcats and len(
                    full_cat.split('/')) - 1 < self.max_depth:  # Проверка вложенности
                self._build_category_tree(subcats, cat_id)  # Рекурсивный вызов для подкатегорий

    def sort_documents(self, source_dir, dest_dir):
        """
        Сортирует документы из исходной папки в целевую с учётом дубликатов и категорий.

        Аргументы:
            source_dir (str): Путь к исходной папке.
            dest_dir (str): Путь к целевой папке.
        """
        try:
            files = []  # Список файлов для сортировки
            if self.google_drive_service or self.dropbox_client:  # Если используется облако
                files = self.get_cloud_files(source_dir)  # Получение файлов из облака
            else:
                files = [os.path.join(source_dir, f) for f in os.listdir(source_dir) if
                         os.path.isfile(os.path.join(source_dir, f))]  # Получение локальных файлов
            if not files:  # Если файлов нет
                self.log_message(_("No files found"))  # Логирование отсутствия файлов
                self.complete_sorting()  # Завершение сортировки
                return
            self.log_message(_(f"Found {len(files)} files to process"))  # Логирование количества файлов

            if self.auto_sort_var.get():  # Если включена автоматическая сортировка
                self.generate_auto_categories(files)  # Генерация категорий
            for category in self.category_list:  # Создание папок для категорий
                os.makedirs(os.path.join(dest_dir, category), exist_ok=True)  # Создание директории, если её нет

            dedupe_mode = self.dedupe_mode.get()  # Получение режима удаления дубликатов
            files = self.find_and_remove_duplicates(files, dedupe_mode)  # Удаление дубликатов
            self.log_message(_(f"After deduplication: {len(files)} files remain"))  # Логирование оставшихся файлов

            with concurrent.futures.ThreadPoolExecutor(max_workers=4) as executor:  # Создание пула потоков
                futures = []  # Список задач для выполнения
                for file_path in files:  # Обход файлов
                    if self.cancel_requested:  # Проверка запроса отмены
                        self.log_message(_("Sorting cancelled"))  # Логирование отмены
                        break
                    futures.append(executor.submit(self.process_file, file_path, dest_dir))  # Добавление задачи в пул
                for i, future in enumerate(concurrent.futures.as_completed(futures)):  # Обработка завершённых задач
                    if self.cancel_requested:  # Проверка запроса отмены
                        break
                    progress = (i + 1) / len(files) * 100  # Вычисление прогресса
                    self.progress_var.set(progress)  # Обновление прогресс-бара
                    self.root.update_idletasks()  # Обновление интерфейса
            self.log_message(_("Sorting completed"))  # Логирование завершения сортировки
        except Exception as e:
            self.log_message(_(f"Sorting error: {str(e)}"))  # Логирование ошибки
        finally:
            self.complete_sorting()  # Завершение сортировки

    def get_cloud_files(self, source_dir):
        """
        Получает файлы из облачных хранилищ (Google Drive или Dropbox).

        Аргументы:
            source_dir (str): Путь к папке в облаке.

        Возвращает:
            list: Список путей к скачанным файлам.
        """
        files = []  # Список для хранения путей к файлам
        temp_dir = os.path.join(os.path.expanduser("~"), "DocumentSorterTemp")  # Временная папка для скачивания
        os.makedirs(temp_dir, exist_ok=True)  # Создание временной папки, если её нет
        if self.google_drive_service:  # Если подключён Google Drive
            results = self.google_drive_service.files().list().execute()  # Получение списка файлов
            for file in results.get('files', []):  # Обход файлов
                request = self.google_drive_service.files().get_media(fileId=file['id'])  # Запрос на скачивание
                file_path = os.path.join(temp_dir, file['name'])  # Путь для сохранения файла
                with open(file_path, 'wb') as f:  # Открытие файла для записи
                    downloader = MediaIoBaseDownload(f, request)  # Инициализация загрузчика
                    done = False  # Флаг завершения загрузки
                    while not done:  # Пока загрузка не завершена
                        _, done = downloader.next_chunk()  # Скачивание следующей части файла
                files.append(file_path)  # Добавление пути в список
        elif self.dropbox_client:  # Если подключён Dropbox
            result = self.dropbox_client.files_list_folder(source_dir)  # Получение списка файлов
            for entry in result.entries:  # Обход файлов
                if isinstance(entry, dropbox.files.FileMetadata):  # Проверка, что это файл
                    file_path = os.path.join(temp_dir, entry.name)  # Путь для сохранения файла
                    self.dropbox_client.files_download_to_file(file_path, entry.path_lower)  # Скачивание файла
                    files.append(file_path)  # Добавление пути в список
        return files  # Возвращение списка файлов

    def process_file(self, file_path, dest_dir):
        """
        Обрабатывает один файл: классифицирует и перемещает его в целевую папку.

        Аргументы:
            file_path (str): Путь к файлу.
            dest_dir (str): Целевая папка для сортировки.
        """
        filename = os.path.basename(file_path)  # Получение имени файла
        file_info = {"filename": filename, "extension": os.path.splitext(filename)[1].lower(),
                     "size_bytes": os.path.getsize(file_path)}  # Формирование информации о файле
        self.log_message(_(f"Processing: {filename}"))  # Логирование начала обработки
        category = self.classify_file(file_info)  # Классификация файла
        if category:  # Если категория определена
            dest_path = os.path.join(dest_dir, category, filename)  # Формирование пути назначения
            os.makedirs(os.path.dirname(dest_path), exist_ok=True)  # Создание папки, если её нет
            if os.path.exists(dest_path):  # Если файл уже существует
                base, ext = os.path.splitext(filename)  # Разделение имени и расширения
                dest_path = os.path.join(dest_dir, category, f"{base}_copy{ext}")  # Добавление "_copy" к имени
            shutil.move(file_path, dest_path)  # Перемещение файла
            self.log_message(_(f"Moved '{filename}' to '{category}'"))  # Логирование перемещения

    def classify_file(self, file_info):
        """
        Классифицирует файл с помощью Ollama и возвращает категорию.

        Аргументы:
            file_info (dict): Информация о файле (имя, расширение, размер).

        Возвращает:
            str: Название категории для файла.
        """
        try:
            prompt = f"""
            {_('Classify the file into ONE of these categories:')} {', '.join(self.category_list)}
            {_('File:')} {file_info['filename']}
            {_('Extension:')} {file_info['extension']}
            {_('Size:')} {file_info['size_bytes']} {_('bytes')}
            {_('Respond with ONLY the category name.')}
            """  # Формирование запроса для Ollama
            response = requests.post(f"{self.ollama_url}/generate",
                                     json={"model": self.model, "prompt": prompt, "stream": False})  # Отправка запроса
            if response.status_code == 200:  # Если запрос успешен
                category = response.json().get("response", "").strip()  # Получение категории из ответа
                return category if category in self.category_list else self.category_list[0]  # Проверка категории
            return self.category_list[0]  # Возвращение категории по умолчанию при ошибке
        except Exception as e:
            self.log_message(_(f"Classification error: {str(e)}"))  # Логирование ошибки
            return self.category_list[0]  # Возвращение категории по умолчанию

    def complete_sorting(self):
        """
        Завершает процесс сортировки и восстанавливает интерфейс.
        """
        self.is_processing = False  # Сброс флага выполнения
        self.sort_button.config(state=tk.NORMAL)  # Включение кнопки сортировки
        self.backup_button.config(state=tk.NORMAL)  # Включение кнопки резервного копирования
        self.progress_var.set(0)  # Сброс прогресс-бара


def main():
    """
    Главная функция для запуска приложения с поддержкой командной строки или GUI.
    """
    parser = argparse.ArgumentParser(description="Document Sorter")  # Инициализация парсера аргументов
    parser.add_argument("--source", help="Source directory")  # Аргумент для исходной папки
    parser.add_argument("--dest", help="Destination directory")  # Аргумент для целевой папки
    parser.add_argument("--categories",
                        help="Comma-separated categories (disables auto-sorting if provided)")  # Аргумент для категорий
    parser.add_argument("--dedupe", choices=["none", "normal", "hardcore"], default="none",
                        help="Duplicate removal mode")  # Аргумент для режима дубликатов
    args = parser.parse_args()  # Парсинг аргументов

    if args.source and args.dest:  # Если указаны исходная и целевая папки
        root = TkinterDnD.Tk()  # Создание окна с поддержкой Drag-and-Drop
        sorter = DocumentSorter(root)  # Инициализация сортировщика
        sorter.source_dir_var.set(args.source)  # Установка исходной папки
        sorter.dest_dir_var.set(args.dest)  # Установка целевой папки
        if args.categories:  # Если категории указаны через аргументы
            sorter.auto_sort_var.set(False)  # Отключение автоматической сортировки
            sorter.category_list = args.categories.split(",")  # Разделение категорий по запятой
            for cat in sorter.category_list:  # Обход категорий
                parts = cat.split("/")  # Разделение на подкатегории
                parent = ""  # Идентификатор родителя
                for part in parts:  # Обход частей пути
                    full_part = f"{parent}/{part}" if parent else part  # Формирование полного пути
                    if full_part not in [sorter.category_tree.item(i, "text") for i in
                                         sorter.category_tree.get_children(parent)]:  # Проверка уникальности
                        parent_id = sorter.category_tree.insert(parent if parent else "", tk.END,
                                                                text=part)  # Добавление в дерево
                        parent = parent_id if parent else parent_id  # Обновление родителя
        sorter.dedupe_mode.set(args.dedupe)  # Установка режима дубликатов
        sorter.sort_documents(args.source, args.dest)  # Запуск сортировки
    else:
        root = TkinterDnD.Tk()  # Создание окна для GUI-режима
        app = DocumentSorter(root)  # Инициализация сортировщика
        root.mainloop()  # Запуск главного цикла приложения


if __name__ == "__main__":
    main()  # Запуск главной функции