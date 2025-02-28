import os
import shutil
import tkinter as tk
from tkinter import filedialog, ttk, messagebox, simpledialog
import threading
import requests
import json
import time
from typing import List, Dict, Any
import mimetypes
import datetime
import zipfile
import hashlib
import locale

# Для извлечения текста из документов
import docx
import PyPDF2
import csv
from odf import text, teletype  # Для OpenDocument форматов

# Для многопоточности
import concurrent.futures

# Настройка локали для правильного отображения дат
locale.setlocale(locale.LC_ALL, '')


class DocumentSorter:
    """
    Класс для сортировки документов с использованием модели Ollama.

    Атрибуты:
        root (tk.Tk): Корневое окно Tkinter.
        ollama_url (str): URL для подключения к Ollama API.
        model (str): Название модели Ollama по умолчанию.
        available_models (list): Список доступных моделей Ollama.
        category_list (list): Список категорий для сортировки.
        cache (dict): Кэш для хранения результатов анализа файлов.
    """

    def __init__(self, root):
        self.root = root
        self.root.title("Document Sorter with Ollama")
        self.root.geometry("900x700")
        self.root.resizable(True, True)

        # Базовые настройки
        self.ollama_url = "http://localhost:11434/api"  # URL для Ollama API
        self.model = "deepseek-coder"  # Модель по умолчанию
        self.available_models = []  # Список моделей Ollama
        self.category_list = []  # Список категорий
        self.cache = {}  # Кэш для результатов анализа

        self.setup_ui()  # Настройка интерфейса
        self.check_ollama_status()  # Проверка статуса Ollama

    def setup_ui(self):
        """
        Настраивает пользовательский интерфейс приложения.
        """
        # Главный фрейм
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Индикатор статуса Ollama
        status_frame = ttk.Frame(main_frame)
        status_frame.pack(fill=tk.X, pady=5)

        ttk.Label(status_frame, text="Статус Ollama:").pack(side=tk.LEFT, padx=5)
        self.status_label = ttk.Label(status_frame, text="Проверка...", foreground="orange")
        self.status_label.pack(side=tk.LEFT, padx=5)

        # Выбор модели
        model_frame = ttk.Frame(main_frame)
        model_frame.pack(fill=tk.X, pady=5)

        ttk.Label(model_frame, text="Выберите модель:").pack(side=tk.LEFT, padx=5)
        self.model_combobox = ttk.Combobox(model_frame, state="readonly")
        self.model_combobox.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        self.model_combobox.bind("<<ComboboxSelected>>", self.on_model_selected)

        refresh_button = ttk.Button(model_frame, text="Обновить модели", command=self.fetch_models)
        refresh_button.pack(side=tk.RIGHT, padx=5)

        # Настройки параметров модели
        params_frame = ttk.LabelFrame(main_frame, text="Параметры модели", padding="10")
        params_frame.pack(fill=tk.X, pady=10)

        ttk.Label(params_frame, text="Температура:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.temperature_var = tk.DoubleVar(value=0.1)
        ttk.Scale(params_frame, from_=0.0, to=1.0, orient=tk.HORIZONTAL, variable=self.temperature_var).grid(row=0, column=1, padx=5, pady=5, sticky=tk.EW)

        ttk.Label(params_frame, text="Top P:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.top_p_var = tk.DoubleVar(value=0.9)
        ttk.Scale(params_frame, from_=0.0, to=1.0, orient=tk.HORIZONTAL, variable=self.top_p_var).grid(row=1, column=1, padx=5, pady=5, sticky=tk.EW)

        # Выбор каталогов
        dir_frame = ttk.LabelFrame(main_frame, text="Выбор каталогов", padding="10")
        dir_frame.pack(fill=tk.X, pady=10)

        ttk.Label(dir_frame, text="Исходный каталог:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.source_dir_var = tk.StringVar()
        ttk.Entry(dir_frame, textvariable=self.source_dir_var, width=50).grid(row=0, column=1, padx=5, pady=5, sticky=tk.EW)
        ttk.Button(dir_frame, text="Обзор", command=self.browse_source_dir).grid(row=0, column=2, padx=5, pady=5)

        ttk.Label(dir_frame, text="Целевой каталог:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.dest_dir_var = tk.StringVar()
        ttk.Entry(dir_frame, textvariable=self.dest_dir_var, width=50).grid(row=1, column=1, padx=5, pady=5, sticky=tk.EW)
        ttk.Button(dir_frame, text="Обзор", command=self.browse_dest_dir).grid(row=1, column=2, padx=5, pady=5)

        dir_frame.columnconfigure(1, weight=1)

        # Настройки категорий
        category_frame = ttk.LabelFrame(main_frame, text="Категории", padding="10")
        category_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        self.category_listbox = tk.Listbox(category_frame, height=5)
        self.category_listbox.pack(fill=tk.BOTH, expand=True, pady=5)

        category_buttons_frame = ttk.Frame(category_frame)
        category_buttons_frame.pack(fill=tk.X, pady=5)

        ttk.Button(category_buttons_frame, text="Добавить", command=self.add_category).pack(side=tk.LEFT, padx=5)
        ttk.Button(category_buttons_frame, text="Удалить", command=self.remove_category).pack(side=tk.LEFT, padx=5)

        # Настройки анализа документов
        analysis_frame = ttk.LabelFrame(main_frame, text="Настройки анализа", padding="10")
        analysis_frame.pack(fill=tk.X, pady=10)

        self.analyze_content_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(analysis_frame, text="Анализировать содержимое документов", variable=self.analyze_content_var).pack(anchor=tk.W)

        max_size_frame = ttk.Frame(analysis_frame)
        max_size_frame.pack(fill=tk.X, pady=5)
        ttk.Label(max_size_frame, text="Макс. размер файла для анализа (МБ):").pack(side=tk.LEFT, padx=5)
        self.max_size_var = tk.StringVar(value="10")
        ttk.Entry(max_size_frame, textvariable=self.max_size_var, width=5).pack(side=tk.LEFT, padx=5)

        # Фильтр по дате
        date_filter_frame = ttk.Frame(analysis_frame)
        date_filter_frame.pack(fill=tk.X, pady=5)
        ttk.Label(date_filter_frame, text="Фильтр по дате (с):").pack(side=tk.LEFT, padx=5)
        self.date_from_var = tk.StringVar()
        ttk.Entry(date_filter_frame, textvariable=self.date_from_var, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Label(date_filter_frame, text="по:").pack(side=tk.LEFT, padx=5)
        self.date_to_var = tk.StringVar()
        ttk.Entry(date_filter_frame, textvariable=self.date_to_var, width=10).pack(side=tk.LEFT, padx=5)

        # Лог
        log_frame = ttk.LabelFrame(main_frame, text="Лог", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        self.log_text = tk.Text(log_frame, height=10, state=tk.DISABLED)
        self.log_text.pack(fill=tk.BOTH, expand=True, pady=5)

        # Кнопка экспорта лога
        ttk.Button(log_frame, text="Экспорт лога", command=self.export_log).pack(side=tk.BOTTOM, pady=5)

        # Прогресс
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(main_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, pady=10)

        # Кнопки
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10)

        self.sort_button = ttk.Button(button_frame, text="Начать сортировку", command=self.start_sorting)
        self.sort_button.pack(side=tk.RIGHT, padx=5)

        self.preview_button = ttk.Button(button_frame, text="Предварительный просмотр", command=self.preview_structure)
        self.preview_button.pack(side=tk.RIGHT, padx=5)

        self.cancel_button = ttk.Button(button_frame, text="Отмена", command=self.cancel_sorting, state=tk.DISABLED)
        self.cancel_button.pack(side=tk.RIGHT, padx=5)

        # Флаги обработки
        self.is_processing = False
        self.cancel_requested = False

    def check_ollama_status(self):
        """
        Проверяет статус подключения к Ollama API.
        """
        try:
            response = requests.get(f"{self.ollama_url}/version")
            if response.status_code == 200:
                self.status_label.config(text="Подключено", foreground="green")
                self.fetch_models()  # Загрузить доступные модели
            else:
                self.status_label.config(text="Ошибка: API не отвечает", foreground="red")
        except requests.exceptions.ConnectionError:
            self.status_label.config(text="Отключено (Ollama запущена?)", foreground="red")
            self.root.after(5000, self.check_ollama_status)  # Повторная проверка через 5 секунд

    def fetch_models(self):
        """
        Получает список доступных моделей от Ollama API.
        """
        try:
            response = requests.get(f"{self.ollama_url}/tags")
            if response.status_code == 200:
                models_data = response.json()
                self.available_models = [model["name"] for model in models_data.get("models", [])]

                if not self.available_models:
                    self.log_message("Модели не найдены. Пожалуйста, загрузите модель в Ollama.")
                    self.available_models = ["deepseek-coder", "llama3", "codellama", "mistral"]  # Резервные модели

                self.model_combobox["values"] = self.available_models

                # Установить модель по умолчанию
                if self.model in self.available_models:
                    self.model_combobox.set(self.model)
                elif self.available_models:
                    self.model_combobox.set(self.available_models[0])
                    self.model = self.available_models[0]
            else:
                self.log_message(f"Ошибка при получении моделей: {response.status_code}")
        except requests.exceptions.ConnectionError:
            self.log_message("Не удаётся подключиться к Ollama. Проверьте, запущена ли она.")
            self.status_label.config(text="Отключено", foreground="red")

    def on_model_selected(self, event):
        """
        Обрабатывает выбор модели пользователем.
        """
        self.model = self.model_combobox.get()
        self.log_message(f"Выбрана модель: {self.model}")

    def browse_source_dir(self):
        """
        Открывает диалог для выбора исходного каталога.
        """
        dir_path = filedialog.askdirectory(title="Выберите исходный каталог")
        if dir_path:
            self.source_dir_var.set(dir_path)

    def browse_dest_dir(self):
        """
        Открывает диалог для выбора целевого каталога.
        """
        dir_path = filedialog.askdirectory(title="Выберите целевой каталог")
        if dir_path:
            self.dest_dir_var.set(dir_path)

    def add_category(self):
        """
        Добавляет новую категорию в список.
        """
        category = simpledialog.askstring("Добавить категорию", "Введите название категории:")
        if category and category not in self.category_list:
            self.category_listbox.insert(tk.END, category)
            self.category_list.append(category)

    def remove_category(self):
        """
        Удаляет выбранную категорию из списка.
        """
        selected = self.category_listbox.curselection()
        if selected:
            index = selected[0]
            category = self.category_listbox.get(index)
            self.category_listbox.delete(index)
            self.category_list.remove(category)

    def log_message(self, message):
        """
        Выводит сообщение в лог с временной меткой.
        """
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"{time.strftime('%H:%M:%S')} - {message}\n")
        self.log_text.see(tk.END)  # Автоматическая прокрутка вниз
        self.log_text.config(state=tk.DISABLED)

    def export_log(self):
        """
        Экспортирует содержимое лога в текстовый файл.
        """
        log_content = self.log_text.get("1.0", tk.END)
        file_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt")])
        if file_path:
            with open(file_path, "w", encoding="utf-8") as file:
                file.write(log_content)
            self.log_message(f"Лог сохранён в {file_path}")

    def start_sorting(self):
        """
        Запускает процесс сортировки в отдельном потоке.
        """
        if self.is_processing:
            return

        source_dir = self.source_dir_var.get()
        dest_dir = self.dest_dir_var.get()

        if not source_dir or not os.path.isdir(source_dir):
            messagebox.showerror("Ошибка", "Выберите действительный исходный каталог")
            return

        if not dest_dir or not os.path.isdir(dest_dir):
            messagebox.showerror("Ошибка", "Выберите действительный целевой каталог")
            return

        categories = self.category_list
        if not categories:
            messagebox.showerror("Ошибка", "Определите хотя бы одну категорию")
            return

        # Проверка максимального размера файла
        try:
            max_size = float(self.max_size_var.get())
            if max_size <= 0:
                raise ValueError("Размер должен быть положительным")
        except ValueError:
            messagebox.showerror("Ошибка", "Введите действительный размер файла (МБ)")
            return

        # Фильтр по дате
        date_from = self.date_from_var.get()
        date_to = self.date_to_var.get()
        date_from_dt = None
        date_to_dt = None
        if date_from:
            try:
                date_from_dt = datetime.datetime.strptime(date_from, "%Y-%m-%d")
            except ValueError:
                messagebox.showerror("Ошибка", "Неверный формат даты (от). Используйте ГГГГ-ММ-ДД")
                return
        if date_to:
            try:
                date_to_dt = datetime.datetime.strptime(date_to, "%Y-%m-%d")
            except ValueError:
                messagebox.showerror("Ошибка", "Неверный формат даты (до). Используйте ГГГГ-ММ-ДД")
                return

        # Создать папки категорий
        for category in categories:
            category_path = os.path.join(dest_dir, category)
            if not os.path.exists(category_path):
                os.makedirs(category_path)

        # Начать сортировку в отдельном потоке
        self.sort_button.config(state=tk.DISABLED)
        self.preview_button.config(state=tk.DISABLED)
        self.cancel_button.config(state=tk.NORMAL)
        self.is_processing = True
        self.cancel_requested = False

        thread = threading.Thread(
            target=self.sort_documents,
            args=(source_dir, dest_dir, categories, max_size, date_from_dt, date_to_dt)
        )
        thread.daemon = True
        thread.start()

    def cancel_sorting(self):
        """
        Отменяет процесс сортировки.
        """
        if self.is_processing:
            self.cancel_requested = True
            self.log_message("Запрошена отмена. Ожидание завершения текущих операций...")

    def preview_structure(self):
        """
        Показывает предварительный просмотр структуры каталога после сортировки.
        """
        source_dir = self.source_dir_var.get()
        dest_dir = self.dest_dir_var.get()
        categories = self.category_list

        if not source_dir or not os.path.isdir(source_dir):
            messagebox.showerror("Ошибка", "Выберите действительный исходный каталог")
            return

        if not dest_dir or not os.path.isdir(dest_dir):
            messagebox.showerror("Ошибка", "Выберите действительный целевой каталог")
            return

        if not categories:
            messagebox.showerror("Ошибка", "Определите хотя бы одну категорию")
            return

        files = [f for f in os.listdir(source_dir) if os.path.isfile(os.path.join(source_dir, f))]
        if not files:
            messagebox.showinfo("Информация", "В исходном каталоге нет файлов")
            return

        # Простая симуляция классификации по расширению
        structure = {category: [] for category in categories}
        for filename in files:
            ext = os.path.splitext(filename)[1].lower()
            if ext in ['.doc', '.docx', '.txt', '.rtf', '.odt']:
                category = 'Documents'
            elif ext in ['.jpg', '.jpeg', '.png', '.gif', '.bmp']:
                category = 'Images'
            elif ext in ['.xls', '.xlsx', '.csv', '.ods']:
                category = 'Spreadsheets'
            elif ext in ['.ppt', '.pptx', '.odp']:
                category = 'Presentations'
            elif ext == '.pdf':
                category = 'PDFs'
            elif ext in ['.zip', '.rar']:
                category = 'Archives'
            else:
                category = 'Other'

            if category in categories:
                structure[category].append(filename)
            else:
                structure[categories[0]].append(filename)  # По умолчанию в первую категорию

        # Отобразить предварительный просмотр
        preview_window = tk.Toplevel(self.root)
        preview_window.title("Предварительный просмотр структуры")
        preview_window.geometry("600x400")

        tree = ttk.Treeview(preview_window)
        tree.pack(fill=tk.BOTH, expand=True)

        for category, files in structure.items():
            cat_id = tree.insert("", "end", text=category)
            for file in files:
                tree.insert(cat_id, "end", text=file)

    def extract_document_content(self, file_path, file_info):
        """
        Извлекает текстовое содержимое из файла, включая заголовок и автора, если доступно.

        Аргументы:
            file_path (str): Путь к файлу.
            file_info (dict): Информация о файле (имя, расширение и т.д.).

        Возвращает:
            dict: Словарь с извлечённой информацией (content_sample, title, author).
        """
        try:
            extension = file_info['extension'].lower()
            content = ""
            title = ""
            author = ""

            # Обработка DOCX
            if extension == '.docx':
                doc = docx.Document(file_path)
                content = "\n".join([p.text for p in doc.paragraphs if p.text])
                properties = doc.core_properties
                title = properties.title or ""
                author = properties.author or ""

            # Обработка PDF
            elif extension == '.pdf':
                with open(file_path, 'rb') as file:
                    reader = PyPDF2.PdfReader(file)
                    page_texts = [reader.pages[i].extract_text() for i in range(min(5, len(reader.pages)))]
                    content = "\n".join(page_texts)
                    metadata = reader.metadata or {}
                    title = metadata.get('/Title', '')
                    author = metadata.get('/Author', '')

            # Обработка текстовых файлов
            elif extension in ['.txt', '.md', '.rtf']:
                with open(file_path, 'r', encoding='utf-8', errors='ignore') as file:
                    content = file.read(10000)  # Ограничение до 10к символов
                    lines = content.split('\n')
                    if lines and lines[0].strip():
                        title = lines[0].strip()

            # Обработка CSV
            elif extension == '.csv':
                with open(file_path, 'r', encoding='utf-8', errors='ignore') as file:
                    csv_reader = csv.reader(file)
                    rows = [','.join(row) for i, row in enumerate(csv_reader) if i < 10]
                    content = '\n'.join(rows)

            # Обработка OpenDocument форматов
            elif extension in ['.odt', '.ods', '.odp']:
                doc = text.load(file_path)
                content = teletype.extractText(doc)
                headers = doc.getElementsByType(text.H)
                if headers:
                    title = teletype.extractText(headers[0])

            # Обработка архивов
            elif extension in ['.zip']:
                with zipfile.ZipFile(file_path, 'r') as zip_ref:
                    content = "\n".join(zip_ref.namelist())

            # Если заголовок не найден, взять первую подходящую строку
            if not title and content:
                lines = content.split('\n')
                for line in lines[:5]:
                    if line.strip() and len(line) < 100:
                        title = line
                        break

            return {
                "content_sample": content[:1000] if content else "",
                "title": title,
                "author": author
            }

        except Exception as e:
            self.log_message(f"Ошибка при извлечении содержимого {file_info['filename']}: {str(e)}")
            return {"content_sample": "", "title": "", "author": ""}

    def sort_documents(self, source_dir, dest_dir, categories, max_size_mb, date_from, date_to):
        """
        Сортирует документы в многопоточном режиме.

        Аргументы:
            source_dir (str): Исходный каталог.
            dest_dir (str): Целевой каталог.
            categories (list): Список категорий.
            max_size_mb (float): Максимальный размер файла для анализа (МБ).
            date_from (datetime): Начальная дата фильтра.
            date_to (datetime): Конечная дата фильтра.
        """
        try:
            files = [f for f in os.listdir(source_dir) if os.path.isfile(os.path.join(source_dir, f))]
            if not files:
                self.log_message("В исходном каталоге нет файлов")
                self.complete_sorting()
                return

            self.log_message(f"Найдено {len(files)} файлов для обработки")
            max_size_bytes = max_size_mb * 1024 * 1024  # Перевод МБ в байты
            analyze_content = self.analyze_content_var.get()

            # Фильтрация по дате
            if date_from or date_to:
                files = [f for f in files if self.is_file_in_date_range(os.path.join(source_dir, f), date_from, date_to)]

            # Многопоточная обработка файлов
            with concurrent.futures.ThreadPoolExecutor(max_workers=4) as executor:
                futures = []
                for filename in files:
                    if self.cancel_requested:
                        self.log_message("Сортировка отменена")
                        break
                    file_path = os.path.join(source_dir, filename)
                    futures.append(executor.submit(self.process_file, file_path, dest_dir, categories, max_size_bytes, analyze_content))

                for i, future in enumerate(concurrent.futures.as_completed(futures)):
                    if self.cancel_requested:
                        break
                    progress = (i + 1) / len(files) * 100
                    self.progress_var.set(progress)
                    self.root.update_idletasks()

            self.log_message("Сортировка завершена")
        except Exception as e:
            self.log_message(f"Ошибка при сортировке: {str(e)}")
        finally:
            self.complete_sorting()

    def is_file_in_date_range(self, file_path, date_from, date_to):
        """
        Проверяет, находится ли дата изменения файла в заданном диапазоне.

        Аргументы:
            file_path (str): Путь к файлу.
            date_from (datetime): Начальная дата.
            date_to (datetime): Конечная дата.

        Возвращает:
            bool: True, если файл в диапазоне, иначе False.
        """
        file_mtime = datetime.datetime.fromtimestamp(os.path.getmtime(file_path))
        if date_from and file_mtime < date_from:
            return False
        if date_to and file_mtime > date_to:
            return False
        return True

    def process_file(self, file_path, dest_dir, categories, max_size_bytes, analyze_content):
        """
        Обрабатывает один файл: классифицирует и перемещает его.

        Аргументы:
            file_path (str): Путь к файлу.
            dest_dir (str): Целевой каталог.
            categories (list): Список категорий.
            max_size_bytes (int): Максимальный размер файла для анализа.
            analyze_content (bool): Анализировать ли содержимое.
        """
        filename = os.path.basename(file_path)
        file_size = os.path.getsize(file_path)
        _, ext = os.path.splitext(filename)
        file_info = {
            "filename": filename,
            "extension": ext.lower(),
            "size_bytes": file_size,
            "mime_type": mimetypes.guess_type(filename)[0] or "unknown/unknown"
        }

        self.log_message(f"Обработка: {filename} ({file_size / 1024 / 1024:.2f} МБ)")

        # Проверка кэша
        file_hash = self.get_file_hash(file_path)
        if file_hash in self.cache:
            category = self.cache[file_hash]
            self.log_message(f"Использован кэш для {filename}: {category}")
        else:
            # Извлечение содержимого
            content_info = {"content_sample": "", "title": "", "author": ""}
            if analyze_content and file_size <= max_size_bytes and ext.lower() in ['.docx', '.pdf', '.txt', '.md', '.rtf', '.csv', '.odt', '.ods', '.odp']:
                content_info = self.extract_document_content(file_path, file_info)
                if content_info["title"]:
                    self.log_message(f"Найден заголовок: {content_info['title']}")
                if content_info["author"]:
                    self.log_message(f"Найден автор: {content_info['author']}")

            # Классификация файла
            category = self.classify_file(file_info, categories, content_info)
            self.cache[file_hash] = category  # Сохранение в кэш

        if category and category in categories:
            dest_path = os.path.join(dest_dir, category, filename)
            if os.path.exists(dest_path):
                base, ext = os.path.splitext(filename)
                dest_path = os.path.join(dest_dir, category, f"{base}_copy{ext}")
            shutil.move(file_path, dest_path)
            self.log_message(f"Перемещён '{filename}' в категорию '{category}'")
        else:
            self.log_message(f"Не удалось классифицировать '{filename}', оставлен в исходной папке")

    def get_file_hash(self, file_path):
        """
        Вычисляет MD5 хэш файла для кэширования.

        Аргументы:
            file_path (str): Путь к файлу.

        Возвращает:
            str: Хэш файла.
        """
        hasher = hashlib.md5()
        with open(file_path, 'rb') as file:
            hasher.update(file.read())
        return hasher.hexdigest()

    def classify_file(self, file_info, categories, content_info=None):
        """
        Классифицирует файл с помощью Ollama API.

        Аргументы:
            file_info (dict): Информация о файле.
            categories (list): Список категорий.
            content_info (dict): Информация о содержимом (если доступно).

        Возвращает:
            str: Название категории или None.
        """
        try:
            prompt = f"""
            Пожалуйста, классифицируйте следующий файл в ОДНУ из этих категорий: {', '.join(categories)}

            Информация о файле:
            - Имя файла: {file_info['filename']}
            - Расширение: {file_info['extension']}
            - Размер: {file_info['size_bytes']} байт
            - MIME-тип: {file_info['mime_type']}
            """

            if content_info and any(content_info.values()):
                prompt += "\nИнформация о содержимом документа:\n"
                if content_info["title"]:
                    prompt += f"- Заголовок: {content_info['title']}\n"
                if content_info["author"]:
                    prompt += f"- Автор: {content_info['author']}\n"
                if content_info["content_sample"]:
                    sample = content_info["content_sample"][:500] + "..." if len(content_info["content_sample"]) > 500 else content_info["content_sample"]
                    prompt += f"- Образец содержимого: {sample}\n"

            prompt += "\nОтветьте ТОЛЬКО названием категории."

            # Запрос к Ollama API
            response = requests.post(
                f"{self.ollama_url}/generate",
                json={
                    "model": self.model,
                    "prompt": prompt,
                    "stream": False,
                    "options": {
                        "temperature": self.temperature_var.get(),
                        "top_p": self.top_p_var.get()
                    }
                }
            )

            if response.status_code == 200:
                prediction = response.json().get("response", "").strip()
                if prediction in categories:
                    return prediction

                # Резервная классификация по расширению
                ext = file_info['extension'].lower()
                extension_map = {
                    '.doc': 'Documents', '.docx': 'Documents', '.txt': 'Documents', '.rtf': 'Documents', '.odt': 'Documents',
                    '.jpg': 'Images', '.jpeg': 'Images', '.png': 'Images', '.gif': 'Images', '.bmp': 'Images',
                    '.xls': 'Spreadsheets', '.xlsx': 'Spreadsheets', '.csv': 'Spreadsheets', '.ods': 'Spreadsheets',
                    '.ppt': 'Presentations', '.pptx': 'Presentations', '.odp': 'Presentations',
                    '.pdf': 'PDFs',
                    '.zip': 'Archives'
                }
                category = extension_map.get(ext, 'Other')
                return category if category in categories else categories[0]
            else:
                self.log_message(f"Ошибка Ollama API: {response.status_code}")
                return None
        except Exception as e:
            self.log_message(f"Ошибка классификации: {str(e)}")
            return None

    def complete_sorting(self):
        """
        Завершает сортировку и восстанавливает интерфейс.
        """
        self.is_processing = False
        self.sort_button.config(state=tk.NORMAL)
        self.preview_button.config(state=tk.NORMAL)
        self.cancel_button.config(state=tk.DISABLED)
        self.progress_var.set(0)


def main():
    mimetypes.init()  # Инициализация MIME-типов
    root = tk.Tk()
    app = DocumentSorter(root)
    root.mainloop()


if __name__ == "__main__":
    main()
