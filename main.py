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
import subprocess
import sys
import argparse
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import io
from dropbox import Dropbox
from dropbox.files import WriteMode
from dropbox.exceptions import ApiError, AuthError

# Для извлечения текста из документов
import docx
import PyPDF2
import csv
from odf import text, teletype  # Для OpenDocument форматов
from langdetect import detect  # Для определения языка

# Для многопоточности
import concurrent.futures

# Для многоязычности
import gettext

# Настройка локали для правильного отображения дат
locale.setlocale(locale.LC_ALL, '')

# Настройка многоязычности
def setup_localization(lang="en"):
    languages = {'en': 'en_US', 'ru': 'ru_RU'}
    loc = languages.get(lang, 'en_US')
    translation = gettext.translation('sorter', localedir='locale', languages=[loc], fallback=True)
    translation.install()
    return translation.gettext

_ = setup_localization("en")  # Функция перевода по умолчанию

class DocumentSorter:
    """
    Класс для сортировки документов с использованием модели Ollama и дополнительных функций, включая удаление дубликатов.

    Атрибуты:
        root (tk.Tk): Корневое окно Tkinter.
        ollama_url (str): URL для подключения к Ollama API.
        model (str): Название модели Ollama по умолчанию.
        category_list (list): Список категорий для сортировки.
        cache (dict): Кэш для хранения результатов анализа файлов.
        language (str): Текущий язык интерфейса.
    """

    def __init__(self, root):
        self.root = root
        self.root.title(_("Document Sorter with Ollama"))
        self.root.geometry("900x700")
        self.root.resizable(True, True)

        # Базовые настройки
        self.ollama_url = "http://localhost:11434/api"  # URL для Ollama API
        self.model = "deepseek-coder"  # Модель по умолчанию
        self.available_models = []  # Список моделей Ollama
        self.category_list = []  # Список категорий
        self.cache = {}  # Кэш для результатов анализа
        self.language = "en"  # Язык по умолчанию

        # Поддержка облачных сервисов
        self.google_drive_service = None
        self.dropbox_client = None

        self.setup_ui()  # Настройка интерфейса
        self.check_ollama_status()  # Проверка статуса Ollama

        # Поддержка drag-and-drop
        self.root.drop_target_register(tk.DND_FILES)
        self.root.dnd_bind('<<Drop>>', self.handle_drop)

    def setup_ui(self):
        """
        Настраивает пользовательский интерфейс приложения.
        """
        # Главное меню
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        lang_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label=_("Language"), menu=lang_menu)
        lang_menu.add_command(label="English", command=lambda: self.change_language("en"))
        lang_menu.add_command(label="Русский", command=lambda: self.change_language("ru"))

        # Главный фрейм
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Индикатор статуса Ollama
        status_frame = ttk.Frame(main_frame)
        status_frame.pack(fill=tk.X, pady=5)

        ttk.Label(status_frame, text=_("Ollama Status:")).pack(side=tk.LEFT, padx=5)
        self.status_label = ttk.Label(status_frame, text=_("Checking..."), foreground="orange")
        self.status_label.pack(side=tk.LEFT, padx=5)

        # Выбор модели
        model_frame = ttk.Frame(main_frame)
        model_frame.pack(fill=tk.X, pady=5)

        ttk.Label(model_frame, text=_("Select Model:")).pack(side=tk.LEFT, padx=5)
        self.model_combobox = ttk.Combobox(model_frame, state="readonly")
        self.model_combobox.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        self.model_combobox.bind("<<ComboboxSelected>>", self.on_model_selected)

        refresh_button = ttk.Button(model_frame, text=_("Refresh Models"), command=self.fetch_models)
        refresh_button.pack(side=tk.RIGHT, padx=5)

        # Интеграция с облаком
        cloud_frame = ttk.LabelFrame(main_frame, text=_("Cloud Storage"), padding="10")
        cloud_frame.pack(fill=tk.X, pady=10)

        ttk.Button(cloud_frame, text=_("Connect Google Drive"), command=self.connect_google_drive).pack(side=tk.LEFT, padx=5)
        ttk.Button(cloud_frame, text=_("Connect Dropbox"), command=self.connect_dropbox).pack(side=tk.LEFT, padx=5)

        # Выбор каталогов
        dir_frame = ttk.LabelFrame(main_frame, text=_("Directory Selection"), padding="10")
        dir_frame.pack(fill=tk.X, pady=10)

        ttk.Label(dir_frame, text=_("Source Directory:")).grid(row=0, column=0, sticky=tk.W, pady=5)
        self.source_dir_var = tk.StringVar()
        ttk.Entry(dir_frame, textvariable=self.source_dir_var, width=50).grid(row=0, column=1, padx=5, pady=5, sticky=tk.EW)
        ttk.Button(dir_frame, text=_("Browse"), command=self.browse_source_dir).grid(row=0, column=2, padx=5, pady=5)

        ttk.Label(dir_frame, text=_("Destination Directory:")).grid(row=1, column=0, sticky=tk.W, pady=5)
        self.dest_dir_var = tk.StringVar()
        ttk.Entry(dir_frame, textvariable=self.dest_dir_var, width=50).grid(row=1, column=1, padx=5, pady=5, sticky=tk.EW)
        ttk.Button(dir_frame, text=_("Browse"), command=self.browse_dest_dir).grid(row=1, column=2, padx=5, pady=5)

        dir_frame.columnconfigure(1, weight=1)

        # Настройки категорий
        category_frame = ttk.LabelFrame(main_frame, text=_("Categories"), padding="10")
        category_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        self.category_tree = ttk.Treeview(category_frame, height=5)
        self.category_tree.pack(fill=tk.BOTH, expand=True, pady=5)

        category_buttons_frame = ttk.Frame(category_frame)
        category_buttons_frame.pack(fill=tk.X, pady=5)

        ttk.Button(category_buttons_frame, text=_("Add Category"), command=self.add_category).pack(side=tk.LEFT, padx=5)
        ttk.Button(category_buttons_frame, text=_("Add Subcategory"), command=self.add_subcategory).pack(side=tk.LEFT, padx=5)
        ttk.Button(category_buttons_frame, text=_("Remove"), command=self.remove_category).pack(side=tk.LEFT, padx=5)

        # Настройки удаления дубликатов
        dedupe_frame = ttk.LabelFrame(main_frame, text=_("Duplicate Removal Options"), padding="10")
        dedupe_frame.pack(fill=tk.X, pady=10)

        self.dedupe_mode = tk.StringVar(value="none")
        ttk.Radiobutton(dedupe_frame, text=_("No Deduplication"), value="none", variable=self.dedupe_mode).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(dedupe_frame, text=_("Normal (Exact Matches)"), value="normal", variable=self.dedupe_mode).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(dedupe_frame, text=_("Hardcore (Similar Files)"), value="hardcore", variable=self.dedupe_mode).pack(side=tk.LEFT, padx=5)

        # Лог
        log_frame = ttk.LabelFrame(main_frame, text=_("Log"), padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        self.log_text = tk.Text(log_frame, height=10, state=tk.DISABLED)
        self.log_text.pack(fill=tk.BOTH, expand=True, pady=5)

        ttk.Button(log_frame, text=_("Export Log"), command=self.export_log).pack(side=tk.BOTTOM, pady=5)

        # Прогресс
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(main_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, pady=10)

        # Кнопки
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10)

        self.sort_button = ttk.Button(button_frame, text=_("Start Sorting"), command=self.start_sorting)
        self.sort_button.pack(side=tk.RIGHT, padx=5)

        self.backup_button = ttk.Button(button_frame, text=_("Create Backup"), command=self.create_backup)
        self.backup_button.pack(side=tk.RIGHT, padx=5)

        # Флаги обработки
        self.is_processing = False
        self.cancel_requested = False

    def change_language(self, lang):
        """
        Меняет язык интерфейса.
        """
        global _
        self.language = lang
        _ = setup_localization(lang)
        self.root.title(_("Document Sorter with Ollama"))
        self.setup_ui()  # Перерисовка интерфейса

    def connect_google_drive(self):
        """
        Подключается к Google Drive через OAuth.
        """
        try:
            creds = service_account.Credentials.from_service_account_file("credentials.json", scopes=["https://www.googleapis.com/auth/drive"])
            self.google_drive_service = build('drive', 'v3', credentials=creds)
            self.log_message(_("Connected to Google Drive"))
        except Exception as e:
            self.log_message(_(f"Google Drive connection error: {str(e)}"))

    def connect_dropbox(self):
        """
        Подключается к Dropbox через токен доступа.
        """
        token = simpledialog.askstring(_("Dropbox"), _("Enter Dropbox Access Token:"))
        if token:
            try:
                self.dropbox_client = Dropbox(token)
                self.log_message(_("Connected to Dropbox"))
            except AuthError as e:
                self.log_message(_(f"Dropbox authentication error: {str(e)}"))

    def handle_drop(self, event):
        """
        Обрабатывает событие drag-and-drop.
        """
        dropped = event.data
        if os.path.isdir(dropped):
            self.source_dir_var.set(dropped)
            self.log_message(_(f"Dropped directory: {dropped}"))

    def check_ollama_status(self):
        """
        Проверяет статус подключения к Ollama API.
        """
        try:
            response = requests.get(f"{self.ollama_url}/version")
            if response.status_code == 200:
                self.status_label.config(text=_("Connected"), foreground="green")
                self.fetch_models()
            else:
                self.status_label.config(text=_("Error: API not responding"), foreground="red")
        except requests.exceptions.ConnectionError:
            self.status_label.config(text=_("Disconnected (Is Ollama running?)"), foreground="red")
            self.root.after(5000, self.check_ollama_status)

    def fetch_models(self):
        """
        Получает список доступных моделей от Ollama API.
        """
        try:
            response = requests.get(f"{self.ollama_url}/tags")
            if response.status_code == 200:
                models_data = response.json()
                self.available_models = [model["name"] for model in models_data.get("models", [])]
                self.model_combobox["values"] = self.available_models
                if self.model in self.available_models:
                    self.model_combobox.set(self.model)
                elif self.available_models:
                    self.model_combobox.set(self.available_models[0])
                    self.model = self.available_models[0]
        except requests.exceptions.ConnectionError:
            self.log_message(_("Cannot connect to Ollama"))

    def on_model_selected(self, event):
        """
        Обрабатывает выбор модели пользователем.
        """
        self.model = self.model_combobox.get()
        self.log_message(_(f"Selected model: {self.model}"))

    def browse_source_dir(self):
        """
        Открывает диалог для выбора исходного каталога.
        """
        dir_path = filedialog.askdirectory(title=_("Select Source Directory"))
        if dir_path:
            self.source_dir_var.set(dir_path)

    def browse_dest_dir(self):
        """
        Открывает диалог для выбора целевого каталога.
        """
        dir_path = filedialog.askdirectory(title=_("Select Destination Directory"))
        if dir_path:
            self.dest_dir_var.set(dir_path)

    def add_category(self):
        """
        Добавляет новую категорию в список.
        """
        category = simpledialog.askstring(_("Add Category"), _("Enter category name:"))
        if category and category not in self.category_list:
            self.category_tree.insert("", tk.END, text=category)
            self.category_list.append(category)

    def add_subcategory(self):
        """
        Добавляет подкатегорию к выбранной категории.
        """
        selected = self.category_tree.selection()
        if not selected:
            messagebox.showwarning(_("Warning"), _("Select a category first"))
            return
        parent = self.category_tree.item(selected[0])["text"]
        subcategory = simpledialog.askstring(_("Add Subcategory"), _(f"Enter subcategory for {parent}:"))
        if subcategory:
            self.category_tree.insert(selected[0], tk.END, text=subcategory)
            self.category_list.append(f"{parent}/{subcategory}")

    def remove_category(self):
        """
        Удаляет выбранную категорию или подкатегорию.
        """
        selected = self.category_tree.selection()
        if selected:
            item = self.category_tree.item(selected[0])["text"]
            parent = self.category_tree.parent(selected[0])
            if parent:
                item = f"{self.category_tree.item(parent)['text']}/{item}"
            self.category_tree.delete(selected[0])
            self.category_list.remove(item)

    def log_message(self, message):
        """
        Выводит сообщение в лог с временной меткой.
        """
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"{time.strftime('%H:%M:%S')} - {message}\n")
        self.log_text.see(tk.END)
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
            self.log_message(_(f"Log saved to {file_path}"))

    def create_backup(self):
        """
        Создаёт резервную копию исходного каталога.
        """
        source_dir = self.source_dir_var.get()
        if not source_dir or not os.path.isdir(source_dir):
            messagebox.showerror(_("Error"), _("Please select a valid source directory"))
            return

        backup_path = filedialog.asksaveasfilename(
            defaultextension=".zip",
            filetypes=[("ZIP files", "*.zip")],
            title=_("Save Backup As")
        )
        if backup_path:
            try:
                with zipfile.ZipFile(backup_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                    for root, _, files in os.walk(source_dir):
                        for file in files:
                            file_path = os.path.join(root, file)
                            arcname = os.path.relpath(file_path, source_dir)
                            zipf.write(file_path, arcname)
                self.log_message(_(f"Backup created at {backup_path}"))
            except Exception as e:
                self.log_message(_(f"Backup error: {str(e)}"))

    def start_sorting(self):
        """
        Запускает процесс сортировки в отдельном потоке.
        """
        if self.is_processing:
            return

        source_dir = self.source_dir_var.get()
        dest_dir = self.dest_dir_var.get()

        if not source_dir or not os.path.isdir(source_dir):
            messagebox.showerror(_("Error"), _("Please select a valid source directory"))
            return

        if not dest_dir or not os.path.isdir(dest_dir):
            messagebox.showerror(_("Error"), _("Please select a valid destination directory"))
            return

        categories = self.category_list
        if not categories:
            messagebox.showerror(_("Error"), _("Please define at least one category"))
            return

        # Создать структуру каталогов
        for category in categories:
            category_path = os.path.join(dest_dir, category)
            os.makedirs(category_path, exist_ok=True)

        self.sort_button.config(state=tk.DISABLED)
        self.backup_button.config(state=tk.DISABLED)
        self.is_processing = True
        self.cancel_requested = False

        thread = threading.Thread(
            target=self.sort_documents,
            args=(source_dir, dest_dir, categories)
        )
        thread.daemon = True
        thread.start()

    def get_file_hash(self, file_path):
        """
        Вычисляет хэш MD5 файла для поиска дубликатов.

        Аргументы:
            file_path (str): Путь к файлу.

        Возвращает:
            str: Хэш файла.
        """
        hasher = hashlib.md5()
        with open(file_path, 'rb') as f:
            hasher.update(f.read())
        return hasher.hexdigest()

    def find_and_remove_duplicates(self, files, mode="normal"):
        """
        Находит и удаляет дубликаты файлов в зависимости от выбранного режима.

        Аргументы:
            files (list): Список путей к файлам.
            mode (str): Режим удаления дубликатов ("normal" или "hardcore").

        Возвращает:
            list: Список файлов после удаления дубликатов.
        """
        if mode == "none":
            return files

        file_info = {}
        for file_path in files:
            file_hash = self.get_file_hash(file_path)
            file_size = os.path.getsize(file_path)
            mod_time = os.path.getmtime(file_path)
            file_name = os.path.basename(file_path)
            file_info[file_path] = {
                "hash": file_hash,
                "size": file_size,
                "mod_time": mod_time,
                "name": file_name
            }

        # Группировка файлов
        duplicates = {}
        if mode == "normal":
            # Только абсолютно идентичные файлы (одинаковый хэш)
            for path, info in file_info.items():
                key = info["hash"]
                if key not in duplicates:
                    duplicates[key] = []
                duplicates[key].append(path)
        elif mode == "hardcore":
            # Файлы с одинаковым именем и размером (допускается разная чексумма)
            for path, info in file_info.items():
                key = (info["name"], info["size"])
                if key not in duplicates:
                    duplicates[key] = []
                duplicates[key].append(path)

        # Удаление дубликатов, оставляя самый новый файл
        unique_files = []
        for group in duplicates.values():
            if len(group) > 1:
                # Сортировка по времени изменения (от нового к старому)
                sorted_group = sorted(group, key=lambda x: file_info[x]["mod_time"], reverse=True)
                keep_file = sorted_group[0]  # Оставляем самый новый
                unique_files.append(keep_file)
                for duplicate in sorted_group[1:]:  # Удаляем остальные
                    os.remove(duplicate)
                    self.log_message(_(f"Removed duplicate: {os.path.basename(duplicate)}"))
            else:
                unique_files.append(group[0])

        return unique_files

    def sort_documents(self, source_dir, dest_dir, categories):
        """
        Сортирует документы из локального или облачного источника с удалением дубликатов.
        """
        try:
            files = []
            if self.google_drive_service or self.dropbox_client:
                files = self.get_cloud_files(source_dir)
            else:
                files = [os.path.join(source_dir, f) for f in os.listdir(source_dir) if os.path.isfile(os.path.join(source_dir, f))]

            if not files:
                self.log_message(_("No files found"))
                self.complete_sorting()
                return

            self.log_message(_(f"Found {len(files)} files to process"))

            # Удаление дубликатов перед сортировкой
            dedupe_mode = self.dedupe_mode.get()
            files = self.find_and_remove_duplicates(files, dedupe_mode)
            self.log_message(_(f"After deduplication: {len(files)} files remain"))

            with concurrent.futures.ThreadPoolExecutor(max_workers=4) as executor:
                futures = []
                for file_path in files:
                    if self.cancel_requested:
                        self.log_message(_("Sorting cancelled"))
                        break
                    futures.append(executor.submit(self.process_file, file_path, dest_dir, categories))

                for i, future in enumerate(concurrent.futures.as_completed(futures)):
                    if self.cancel_requested:
                        break
                    progress = (i + 1) / len(files) * 100
                    self.progress_var.set(progress)
                    self.root.update_idletasks()

            self.log_message(_("Sorting completed"))
        except Exception as e:
            self.log_message(_(f"Sorting error: {str(e)}"))
        finally:
            self.complete_sorting()

    def get_cloud_files(self, source_dir):
        """
        Получает файлы из облачного хранилища.
        """
        files = []
        temp_dir = os.path.join(os.path.expanduser("~"), "DocumentSorterTemp")
        os.makedirs(temp_dir, exist_ok=True)

        if self.google_drive_service:
            results = self.google_drive_service.files().list().execute()
            for file in results.get('files', []):
                request = self.google_drive_service.files().get_media(fileId=file['id'])
                file_path = os.path.join(temp_dir, file['name'])
                with open(file_path, 'wb') as f:
                    downloader = MediaIoBaseDownload(f, request)
                    done = False
                    while not done:
                        _, done = downloader.next_chunk()
                files.append(file_path)

        elif self.dropbox_client:
            result = self.dropbox_client.files_list_folder(source_dir)
            for entry in result.entries:
                if isinstance(entry, dropbox.files.FileMetadata):
                    file_path = os.path.join(temp_dir, entry.name)
                    self.dropbox_client.files_download_to_file(file_path, entry.path_lower)
                    files.append(file_path)

        return files

    def process_file(self, file_path, dest_dir, categories):
        """
        Обрабатывает один файл: классифицирует и перемещает его.
        """
        filename = os.path.basename(file_path)
        file_info = {
            "filename": filename,
            "extension": os.path.splitext(filename)[1].lower(),
            "size_bytes": os.path.getsize(file_path)
        }

        self.log_message(_(f"Processing: {filename}"))
        category = self.classify_file(file_info, categories)
        if category:
            dest_path = os.path.join(dest_dir, category, filename)
            os.makedirs(os.path.dirname(dest_path), exist_ok=True)
            if os.path.exists(dest_path):
                base, ext = os.path.splitext(filename)
                dest_path = os.path.join(dest_dir, category, f"{base}_copy{ext}")
            shutil.move(file_path, dest_path)
            self.log_message(_(f"Moved '{filename}' to '{category}'"))

    def classify_file(self, file_info, categories):
        """
        Классифицирует файл с помощью Ollama API.
        """
        try:
            prompt = f"""
            {_('Classify the file into ONE of these categories:')} {', '.join(categories)}
            {_('File:')} {file_info['filename']}
            {_('Extension:')} {file_info['extension']}
            {_('Size:')} {file_info['size_bytes']} {_('bytes')}
            {_('Respond with ONLY the category name.')}
            """
            response = requests.post(
                f"{self.ollama_url}/generate",
                json={"model": self.model, "prompt": prompt, "stream": False}
            )
            if response.status_code == 200:
                category = response.json().get("response", "").strip()
                return category if category in categories else categories[0]
            return categories[0]  # Резервная категория
        except Exception as e:
            self.log_message(_(f"Classification error: {str(e)}"))
            return categories[0]

    def complete_sorting(self):
        """
        Завершает сортировку и восстанавливает интерфейс.
        """
        self.is_processing = False
        self.sort_button.config(state=tk.NORMAL)
        self.backup_button.config(state=tk.NORMAL)
        self.progress_var.set(0)

def main():
    # Поддержка командной строки
    parser = argparse.ArgumentParser(description="Document Sorter")
    parser.add_argument("--source", help="Source directory")
    parser.add_argument("--dest", help="Destination directory")
    parser.add_argument("--categories", help="Comma-separated categories")
    parser.add_argument("--dedupe", choices=["none", "normal", "hardcore"], default="none", help="Duplicate removal mode")
    args = parser.parse_args()

    if args.source and args.dest and args.categories:
        categories = args.categories.split(",")
        sorter = DocumentSorter(tk.Tk())
        sorter.source_dir_var.set(args.source)
        sorter.dest_dir_var.set(args.dest)
        sorter.category_list = categories
        sorter.dedupe_mode.set(args.dedupe)
        sorter.sort_documents(args.source, args.dest, categories)
    else:
        root = tk.Tk()
        app = DocumentSorter(root)
        root.mainloop()

if __name__ == "__main__":
    main()
