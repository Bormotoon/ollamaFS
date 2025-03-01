# Document Sorter with Ollama

## English

### Overview

"Document Sorter with Ollama" is a Python-based application designed to organize files into categories using the Ollama
AI model. It supports both local and cloud-based file sorting (Google Drive and Dropbox), with features like duplicate
removal, automatic category generation, and a user-friendly GUI built with Tkinter.

### Features

- **AI-Powered Sorting**: Uses the Ollama model to classify files based on their names, extensions, and sizes.
- **Automatic Categories**: Option to let Ollama generate a hierarchical category structure (up to a specified depth,
  default is 3).
- **Manual Categories**: Users can define custom categories and subcategories via a tree-like interface.
- **Duplicate Removal**: Two modes:
    - **Normal**: Removes exact duplicates based on MD5 hash, keeping the newest file.
    - **Hardcore**: Removes files with identical names and sizes (allowing minor content differences), keeping the
      newest.
- **Cloud Integration**: Supports Google Drive and Dropbox for sorting files directly from cloud storage.
- **Backup**: Creates ZIP backups of the source directory before sorting.
- **Localization**: Available in English and Russian, switchable via the GUI.
- **Drag-and-Drop**: Supports drag-and-drop for selecting the source directory.
- **Command-Line Interface**: Can be run from the terminal with predefined options.

### Installation

1. **Clone the Repository**:
   ```bash
   git clone https://github.com/yourusername/document-sorter-ollama.git
   cd document-sorter-ollama

2. **Set Up Virtual Environment** (optional but recommended):

   bash



`python -m venv venv source venv/bin/activate # Linux/Mac venv\Scripts\activate # Windows`

4. **Install Dependencies**:

   bash



`pip install -r requirements.txt`

6. **Install Ollama**: Ensure Ollama is running locally at http://localhost:11434. Download and install
   from [Ollama's official site](https://ollama.ai/).

### Usage

#### GUI Mode

- Run the script without arguments:

  bash

  

  `python main.py`

- Use the interface to:

    - Select source and destination directories.

    - Enable "Automatic Sorting" or manually define categories.

    - Choose a duplicate removal mode.

    - Start sorting with the "Start Sorting" button.

#### Command-Line Mode

- Run with arguments:

  bash

  

  `python main.py --source /path/to/source --dest /path/to/dest --categories "Docs,Images" --dedupe normal`

- --source: Source directory.

- --dest: Destination directory.

- --categories: Optional comma-separated list (disables auto-sorting).

- --dedupe: "none", "normal", or "hardcore".

### Cloud Integration

- **Google Drive**: Requires a credentials.json file from Google Cloud Console. Place it in the project root.

- **Dropbox**: Requires an access token, prompted via the GUI or pre-configured.

### Requirements

See requirements.txt for a complete list of Python dependencies.

---

## Русский

### Обзор

"Document Sorter with Ollama" — это приложение на Python для сортировки файлов по категориям с использованием модели ИИ
Ollama. Поддерживает работу с локальными файлами и облачными хранилищами (Google Drive и Dropbox), включает удаление
дубликатов, автоматическую генерацию категорий и удобный интерфейс на Tkinter.

### Возможности

- **Сортировка с ИИ**: Использует модель Ollama для классификации файлов по именам, расширениям и размерам.

- **Автоматические категории**: Опция автоматической генерации иерархической структуры категорий (глубина до заданного
  уровня, по умолчанию 3).

- **Ручные категории**: Пользователь может задавать категории и подкатегории через древовидный интерфейс.

- **Удаление дубликатов**: Два режима:

    - **Обычный**: Удаляет точные копии по MD5-хэшу, сохраняя самый новый файл.

    - **Жёсткий**: Удаляет файлы с одинаковыми именами и размерами (допуская небольшие различия), сохраняя самый новый.

- **Интеграция с облаком**: Поддержка Google Drive и Dropbox для сортировки файлов из облака.

- **Резервное копирование**: Создаёт ZIP-архив исходной папки перед сортировкой.

- **Локализация**: Доступно на английском и русском языках, переключение через интерфейс.

- **Drag-and-Drop**: Поддержка перетаскивания для выбора исходной папки.

- **Командная строка**: Запуск из терминала с заданными параметрами.

### Установка

1. **Клонируйте репозиторий**:

   bash



`git clone https://github.com/yourusername/document-sorter-ollama.git cd document-sorter-ollama`

3. **Настройте виртуальное окружение** (рекомендуется):

   bash



`python -m venv venv source venv/bin/activate # Linux/Mac venv\Scripts\activate # Windows`

5. **Установите зависимости**:

   bash



`pip install -r requirements.txt`

7. **Установите Ollama**: Убедитесь, что Ollama работает локально на http://localhost:11434. Скачайте и установите
   с [официального сайта Ollama](https://ollama.ai/).

### Использование

#### Режим GUI

- Запустите скрипт без аргументов:

  bash

  

  `python main.py`

- Используйте интерфейс для:

    - Выбора исходной и целевой папок.

    - Включения "Автоматической сортировки" или задания категорий вручную.

    - Выбора режима удаления дубликатов.

    - Запуска сортировки кнопкой "Start Sorting".

#### Режим командной строки

- Запустите с аргументами:

  bash

  

  `python main.py --source /путь/к/исходной --dest /путь/к/целевой --categories "Документы,Изображения" --dedupe normal`

- --source: Исходная папка.

- --dest: Целевая папка.

- --categories: Список категорий через запятую (отключает авто-сортировку).

- --dedupe: "none", "normal" или "hardcore".

### Интеграция с облаком

- **Google Drive**: Требуется файл credentials.json из Google Cloud Console. Поместите его в корень проекта.

- **Dropbox**: Требуется токен доступа, запрашивается через интерфейс или задаётся заранее.

### Зависимости

См. файл requirements.txt для полного списка зависимостей Python.
