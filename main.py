# --- START OF FILE main.py ---

import argparse
import concurrent.futures
import gettext
import hashlib
import json
import locale
import os
import shutil
import threading
import time
import tkinter as tk
import zipfile
from tkinter import filedialog, ttk, messagebox, simpledialog

import logger
from jinja2 import Environment, FileSystemLoader
import logging
from logging.handlers import RotatingFileHandler
import asyncio
import aiohttp
import requests  # Keep for initial synchronous check
from multiprocessing import Pool, cpu_count  # Import cpu_count
import traceback  # For detailed error logging

# --- Library Imports for File Types ---
try:
	import PyPDF2
except ImportError:
	PyPDF2 = None
try:
	import docx
except ImportError:
	docx = None
try:
	import openpyxl
except ImportError:
	openpyxl = None
try:
	from odf import text as odf_text, teletype as odf_teletype
except ImportError:
	odf_text, odf_teletype = None, None
try:
	from dropbox import Dropbox
	from dropbox.exceptions import ApiError, AuthError
	from dropbox.files import WriteMode
except ImportError:
	Dropbox, ApiError, AuthError, WriteMode = None, None, None, None
try:
	from google.oauth2 import service_account
	from googleapiclient.discovery import build
	from googleapiclient.http import MediaIoBaseDownload
except ImportError:
	service_account, build, MediaIoBaseDownload = None, None, None
# langdetect is less critical for core function if Ollama handles classification
# try:
#     from langdetect import detect
# except ImportError:
#     detect = None
try:
	from tkinterdnd2 import *  # Requires separate installation
except ImportError:
	messagebox.showerror("Error", "tkinterdnd2 not found. Please install it: pip install tkinterdnd2-universal")
	exit()
try:
	import msal
except ImportError:
	msal = None

# --- Basic Setup ---
try:
	locale.setlocale(locale.LC_ALL, '')
except locale.Error:
	logger.warning("Could not set default locale. Using system default.")

# --- Logging Setup ---
log_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
log_handler = RotatingFileHandler('sorter.log', maxBytes=1024 * 1024, backupCount=3, encoding='utf-8')
log_handler.setFormatter(log_formatter)

logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)  # Set default level
logger.addHandler(log_handler)


# Optional: Add console handler for debugging
# console_handler = logging.StreamHandler()
# console_handler.setFormatter(log_formatter)
# logger.addHandler(console_handler)
# logger.setLevel(logging.DEBUG) # Set DEBUG for more verbose output

# --- Helper for Deduplication (Top Level for Multiprocessing) ---
def process_file_for_deduplication(file_path):
	"""Обрабатывает файл для получения информации о нём в multiprocessing."""
	try:
		hasher = hashlib.md5()
		with open(file_path, 'rb') as f:
			# Read in chunks for potentially large files
			while chunk := f.read(8192):
				hasher.update(chunk)
		return (file_path, {
			"hash": hasher.hexdigest(),
			"size": os.path.getsize(file_path),
			"mod_time": os.path.getmtime(file_path),
			"name": os.path.basename(file_path)
		})
	except Exception as e:
		logger.error(f"Error processing {file_path} for deduplication: {e}")
		return (file_path, None)  # Return None on error


# --- Localization ---
def setup_localization(lang="en"):
	"""Настраивает локализацию приложения."""
	# Simplified setup assuming standard locale directory structure
	languages = {'en': 'en_US', 'ru': 'ru_RU'}
	loc = languages.get(lang, 'en_US')
	try:
		translation = gettext.translation('sorter', localedir='locale', languages=[loc], fallback=True)
		translation.install()
		return translation.gettext
	except FileNotFoundError:
		logger.warning(f"Locale directory 'locale' or translation files not found. Falling back to English.")
		return lambda s: s  # Return identity function if locales missing


_ = setup_localization("en")  # Default to English


# --- Main Application Class ---
class DocumentSorter:
	# (Keep __init__ mostly the same, just add error checks for libraries)
	def __init__(self, root, ollama_url="http://localhost:11434"):
		"""Инициализирует экземпляр класса DocumentSorter."""
		logger.info(f"DEBUG: __init__ called with ollama_url argument: {ollama_url}")
		self.root = root
		self.root.title(_("Document Sorter with Ollama"))
		self.root.geometry("768x1024")
		self.root.resizable(True, True)

		self.ollama_url = ollama_url
		logger.info(f"DEBUG: self.ollama_url immediately after assignment in __init__: {self.ollama_url}")
		self.model = "qwen2.5:7b"  # Default model
		self.available_models = []
		self.category_list = []
		self.cache = self.load_cache()
		self.language = "en"
		self.google_drive_service = None
		self.dropbox_client = None
		self.onedrive_client = None
		self.is_paused = False
		self.cancel_requested = False
		self.is_processing = False
		self.max_depth = 3  # Default max depth

		# Check required libraries
		self.check_libraries()

		self.setup_ui()
		logger.info(f"DEBUG: self.ollama_url BEFORE load_config: {self.ollama_url}")
		self.load_config()
		logger.info(f"DEBUG: self.ollama_url AFTER load_config: {self.ollama_url}")
		self.check_ollama_status()

		# Setup Drag-and-Drop
		try:
			self.root.drop_target_register(DND_FILES)
			self.root.dnd_bind('<<Drop>>', self.handle_drop)
		except tk.TclError as e:
			logger.error(f"Failed to initialize Drag and Drop (tkinterdnd2): {e}")
			messagebox.showerror(_("Initialization Error"),
								 _("Failed to set up Drag and Drop. Ensure tkinterdnd2 is correctly installed."))

		# Setup asyncio loop in a separate thread
		self.loop = asyncio.new_event_loop()
		self.loop_thread = threading.Thread(target=self.run_loop, daemon=True)
		self.loop_thread.start()
		self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

	def check_libraries(self):
		"""Checks if optional libraries for file types and cloud are loaded."""
		if not PyPDF2: logger.warning("PyPDF2 not found. PDF processing disabled. Install with: pip install pypdf2")
		if not docx: logger.warning(
			"python-docx not found. DOCX processing disabled. Install with: pip install python-docx")
		if not openpyxl: logger.warning(
			"openpyxl not found. XLSX processing disabled. Install with: pip install openpyxl")
		if not odf_text: logger.warning("odfpy not found. ODT processing disabled. Install with: pip install odfpy")
		if not Dropbox: logger.warning(
			"dropbox not found. Dropbox integration disabled. Install with: pip install dropbox")
		if not service_account: logger.warning(
			"google-api-python-client and google-auth-oauthlib not found. Google Drive disabled. Install with: pip install google-api-python-client google-auth-oauthlib google-auth-httplib2")
		if not msal: logger.warning("msal not found. OneDrive integration disabled. Install with: pip install msal")

	def run_loop(self):
		"""Runs the asyncio event loop."""
		try:
			asyncio.set_event_loop(self.loop)
			self.loop.run_forever()
		finally:
			self.loop.close()
			logger.info("Asyncio event loop closed.")

	def on_closing(self):
		"""Handles window closing."""
		logger.info("Close requested. Stopping asyncio loop and saving cache/config.")
		self.cancel_sorting(force=True)  # Attempt to cancel if running
		if self.loop.is_running():
			self.loop.call_soon_threadsafe(self.loop.stop)
		# Give the loop thread time to finish
		self.loop_thread.join(timeout=2.0)
		if self.loop_thread.is_alive():
			logger.warning("Asyncio loop thread did not exit cleanly.")
		self.save_cache()
		self.save_config()
		self.root.destroy()

	def load_cache(self):
		"""Loads classification cache."""
		cache_file = "cache.json"
		try:
			if os.path.exists(cache_file):
				with open(cache_file, 'r', encoding='utf-8') as f:
					return json.load(f)
		except (json.JSONDecodeError, IOError) as e:
			logger.error(f"Error loading cache file {cache_file}: {e}")
		return {}

	def save_cache(self):
		"""Saves classification cache."""
		try:
			with open("cache.json", 'w', encoding='utf-8') as f:
				json.dump(self.cache, f, indent=2, ensure_ascii=False)
		except IOError as e:
			logger.error(f"Error saving cache file: {e}")

	def load_config(self):
		"""Loads user configuration."""
		config_file = "config.json"
		try:
			if os.path.exists(config_file):
				with open(config_file, 'r', encoding='utf-8') as f:
					config = json.load(f)
					self.source_dir_var.set(config.get("source_dir", ""))
					self.dest_dir_var.set(config.get("dest_dir", ""))
					self.dedupe_mode.set(config.get("dedupe_mode", "none"))
					self.ollama_url = config.get("ollama_url", self.ollama_url)
					logger.info(f"DEBUG: ollama_url after loading config: {self.ollama_url}")
					# Ensure model from config exists, else reset
					loaded_model = config.get("model")
					if loaded_model and loaded_model in self.available_models:
						self.model = loaded_model
						self.model_combobox.set(self.model)
					elif self.available_models:
						self.model = self.available_models[0]
						self.model_combobox.set(self.model)
					else:
						# Keep default if no models loaded yet
						pass

					# Load categories *after* UI is setup
					if "categories" in config and config["categories"]:
						self.auto_sort_var.set(False)
						self.category_list = config["categories"]
						# Clear existing tree before loading
						for item in self.category_tree.get_children():
							self.category_tree.delete(item)
						# Use a helper to build the tree from the list
						self._rebuild_category_tree_from_list()
						self.toggle_auto_sort()  # Update button states
					else:
						self.auto_sort_var.set(True)
						self.toggle_auto_sort()

					self.max_depth_var.set(str(config.get("max_depth", 3)))
					self.max_depth = int(self.max_depth_var.get())

		except (json.JSONDecodeError, IOError) as e:
			logger.error(f"Error loading config file {config_file}: {e}")
		except Exception as e:
			logger.error(f"Unexpected error loading config: {e}\n{traceback.format_exc()}")

	def save_config(self):
		"""Saves user configuration."""
		config = {
			"source_dir": self.source_dir_var.get(),
			"dest_dir": self.dest_dir_var.get(),
			"dedupe_mode": self.dedupe_mode.get(),
			"ollama_url": self.ollama_url,
			"model": self.model,  # Save selected model
			"max_depth": self.max_depth,  # Save max depth
			# Save categories only if manual sorting is enabled
			"categories": self.category_list if not self.auto_sort_var.get() else []
		}
		try:
			with open("config.json", 'w', encoding='utf-8') as f:
				json.dump(config, f, indent=2, ensure_ascii=False)
		except IOError as e:
			logger.error(f"Error saving config file: {e}")

	def _rebuild_category_tree_from_list(self):
		"""Helper to rebuild the Treeview from self.category_list."""
		self.category_tree.delete(*self.category_tree.get_children())
		items = {}  # Keep track of inserted items by full path
		# Sort to ensure parents are created before children
		sorted_categories = sorted(self.category_list, key=lambda x: x.count('/'))
		for cat_path in sorted_categories:
			parts = cat_path.split('/')
			parent_id = ""
			current_path = ""
			for i, part in enumerate(parts):
				if current_path:
					current_path += "/" + part
				else:
					current_path = part

				if current_path not in items:
					item_id = self.category_tree.insert(parent_id, tk.END, text=part)
					items[current_path] = item_id
				parent_id = items[current_path]

	# --- UI Setup (Keep mostly the same, add library checks to buttons) ---
	def setup_ui(self):
		"""Sets up the user interface."""
		# Clear existing widgets if re-drawing (e.g., language change)
		for widget in self.root.winfo_children():
			widget.destroy()

		# --- Menu ---
		menubar = tk.Menu(self.root)
		self.root.config(menu=menubar)
		lang_menu = tk.Menu(menubar, tearoff=0)
		menubar.add_cascade(label=_("Language"), menu=lang_menu)
		lang_menu.add_command(label="English", command=lambda: self.change_language("en"))
		lang_menu.add_command(label="Русский", command=lambda: self.change_language("ru"))
		# Add settings menu maybe?
		settings_menu = tk.Menu(menubar, tearoff=0)
		menubar.add_cascade(label=_("Settings"), menu=settings_menu)
		settings_menu.add_command(label=_("Set Ollama URL"), command=self.set_ollama_url)

		# --- Main Frame ---
		main_frame = ttk.Frame(self.root, padding="10")
		main_frame.pack(fill=tk.BOTH, expand=True)
		main_frame.columnconfigure(0, weight=1)  # Make content expand horizontally

		# --- Ollama Status & Model ---
		status_model_frame = ttk.Frame(main_frame)
		status_model_frame.pack(fill=tk.X, pady=5)
		status_model_frame.columnconfigure(1, weight=1)  # Make combobox expand

		ttk.Label(status_model_frame, text=_("Ollama Status:")).grid(row=0, column=0, padx=5, sticky=tk.W)
		self.status_label = ttk.Label(status_model_frame, text=_("Checking..."), foreground="orange",
									  width=25)  # Fixed width
		self.status_label.grid(row=0, column=1, padx=5, sticky=tk.W)

		ttk.Label(status_model_frame, text=_("Select Model:")).grid(row=1, column=0, padx=5, sticky=tk.W)
		self.model_combobox = ttk.Combobox(status_model_frame, state="readonly",
										   postcommand=self.fetch_models)  # Refresh on dropdown
		self.model_combobox.grid(row=1, column=1, padx=5, pady=5, sticky=tk.EW)
		self.model_combobox.bind("<<ComboboxSelected>>", self.on_model_selected)
		refresh_button = ttk.Button(status_model_frame, text=_("Refresh"), command=self.fetch_models,
									width=10)  # Smaller button
		refresh_button.grid(row=1, column=2, padx=5, pady=5)

		# --- Cloud Storage Frame ---
		cloud_frame = ttk.LabelFrame(main_frame, text=_("Cloud Storage"), padding="10")
		cloud_frame.pack(fill=tk.X, pady=10)
		# Disable buttons if libraries are missing
		gdrive_state = tk.NORMAL if service_account and build else tk.DISABLED
		dropbox_state = tk.NORMAL if Dropbox else tk.DISABLED
		onedrive_state = tk.NORMAL if msal else tk.DISABLED
		ttk.Button(cloud_frame, text=_("Connect Google Drive"), command=self.connect_google_drive,
				   state=gdrive_state).pack(side=tk.LEFT, padx=5)
		ttk.Button(cloud_frame, text=_("Connect Dropbox"), command=self.connect_dropbox, state=dropbox_state).pack(
			side=tk.LEFT, padx=5)
		ttk.Button(cloud_frame, text=_("Connect OneDrive"), command=self.connect_onedrive, state=onedrive_state).pack(
			side=tk.LEFT, padx=5)

		# --- Directory Selection ---
		dir_frame = ttk.LabelFrame(main_frame, text=_("Directory Selection"), padding="10")
		dir_frame.pack(fill=tk.X, pady=10)
		dir_frame.columnconfigure(1, weight=1)

		ttk.Label(dir_frame, text=_("Source Directory:")).grid(row=0, column=0, sticky=tk.W, pady=2)
		self.source_dir_var = tk.StringVar()
		ttk.Entry(dir_frame, textvariable=self.source_dir_var, width=60).grid(row=0, column=1, padx=5, pady=2,
																			  sticky=tk.EW)
		ttk.Button(dir_frame, text=_("Browse"), command=self.browse_source_dir).grid(row=0, column=2, padx=5, pady=2)

		ttk.Label(dir_frame, text=_("Destination Directory:")).grid(row=1, column=0, sticky=tk.W, pady=2)
		self.dest_dir_var = tk.StringVar()
		ttk.Entry(dir_frame, textvariable=self.dest_dir_var, width=60).grid(row=1, column=1, padx=5, pady=2,
																			sticky=tk.EW)
		ttk.Button(dir_frame, text=_("Browse"), command=self.browse_dest_dir).grid(row=1, column=2, padx=5, pady=2)

		# --- Categories Frame ---
		category_frame = ttk.LabelFrame(main_frame, text=_("Categories & Sorting"), padding="10")
		category_frame.pack(fill=tk.BOTH, expand=True, pady=10)
		category_frame.rowconfigure(1, weight=1)  # Make treeview expand vertically
		category_frame.columnconfigure(0, weight=1)  # Make treeview expand horizontally

		# Auto-sort options row
		auto_frame = ttk.Frame(category_frame)
		auto_frame.grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=5)
		self.auto_sort_var = tk.BooleanVar(value=True)
		self.auto_sort_check = ttk.Checkbutton(auto_frame, text=_("Automatic Category Generation (via Ollama)"),
											   variable=self.auto_sort_var, command=self.toggle_auto_sort)
		self.auto_sort_check.pack(side=tk.LEFT, padx=5)
		ttk.Label(auto_frame, text=_("Max Depth:")).pack(side=tk.LEFT, padx=5)
		self.max_depth_var = tk.StringVar(value="3")
		self.max_depth_entry = ttk.Entry(auto_frame, textvariable=self.max_depth_var, width=3)
		self.max_depth_entry.pack(side=tk.LEFT, padx=5)
		# Add validation command later if needed

		# Category Treeview
		self.category_tree = ttk.Treeview(category_frame, height=8)  # Adjusted height
		self.category_tree.grid(row=1, column=0, pady=5, sticky="nsew")
		# Add scrollbar for Treeview
		tree_scrollbar = ttk.Scrollbar(category_frame, orient="vertical", command=self.category_tree.yview)
		self.category_tree.configure(yscrollcommand=tree_scrollbar.set)
		tree_scrollbar.grid(row=1, column=1, sticky="ns", pady=5)

		# Category buttons row
		category_buttons_frame = ttk.Frame(category_frame)
		category_buttons_frame.grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=5)
		self.add_category_btn = ttk.Button(category_buttons_frame, text=_("Add Root Cat."), command=self.add_category,
										   state=tk.DISABLED)
		self.add_category_btn.pack(side=tk.LEFT, padx=5)
		self.add_subcategory_btn = ttk.Button(category_buttons_frame, text=_("Add Subcat."),
											  command=self.add_subcategory, state=tk.DISABLED)
		self.add_subcategory_btn.pack(side=tk.LEFT, padx=5)
		self.remove_category_btn = ttk.Button(category_buttons_frame, text=_("Remove Sel."),
											  command=self.remove_category, state=tk.DISABLED)
		self.remove_category_btn.pack(side=tk.LEFT, padx=5)

		# --- Deduplication Frame ---
		dedupe_frame = ttk.LabelFrame(main_frame, text=_("Duplicate Handling"), padding="10")
		dedupe_frame.pack(fill=tk.X, pady=10)
		self.dedupe_mode = tk.StringVar(value="none")
		ttk.Radiobutton(dedupe_frame, text=_("Keep All"), value="none", variable=self.dedupe_mode).pack(side=tk.LEFT,
																										padx=10)
		ttk.Radiobutton(dedupe_frame, text=_("Remove Exact Duplicates (Hash)"), value="normal",
						variable=self.dedupe_mode).pack(side=tk.LEFT, padx=10)
		ttk.Radiobutton(dedupe_frame, text=_("Remove Duplicates (Name+Size, Keep Newest)"), value="hardcore",
						variable=self.dedupe_mode).pack(side=tk.LEFT, padx=10)

		# --- Log Frame ---
		log_frame = ttk.LabelFrame(main_frame, text=_("Log"), padding="10")
		log_frame.pack(fill=tk.BOTH, expand=True, pady=10)
		log_frame.rowconfigure(0, weight=1)
		log_frame.columnconfigure(0, weight=1)
		self.log_text = tk.Text(log_frame, height=8, state=tk.DISABLED, wrap=tk.WORD)  # Adjusted height, word wrap
		self.log_text.grid(row=0, column=0, sticky="nsew", pady=5)
		# Add scrollbar for Log
		log_scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
		self.log_text.configure(yscrollcommand=log_scrollbar.set)
		log_scrollbar.grid(row=0, column=1, sticky="ns", pady=5)

		ttk.Button(log_frame, text=_("Export Log"), command=self.export_log).grid(row=1, column=0, columnspan=2, pady=5)

		# --- Progress Bar ---
		self.progress_var = tk.DoubleVar()
		self.progress_bar = ttk.Progressbar(main_frame, variable=self.progress_var, maximum=100)
		self.progress_bar.pack(fill=tk.X, pady=5)

		# --- Control Buttons ---
		button_frame = ttk.Frame(main_frame, padding="5 0 0 0")  # Padding top
		button_frame.pack(fill=tk.X)
		# Right-align buttons
		self.backup_button = ttk.Button(button_frame, text=_("Create Backup"), command=self.create_backup)
		self.backup_button.pack(side=tk.LEFT, padx=5)

		self.cancel_button = ttk.Button(button_frame, text=_("Cancel"), command=self.cancel_sorting, state=tk.DISABLED)
		self.cancel_button.pack(side=tk.RIGHT, padx=5)
		self.pause_button = ttk.Button(button_frame, text=_("Pause"), command=self.pause_sorting, state=tk.DISABLED)
		self.pause_button.pack(side=tk.RIGHT, padx=5)
		self.sort_button = ttk.Button(button_frame, text=_("Start Sorting"), command=self.start_sorting)
		self.sort_button.pack(side=tk.RIGHT, padx=5)

		# Apply initial state based on auto_sort
		self.toggle_auto_sort()

	def toggle_auto_sort(self):
		"""Enables/disables category editing based on auto_sort setting."""
		is_auto = self.auto_sort_var.get()
		cat_button_state = tk.DISABLED if is_auto else tk.NORMAL
		tree_state = tk.DISABLED if is_auto else tk.NORMAL  # Maybe allow viewing tree in auto? Let's disable for now.
		depth_entry_state = tk.NORMAL if is_auto else tk.DISABLED

		self.add_category_btn.config(state=cat_button_state)
		self.add_subcategory_btn.config(state=cat_button_state)
		self.remove_category_btn.config(state=cat_button_state)
		# self.category_tree.config(state=tree_state) # Treeview doesn't have a simple state=DISABLED
		# For Treeview, we just prevent edits via button states. We clear it if Auto is checked.
		self.max_depth_entry.config(state=depth_entry_state)

		if is_auto:
			# Clear manual categories when switching to auto
			self.category_list.clear()
			self.category_tree.delete(*self.category_tree.get_children())
		else:
			# Optional: Load last known manual categories if switching back?
			# For now, just enable buttons. User needs to add cats manually.
			self._rebuild_category_tree_from_list()  # Rebuild tree if switching to manual

	# No need to save config here, will be saved on exit or relevant actions

	def change_language(self, lang):
		"""Changes the application language."""
		global _
		logger.info(f"Changing language to: {lang}")
		self.language = lang
		_ = setup_localization(lang)
		# Save current config before redrawing UI
		self.save_config()
		# Re-setup UI
		self.setup_ui()
		# Reload config to populate new UI elements correctly
		self.load_config()
		# Re-check status as labels are recreated
		self.check_ollama_status()

	def set_ollama_url(self):
		"""Opens a dialog to set the Ollama API URL."""
		new_url = simpledialog.askstring(
			_("Ollama Base URL"),  # Changed title
			_("Enter the BASE URL for the Ollama server (e.g., http://localhost:11434):"),  # Changed prompt
			initialvalue=self.ollama_url
		)
		if new_url:
			# Basic validation
			if new_url.startswith("http://") or new_url.startswith("https://"):
				self.ollama_url = new_url.strip('/')  # Remove trailing slash
				logger.info(f"Ollama URL set to: {self.ollama_url}")
				self.log_message(f"Ollama URL set to: {self.ollama_url}")
				self.check_ollama_status()  # Check connection with new URL
				self.save_config()  # Save the new URL
			else:
				messagebox.showerror(_("Invalid URL"), _("URL must start with http:// or https://"))

	# --- Cloud Connection Methods (Add library checks) ---
	def connect_google_drive(self):
		"""Connects to Google Drive."""
		if not service_account or not build:
			messagebox.showerror(_("Error"), _("Google Drive library not found. See logs for install instructions."))
			return
		# Use filedialog to ask for credentials.json
		creds_path = filedialog.askopenfilename(
			title=_("Select Google Drive credentials.json"),
			filetypes=[("JSON files", "*.json")]
		)
		if not creds_path:
			return  # User cancelled

		try:
			creds = service_account.Credentials.from_service_account_file(creds_path, scopes=[
				"https://www.googleapis.com/auth/drive"])  # Full control scope needed for upload/create
			self.google_drive_service = build('drive', 'v3', credentials=creds)
			# Perform a simple test call
			self.google_drive_service.files().list(pageSize=1, fields="files(id)").execute()
			logger.info(_("Successfully connected to Google Drive"))
			self.log_message(_("Successfully connected to Google Drive"))
		except FileNotFoundError:
			messagebox.showerror(_("Error"), _(f"Credentials file not found at: {creds_path}"))
			logger.error(_(f"Google Drive credentials file not found: {creds_path}"))
		except Exception as e:
			logger.error(_(f"Google Drive connection error: {str(e)}"), exc_info=True)
			self.log_message(_(f"Google Drive connection error: Check logs for details."))
			messagebox.showerror(_("Google Drive Error"), _(f"Connection failed: {str(e)}"))
			self.google_drive_service = None  # Reset on failure

	def connect_dropbox(self):
		"""Connects to Dropbox."""
		if not Dropbox:
			messagebox.showerror(_("Error"), _("Dropbox library not found. See logs for install instructions."))
			return
		token = simpledialog.askstring(_("Dropbox"), _("Enter Dropbox Access Token:"))
		if token:
			try:
				self.dropbox_client = Dropbox(token)
				# Test connection
				self.dropbox_client.users_get_current_account()
				logger.info(_("Successfully connected to Dropbox"))
				self.log_message(_("Successfully connected to Dropbox"))
			except AuthError as e:
				logger.error(_(f"Dropbox authentication error: {str(e)}"), exc_info=True)
				self.log_message(_(f"Dropbox authentication error: Invalid token?"))
				messagebox.showerror(_("Dropbox Error"), _("Authentication failed. Please check your token."))
				self.dropbox_client = None
			except Exception as e:
				logger.error(_(f"Dropbox connection error: {str(e)}"), exc_info=True)
				self.log_message(_(f"Dropbox connection error: Check logs."))
				messagebox.showerror(_("Dropbox Error"), _(f"Connection failed: {str(e)}"))
				self.dropbox_client = None

	def connect_onedrive(self):
		"""Connects to OneDrive."""
		if not msal:
			messagebox.showerror(_("Error"), _("MSAL library not found. See logs for install instructions."))
			return
		# Consider a more robust OAuth flow later if needed. For now, stick to client credentials.
		client_id = simpledialog.askstring(_("OneDrive"), _("Enter Azure App Client ID:"))
		client_secret = simpledialog.askstring(_("OneDrive"), _("Enter Azure App Client Secret:"), show='*')
		tenant_id = simpledialog.askstring(_("OneDrive"), _("Enter Azure Tenant ID (or 'common' for multi-tenant):"),
										   initialvalue="common")

		if client_id and client_secret and tenant_id:
			try:
				authority = f"https://login.microsoftonline.com/{tenant_id}"
				app = msal.ConfidentialClientApplication(
					client_id,
					authority=authority,
					client_credential=client_secret
				)
				# Acquire token for Microsoft Graph API
				result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])

				if "access_token" in result:
					self.onedrive_client = {"token": result["access_token"]}
					# Test connection (optional: get drive info)
					# headers = {"Authorization": f"Bearer {self.onedrive_client['token']}"}
					# test_url = "https://graph.microsoft.com/v1.0/me/drive"
					# response = requests.get(test_url, headers=headers)
					# if response.status_code == 200:
					#      logger.info(f"OneDrive Drive ID: {response.json().get('id')}")
					# else:
					#      logger.warning(f"OneDrive test call failed: {response.status_code} - {response.text}")

					logger.info(_("Successfully connected to OneDrive"))
					self.log_message(_("Successfully connected to OneDrive"))
				else:
					error_desc = result.get("error_description", "No access token received.")
					logger.error(_(f"OneDrive authentication failed: {error_desc}"))
					self.log_message(_(f"OneDrive authentication failed: {error_desc}"))
					messagebox.showerror(_("OneDrive Error"), _(f"Authentication failed: {error_desc}"))
					self.onedrive_client = None

			except Exception as e:
				logger.error(_(f"OneDrive connection error: {str(e)}"), exc_info=True)
				self.log_message(_(f"OneDrive connection error: Check logs."))
				messagebox.showerror(_("OneDrive Error"), _(f"Connection failed: {str(e)}"))
				self.onedrive_client = None

	def handle_drop(self, event):
		"""Handles dropped files/dirs."""
		# TkinterDND returns a string, potentially with spaces and braces for multiple files
		# We are interested only if it's a single directory
		dropped = event.data.strip()
		# Simple check: if it contains spaces and isn't clearly quoted, it might be multiple files
		if '{' in dropped or '}' in dropped:
			# Could try to parse, but safer to just ignore multi-file drops on the main window
			self.log_message(_("Multiple items dropped. Please drop a single source directory."))
			logger.warning(f"Ignoring multi-item drop: {event.data}")
			return

		# Remove potential surrounding braces added by TkinterDND on some platforms
		dropped = dropped.strip('{}')

		if os.path.isdir(dropped):
			self.source_dir_var.set(dropped)
			logger.info(_(f"Source directory set by drop: {dropped}"))
			self.log_message(_(f"Source directory set: {dropped}"))
			self.save_config()
		else:
			self.log_message(_("Dropped item is not a directory."))
			logger.warning(f"Ignoring non-directory drop: {dropped}")

	# --- Ollama Interaction (Increased Timeouts, Better Error Handling) ---
	def check_ollama_status(self):
		"""Checks Ollama API status."""
		if not hasattr(self, 'status_label') or not self.status_label.winfo_exists():
			return  # UI not ready yet

		url = f"{self.ollama_url.strip('/')}/api/tags"  # Use /api/tags which lists models
		try:
			# Use a short timeout for status check
			response = requests.get(url, timeout=5)
			if response.status_code == 200:
				self.status_label.config(text=_("Connected"), foreground="green")
			# Don't fetch models here, let postcommand or refresh button do it
			# self.fetch_models()
			else:
				self.status_label.config(text=_("Error: API Status {status}").format(status=response.status_code),
										 foreground="red")
				logger.warning(f"Ollama API check failed: {response.status_code} - {response.text[:100]}")
		except requests.exceptions.ConnectionError:
			self.status_label.config(text=_("Disconnected"), foreground="red")
		# logger.warning(_("Cannot connect to Ollama API. Is it running?")) # Less verbose logging
		except requests.exceptions.Timeout:
			self.status_label.config(text=_("Timeout"), foreground="orange")
			logger.warning(_("Ollama API check timed out."))
		except Exception as e:
			self.status_label.config(text=_("Error"), foreground="red")
			logger.error(f"Error checking Ollama status: {e}", exc_info=False)  # Avoid stack trace for common errors

	# # Schedule next check (optional, can be resource intensive)
	# self.root.after(15000, self.check_ollama_status) # Check every 15s

	def fetch_models(self):
		"""Fetches available models from Ollama."""
		if not hasattr(self, 'model_combobox') or not self.model_combobox.winfo_exists():
			return  # UI not ready yet

		url = f"{self.ollama_url.strip('/')}/api/tags"
		logger.info(f"DEBUG: Using Ollama URL base: {self.ollama_url}")  # <-- Add this
		logger.info(f"Fetching models from {url}...")  # Keep this
		try:
			# Slightly longer timeout for fetching models
			response = requests.get(url, timeout=10)
			if response.status_code == 200:
				models_data = response.json()
				self.available_models = sorted([model["name"] for model in models_data.get("models", [])])
				self.model_combobox["values"] = self.available_models
				logger.info(f"Found models: {self.available_models}")

				# Preserve current selection if possible, otherwise select first
				current_selection = self.model_combobox.get()
				if current_selection in self.available_models:
					self.model_combobox.set(current_selection)
					self.model = current_selection
				elif self.model in self.available_models:
					self.model_combobox.set(self.model)
				elif self.available_models:
					self.model_combobox.set(self.available_models[0])
					self.model = self.available_models[0]
				else:
					logger.warning("No models found in Ollama response.")
					self.model_combobox.set("")  # Clear selection

				self.status_label.config(text=_("Connected"), foreground="green")  # Update status on success
			else:
				logger.error(f"Failed to fetch models: {response.status_code} - {response.text[:100]}")
				self.log_message(_("Failed to fetch models from Ollama."))
				self.status_label.config(text=_("Error Fetching"), foreground="red")  # Indicate fetch error

		except requests.exceptions.RequestException as e:
			logger.error(f"Error fetching Ollama models: {e}", exc_info=False)
			self.log_message(_("Error connecting to Ollama to fetch models."))
			self.status_label.config(text=_("Connection Error"), foreground="red")
			# Clear models if connection fails
			self.available_models = []
			self.model_combobox["values"] = []
			self.model_combobox.set("")

	def on_model_selected(self, event=None):
		"""Handles model selection."""
		new_model = self.model_combobox.get()
		if new_model and new_model != self.model:
			self.model = new_model
			logger.info(_(f"Selected model: {self.model}"))
			self.log_message(_(f"Selected model: {self.model}"))
			self.save_config()  # Save the newly selected model

	# --- Directory Browsing ---
	def browse_source_dir(self):
		dir_path = filedialog.askdirectory(title=_("Select Source Directory"), initialdir=self.source_dir_var.get())
		if dir_path:
			self.source_dir_var.set(dir_path)
			self.save_config()

	def browse_dest_dir(self):
		dir_path = filedialog.askdirectory(title=_("Select Destination Directory"), initialdir=self.dest_dir_var.get())
		if dir_path:
			self.dest_dir_var.set(dir_path)
			self.save_config()

	# --- Category Management (Improved tree handling) ---
	def _get_full_path_from_tree_item(self, item_id):
		"""Gets the full category path string from a Treeview item ID."""
		path = self.category_tree.item(item_id, "text")
		parent_id = self.category_tree.parent(item_id)
		while parent_id:
			path = self.category_tree.item(parent_id, "text") + "/" + path
			parent_id = self.category_tree.parent(parent_id)
		return path

	def add_category(self):
		"""Adds a new root category."""
		category = simpledialog.askstring(_("Add Root Category"), _("Enter category name:"))
		if category:
			category = category.strip().replace('/', '-')  # Sanitize name
			if not category: return

			# Check if root category already exists
			for item_id in self.category_tree.get_children(""):
				if self.category_tree.item(item_id, "text") == category:
					messagebox.showwarning(_("Duplicate"), _("This root category already exists."))
					return

			new_id = self.category_tree.insert("", tk.END, text=category)
			self.category_list.append(category)
			self.category_list.sort()  # Keep list sorted
			self.save_config()
			logger.info(f"Added root category: {category}")

	def add_subcategory(self):
		"""Adds a subcategory to the selected category."""
		selected_ids = self.category_tree.selection()
		if not selected_ids:
			messagebox.showwarning(_("Warning"), _("Select a parent category first."))
			return
		selected_id = selected_ids[0]  # Use the first selected item

		parent_path = self._get_full_path_from_tree_item(selected_id)
		parent_text = self.category_tree.item(selected_id, "text")

		# Check depth limit
		current_depth = parent_path.count('/')
		# Load max_depth from config or default
		try:
			max_d = int(self.max_depth_var.get()) if not self.auto_sort_var.get() else self.max_depth
		except ValueError:
			max_d = 3  # Fallback
		if current_depth >= max_d:
			messagebox.showwarning(_("Depth Limit"),
								   _("Maximum subcategory depth ({max_d}) reached.").format(max_d=max_d))
			return

		subcategory = simpledialog.askstring(_("Add Subcategory"), _(f"Enter subcategory name for '{parent_text}':"))
		if subcategory:
			subcategory = subcategory.strip().replace('/', '-')  # Sanitize name
			if not subcategory: return

			# Check if subcategory already exists under this parent
			for item_id in self.category_tree.get_children(selected_id):
				if self.category_tree.item(item_id, "text") == subcategory:
					messagebox.showwarning(_("Duplicate"), _("This subcategory already exists here."))
					return

			full_path = f"{parent_path}/{subcategory}"
			new_id = self.category_tree.insert(selected_id, tk.END, text=subcategory)
			self.category_list.append(full_path)
			self.category_list.sort()  # Keep list sorted
			self.save_config()
			logger.info(f"Added subcategory: {full_path}")

	def remove_category(self):
		"""Removes the selected category and its children."""
		selected_ids = self.category_tree.selection()
		if not selected_ids:
			messagebox.showwarning(_("Warning"), _("Select a category or subcategory to remove."))
			return
		selected_id = selected_ids[0]

		full_path_to_remove = self._get_full_path_from_tree_item(selected_id)

		if messagebox.askyesno(_("Confirm Removal"),
							   _("Are you sure you want to remove '{path}' and all its subcategories?").format(
									   path=full_path_to_remove)):
			# Recursively collect all paths to remove
			paths_to_remove = set()
			items_to_process = [selected_id]
			while items_to_process:
				current_id = items_to_process.pop(0)
				paths_to_remove.add(self._get_full_path_from_tree_item(current_id))
				items_to_process.extend(self.category_tree.get_children(current_id))

			# Remove from list
			self.category_list = [cat for cat in self.category_list if cat not in paths_to_remove]

			# Remove from tree
			self.category_tree.delete(selected_id)

			self.category_list.sort()  # Keep list sorted
			self.save_config()
			logger.info(f"Removed category and subcategories starting from: {full_path_to_remove}")

	# --- Logging & Reporting ---
	def log_message(self, message):
		"""Adds a message to the GUI log."""

		# Ensure this runs in the main Tkinter thread
		def update_log():
			if not hasattr(self, 'log_text') or not self.log_text.winfo_exists():
				return  # Avoid errors if UI closed prematurely
			try:
				self.log_text.config(state=tk.NORMAL)
				timestamp = time.strftime('%H:%M:%S')
				self.log_text.insert(tk.END, f"{timestamp} - {message}\n")
				self.log_text.see(tk.END)  # Scroll to the end
				self.log_text.config(state=tk.DISABLED)
			except tk.TclError as e:
				# Can happen if widget is destroyed during update
				logger.warning(f"GUI log update failed: {e}")

		if threading.current_thread() is threading.main_thread():
			update_log()
		else:
			# Schedule the update in the main thread
			self.root.after(0, update_log)

	def export_log(self):
		"""Exports the GUI log content."""
		try:
			log_content = self.log_text.get("1.0", tk.END)
			file_path = filedialog.asksaveasfilename(
				defaultextension=".txt",
				filetypes=[("Text files", "*.txt"), ("Log files", "*.log")],
				title=_("Export Log As")
			)
			if file_path:
				with open(file_path, "w", encoding="utf-8") as f:
					f.write(log_content)
				logger.info(_(f"GUI log exported to {file_path}"))
				self.log_message(_(f"Log exported to {file_path}"))
		except Exception as e:
			logger.error(f"Failed to export log: {e}", exc_info=True)
			messagebox.showerror(_("Error"), _("Failed to export log. See application log file for details."))

	def generate_report(self, stats):
		"""Generates an HTML report."""
		try:
			env = Environment(loader=FileSystemLoader('.'))
			# More detailed template
			template_str = """
            <!DOCTYPE html>
            <html lang="{{ lang_code }}">
            <head>
                <meta charset="UTF-8">
                <title>{{ _('Sorting Report') }}</title>
                <style>
                    body { font-family: sans-serif; margin: 20px; }
                    h1 { color: #333; }
                    p { margin: 5px 0; }
                    strong { color: #555; }
                    .summary { border: 1px solid #ccc; padding: 15px; margin-top: 15px; background-color: #f9f9f9; }
                </style>
            </head>
            <body>
                <h1>{{ _('Document Sorter Report') }}</h1>
                <p><strong>{{ _('Report Generated:') }}</strong> {{ timestamp }}</p>
                <p><strong>{{ _('Source Directory:') }}</strong> {{ source_dir }}</p>
                <p><strong>{{ _('Destination Directory:') }}</strong> {{ dest_dir }}</p>

                <div class="summary">
                    <h2>{{ _('Summary') }}</h2>
                    <p><strong>{{ _('Total Files Processed:') }}</strong> {{ stats.processed_files }}</p>
                    <p><strong>{{ _('Categories Used:') }}</strong> {{ stats.categories_used }}</p>
                    <p><strong>{{ _('Duplicate Files Removed:') }}</strong> {{ stats.duplicates_removed }} (Mode: {{ dedupe_mode }})</p>
                    <p><strong>{{ _('Sorting Mode:') }}</strong> {{ sorting_mode }}</p>
                    <p><strong>{{ _('Elapsed Time:') }}</strong> {{ stats.elapsed_time }}</p>
                </div>

                <!-- Optional: Add list of categories created/used? -->
                <!-- Optional: Add list of moved files? (Could be very long) -->

            </body>
            </html>
            """
			template = env.from_string(template_str)
			# Add more context to stats
			stats['timestamp'] = time.strftime('%Y-%m-%d %H:%M:%S')
			stats['source_dir'] = self.source_dir_var.get()
			stats['dest_dir'] = self.dest_dir_var.get()
			stats['dedupe_mode'] = self.dedupe_mode.get()
			stats['sorting_mode'] = _("Automatic") if self.auto_sort_var.get() else _("Manual Categories")
			# Pass translation function and lang code to template
			report_html = template.render(
				stats=stats,
				lang_code=self.language,
				_=_)  # Pass translation func

			report_filename = "sorting_report.html"
			with open(report_filename, "w", encoding="utf-8") as f:
				f.write(report_html)
			logger.info(_(f"Report generated: {report_filename}"))
			self.log_message(_(f"Report generated: {report_filename}"))
			# Ask user if they want to open the report
			if messagebox.askyesno(_("Report Generated"), _("Sorting report saved as {filename}. Open it now?").format(
					filename=report_filename)):
				import webbrowser
				webbrowser.open(f"file://{os.path.abspath(report_filename)}")

		except Exception as e:
			logger.error(f"Failed to generate report: {e}", exc_info=True)
			self.log_message(_("Failed to generate report. See logs."))

	# --- Backup ---
	def create_backup(self):
		"""Creates a ZIP backup of the source directory."""
		source_dir = self.source_dir_var.get()
		if not source_dir or not os.path.isdir(source_dir):
			messagebox.showerror(_("Error"), _("Please select a valid source directory first."))
			return

		# Suggest a filename based on source dir name and date
		default_filename = f"backup_{os.path.basename(source_dir)}_{time.strftime('%Y%m%d_%H%M%S')}.zip"
		backup_path = filedialog.asksaveasfilename(
			defaultextension=".zip",
			filetypes=[("ZIP files", "*.zip")],
			title=_("Save Backup As"),
			initialfile=default_filename
		)

		if backup_path:
			logger.info(f"Starting backup of '{source_dir}' to '{backup_path}'...")
			self.log_message(_("Starting backup... This may take a while."))
			self.progress_var.set(0)  # Use progress bar for backup too
			self.sort_button.config(state=tk.DISABLED)
			self.backup_button.config(state=tk.DISABLED)

			# Run backup in a separate thread to keep UI responsive
			backup_thread = threading.Thread(target=self._execute_backup, args=(source_dir, backup_path), daemon=True)
			backup_thread.start()

	def _execute_backup(self, source_dir, backup_path):
		"""Actual backup logic running in a thread."""
		total_files = sum(len(files) for _, _, files in os.walk(source_dir))
		files_added = 0
		try:
			with zipfile.ZipFile(backup_path, 'w', zipfile.ZIP_DEFLATED, allowZip64=True) as zipf:
				for root, _, files in os.walk(source_dir):
					for file in files:
						if self.cancel_requested:  # Check for cancellation
							raise InterruptedError("Backup cancelled by user.")
						file_path = os.path.join(root, file)
						arcname = os.path.relpath(file_path, source_dir)
						try:
							zipf.write(file_path, arcname)
							files_added += 1
							# Update progress (less frequently to avoid GUI overload)
							if files_added % 50 == 0 or files_added == total_files:
								progress = (files_added / total_files) * 100 if total_files > 0 else 100
								self.root.after(0, self.progress_var.set, progress)
						except Exception as write_err:
							logger.warning(f"Could not add file to backup: {file_path} - {write_err}")
						# Optionally log to GUI as well
						# self.log_message(_("Warning: Could not back up {file}").format(file=os.path.basename(file_path)))

			logger.info(_(f"Backup successfully created at {backup_path}"))
			self.log_message(
				_("Backup created successfully: {filename}").format(filename=os.path.basename(backup_path)))

		except InterruptedError:
			logger.warning(f"Backup cancelled. Partial backup might exist at {backup_path}")
			self.log_message(_("Backup cancelled."))
		# Optionally remove partial backup? os.remove(backup_path)

		except Exception as e:
			logger.error(f"Backup failed: {e}", exc_info=True)
			self.log_message(_("Backup failed. See logs for details."))
			messagebox.showerror(_("Backup Error"), _("Failed to create backup. Check logs."))
			# Attempt to remove potentially corrupt zip file
			if os.path.exists(backup_path):
				try:
					os.remove(backup_path)
				except OSError:
					pass
		finally:
			# Reset UI elements in the main thread
			def reset_ui():
				self.progress_var.set(0)
				self.sort_button.config(state=tk.NORMAL)
				self.backup_button.config(state=tk.NORMAL)
				self.cancel_requested = False  # Reset cancel flag

			self.root.after(0, reset_ui)

	# --- Sorting Control ---
	def start_sorting(self):
		"""Starts the sorting process."""
		if self.is_processing:
			logger.warning("Sorting process already running.")
			return

		# --- Pre-flight Checks ---
		source_dir = self.source_dir_var.get()
		dest_dir = self.dest_dir_var.get()

		if not source_dir or not os.path.isdir(source_dir):
			messagebox.showerror(_("Error"), _("Please select a valid source directory."))
			return
		if not dest_dir or not os.path.isdir(dest_dir):
			# Ask to create dest_dir?
			if messagebox.askyesno(_("Create Directory?"), _("Destination directory does not exist. Create it?")):
				try:
					os.makedirs(dest_dir, exist_ok=True)
					logger.info(f"Created destination directory: {dest_dir}")
				except OSError as e:
					messagebox.showerror(_("Error"),
										 _("Could not create destination directory: {error}").format(error=e))
					logger.error(f"Failed to create destination directory {dest_dir}: {e}")
					return
			else:
				messagebox.showerror(_("Error"), _("Please select or create a valid destination directory."))
				return

		if source_dir == dest_dir or os.path.abspath(source_dir) == os.path.abspath(dest_dir):
			messagebox.showerror(_("Error"), _("Source and destination directories cannot be the same."))
			return
		# Check if dest is inside source (dangerous)
		if os.path.abspath(dest_dir).startswith(os.path.abspath(source_dir) + os.sep):
			messagebox.showerror(_("Error"), _("Destination directory cannot be inside the source directory."))
			return

		# Check categories if manual sorting
		is_auto = self.auto_sort_var.get()
		if not is_auto and not self.category_list:
			messagebox.showerror(_("Error"),
								 _("Manual sorting selected, but no categories defined. Please add categories first."))
			return

		# Check Ollama connection only if needed (auto-sort or classification)
		# Let the check_ollama_status handle the visual indication. Maybe add check here?
		if self.status_label.cget("foreground") == "red":
			if messagebox.askretrycancel(_("Ollama Disconnected"),
										 _("Cannot connect to Ollama. Sorting might fail or use basic fallback. Retry connection check?")):
				self.check_ollama_status()  # Try again
				return  # Let user retry starting if connection succeeds
			else:
				# Allow proceeding but warn
				logger.warning("Proceeding with sorting despite Ollama connection issues.")
				self.log_message(_("Warning: Ollama disconnected. Classification may fail."))

		# Parse max_depth
		try:
			self.max_depth = int(self.max_depth_var.get())
			if self.max_depth < 1: raise ValueError("Depth must be >= 1")
		except ValueError:
			messagebox.showerror(_("Error"), _("Invalid Max Subcategory Depth. Must be a number >= 1."))
			return

		# --- Start Processing ---
		logger.info("=" * 20 + " Starting Sorting Process " + "=" * 20)
		self.log_message(_("Sorting started..."))
		self.save_config()  # Save settings before starting

		self.is_processing = True
		self.cancel_requested = False
		self.is_paused = False

		# Update button states
		self.sort_button.config(state=tk.DISABLED)
		self.backup_button.config(state=tk.DISABLED)
		self.pause_button.config(state=tk.NORMAL, text=_("Pause"))
		self.cancel_button.config(state=tk.NORMAL)
		# Disable config changes during sorting
		self.auto_sort_check.config(state=tk.DISABLED)
		self.max_depth_entry.config(state=tk.DISABLED)
		self.add_category_btn.config(state=tk.DISABLED)
		self.add_subcategory_btn.config(state=tk.DISABLED)
		self.remove_category_btn.config(state=tk.DISABLED)

		self.progress_var.set(0)

		# Run the main sorting logic in a separate thread
		thread = threading.Thread(target=self.sort_documents, args=(source_dir, dest_dir), daemon=True)
		thread.start()

	def pause_sorting(self):
		"""Pauses or resumes the sorting process."""
		if not self.is_processing:
			return

		self.is_paused = not self.is_paused
		new_text = _("Resume") if self.is_paused else _("Pause")
		self.pause_button.config(text=new_text)
		status_msg = _("Sorting paused.") if self.is_paused else _("Sorting resumed.")
		logger.info(status_msg)
		self.log_message(status_msg)

	def cancel_sorting(self, force=False):
		"""Requests cancellation of the sorting process."""
		if not self.is_processing and not force:
			return

		if force or messagebox.askyesno(_("Confirm Cancellation"),
										_("Are you sure you want to cancel the current sorting process?")):
			self.cancel_requested = True
			logger.info("Cancellation requested by user.")
			self.log_message(_("Cancellation requested..."))
			self.cancel_button.config(state=tk.DISABLED)  # Disable further cancel clicks
			self.pause_button.config(state=tk.DISABLED)  # Disable pause during cancel

	def complete_sorting(self, status_message="Sorting finished."):
		"""Finalizes the sorting process and resets the UI."""

		# Ensure this runs in the main Tkinter thread
		def reset_ui():
			logger.info(status_message)
			self.log_message(status_message)

			self.is_processing = False
			self.cancel_requested = False
			self.is_paused = False

			self.progress_var.set(0)  # Reset progress bar

			# Restore button states
			self.sort_button.config(state=tk.NORMAL)
			self.backup_button.config(state=tk.NORMAL)
			self.pause_button.config(state=tk.DISABLED, text=_("Pause"))
			self.cancel_button.config(state=tk.DISABLED)

			# Re-enable config changes
			self.auto_sort_check.config(state=tk.NORMAL)
			self.toggle_auto_sort()  # Correctly sets dependent states

			# Re-check Ollama status after processing
			self.check_ollama_status()

		if threading.current_thread() is threading.main_thread():
			reset_ui()
		else:
			self.root.after(0, reset_ui)

	# --- Deduplication (Using multiprocessing helper) ---
	def get_file_hash(self, file_path):
		"""Computes MD5 hash for a file (helper)."""
		# This is kept for single file hashing in cache check, dedupe uses the top-level func
		try:
			hasher = hashlib.md5()
			with open(file_path, 'rb') as f:
				while chunk := f.read(8192):
					hasher.update(chunk)
			return hasher.hexdigest()
		except IOError as e:
			logger.error(f"Error hashing file {file_path}: {e}")
			return None

	def find_and_remove_duplicates(self, files_to_check, mode="normal"):
		"""Finds and removes duplicates using multiprocessing."""
		if mode == "none":
			logger.info("Deduplication skipped.")
			return files_to_check, 0

		logger.info(f"Starting deduplication (mode: {mode}) for {len(files_to_check)} files...")
		self.log_message(_("Checking for duplicates..."))

		start_time = time.time()
		duplicates_removed_count = 0
		unique_files = []

		# Use multiprocessing Pool
		# Limit processes to avoid overwhelming older systems, but use more than 1 if possible
		num_processes = min(max(1, cpu_count() - 1), 4)  # Use N-1 cores, max 4, min 1
		logger.debug(f"Using {num_processes} processes for hashing.")

		file_info_map = {}
		try:
			with Pool(processes=num_processes) as pool:
				# Use imap_unordered for potentially better memory usage and responsiveness
				results = pool.imap_unordered(process_file_for_deduplication, files_to_check)
				processed_count = 0
				for i, result in enumerate(results):
					processed_count += 1
					if self.cancel_requested: raise InterruptedError("Deduplication cancelled.")
					if result and result[1]:  # Check if result is valid
						file_path, info = result
						file_info_map[file_path] = info
					# Update progress occasionally (e.g., every 5%)
					if i % max(1, len(files_to_check) // 20) == 0:
						progress = (i + 1) / len(files_to_check) * 100
						self.root.after(0, self.progress_var.set, progress)
						self.root.after(0, self.log_message,
										_("Hashing files for deduplication ({:.0f}%)...").format(progress))


		except InterruptedError:
			logger.warning("Deduplication hashing cancelled.")
			self.log_message(_("Deduplication cancelled."))
			return files_to_check, 0  # Return original list if cancelled during hashing
		except Exception as e:
			logger.error(f"Error during multiprocessing hashing: {e}", exc_info=True)
			self.log_message(_("Error during deduplication hashing. See logs."))
			# Proceed without deduplication if hashing failed
			return files_to_check, 0

		# Reset progress for removal phase
		self.root.after(0, self.progress_var.set, 0)
		logger.info(f"Hashing complete in {time.time() - start_time:.2f}s. Identifying duplicates...")

		# Group files by chosen key
		groups = {}
		if mode == "normal":  # Hash-based
			for path, info in file_info_map.items():
				key = info["hash"]
				if key not in groups: groups[key] = []
				groups[key].append(path)
		elif mode == "hardcore":  # Name + Size based
			for path, info in file_info_map.items():
				key = (info["name"], info["size"])
				if key not in groups: groups[key] = []
				groups[key].append(path)

		# Process groups: keep one, mark others for removal
		files_to_remove = set()
		for key, group in groups.items():
			if len(group) > 1:
				# Sort by modification time (newest first), then by path as tie-breaker
				sorted_group = sorted(
					group,
					key=lambda p: (file_info_map[p]["mod_time"], p),
					reverse=True
				)
				keep_file = sorted_group[0]
				unique_files.append(keep_file)
				# Add others to removal set
				files_to_remove.update(sorted_group[1:])
			elif group:  # Single file in group
				unique_files.append(group[0])

		# Perform removal
		duplicates_removed_count = len(files_to_remove)
		if duplicates_removed_count > 0:
			logger.info(f"Found {duplicates_removed_count} duplicate files to remove.")
			self.log_message(_("Removing {count} duplicate files...").format(count=duplicates_removed_count))
			removed_success = 0
			for i, duplicate_path in enumerate(files_to_remove):
				if self.cancel_requested: raise InterruptedError("Deduplication cancelled.")
				try:
					os.remove(duplicate_path)
					logger.debug(f"Removed duplicate: {duplicate_path}")
					removed_success += 1
					# Update progress
					progress = (i + 1) / duplicates_removed_count * 100
					self.root.after(0, self.progress_var.set, progress)

				except OSError as e:
					logger.error(f"Failed to remove duplicate {duplicate_path}: {e}")
					self.log_message(
						_("Error removing {file}: {err}").format(file=os.path.basename(duplicate_path), err=e))
					# If removal failed, keep it in the list to be processed? Or skip?
					# Let's keep it for safety, user might want to handle it manually.
					if duplicate_path not in unique_files:
						unique_files.append(duplicate_path)  # Re-add if removal failed
			duplicates_removed_count = removed_success  # Update count to actual removed
			logger.info(f"Successfully removed {removed_success} duplicates.")
			self.log_message(_("Removed {count} duplicates.").format(count=removed_success))
		else:
			logger.info("No duplicates found to remove.")
			self.log_message(_("No duplicates found."))

		end_time = time.time()
		logger.info(
			f"Deduplication finished in {end_time - start_time:.2f}s. Removed: {duplicates_removed_count} files.")
		self.root.after(0, self.progress_var.set, 0)  # Reset progress bar

		return unique_files, duplicates_removed_count

	# --- Automatic Category Generation (Async) ---
	async def async_generate_auto_categories(self, files_sample):
		"""Async generates categories via Ollama based on a sample of file info."""
		logger.info("Attempting automatic category generation...")
		self.log_message(_("Generating categories with Ollama..."))

		# Clear existing manual/previous auto categories
		self.category_list.clear()
		# Clear tree in main thread
		self.root.after(0, lambda: self.category_tree.delete(*self.category_tree.get_children()))

		if not files_sample:
			logger.warning("No files provided for auto-category generation.")
			return False  # Indicate failure

		# Prepare file info list (limit sample size for prompt)
		max_sample = 10  # Limit number of files sent in prompt
		file_info_list = [
			{"filename": os.path.basename(f),
			 "extension": os.path.splitext(f)[1].lower(),
			 "size_bytes": os.path.getsize(f)}
			for f in files_sample[:max_sample]
		]

		# Improved prompt
		prompt = f"""Analyze the following file list and propose a hierarchical category structure suitable for organizing them.
Use common sense categories based on file names, extensions, and typical usage.
Maximum category depth: {self.max_depth}.
Prioritize broader categories first. Be concise.

File Sample ({len(file_info_list)} files):
{json.dumps(file_info_list, indent=2)}

Respond ONLY with a JSON object representing the category tree. Example:
{{
  "Documents": {{
    "Reports": {{}},
    "Invoices": {{}}
  }},
  "Images": {{
    "Photos": {{}},
    "Screenshots": {{}}
  }},
  "Code": {{}},
  "Archives": {{}}
}}"""

		generated_categories = {}
		url = f"{self.ollama_url.strip('/')}/api/generate"
		payload = {
			"model": self.model,
			"prompt": prompt,
			"stream": False,
			"format": "json",  # Request JSON output format if model/Ollama supports it
			"options": {
				"num_predict": 512  # Increase predict limit for potentially larger JSON
			}
		}
		# Use a longer timeout for generation (can take time)
		timeout = aiohttp.ClientTimeout(total=300.0)  # 5 minutes total timeout

		try:
			logger.debug(f"Sending auto-category prompt to Ollama (model: {self.model})")
			async with aiohttp.ClientSession(timeout=timeout) as session:
				async with session.post(url, json=payload) as response:
					logger.info(f"Ollama auto-category response status: {response.status}")
					if response.status == 200:
						data = await response.json()
						response_text = data.get("response", "").strip()
						logger.debug(f"Ollama auto-category raw response: {response_text}")

						# Attempt to parse the JSON response
						try:
							# Sometimes models add markdown backticks, try removing them
							if response_text.startswith("```json"): response_text = response_text[7:]
							if response_text.endswith("```"): response_text = response_text[:-3]
							response_text = response_text.strip()

							generated_categories = json.loads(response_text)
							if not isinstance(generated_categories, dict):
								logger.warning(
									f"Ollama returned valid JSON, but not a dictionary: {type(generated_categories)}")
								generated_categories = {}  # Reset if not dict

						except json.JSONDecodeError as json_err:
							logger.error(f"Failed to parse Ollama JSON response for auto-categories: {json_err}")
							logger.error(f"Raw response was: {response_text}")
							self.log_message(_("Error: Ollama returned invalid format for categories."))
							return False  # Indicate failure
					else:
						error_text = await response.text()
						logger.error(f"Ollama auto-category request failed: {response.status} - {error_text[:200]}")
						self.log_message(_("Error: Ollama failed to generate categories (Status: {status})").format(
							status=response.status))
						return False  # Indicate failure

		except asyncio.TimeoutError:
			logger.error("Ollama auto-category request timed out.")
			self.log_message(_("Error: Ollama request timed out during category generation."))
			return False
		except aiohttp.ClientError as client_err:
			logger.error(f"Network error during Ollama auto-category request: {client_err}")
			self.log_message(_("Error: Network problem connecting to Ollama."))
			return False
		except Exception as e:
			logger.error(f"Unexpected error during Ollama auto-category generation: {e}", exc_info=True)
			self.log_message(_("Error: Unexpected problem during category generation."))
			return False

		# If categories were generated, build the tree and list
		if generated_categories:
			logger.info("Successfully received category structure from Ollama.")
			self._build_category_tree_and_list(generated_categories)
			# Log generated categories? Maybe too verbose.
			self.log_message(_("Automatic categories generated successfully."))
			return True
		else:
			logger.warning("Ollama did not return usable categories.")
			self.log_message(_("Warning: Ollama did not provide categories."))
			return False  # Indicate failure

	def _build_category_tree_and_list(self, categories_dict, parent_id="", current_path=""):
		"""Recursively builds treeview and category_list from Ollama's dict."""
		for name, subcategories in categories_dict.items():
			# Sanitize category names from Ollama
			safe_name = name.strip().replace('/', '-')
			if not safe_name: continue  # Skip empty names

			full_path = f"{current_path}/{safe_name}" if current_path else safe_name

			# Avoid duplicates at the same level (visual check in Treeview)
			child_exists = False
			for child_id in self.category_tree.get_children(parent_id):
				if self.category_tree.item(child_id, "text") == safe_name:
					child_exists = True
					item_id = child_id  # Use existing item ID for recursion
					break

			if not child_exists:
				# Add to tree in main thread
				def add_item():
					nonlocal item_id
					try:
						item_id = self.category_tree.insert(parent_id, tk.END, text=safe_name)
					except tk.TclError as e:  # Handle cases where parent might be gone
						logger.warning(f"Failed to insert tree item {safe_name} under {parent_id}: {e}")
						item_id = None

				item_id = None
				self.root.after(0, add_item)
			# Give Tkinter a moment to process the insertion - needed? Maybe not.
			# time.sleep(0.01) # Small delay - AVOID SLEEP IN ASYNC/MAIN THREAD HELPERS

			# Add to list (ensure no duplicates in the flat list)
			if full_path not in self.category_list:
				self.category_list.append(full_path)

			# Recurse for subcategories if depth allows and item was created/found
			current_depth = full_path.count('/')
			if isinstance(subcategories,
						  dict) and subcategories and current_depth < self.max_depth and item_id is not None:
				self._build_category_tree_and_list(subcategories, item_id, full_path)

		# Sort the final list after building
		self.category_list.sort()

	# --- Main Sorting Logic ---
	def sort_documents(self, source_dir, dest_dir):
		"""Main function orchestrating the sorting process."""
		start_time = time.time()
		processed_files_count = 0
		duplicates_removed_count = 0
		categories_created = set()  # Track unique category paths used
		all_files = []

		try:
			# --- 1. Collect Files ---
			# Currently only supports local files based on UI flow.
			# Cloud integration would need adjustment here (download first or process in place?)
			# For simplicity, let's assume local source_dir for now.
			# Add cloud download logic here if needed based on connection status
			if self.google_drive_service or self.dropbox_client or self.onedrive_client:
				# Placeholder: Implement cloud download/listing if required
				logger.warning("Cloud source directory not yet fully implemented. Processing local source directory.")
				self.log_message(_("Warning: Cloud source not fully implemented. Using local."))
			# Fall through to local file listing

			self.log_message(_("Scanning source directory..."))
			logger.info(f"Scanning source directory: {source_dir}")
			for entry in os.scandir(source_dir):
				if entry.is_file(follow_symlinks=False):  # Don't follow symlinks out of source
					# Basic check for file readability
					try:
						# Try opening briefly to catch permission errors early
						with open(entry.path, 'rb') as f:
							f.read(1)
						all_files.append(entry.path)
					except OSError as e:
						logger.warning(f"Skipping unreadable file: {entry.path} - {e}")
						self.log_message(_("Skipping unreadable file: {filename}").format(filename=entry.name))
				if self.cancel_requested: raise InterruptedError("Scan cancelled.")

			if not all_files:
				logger.warning("No files found in the source directory.")
				self.complete_sorting(_("No files found to process."))
				return

			logger.info(f"Found {len(all_files)} files to process.")
			self.log_message(_("Found {count} files.").format(count=len(all_files)))

			# --- 2. Handle Duplicates ---
			dedupe_mode = self.dedupe_mode.get()
			files_to_process, duplicates_removed_count = self.find_and_remove_duplicates(all_files, dedupe_mode)

			if not files_to_process:
				logger.warning("No files remaining after deduplication.")
				self.complete_sorting(_("No files left after duplicate removal."))
				return

			logger.info(f"{len(files_to_process)} unique files remaining for sorting.")
			self.log_message(_("{count} files remaining after deduplication.").format(count=len(files_to_process)))

			# --- 3. Generate/Verify Categories ---
			if self.auto_sort_var.get():
				# Use a sample of remaining files for generation
				sample_size = min(len(files_to_process), 100)  # Sample up to 100 files
				sample_files = files_to_process[:sample_size]  # Or random sample?
				# Run category generation asynchronously and wait for result
				future = asyncio.run_coroutine_threadsafe(self.async_generate_auto_categories(sample_files), self.loop)
				generation_success = future.result()  # Wait for completion

				if not generation_success or not self.category_list:
					logger.warning("Auto-category generation failed or yielded no categories. Using fallback.")
					self.log_message(_("Warning: Auto-category generation failed. Using 'Uncategorized'."))
					# Clear tree in main thread
					self.root.after(0, lambda: self.category_tree.delete(*self.category_tree.get_children()))
					self.category_list = ["Uncategorized"]
					# Add fallback to tree in main thread
					self.root.after(0, lambda: self.category_tree.insert("", tk.END, text="Uncategorized"))
			else:
				# Manual mode: Ensure destination category folders exist
				self.log_message(_("Using manually defined categories."))
				logger.info(f"Using manual categories: {self.category_list}")
				for category_path in self.category_list:
					try:
						# Ensure all levels of the path exist
						full_dest_path = os.path.join(dest_dir, *category_path.split('/'))
						os.makedirs(full_dest_path, exist_ok=True)
						categories_created.add(category_path)  # Track for report
					except OSError as e:
						logger.error(f"Could not create manual category directory: {full_dest_path} - {e}")
						self.log_message(
							_("Error creating directory for category '{cat}'. Skipping.").format(cat=category_path))
					# Maybe remove category from list if dir fails? Risky.

			if not self.category_list:
				# Should not happen if fallback works, but as a final check
				logger.error("No categories available (manual or auto). Aborting.")
				self.complete_sorting(_("Error: No categories defined or generated. Cannot sort."))
				return

			# --- 4. Process Files (Classification & Move) ---
			total_files_to_process = len(files_to_process)
			logger.info(f"Starting classification and moving {total_files_to_process} files...")
			self.log_message(_("Classifying and moving files..."))

			# Use ThreadPoolExecutor for I/O bound tasks (network classification, file move)
			num_workers = min(max(1, cpu_count()), 4)  # Limit workers on older systems
			logger.debug(f"Using {num_workers} worker threads for file processing.")

			with concurrent.futures.ThreadPoolExecutor(max_workers=num_workers) as executor:
				# Submit tasks
				future_to_file = {executor.submit(self.process_single_file, file_path, dest_dir): file_path for
								  file_path in files_to_process}

				for i, future in enumerate(concurrent.futures.as_completed(future_to_file)):
					file_path = future_to_file[future]
					try:
						category_used = future.result()  # Get the category path used (or None if failed)
						if category_used:
							processed_files_count += 1
							categories_created.add(category_used)  # Track unique categories used
					except InterruptedError:
						# Caught if cancel_requested was set within process_single_file
						logger.info(f"Processing cancelled for {os.path.basename(file_path)} and subsequent files.")
						break  # Stop processing more files
					except Exception as exc:
						logger.error(f"Error processing file {os.path.basename(file_path)}: {exc}", exc_info=True)
						self.log_message(
							_("Error processing {filename}. See logs.").format(filename=os.path.basename(file_path)))

					# Update progress bar (runs in main thread via root.after)
					progress = (i + 1) / total_files_to_process * 100
					self.root.after(0, self.progress_var.set, progress)

					# Check for cancellation between files
					if self.cancel_requested:
						logger.info("Cancellation detected, shutting down worker threads...")
						# Attempt to cancel pending futures (may not work for already running tasks)
						for f in future_to_file:
							if not f.done(): f.cancel()
						executor.shutdown(wait=False, cancel_futures=True)  # Python 3.9+
						# executor.shutdown(wait=False) # Older Python
						break

					# Handle pause
					while self.is_paused:
						if self.cancel_requested: break  # Allow cancelling while paused
						time.sleep(0.5)  # Sleep briefly while paused

			# --- 5. Finalization ---
			end_time = time.time()
			elapsed_time = end_time - start_time

			final_status = _("Sorting cancelled.") if self.cancel_requested else _("Sorting completed.")

			logger.info("-" * 50)
			logger.info(final_status)
			logger.info(f"Total files processed: {processed_files_count}/{total_files_to_process}")
			logger.info(f"Duplicates removed: {duplicates_removed_count}")
			logger.info(f"Unique categories used/created: {len(categories_created)}")
			logger.info(f"Total time: {elapsed_time:.2f} seconds")
			logger.info("-" * 50)

			# Prepare stats for report
			stats = {
				"processed_files": processed_files_count,
				"categories_used": len(categories_created),
				"duplicates_removed": duplicates_removed_count,
				"elapsed_time": f"{elapsed_time:.2f} seconds"
			}
			# Generate report in main thread
			self.root.after(0, self.generate_report, stats)

		# Cloud sync would go here if implemented
		# if self.google_drive_service or self.dropbox_client or self.onedrive_client:
		#     logger.info("Starting cloud synchronization...")
		#     self.log_message(_("Syncing results to cloud..."))
		#     sync_future = asyncio.run_coroutine_threadsafe(self.sync_to_cloud(dest_dir), self.loop)
		#     try:
		#         sync_future.result(timeout=300) # Wait up to 5 mins for sync
		#         logger.info("Cloud synchronization complete.")
		#         self.log_message(_("Cloud sync finished."))
		#     except asyncio.TimeoutError:
		#         logger.error("Cloud synchronization timed out.")
		#         self.log_message(_("Error: Cloud sync timed out."))
		#     except Exception as sync_err:
		#         logger.error(f"Cloud synchronization failed: {sync_err}", exc_info=True)
		#         self.log_message(_("Error: Cloud sync failed. See logs."))


		except InterruptedError:
			final_status = _("Sorting cancelled by user.")
			logger.warning(final_status)
		except Exception as e:
			final_status = _("Sorting failed with error.")
			logger.critical(f"Critical error during sorting process: {e}", exc_info=True)
			self.log_message(_("Critical Error: Sorting failed. Check logs."))
			# Show error in UI as well
			self.root.after(0, messagebox.showerror, _("Sorting Error"),
							_("An unexpected error occurred: {error}. Check logs.").format(error=e))
		finally:
			# Ensure UI is reset regardless of how the process ended
			self.complete_sorting(final_status)

	def process_single_file(self, file_path, dest_dir):
		"""Processes one file: classify, create dir, move. Returns category path or None."""
		# Check cancellation flag at the start
		if self.cancel_requested: raise InterruptedError("Cancelled")

		filename = os.path.basename(file_path)
		logger.debug(f"Processing: {filename}")
		# Don't log every file to GUI by default, too verbose
		# self.log_message(_("Processing: {filename}").format(filename=filename))

		# --- 1. Classify ---
		# Prepare file info (avoid reading full file here if possible)
		try:
			file_info = {
				"filename": filename,
				"extension": os.path.splitext(filename)[1].lower(),
				"size_bytes": os.path.getsize(file_path),
				"path": file_path  # Pass full path for content sampling
			}
		except OSError as e:
			logger.error(f"Cannot get info for file {file_path}: {e}")
			self.log_message(_("Error getting info for {filename}. Skipping.").format(filename=filename))
			return None  # Skip file

		# Use asyncio.run_coroutine_threadsafe to call the async classification
		classify_future = asyncio.run_coroutine_threadsafe(self.async_classify_file(file_info), self.loop)
		try:
			# Add a timeout for classification per file
			category_path = classify_future.result(timeout=60.0)  # 60s timeout per file classification
		except asyncio.TimeoutError:
			logger.warning(f"Classification timed out for {filename}. Using fallback.")
			self.log_message(_("Timeout classifying {filename}. Using fallback.").format(filename=filename))
			category_path = self.category_list[0]  # Use first category as fallback
		except Exception as classify_err:
			logger.error(f"Classification failed for {filename}: {classify_err}", exc_info=True)
			self.log_message(_("Error classifying {filename}. Using fallback.").format(filename=filename))
			category_path = self.category_list[0]  # Use first category as fallback

		if not category_path or category_path not in self.category_list:
			logger.warning(
				f"Invalid or empty category '{category_path}' returned for {filename}. Using fallback '{self.category_list[0]}'.")
			category_path = self.category_list[0]  # Fallback

		# --- 2. Prepare Destination ---
		try:
			# Split category path for directory creation
			dest_subdirs = category_path.split('/')
			final_dest_dir = os.path.join(dest_dir, *dest_subdirs)

			# Create directory if it doesn't exist (thread-safe)
			os.makedirs(final_dest_dir, exist_ok=True)

			dest_path = os.path.join(final_dest_dir, filename)

			# Handle potential naming conflicts
			counter = 1
			base, ext = os.path.splitext(filename)
			while os.path.exists(dest_path):
				# Check if existing file is identical (hash check?) - potentially slow
				# Simple approach: rename the file being moved
				logger.warning(f"Destination file exists: {dest_path}. Renaming.")
				dest_path = os.path.join(final_dest_dir, f"{base}_{counter}{ext}")
				counter += 1
				if counter > 100:  # Safety break
					logger.error(
						f"Could not find unique name for {filename} in {final_dest_dir} after 100 attempts. Skipping.")
					self.log_message(
						_("Error: Too many name conflicts for {filename}. Skipping.").format(filename=filename))
					return None  # Skip file

		except OSError as e:
			logger.error(f"Error preparing destination for {filename} (category: {category_path}): {e}")
			self.log_message(_("Error creating directory for {filename}. Skipping.").format(filename=filename))
			return None  # Skip file

		# --- 3. Move File ---
		# Check cancellation flag again before moving
		if self.cancel_requested: raise InterruptedError("Cancelled")

		try:
			shutil.move(file_path, dest_path)
			logger.info(f"Moved '{filename}' -> '{category_path}'")
			# Maybe log moves less frequently to GUI?
			# if processed_files_count % 10 == 0: # Example: Log every 10 moves
			#      self.log_message(_("Moved {count} files...").format(count=processed_files_count))
			return category_path  # Return category used

		except Exception as move_err:
			logger.error(f"Failed to move {filename} to {dest_path}: {move_err}", exc_info=True)
			self.log_message(_("Error moving {filename}: {err}").format(filename=filename, err=move_err))
			# Attempt to copy and then delete? More robust but slower.
			# try:
			#     shutil.copy2(file_path, dest_path) # copy2 preserves metadata
			#     os.remove(file_path)
			#     logger.info(f"Copied and removed '{filename}' -> '{category_path}' after move failed.")
			#     return category_path
			# except Exception as copy_err:
			#     logger.error(f"Copy also failed for {filename}: {copy_err}")
			#     # File remains in source
			return None  # Indicate failure

	async def async_classify_file(self, file_info):
		"""Async classifies a single file using Ollama, with caching and content sampling."""
		file_path = file_info["path"]
		filename = file_info["filename"]

		# --- 1. Check Cache ---
		file_hash = self.get_file_hash(file_path)  # Hash check still useful
		if file_hash and file_hash in self.cache:
			cached_category = self.cache[file_hash]
			# Verify cached category still exists in current list
			if cached_category in self.category_list:
				logger.debug(f"Using cached category '{cached_category}' for '{filename}'")
				# self.log_message(_("Using cache for {filename}").format(filename=filename)) # Too verbose for GUI
				return cached_category
			else:
				logger.debug(f"Cached category '{cached_category}' for '{filename}' no longer valid. Re-classifying.")
			# Remove invalid entry from cache?
			# del self.cache[file_hash]

		# --- 2. Prepare Prompt ---
		content_sample = await self.get_content_sample(file_path, file_info["extension"])

		# Simplified prompt, relying more on file info + short sample
		prompt = f"""Classify the following file into ONE category from the list provided.
Respond with ONLY the category name.

Categories: {', '.join(self.category_list)}

File Information:
- Name: {filename}
- Extension: {file_info['extension']}
- Size: {file_info['size_bytes']} bytes

Content Sample (up to 500 chars):
{content_sample[:500]}

Category:"""  # Let the model complete this

		# --- 3. Call Ollama ---
		url = f"{self.ollama_url.strip('/')}/api/generate"
		payload = {
			"model": self.model,
			"prompt": prompt,
			"stream": False,
			# "format": "json", # Not requesting JSON here, just raw string category
			"options": {
				"num_predict": 32,  # Limit prediction length, category names are short
				"temperature": 0.2  # Lower temperature for more deterministic category choice
			}
		}
		# Use a moderate timeout for classification
		timeout = aiohttp.ClientTimeout(total=30.0)  # 30s total timeout

		category = None
		try:
			async with aiohttp.ClientSession(timeout=timeout) as session:
				async with session.post(url, json=payload) as response:
					if response.status == 200:
						data = await response.json()
						raw_category = data.get("response", "").strip()
						# Clean up potential model verbosity ("Category: X" -> "X")
						if ":" in raw_category: raw_category = raw_category.split(":")[-1].strip()
						# Remove potential quotes
						raw_category = raw_category.strip('"`\'')

						logger.debug(f"Ollama classification for '{filename}': '{raw_category}'")

						# Find the best match in our list (case-insensitive partial match?)
						# Stricter matching is safer: exact match or find if response is a sub-path
						found_match = None
						if raw_category in self.category_list:
							found_match = raw_category
						else:
							# Check if Ollama returned a sub-path like "Work/Reports" when only "Work" exists
							# Or if it returned "Report" instead of "Reports"
							# Simple approach: find first category name containing the response (or vice versa) - risky
							# Safer: Use exact match from list. If model hallucinates, use fallback.
							logger.warning(
								f"Ollama returned category '{raw_category}' not in list {self.category_list} for file '{filename}'.")
							found_match = None  # Force fallback later

						category = found_match

					else:
						error_text = await response.text()
						logger.error(
							f"Ollama classification request failed for {filename}: {response.status} - {error_text[:200]}")
					# Fall through to return None (will trigger fallback)

		except asyncio.TimeoutError:
			logger.warning(f"Ollama classification request timed out for {filename}.")
		# Fall through
		except aiohttp.ClientError as client_err:
			logger.error(f"Network error during Ollama classification for {filename}: {client_err}")
		# Fall through
		except Exception as e:
			logger.error(f"Unexpected error during Ollama classification for {filename}: {e}", exc_info=True)
		# Fall through

		# --- 4. Update Cache and Return ---
		if category and file_hash:
			self.cache[file_hash] = category
			# Save cache periodically? Or on completion? Saving every time is slow.
			# Consider a separate thread or async task for saving cache less often.
			# For now, let's save on completion in `on_closing` and `complete_sorting`.
			# self.save_cache() # Avoid saving here
			return category
		else:
			# Return None to indicate classification failed or returned invalid category
			# The calling function (process_single_file) will handle the fallback.
			return None

	async def get_content_sample(self, file_path, extension):
		"""Async helper to get a small content sample from different file types."""
		# Keep sampling limited to avoid performance hits
		sample_size_kb = 10
		max_chars = 500  # Max chars to return

		try:
			loop = asyncio.get_running_loop()
			# Run blocking I/O and parsing in a default executor
			return await loop.run_in_executor(None, self._read_content_sample_sync, file_path, extension,
											  sample_size_kb * 1024, max_chars)
		except Exception as e:
			logger.warning(f"Failed to get content sample for {os.path.basename(file_path)}: {e}")
			return ""  # Return empty string on error

	def _read_content_sample_sync(self, file_path, extension, max_bytes, max_chars):
		"""Synchronous part of reading content samples (runs in executor)."""
		content_sample = ""
		try:
			ext = extension.lower()
			if ext == '.pdf' and PyPDF2:
				try:
					with open(file_path, 'rb') as f:
						reader = PyPDF2.PdfReader(f)
						num_pages = len(reader.pages)
						text = ""
						for i in range(min(num_pages, 2)):  # Sample first 2 pages
							page = reader.pages[i]
							page_text = page.extract_text()
							if page_text:
								text += page_text + "\n"
								if len(text) >= max_chars: break
						content_sample = text[:max_chars]
				except Exception as pdf_err:
					logger.debug(f"PyPDF2 failed for {os.path.basename(file_path)}: {pdf_err}")
					# Fallback to binary read
					with open(file_path, 'rb') as f:
						content_sample = f.read(max_bytes).decode('utf-8', errors='ignore')[:max_chars]

			elif ext == '.docx' and docx:
				try:
					doc = docx.Document(file_path)
					text = ""
					for para in doc.paragraphs[:10]:  # Sample first 10 paragraphs
						text += para.text + "\n"
						if len(text) >= max_chars: break
					content_sample = text[:max_chars]
				except Exception as docx_err:
					logger.debug(f"python-docx failed for {os.path.basename(file_path)}: {docx_err}")
					with open(file_path, 'rb') as f:
						content_sample = f.read(max_bytes).decode('utf-8', errors='ignore')[:max_chars]

			elif ext == '.xlsx' and openpyxl:
				try:
					wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)  # Read only faster
					sheet = wb.active
					text = ""
					cell_count = 0
					for row in sheet.iter_rows(max_row=20, max_col=10):  # Sample 20x10 cells
						for cell in row:
							if cell.value is not None:
								text += str(cell.value) + " "
								cell_count += 1
								if len(text) >= max_chars or cell_count > 50: break  # Limit cells too
						if len(text) >= max_chars or cell_count > 50: break
					content_sample = text[:max_chars]
					wb.close()  # Close workbook
				except Exception as xlsx_err:
					logger.debug(f"openpyxl failed for {os.path.basename(file_path)}: {xlsx_err}")
					with open(file_path, 'rb') as f:
						content_sample = f.read(max_bytes).decode('utf-8', errors='ignore')[:max_chars]

			elif ext == '.odt' and odf_teletype:
				try:
					text = odf_teletype.extractText(file_path)
					content_sample = text[:max_chars]
				except Exception as odt_err:
					logger.debug(f"odfpy failed for {os.path.basename(file_path)}: {odt_err}")
					with open(file_path, 'rb') as f:
						content_sample = f.read(max_bytes).decode('utf-8', errors='ignore')[:max_chars]

			else:  # General text or binary fallback
				with open(file_path, 'rb') as f:
					content_sample = f.read(max_bytes).decode('utf-8', errors='ignore')[:max_chars]

		except Exception as e:
			logger.warning(f"Error reading sample from {os.path.basename(file_path)}: {e}")
			return ""  # Return empty on error

		return content_sample.strip()

# --- Cloud Sync (Placeholder/Example) ---
# async def sync_to_cloud(self, local_dest_dir):
#     """Placeholder for async cloud synchronization."""
#     if self.dropbox_client:
#         logger.info("Starting Dropbox sync...")
#         # Implement Dropbox upload logic here, potentially using ThreadPoolExecutor for uploads
#         pass
#     elif self.google_drive_service:
#         logger.info("Starting Google Drive sync...")
#         # Implement Google Drive upload logic here
#         pass
#     elif self.onedrive_client:
#          logger.info("Starting OneDrive sync...")
#         # Implement OneDrive upload logic here
#         pass
#     else:
#         logger.info("No cloud service connected for sync.")


# --- Main Execution ---
def main():
	"""Main function to parse args and run the app."""
	# Argument parsing remains the same
	parser = argparse.ArgumentParser(description="Document Sorter with Ollama")
	parser.add_argument("--source", help="Source directory (overrides config)")
	parser.add_argument("--dest", help="Destination directory (overrides config)")
	parser.add_argument("--categories", help="Comma-separated categories for manual mode (overrides config)")
	parser.add_argument("--dedupe", choices=["none", "normal", "hardcore"],
						help="Duplicate removal mode (overrides config)")
	parser.add_argument("--ollama-url", help="URL for Ollama API (overrides config)")
	parser.add_argument("--model", help="Ollama model to use (overrides config)")
	parser.add_argument("--lang", choices=["en", "ru"], default="en", help="Interface language (en or ru)")
	parser.add_argument("--no-gui", action="store_true", help="Run in command-line mode (requires source and dest)")
	parser.add_argument('--debug', action='store_true', help='Enable debug logging')

	args = parser.parse_args()

	# Setup logging level based on args
	if args.debug:
		logger.setLevel(logging.DEBUG)
		# Add console handler if debugging
		console_handler = logging.StreamHandler()
		console_handler.setFormatter(log_formatter)
		logger.addHandler(console_handler)
		logger.debug("Debug logging enabled.")

	# Set language
	global _
	_ = setup_localization(args.lang)

	# --- GUI Mode ---
	if not args.no_gui:
		try:
			root = TkinterDnD.Tk()
		except tk.TclError:
			logger.critical("Could not initialize TkinterDnD. Is it installed and configured correctly?")
			print("ERROR: TkinterDND not found or failed to initialize. Please install tkinterdnd2-universal.")
			# Attempt fallback to basic Tkinter? No, DND is core.
			exit(1)
		except Exception as e:
			logger.critical(f"Failed to initialize GUI: {e}", exc_info=True)
			print(f"ERROR: Failed to initialize GUI: {e}")
			exit(1)

		# Apply overrides from args to the app instance after init
		app = DocumentSorter(root, ollama_url=args.ollama_url or "http://localhost:11434")  # Pass URL override

		# Override config values if provided in args AFTER load_config is called internally
		if args.source: app.source_dir_var.set(args.source)
		if args.dest: app.dest_dir_var.set(args.dest)
		if args.dedupe: app.dedupe_mode.set(args.dedupe)
		if args.model:
			# Check if model exists after fetching
			if args.model in app.available_models:
				app.model = args.model
				app.model_combobox.set(args.model)
			else:
				logger.warning(
					f"Model '{args.model}' from args not found in available models: {app.available_models}. Using default/config model.")
				app.log_message(_("Warning: Model '{model}' not found.").format(model=args.model))

		if args.categories:
			app.auto_sort_var.set(False)
			app.category_list = sorted([c.strip() for c in args.categories.split(',') if c.strip()])
			app._rebuild_category_tree_from_list()  # Update tree
			app.toggle_auto_sort()  # Update button states

		# Check this block carefully:
		if args.ollama_url:  # Check if the argument was provided *at all*
			logger.info(f"DEBUG: Overriding ollama_url from command line arg: {args.ollama_url}")
			# Ensure the override uses the *correct* base URL format
			# It should already be correct if the arg parsing uses the base URL,
			# but let's be explicit or add validation if needed.
			# The value passed to the constructor should handle the override initially.
			# This block might be redundant if the constructor handles it.
			# Let's comment it out for now to test, as the constructor already takes the arg override:
			# app.ollama_url = args.ollama_url.strip('/') # TEMPORARILY COMMENT OUT
			pass  # Let constructor handle the initial override

		app.save_config()  # Save potentially overridden config

		root.mainloop()

	# --- Command-Line Mode ---
	else:
		if not args.source or not args.dest:
			parser.error("--no-gui mode requires --source and --dest arguments.")

		logger.info("Running in command-line mode.")
		# Need a dummy root for some methods? No, create a headless sorter instance.
		# We need to adapt the sorter or create a separate CLI runner class/functions.

		# Let's adapt by creating a minimal sorter instance and calling sort_documents directly.
		# We need to handle config loading/overrides manually here.

		# Load config manually first
		config = {}
		config_file = "config.json"
		try:
			if os.path.exists(config_file):
				with open(config_file, 'r', encoding='utf-8') as f:
					config = json.load(f)
		except Exception as e:
			logger.error(f"Error loading config file {config_file}: {e}")

		# Apply overrides
		source_dir = args.source
		dest_dir = args.dest
		ollama_url = args.ollama_url or config.get("ollama_url", "http://localhost:11434/api")
		model = args.model or config.get("model", "qwen2.5:7b")  # Use a default
		dedupe_mode_cli = args.dedupe or config.get("dedupe_mode", "none")
		max_depth_cli = int(config.get("max_depth", 3))  # Get from config or default
		is_auto_cli = True
		categories_cli = []
		if args.categories:
			is_auto_cli = False
			categories_cli = sorted([c.strip() for c in args.categories.split(',') if c.strip()])
		elif config.get("categories"):  # Check config if not in args
			is_auto_cli = False
			categories_cli = config["categories"]

		# Minimal Sorter for CLI (no UI elements needed)
		class HeadlessSorter:
			def __init__(self, ollama_url_cli, model_cli, categories_list, is_auto_mode, max_depth_val,
						 dedupe_cli_mode):
				self.ollama_url = ollama_url_cli
				self.model = model_cli
				self.category_list = categories_list
				self.auto_sort_var = tk.BooleanVar(value=is_auto_mode)  # Need BooleanVar for logic? Maybe just bool.
				self.is_auto_mode = is_auto_mode
				self.max_depth = max_depth_val
				self.dedupe_mode = tk.StringVar(value=dedupe_cli_mode)  # Need StringVar? Maybe just string.
				self.dedupe_mode_str = dedupe_cli_mode

				self.cache = self.load_cache()
				self.cancel_requested = False  # Basic cancellation via Ctrl+C?
				self.is_paused = False  # Pause not really applicable in CLI

				# Async loop setup needed for async functions
				self.loop = asyncio.new_event_loop()
				asyncio.set_event_loop(self.loop)

				# Need dummy UI vars/methods used by shared functions?
				# Mock log_message, progress updates
				self.root = None  # No root window
				self.source_dir_var = tk.StringVar(value=source_dir)  # Needed by async_classify_file

			def log_message(self, message):
				print(f"{time.strftime('%H:%M:%S')} - {message}")  # Print to console

			def progress_var_set(self, value):
				# Simple console progress bar
				bar_length = 30
				filled_length = int(bar_length * value / 100)
				bar = '#' * filled_length + '-' * (bar_length - filled_length)
				print(f'\rProgress: |{bar}| {value:.1f}%', end='', flush=True)
				if value >= 100: print()  # Newline at end

			# --- Include necessary methods from DocumentSorter ---
			# (Copy/paste or inherit - copy/paste simpler for CLI adaptation)
			load_cache = DocumentSorter.load_cache
			save_cache = DocumentSorter.save_cache
			get_file_hash = DocumentSorter.get_file_hash
			# find_and_remove_duplicates needs Pool, process_file_for_deduplication
			find_and_remove_duplicates = DocumentSorter.find_and_remove_duplicates
			# async_generate_auto_categories needs aiohttp, _build_category_tree_and_list (adapted)
			async_generate_auto_categories = DocumentSorter.async_generate_auto_categories
			_build_category_tree_and_list = DocumentSorter._build_category_tree_and_list  # Needs adaptation for no Treeview
			# sort_documents needs ThreadPoolExecutor, process_single_file, generate_report (adapted)
			sort_documents = DocumentSorter.sort_documents  # Needs heavy adaptation
			process_single_file = DocumentSorter.process_single_file  # Needs adaptation
			async_classify_file = DocumentSorter.async_classify_file  # Needs adaptation (source_dir_var)
			get_content_sample = DocumentSorter.get_content_sample
			_read_content_sample_sync = DocumentSorter._read_content_sample_sync
			generate_report = DocumentSorter.generate_report  # Needs adaptation (no UI context)

			# Need simplified _build_category_tree_and_list for CLI
			def _build_category_tree_and_list(self, categories_dict, parent_id="", current_path=""):
				# Only updates self.category_list
				for name, subcategories in categories_dict.items():
					safe_name = name.strip().replace('/', '-')
					if not safe_name: continue
					full_path = f"{current_path}/{safe_name}" if current_path else safe_name
					if full_path not in self.category_list: self.category_list.append(full_path)
					current_depth = full_path.count('/')
					if isinstance(subcategories, dict) and subcategories and current_depth < self.max_depth:
						self._build_category_tree_and_list(subcategories, "", full_path)  # No parent_id needed
				self.category_list.sort()

			# Simplified generate_report for CLI
			def generate_report(self, stats):
				try:
					# Add context
					stats['timestamp'] = time.strftime('%Y-%m-%d %H:%M:%S')
					stats['source_dir'] = source_dir
					stats['dest_dir'] = dest_dir
					stats['dedupe_mode'] = self.dedupe_mode_str
					stats['sorting_mode'] = _("Automatic") if self.is_auto_mode else _("Manual Categories")

					report_filename = "sorting_report_cli.txt"
					with open(report_filename, "w", encoding="utf-8") as f:
						f.write(f"{_('Document Sorter Report (CLI)')}\n")
						f.write("=" * 30 + "\n")
						f.write(f"{_('Report Generated:')} {stats['timestamp']}\n")
						f.write(f"{_('Source Directory:')} {stats['source_dir']}\n")
						f.write(f"{_('Destination Directory:')} {stats['dest_dir']}\n")
						f.write("\n--- {_('Summary')} ---\n")
						f.write(f"{_('Total Files Processed:')} {stats['processed_files']}\n")
						f.write(f"{_('Categories Used:')} {stats['categories_used']}\n")
						f.write(
							f"{_('Duplicate Files Removed:')} {stats['duplicates_removed']} (Mode: {stats['dedupe_mode']})\n")
						f.write(f"{_('Sorting Mode:')} {stats['sorting_mode']}\n")
						f.write(f"{_('Elapsed Time:')} {stats['elapsed_time']}\n")
					logger.info(_(f"Report generated: {report_filename}"))
					self.log_message(_(f"Report generated: {report_filename}"))
				except Exception as e:
					logger.error(f"Failed to generate CLI report: {e}", exc_info=True)
					self.log_message(_("Failed to generate report."))

			# Need adapted process_single_file and sort_documents for CLI progress/logging
			def process_single_file(self, file_path, dest_dir):
				# Simplified version calling original logic but using self.log_message etc.
				if self.cancel_requested: raise InterruptedError("Cancelled")
				filename = os.path.basename(file_path)
				logger.debug(f"Processing: {filename}")
				try:
					file_info = {"filename": filename, "extension": os.path.splitext(filename)[1].lower(),
								 "size_bytes": os.path.getsize(file_path), "path": file_path}
				except OSError as e:
					return None  # Skip
				classify_future = asyncio.run_coroutine_threadsafe(self.async_classify_file(file_info), self.loop)
				try:
					category_path = classify_future.result(timeout=60.0)
				except Exception:
					category_path = self.category_list[0]  # Fallback
				if not category_path or category_path not in self.category_list: category_path = self.category_list[0]
				try:
					dest_subdirs = category_path.split('/')
					final_dest_dir = os.path.join(dest_dir, *dest_subdirs)
					os.makedirs(final_dest_dir, exist_ok=True)
					dest_path = os.path.join(final_dest_dir, filename)
					counter = 1;
					base, ext = os.path.splitext(filename)
					while os.path.exists(dest_path):
						dest_path = os.path.join(final_dest_dir, f"{base}_{counter}{ext}");
						counter += 1
						if counter > 100: return None  # Skip
				except OSError as e:
					return None  # Skip
				if self.cancel_requested: raise InterruptedError("Cancelled")
				try:
					shutil.move(file_path, dest_path)
					logger.info(f"Moved '{filename}' -> '{category_path}'")
					return category_path  # Return category used
				except Exception as move_err:
					logger.error(f"Failed to move {filename}: {move_err}")
					return None  # Indicate failure

			def sort_documents(self, source_dir_arg, dest_dir_arg):
				# Adapted version of sort_documents for CLI
				start_time = time.time();
				processed_files_count = 0;
				duplicates_removed_count = 0;
				categories_created = set();
				all_files = []
				try:
					self.log_message(_("Scanning source directory..."))
					for entry in os.scandir(source_dir_arg):
						if entry.is_file(follow_symlinks=False):
							try:
								with open(entry.path, 'rb') as f:
									f.read(1)
							except OSError:
								logger.warning(f"Skipping unreadable file: {entry.path}")
						if self.cancel_requested: raise InterruptedError("Scan cancelled.")
					if not all_files: self.log_message(_("No files found.")); return

					self.log_message(_("Found {count} files.").format(count=len(all_files)))
					files_to_process, duplicates_removed_count = self.find_and_remove_duplicates(all_files,
																								 self.dedupe_mode_str)  # Use string mode
					if not files_to_process: self.log_message(_("No files left after deduplication.")); return
					self.log_message(_("{count} files remaining.").format(count=len(files_to_process)))

					if self.is_auto_mode:
						sample_files = files_to_process[:min(len(files_to_process), 100)]
						future = asyncio.run_coroutine_threadsafe(self.async_generate_auto_categories(sample_files),
																  self.loop)
						generation_success = future.result()
						if not generation_success or not self.category_list:
							self.log_message(_("Warning: Auto-category failed. Using 'Uncategorized'."))
							self.category_list = ["Uncategorized"]
					else:
						self.log_message(_("Using manual categories."))
						for category_path in self.category_list:
							try:
								os.makedirs(os.path.join(dest_dir_arg, *category_path.split('/')),
											exist_ok=True); categories_created.add(category_path)
							except OSError as e:
								logger.error(f"Could not create dir: {category_path} - {e}")
					if not self.category_list: self.log_message(_("Error: No categories. Cannot sort.")); return

					total_files_to_process = len(files_to_process)
					self.log_message(_("Classifying and moving files..."))
					num_workers = min(max(1, cpu_count()), 4)
					with concurrent.futures.ThreadPoolExecutor(max_workers=num_workers) as executor:
						future_to_file = {executor.submit(self.process_single_file, file_path, dest_dir_arg): file_path
										  for file_path in files_to_process}
						for i, future in enumerate(concurrent.futures.as_completed(future_to_file)):
							try:
								category_used = future.result();
							except InterruptedError:
								break
							except Exception as exc:
								logger.error(f"Error processing {future_to_file[future]}: {exc}")
							else:
								if category_used: processed_files_count += 1; categories_created.add(category_used)
							self.progress_var_set((i + 1) / total_files_to_process * 100)  # Use console progress
							if self.cancel_requested: break
					end_time = time.time();
					elapsed_time = end_time - start_time
					final_status = _("Sorting cancelled.") if self.cancel_requested else _("Sorting completed.")
					logger.info(final_status)
					stats = {"processed_files": processed_files_count, "categories_used": len(categories_created),
							 "duplicates_removed": duplicates_removed_count,
							 "elapsed_time": f"{elapsed_time:.2f} seconds"}
					self.generate_report(stats)  # Generate CLI report
				except InterruptedError:
					logger.warning(_("Sorting cancelled by user."))
				except Exception as e:
					logger.critical(f"Critical error during sorting: {e}", exc_info=True)
				finally:
					self.save_cache()  # Save cache at the end
					self.loop.call_soon_threadsafe(self.loop.stop)  # Stop the loop

		# Create and run the headless sorter
		cli_sorter = HeadlessSorter(ollama_url, model, categories_cli, is_auto_cli, max_depth_cli, dedupe_mode_cli)

		# Handle Ctrl+C for cancellation
		import signal
		def signal_handler(sig, frame):
			print('\nCtrl+C detected! Requesting cancellation...')
			cli_sorter.cancel_requested = True

		signal.signal(signal.SIGINT, signal_handler)

		# Run the sorting in the main thread (or could launch loop thread)
		try:
			# Run sort_documents which uses the internal loop via run_coroutine_threadsafe
			cli_sorter.sort_documents(source_dir, dest_dir)
			# Keep loop running until sort_documents stops it? Or run forever and let sort_documents handle exit?
			# If sort_documents is blocking, this might not be needed. If it uses threads internally, loop needs running.
			# Since sort_documents uses run_coroutine_threadsafe, the loop needs to be running.
			cli_sorter.loop.run_forever()  # Run until loop.stop() is called in sort_documents.finally
		except Exception as cli_run_err:
			logger.critical(f"CLI run failed: {cli_run_err}", exc_info=True)
		finally:
			if cli_sorter.loop.is_running():
				cli_sorter.loop.call_soon_threadsafe(cli_sorter.loop.stop)
			# Ensure loop closes cleanly
			# Need to manage the loop thread if started separately. Here it runs in main thread.
			logger.info("CLI execution finished.")


if __name__ == "__main__":
	main()

# --- END OF FILE main.py ---