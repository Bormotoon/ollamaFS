import os
import shutil
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import threading
import requests
import json
import time
from typing import List, Dict, Any


class DocumentSorter:
	def __init__(self, root):
		self.root = root
		self.root.title("Document Sorter with Ollama")
		self.root.geometry("800x600")
		self.root.resizable(True, True)

		self.ollama_url = "http://localhost:11434/api"
		self.model = "deepseek-coder"  # Default model
		self.available_models = []

		self.setup_ui()
		self.check_ollama_status()

	def setup_ui(self):
		# Main frame
		main_frame = ttk.Frame(self.root, padding="10")
		main_frame.pack(fill=tk.BOTH, expand=True)

		# Ollama status indicator
		status_frame = ttk.Frame(main_frame)
		status_frame.pack(fill=tk.X, pady=5)

		ttk.Label(status_frame, text="Ollama Status:").pack(side=tk.LEFT, padx=5)
		self.status_label = ttk.Label(status_frame, text="Checking...", foreground="orange")
		self.status_label.pack(side=tk.LEFT, padx=5)

		# Model selection
		model_frame = ttk.Frame(main_frame)
		model_frame.pack(fill=tk.X, pady=5)

		ttk.Label(model_frame, text="Select Model:").pack(side=tk.LEFT, padx=5)
		self.model_combobox = ttk.Combobox(model_frame, state="readonly")
		self.model_combobox.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
		self.model_combobox.bind("<<ComboboxSelected>>", self.on_model_selected)

		refresh_button = ttk.Button(model_frame, text="Refresh Models", command=self.fetch_models)
		refresh_button.pack(side=tk.RIGHT, padx=5)

		# Source and destination directory selection
		dir_frame = ttk.LabelFrame(main_frame, text="Directory Selection", padding="10")
		dir_frame.pack(fill=tk.X, pady=10)

		ttk.Label(dir_frame, text="Source Directory:").grid(row=0, column=0, sticky=tk.W, pady=5)
		self.source_dir_var = tk.StringVar()
		ttk.Entry(dir_frame, textvariable=self.source_dir_var, width=50).grid(row=0, column=1, padx=5, pady=5,
																			  sticky=tk.EW)
		ttk.Button(dir_frame, text="Browse", command=self.browse_source_dir).grid(row=0, column=2, padx=5, pady=5)

		ttk.Label(dir_frame, text="Destination Directory:").grid(row=1, column=0, sticky=tk.W, pady=5)
		self.dest_dir_var = tk.StringVar()
		ttk.Entry(dir_frame, textvariable=self.dest_dir_var, width=50).grid(row=1, column=1, padx=5, pady=5,
																			sticky=tk.EW)
		ttk.Button(dir_frame, text="Browse", command=self.browse_dest_dir).grid(row=1, column=2, padx=5, pady=5)

		dir_frame.columnconfigure(1, weight=1)

		# Category settings
		category_frame = ttk.LabelFrame(main_frame, text="Categories", padding="10")
		category_frame.pack(fill=tk.BOTH, expand=True, pady=10)

		self.category_text = tk.Text(category_frame, height=10)
		self.category_text.pack(fill=tk.BOTH, expand=True, pady=5)
		self.category_text.insert(tk.END, "Documents\nImages\nSpreadsheets\nPresentations\nPDFs\nOther")

		# Log output
		log_frame = ttk.LabelFrame(main_frame, text="Log", padding="10")
		log_frame.pack(fill=tk.BOTH, expand=True, pady=10)

		self.log_text = tk.Text(log_frame, height=10, state=tk.DISABLED)
		self.log_text.pack(fill=tk.BOTH, expand=True, pady=5)

		# Progress bar
		self.progress_var = tk.DoubleVar()
		self.progress_bar = ttk.Progressbar(main_frame, variable=self.progress_var, maximum=100)
		self.progress_bar.pack(fill=tk.X, pady=10)

		# Buttons
		button_frame = ttk.Frame(main_frame)
		button_frame.pack(fill=tk.X, pady=10)

		self.sort_button = ttk.Button(button_frame, text="Start Sorting", command=self.start_sorting)
		self.sort_button.pack(side=tk.RIGHT, padx=5)

		self.cancel_button = ttk.Button(button_frame, text="Cancel", command=self.cancel_sorting, state=tk.DISABLED)
		self.cancel_button.pack(side=tk.RIGHT, padx=5)

		# Processing flag
		self.is_processing = False
		self.cancel_requested = False

	def check_ollama_status(self):
		try:
			response = requests.get(f"{self.ollama_url}/version")
			if response.status_code == 200:
				self.status_label.config(text="Connected", foreground="green")
				self.fetch_models()
			else:
				self.status_label.config(text="Error: API response not OK", foreground="red")
		except requests.exceptions.ConnectionError:
			self.status_label.config(text="Disconnected (Is Ollama running?)", foreground="red")
			self.root.after(5000, self.check_ollama_status)  # Try again in 5 seconds

	def fetch_models(self):
		try:
			response = requests.get(f"{self.ollama_url}/tags")
			if response.status_code == 200:
				models_data = response.json()
				self.available_models = [model["name"] for model in models_data.get("models", [])]

				if not self.available_models:
					self.log_message("No models found in Ollama. Please pull a model first.")
					self.available_models = ["deepseek-coder", "llama3", "codellama", "mistral"]  # Default suggestions

				self.model_combobox["values"] = self.available_models

				# Select deepseek if available, otherwise first model
				if self.model in self.available_models:
					self.model_combobox.set(self.model)
				elif self.available_models:
					self.model_combobox.set(self.available_models[0])
					self.model = self.available_models[0]
			else:
				self.log_message(f"Error fetching models: {response.status_code}")
		except requests.exceptions.ConnectionError:
			self.log_message("Cannot connect to Ollama. Is it running?")
			self.status_label.config(text="Disconnected", foreground="red")

	def on_model_selected(self, event):
		self.model = self.model_combobox.get()
		self.log_message(f"Selected model: {self.model}")

	def browse_source_dir(self):
		dir_path = filedialog.askdirectory(title="Select Source Directory")
		if dir_path:
			self.source_dir_var.set(dir_path)

	def browse_dest_dir(self):
		dir_path = filedialog.askdirectory(title="Select Destination Directory")
		if dir_path:
			self.dest_dir_var.set(dir_path)

	def log_message(self, message):
		self.log_text.config(state=tk.NORMAL)
		self.log_text.insert(tk.END, f"{time.strftime('%H:%M:%S')} - {message}\n")
		self.log_text.see(tk.END)
		self.log_text.config(state=tk.DISABLED)

	def start_sorting(self):
		if self.is_processing:
			return

		source_dir = self.source_dir_var.get()
		dest_dir = self.dest_dir_var.get()

		if not source_dir or not os.path.isdir(source_dir):
			messagebox.showerror("Error", "Please select a valid source directory")
			return

		if not dest_dir or not os.path.isdir(dest_dir):
			messagebox.showerror("Error", "Please select a valid destination directory")
			return

		# Get categories from the text field
		categories_text = self.category_text.get("1.0", tk.END).strip()
		categories = [cat.strip() for cat in categories_text.split("\n") if cat.strip()]

		if not categories:
			messagebox.showerror("Error", "Please define at least one category")
			return

		# Create category folders if they don't exist
		for category in categories:
			category_path = os.path.join(dest_dir, category)
			if not os.path.exists(category_path):
				os.makedirs(category_path)

		# Start sorting thread
		self.sort_button.config(state=tk.DISABLED)
		self.cancel_button.config(state=tk.NORMAL)
		self.is_processing = True
		self.cancel_requested = False

		thread = threading.Thread(
			target=self.sort_documents,
			args=(source_dir, dest_dir, categories)
		)
		thread.daemon = True
		thread.start()

	def cancel_sorting(self):
		if self.is_processing:
			self.cancel_requested = True
			self.log_message("Cancellation requested. Waiting for current operations to complete...")

	def sort_documents(self, source_dir, dest_dir, categories):
		try:
			files = [f for f in os.listdir(source_dir) if os.path.isfile(os.path.join(source_dir, f))]

			if not files:
				self.log_message("No files found in the source directory")
				self.complete_sorting()
				return

			self.log_message(f"Found {len(files)} files to process")

			for i, filename in enumerate(files):
				if self.cancel_requested:
					self.log_message("Sorting cancelled")
					break

				file_path = os.path.join(source_dir, filename)

				# Skip very large files
				if os.path.getsize(file_path) > 1024 * 1024 * 10:  # 10MB
					self.log_message(f"Skipping large file: {filename}")
					continue

				self.log_message(f"Processing: {filename}")

				# Get file extension and basic info
				_, ext = os.path.splitext(filename)
				file_info = {
					"filename": filename,
					"extension": ext.lower(),
					"size_bytes": os.path.getsize(file_path)
				}

				# Classify the file using Ollama
				category = self.classify_file(file_info, categories)

				if category:
					dest_path = os.path.join(dest_dir, category, filename)

					# Check if destination file already exists
					if os.path.exists(dest_path):
						base, ext = os.path.splitext(filename)
						dest_path = os.path.join(dest_dir, category, f"{base}_copy{ext}")

					# Move the file
					shutil.move(file_path, dest_path)
					self.log_message(f"Moved '{filename}' to category '{category}'")
				else:
					self.log_message(f"Could not classify '{filename}', leaving in source folder")

				# Update progress
				progress = (i + 1) / len(files) * 100
				self.progress_var.set(progress)
				self.root.update_idletasks()

			self.log_message("Sorting completed")
		except Exception as e:
			self.log_message(f"Error during sorting: {str(e)}")
		finally:
			self.complete_sorting()

	def classify_file(self, file_info, categories):
		try:
			# Prepare prompt for the model
			prompt = f"""
            Please classify the following file into ONE of these categories: {', '.join(categories)}

            File information:
            - Filename: {file_info['filename']}
            - Extension: {file_info['extension']}
            - Size: {file_info['size_bytes']} bytes

            Respond with ONLY the category name, nothing else.
            """

			# Send request to Ollama
			response = requests.post(
				f"{self.ollama_url}/generate",
				json={
					"model": self.model,
					"prompt": prompt,
					"stream": False,
					"options": {
						"temperature": 0.1,
						"top_p": 0.9
					}
				}
			)

			if response.status_code == 200:
				result = response.json()
				# Extract just the category name from response
				prediction = result.get("response", "").strip()

				# Check if the predicted category exists in our list
				for category in categories:
					if category.lower() in prediction.lower():
						return category

				# Use extension-based fallback classification
				ext = file_info['extension'].lower()

				# Common file extensions mapping
				extension_map = {
					'doc': 'Documents', 'docx': 'Documents', 'txt': 'Documents', 'rtf': 'Documents',
					'jpg': 'Images', 'jpeg': 'Images', 'png': 'Images', 'gif': 'Images', 'bmp': 'Images',
					'xls': 'Spreadsheets', 'xlsx': 'Spreadsheets', 'csv': 'Spreadsheets',
					'ppt': 'Presentations', 'pptx': 'Presentations',
					'pdf': 'PDFs'
				}

				if ext[1:] in extension_map:
					category = extension_map[ext[1:]]
					if category in categories:
						return category

				# Default to "Other" if it exists
				if "Other" in categories:
					return "Other"

				# Use the first category as fallback
				return categories[0]
			else:
				self.log_message(f"Error from Ollama API: {response.status_code}")
				return None
		except Exception as e:
			self.log_message(f"Classification error: {str(e)}")
			return None

	def complete_sorting(self):
		self.is_processing = False
		self.sort_button.config(state=tk.NORMAL)
		self.cancel_button.config(state=tk.DISABLED)
		self.progress_var.set(0)


def main():
	root = tk.Tk()
	app = DocumentSorter(root)
	root.mainloop()


if __name__ == "__main__":
	main()