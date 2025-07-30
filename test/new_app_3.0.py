# Suppress specific UserWarnings from openpyxl (e.g., unsupported Data Validation extensions)
# These warnings are harmless and occur when reading Excel files with features not supported by openpyxl
import warnings

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

import sys
import re
import os
import json
import subprocess
from pathlib import Path
from threading import Thread
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QLabel, QListWidget,
    QFileDialog, QLineEdit, QCheckBox, QListWidgetItem, QComboBox, QTextEdit,
    QHBoxLayout, QMenu, QProgressBar, QSplitter, QDialog, QFormLayout, QDialogButtonBox,
    QMessageBox
)
from PySide6.QtCore import QSettings, Qt, Signal, QObject, QTimer
import pandas as pd
import concurrent.futures
import threading

# Supported file types for searching
SUPPORTED_EXTENSIONS = ['.xlsx', '.xls', '.csv', '.txt']

# Predefined network folder shortcuts
NETWORK_FOLDERS = {
    "Apps": r"\\192.168.0.105\Apps",
    "SIM": r"\\192.168.0.105\Service\Quotation\SIM",
    "Sales": r"\\192.168.0.105\Sales Doc",
    "Scan": r"\\192.168.0.105\Scan",
    "Service": r"\\192.168.0.105\Service"
}


class SettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Search Settings")
        self.setModal(True)
        self.resize(400, 300)

        layout = QVBoxLayout(self)
        form_layout = QFormLayout()

        # Column end keywords
        self.col_end_input = QTextEdit()
        self.col_end_input.setMaximumHeight(80)
        self.col_end_input.setPlaceholderText("Enter keywords separated by semicolons (e.g., E. & O.E.;SUB-TOTAL)")
        form_layout.addRow("Column End Keywords:", self.col_end_input)

        # Row end column
        self.row_end_input = QLineEdit()
        self.row_end_input.setPlaceholderText("Enter column letter (e.g., N)")
        form_layout.addRow("Row End Column:", self.row_end_input)

        # Max rows to scan
        self.max_rows_input = QLineEdit()
        self.max_rows_input.setPlaceholderText("Enter maximum rows to scan (default: 1000)")
        form_layout.addRow("Max Rows to Scan:", self.max_rows_input)

        layout.addLayout(form_layout)

        # Buttons
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

    def get_settings(self):
        col_end_text = self.col_end_input.toPlainText().strip()
        col_end_keywords = set()
        if col_end_text:
            col_end_keywords = {k.strip() for k in col_end_text.split(';') if k.strip()}

        row_end = self.row_end_input.text().strip().upper()
        if not row_end:
            row_end = 'N'

        max_rows = self.max_rows_input.text().strip()
        try:
            max_rows = int(max_rows) if max_rows else 1000
        except ValueError:
            max_rows = 1000

        return col_end_keywords, row_end, max_rows

    def set_settings(self, col_end_keywords, row_end, max_rows):
        self.col_end_input.setPlainText(';'.join(col_end_keywords))
        self.row_end_input.setText(row_end)
        self.max_rows_input.setText(str(max_rows))


class SearchWorker(QObject):
    # Signals to communicate back to the main thread
    update_result = Signal(str, str, str)  # file_path, file_name, found_text
    finished_file = Signal(str, str)  # file_path, file_name
    finished = Signal()
    progress_update = Signal(int, int)  # current, total

    def __init__(self, files, keywords, exact_match, col_end_keywords=None, row_end='N', max_rows=1000):
        super().__init__()
        self.files = files
        self.keywords = keywords
        self.exact_match = exact_match
        self.col_end_keywords = col_end_keywords or set()
        self.row_end = row_end
        self.max_rows = max_rows
        self.should_stop = False
        self.executor = concurrent.futures.ThreadPoolExecutor(max_workers=4)

    def stop(self):
        self.should_stop = True
        self.executor.shutdown(wait=False)

    def get_column_number(self, col_letter):
        """Convert column letter to number (A=1, B=2, etc.)"""
        result = 0
        for char in col_letter:
            result = result * 26 + (ord(char) - ord('A') + 1)
        return result

    def search_file(self, file_path):
        if self.should_stop:
            return None

        try:
            matched_texts = set()
            file_name = file_path.name

            # Handle Excel files
            if file_path.suffix.lower() in ['.xls', '.xlsx']:
                # Read the file with limited rows
                df = pd.read_excel(file_path, header=None, dtype=str, nrows=self.max_rows)

                # Find the actual end row using column end keywords
                end_row = len(df)
                if self.col_end_keywords:
                    for idx, row in df.iterrows():
                        if self.should_stop:
                            return None
                        for value in row:
                            if isinstance(value, str):
                                for keyword in self.col_end_keywords:
                                    if keyword.lower() in value.lower():
                                        end_row = idx
                                        break
                                if end_row != len(df):
                                    break
                        if end_row != len(df):
                            break

                # Limit columns based on row_end setting
                max_col = self.get_column_number(self.row_end)

                # Search only within the limited area
                for row_idx in range(min(end_row, len(df))):
                    if self.should_stop:
                        return None
                    row = df.iloc[row_idx]
                    for col_idx, value in enumerate(row[:max_col]):  # Limit columns
                        if not isinstance(value, str):
                            continue
                        for keyword in self.keywords:
                            if self.exact_match:
                                if re.fullmatch(re.escape(keyword), value.strip()):
                                    matched_texts.add(value)
                            else:
                                if keyword.lower() in value.lower():
                                    matched_texts.add(value)
                            if len(matched_texts) >= 10:  # Limit matches for performance
                                break
                        if len(matched_texts) >= 10:
                            break
                    if len(matched_texts) >= 10:
                        break
            else:
                # Handle plain text or CSV files
                with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                    # Read only first 1MB for performance
                    content = f.read(1024 * 1024)
                    for keyword in self.keywords:
                        if self.should_stop:
                            return None
                        if self.exact_match:
                            matches = re.findall(rf'\b{re.escape(keyword)}\b', content)
                            matched_texts.update(matches[:10])  # Limit matches
                        else:
                            if keyword.lower() in content.lower():
                                matched_texts.add(keyword)

            return (str(file_path), file_name, ", ".join(list(matched_texts)[:10]))
        except Exception as e:
            print(f"Failed to read {file_path}: {e}")
            return None

    def run(self):
        total_files = len(self.files)
        futures = []

        # Submit all files to thread pool
        for i, file_path in enumerate(self.files):
            if self.should_stop:
                break
            future = self.executor.submit(self.search_file, file_path)
            futures.append((future, file_path))

        # Process results as they complete
        for i, (future, file_path) in enumerate(futures):
            if self.should_stop:
                break

            try:
                result = future.result(timeout=30)  # 30 second timeout per file
                file_name = file_path.name

                if result:
                    file_path_str, file_name, matched_texts = result
                    if matched_texts:
                        self.update_result.emit(file_path_str, file_name, matched_texts)

                self.finished_file.emit(str(file_path), file_name)
                self.progress_update.emit(i + 1, total_files)

            except concurrent.futures.TimeoutError:
                print(f"Timeout reading {file_path}")
                self.finished_file.emit(str(file_path), file_path.name)
            except Exception as e:
                print(f"Error processing {file_path}: {e}")
                self.finished_file.emit(str(file_path), file_path.name)

        self.executor.shutdown(wait=True)
        self.finished.emit()


class KeywordSearchApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Keyword File Search")
        self.settings = QSettings("MyCompany", "KeywordSearchApp")

        self.folder_path = Path()  # Ensure folder_path attribute exists
        self.worker = None
        self.search_thread = None

        # Default search settings
        self.col_end_keywords = {'E. & O.E.', 'SUB-TOTAL'}
        self.row_end = 'N'
        self.max_rows = 1000

        self.layout = QVBoxLayout(self)

        # Dropdown for selecting predefined network folders
        self.network_combo = QComboBox()
        self.network_combo.addItem("-- Select Network Folder --")
        for name in NETWORK_FOLDERS:
            self.network_combo.addItem(name)
        self.network_combo.currentIndexChanged.connect(self.open_network_folder)
        self.layout.addWidget(self.network_combo)

        # Shows the currently selected folder path
        self.folder_label = QLabel("No folder selected")
        self.layout.addWidget(self.folder_label)

        # Button to open folder selector
        self.select_button = QPushButton("Select Folder")
        self.select_button.clicked.connect(self.select_folder)
        self.layout.addWidget(self.select_button)

        # Input for keywords
        self.keyword_input = QLineEdit()
        self.keyword_input.setPlaceholderText("Enter keywords (e.g., product;branch)")
        self.layout.addWidget(self.keyword_input)

        # Checkbox to control exact vs partial match
        self.exact_match_checkbox = QCheckBox("Exact Match")
        self.exact_match_checkbox.setChecked(True)
        self.layout.addWidget(self.exact_match_checkbox)

        # Settings button
        self.settings_button = QPushButton("Search Settings")
        self.settings_button.clicked.connect(self.show_settings)
        self.layout.addWidget(self.settings_button)

        # Search and Stop buttons
        button_layout = QHBoxLayout()
        self.search_button = QPushButton("Search")
        self.search_button.clicked.connect(self.search_keywords)
        button_layout.addWidget(self.search_button)

        self.stop_button = QPushButton("Stop")
        self.stop_button.clicked.connect(self.stop_search)
        self.stop_button.setEnabled(False)
        button_layout.addWidget(self.stop_button)

        self.layout.addLayout(button_layout)

        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.layout.addWidget(self.progress_bar)

        # Splitter for resizable panels
        splitter = QSplitter(Qt.Horizontal)

        # Left panel - file list
        self.result_list = QListWidget()
        self.result_list.itemClicked.connect(self.show_context_menu)
        splitter.addWidget(self.result_list)

        # Right panel - preview
        self.preview_box = QTextEdit()
        self.preview_box.setReadOnly(True)
        self.preview_box.setPlaceholderText("File preview will appear here...")
        splitter.addWidget(self.preview_box)

        # Set initial splitter sizes (60% left, 40% right)
        splitter.setSizes([480, 320])

        self.layout.addWidget(splitter)

        self.files = []
        self.file_paths = {}  # Map file names to full paths
        self.load_settings()

    def closeEvent(self, event):
        self.stop_search()
        self.save_settings()
        super().closeEvent(event)

    def select_folder(self):
        folder_path = QFileDialog.getExistingDirectory(self, "Select Folder", str(self.folder_label.text()))
        if folder_path:
            self.folder_path = Path(folder_path)
            self.folder_label.setText(str(self.folder_path))
            self.list_files()

    def open_network_folder(self, index):
        if index <= 0:
            return
        key = self.network_combo.currentText()
        path_str = NETWORK_FOLDERS.get(key, "")
        if path_str:
            self.folder_path = Path(path_str)
            self.folder_label.setText(str(self.folder_path))
            if self.folder_path.exists():
                self.list_files()

    def list_files(self):
        self.result_list.clear()
        self.files = []
        self.file_paths = {}

        # Ensure folder_path is a valid directory before proceeding
        if not hasattr(self, 'folder_path') or not self.folder_path.is_dir():
            return

        for file_path in self.folder_path.iterdir():
            try:
                if file_path.name.startswith("~$"):  # Ignore temporary Excel lock files
                    continue
                if file_path.is_file() and file_path.suffix.lower() in SUPPORTED_EXTENSIONS:
                    self.files.append(file_path)
                    self.file_paths[file_path.name] = file_path
                    item = QListWidgetItem(file_path.name)  # Show only file name
                    self.result_list.addItem(item)
            except (PermissionError, OSError) as e:
                # Skip files that can't be accessed (e.g., locked, permission denied, network delays)
                print(f"Skipped file {file_path}: {e}")
                continue

    def search_keywords(self):
        self.result_list.clear()
        self.preview_box.clear()
        self.file_paths = {}

        keywords = self.keyword_input.text().strip()
        if not keywords:
            return

        keyword_list = [k.strip() for k in re.split('[;,]', keywords) if k.strip()]
        exact_match = self.exact_match_checkbox.isChecked()

        # Setup progress bar
        self.progress_bar.setMaximum(len(self.files))
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(True)

        # Disable search button and enable stop button
        self.search_button.setEnabled(False)
        self.stop_button.setEnabled(True)

        # Create a background worker to handle the file scanning
        self.worker = SearchWorker(self.files, keyword_list, exact_match,
                                   self.col_end_keywords, self.row_end, self.max_rows)
        self.worker.update_result.connect(self.handle_result)
        self.worker.finished_file.connect(self.mark_file_scanned)
        self.worker.finished.connect(self.scan_complete)
        self.worker.progress_update.connect(self.update_progress)

        self.search_thread = Thread(target=self.worker.run)
        self.search_thread.start()

    def show_settings(self):
        dialog = SettingsDialog(self)
        dialog.set_settings(self.col_end_keywords, self.row_end, self.max_rows)

        if dialog.exec() == QDialog.Accepted:
            self.col_end_keywords, self.row_end, self.max_rows = dialog.get_settings()
            self.save_settings()  # Save settings immediately

    def stop_search(self):
        if self.worker:
            self.worker.stop()
        self.search_button.setEnabled(True)
        self.stop_button.setEnabled(False)
        self.progress_bar.setVisible(False)

    def update_progress(self, current, total):
        self.progress_bar.setValue(current)

    def handle_result(self, file_path, file_name, found_text):
        self.file_paths[file_name] = Path(file_path)
        items = self.result_list.findItems(file_name, Qt.MatchExactly)
        if not items:
            item = QListWidgetItem(file_name)
            self.result_list.addItem(item)
        else:
            item = items[0]
        item.setBackground(Qt.green)

    def mark_file_scanned(self, file_path, file_name):
        if file_name not in self.file_paths:
            self.file_paths[file_name] = Path(file_path)
        items = self.result_list.findItems(file_name, Qt.MatchExactly)
        if not items:
            item = QListWidgetItem(file_name)
            self.result_list.addItem(item)

    def scan_complete(self):
        self.search_button.setEnabled(True)
        self.stop_button.setEnabled(False)
        self.progress_bar.setVisible(False)
        self.preview_box.append("\n✅ Scanning complete.")

    def show_context_menu(self, item):
        file_name = item.text()
        if file_name not in self.file_paths:
            return

        menu = QMenu(self)

        open_action = menu.addAction("Open")
        preview_action = menu.addAction("Preview")
        export_json_action = menu.addAction("Export to JSON")

        action = menu.exec(self.result_list.mapToGlobal(self.result_list.pos()))

        if action == open_action:
            self.open_file(file_name)
        elif action == preview_action:
            self.preview_file(file_name)
        elif action == export_json_action:
            self.export_to_json(file_name)

    def open_file(self, file_name):
        if file_name not in self.file_paths:
            return

        path = self.file_paths[file_name]
        if path.exists():
            try:
                if sys.platform == "win32":
                    os.startfile(str(path))
                else:
                    subprocess.call(["xdg-open" if sys.platform.startswith("linux") else "open", str(path)])
            except Exception as e:
                print(f"Failed to open file: {e}")
        else:
            print(f"File not found: {path}")

    def export_to_json(self, file_name):
        if file_name not in self.file_paths:
            return

        path = self.file_paths[file_name]
        if not path.exists():
            QMessageBox.warning(self, "Error", "File not found")
            return

        try:
            # Ask user where to save the JSON file
            json_path, _ = QFileDialog.getSaveFileName(
                self, "Save JSON File",
                f"{path.stem}.json",
                "JSON files (*.json)"
            )

            if not json_path:
                return

            json_data = self.extract_json_data(path)

            # Save to JSON file
            with open(json_path, 'w', encoding='utf-8') as f:
                json.dump(json_data, f, indent=2, ensure_ascii=False)

            QMessageBox.information(self, "Success", f"JSON file saved to:\n{json_path}")

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to export JSON: {e}")

    def extract_json_data(self, file_path):
        """Extract structured data from Excel file and convert to JSON format"""
        json_data = {
            "file_info": {
                "filename": file_path.name,
                "path": str(file_path),
                "extracted_date": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
            },
            "content": {}
        }

        if file_path.suffix.lower() in ['.xls', '.xlsx']:
            # Read Excel file
            df = pd.read_excel(file_path, header=None, dtype=str)

            # Find the actual end row using column end keywords
            end_row = len(df)
            if self.col_end_keywords:
                for idx, row in df.iterrows():
                    for value in row:
                        if isinstance(value, str):
                            for keyword in self.col_end_keywords:
                                if keyword.lower() in value.lower():
                                    end_row = idx
                                    break
                            if end_row != len(df):
                                break
                    if end_row != len(df):
                        break

            # Limit columns based on row_end setting
            max_col = self.get_column_number(self.row_end)

            # Extract structured data
            structured_data = self.extract_structured_sections(df, end_row, max_col)
            json_data["content"] = structured_data

        elif file_path.suffix.lower() == '.csv':
            # Handle CSV files
            df = pd.read_csv(file_path, header=None, dtype=str, nrows=self.max_rows)
            json_data["content"]["raw_data"] = df.fillna("").to_dict(orient='records')

        else:
            # Handle text files
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read(5000)  # Read first 5KB
                json_data["content"]["raw_text"] = content

        return json_data

    def get_column_number(self, col_letter):
        """Convert column letter to number (A=1, B=2, etc.)"""
        result = 0
        for char in col_letter:
            result = result * 26 + (ord(char) - ord('A') + 1)
        return result

    def extract_structured_sections(self, df, end_row, max_col):
        """Extract structured sections from DataFrame"""
        structured_data = {
            "header_info": {},
            "table_data": [],
            "summary_info": {},
            "raw_data": []
        }

        # Convert DataFrame to limited area
        limited_df = df.iloc[:end_row, :max_col].fillna("")

        # Extract header information (first few rows)
        header_rows = min(10, len(limited_df))
        for i in range(header_rows):
            row_data = limited_df.iloc[i].to_dict()
            # Look for key-value pairs in header
            for col_idx, value in enumerate(row_data.values()):
                if isinstance(value, str) and value.strip():
                    if ':' in value:
                        parts = value.split(':', 1)
                        if len(parts) == 2:
                            key = parts[0].strip()
                            val = parts[1].strip()
                            if key and val:
                                structured_data["header_info"][key] = val

        # Extract table data (look for structured tables)
        table_start = -1
        for i in range(len(limited_df)):
            row = limited_df.iloc[i]
            # Look for table headers (rows with multiple non-empty cells)
            non_empty_count = sum(1 for cell in row if str(cell).strip())
            if non_empty_count >= 3:  # Assume table if 3+ columns have data
                table_start = i
                break

        if table_start != -1:
            # Extract table data
            table_end = min(table_start + 20, len(limited_df))  # Limit table rows
            for i in range(table_start, table_end):
                row_data = limited_df.iloc[i].to_list()
                # Skip empty rows
                if any(str(cell).strip() for cell in row_data):
                    structured_data["table_data"].append({
                        f"col_{j}": str(cell).strip() if str(cell).strip() else ""
                        for j, cell in enumerate(row_data)
                    })

        # Extract summary information (last few rows before end keywords)
        summary_start = max(0, end_row - 5)
        for i in range(summary_start, end_row):
            row_data = limited_df.iloc[i].to_dict()
            for col_idx, value in enumerate(row_data.values()):
                if isinstance(value, str) and value.strip():
                    # Look for summary patterns
                    if any(keyword in value.lower() for keyword in ['total', 'subtotal', 'amount', 'gst', 'tax']):
                        structured_data["summary_info"][f"row_{i}_col_{col_idx}"] = value.strip()

        # Raw data for reference
        structured_data["raw_data"] = limited_df.to_dict(orient='records')

        return structured_data

    def preview_file(self, file_name):
        if file_name not in self.file_paths:
            return

        path = self.file_paths[file_name]
        if not path.exists():
            self.preview_box.setText("File not found")
            return

        try:
            self.preview_box.clear()
            self.preview_box.append(f"Preview: {file_name}\n" + "=" * 50 + "\n")

            if path.suffix.lower() in ['.xls', '.xlsx']:
                # Preview Excel file
                df = pd.read_excel(path, header=None, dtype=str, nrows=50)  # Limit to 50 rows
                preview_text = df.to_string(index=False, header=False, max_rows=50)
                self.preview_box.append(preview_text)
            else:
                # Preview text/CSV file
                with open(path, 'r', encoding='utf-8', errors='ignore') as f:
                    content = f.read(5000)  # Read first 5KB
                    self.preview_box.append(content)

        except Exception as e:
            self.preview_box.setText(f"Error previewing file: {e}")


    def load_settings(self):
        folder_str = self.settings.value("last_folder", "")
        exact_match = self.settings.value("exact_match", True, type=bool)

        # Load search settings
        col_end_str = self.settings.value("col_end_keywords", "E. & O.E.;SUB-TOTAL")
        self.col_end_keywords = {k.strip() for k in col_end_str.split(';') if k.strip()}
        self.row_end = self.settings.value("row_end", "N")
        self.max_rows = self.settings.value("max_rows", 1000, type=int)

        self.folder_path = Path(folder_str) if folder_str else Path()
        self.folder_label.setText(str(self.folder_path))
        self.exact_match_checkbox.setChecked(exact_match)
        if self.folder_path.is_dir():
            self.list_files()


    def save_settings(self):
        self.settings.setValue("last_folder", str(self.folder_path))
        self.settings.setValue("exact_match", self.exact_match_checkbox.isChecked())

        # Save search settings
        self.settings.setValue("col_end_keywords", ";".join(self.col_end_keywords))
        self.settings.setValue("row_end", self.row_end)
        self.settings.setValue("max_rows", self.max_rows)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = KeywordSearchApp()
    window.resize(1000, 600)
    window.show()
    sys.exit(app.exec())