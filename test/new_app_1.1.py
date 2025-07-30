import sys
import re
import os
import json
import subprocess
from pathlib import Path
from threading import Thread
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QLabel, QListWidget,
    QFileDialog, QLineEdit, QCheckBox, QListWidgetItem, QComboBox, QTextEdit, QHBoxLayout
)
from PySide6.QtCore import QSettings, Qt, Signal, QObject
import pandas as pd

SUPPORTED_EXTENSIONS = ['.xlsx', '.xls', '.csv', '.txt']
NETWORK_FOLDERS = {
    "Apps": r"\\192.168.0.105\Apps",
    "SIM": r"\\192.168.0.105\Service\Quotation\SIM",
    "Sales": r"\\192.168.0.105\Sales Doc",
    "Scan": r"\\192.168.0.105\Scan",
    "Service": r"\\192.168.0.105\Service"
}

class SearchWorker(QObject):
    update_result = Signal(str, str)
    finished_file = Signal(str)
    finished = Signal()

    def __init__(self, files, keywords, exact_match):
        super().__init__()
        self.files = files
        self.keywords = keywords
        self.exact_match = exact_match

    def run(self):
        for file_path in self.files:
            try:
                matched_texts = set()
                if file_path.suffix.lower() in ['.xls', '.xlsx']:
                    df = pd.read_excel(file_path, header=None, dtype=str)
                    for row in df.itertuples(index=True):
                        for value in row[1:]:
                            if not isinstance(value, str):
                                continue
                            for keyword in self.keywords:
                                if self.exact_match:
                                    if re.fullmatch(re.escape(keyword), value.strip()):
                                        matched_texts.add(value)
                                else:
                                    if keyword.lower() in value.lower():
                                        matched_texts.add(value)
                else:
                    with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                        content = f.read()
                        for keyword in self.keywords:
                            if self.exact_match:
                                if re.search(rf'\b{re.escape(keyword)}\b', content):
                                    matched_texts.add(keyword)
                            else:
                                if keyword.lower() in content.lower():
                                    matched_texts.add(keyword)

                if matched_texts:
                    self.update_result.emit(str(file_path), ", ".join(matched_texts))
            except Exception as e:
                print(f"Failed to read {file_path}: {e}")
            finally:
                self.finished_file.emit(str(file_path))

        self.finished.emit()

class KeywordSearchApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Keyword File Search")
        self.settings = QSettings("MyCompany", "KeywordSearchApp")

        self.layout = QVBoxLayout(self)

        self.network_combo = QComboBox()
        self.network_combo.addItem("-- Select Network Folder --")
        for name in NETWORK_FOLDERS:
            self.network_combo.addItem(name)
        self.network_combo.currentIndexChanged.connect(self.open_network_folder)
        self.layout.addWidget(self.network_combo)

        self.folder_label = QLabel("No folder selected")
        self.layout.addWidget(self.folder_label)

        self.select_button = QPushButton("Select Folder")
        self.select_button.clicked.connect(self.select_folder)
        self.layout.addWidget(self.select_button)

        self.keyword_input = QLineEdit()
        self.keyword_input.setPlaceholderText("Enter keywords (e.g., product;branch)")
        self.layout.addWidget(self.keyword_input)

        self.exact_match_checkbox = QCheckBox("Exact Match")
        self.exact_match_checkbox.setChecked(True)
        self.layout.addWidget(self.exact_match_checkbox)

        self.search_button = QPushButton("Search")
        self.search_button.clicked.connect(self.search_keywords)
        self.layout.addWidget(self.search_button)

        # Display results
        self.result_layout = QHBoxLayout()
        self.result_list = QListWidget()
        self.result_list.itemClicked.connect(self.open_file)
        self.result_layout.addWidget(self.result_list)

        self.text_box = QTextEdit()
        self.text_box.setReadOnly(True)
        self.result_layout.addWidget(self.text_box)

        self.layout.addLayout(self.result_layout)

        self.files = []
        self.load_settings()

    def closeEvent(self, event):
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
        if not self.folder_path or not self.folder_path.is_dir():
            return

        for file_path in self.folder_path.iterdir():
            if file_path.name.startswith("~$"):
                continue
            if file_path.is_file() and file_path.suffix.lower() in SUPPORTED_EXTENSIONS:
                self.files.append(file_path)
                item = QListWidgetItem(str(file_path))
                self.result_list.addItem(item)

    def search_keywords(self):
        self.result_list.clear()
        self.text_box.clear()

        keywords = self.keyword_input.text().strip()
        if not keywords:
            return

        keyword_list = [k.strip() for k in re.split('[;,]', keywords) if k.strip()]
        exact_match = self.exact_match_checkbox.isChecked()

        self.worker = SearchWorker(self.files, keyword_list, exact_match)
        self.worker.update_result.connect(self.handle_result)
        self.worker.finished_file.connect(self.mark_file_scanned)
        self.worker.finished.connect(self.scan_complete)

        thread = Thread(target=self.worker.run)
        thread.start()

    def handle_result(self, file_path, found_text):
        items = self.result_list.findItems(file_path, Qt.MatchExactly)
        if not items:
            item = QListWidgetItem(file_path)
            self.result_list.addItem(item)
        else:
            item = items[0]
        item.setBackground(Qt.green)
        self.text_box.append(f"{file_path} => {found_text}")

    def mark_file_scanned(self, file_path):
        items = self.result_list.findItems(file_path, Qt.MatchExactly)
        if not items:
            item = QListWidgetItem(file_path)
            self.result_list.addItem(item)

    def scan_complete(self):
        self.text_box.append("\n✅ Scanning complete.")

    def open_file(self, item):
        path = Path(item.text())
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

    def load_settings(self):
        folder_str = self.settings.value("last_folder", "")
        exact_match = self.settings.value("exact_match", True, type=bool)
        self.folder_path = Path(folder_str) if folder_str else Path()
        self.folder_label.setText(str(self.folder_path))
        self.exact_match_checkbox.setChecked(exact_match)
        if self.folder_path.is_dir():
            self.list_files()

    def save_settings(self):
        self.settings.setValue("last_folder", str(self.folder_path))
        self.settings.setValue("exact_match", self.exact_match_checkbox.isChecked())

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = KeywordSearchApp()
    window.resize(800, 500)
    window.show()
    sys.exit(app.exec())
