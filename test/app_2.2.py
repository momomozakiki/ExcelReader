import sys
import json
import pandas as pd
import os
from pathlib import Path
from PyQt6.QtWidgets import (
    QApplication, QWidget, QPushButton, QLineEdit, QLabel, QTextEdit,
    QFileDialog, QMessageBox, QGridLayout, QComboBox
)
from openpyxl.utils import column_index_from_string


class ExcelScannerGUI(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Keyword Scanner")
        self.resize(700, 550)
        # Load network folders from config file
        self.network_folders = self.load_network_folders("network_folders.json")
        self.init_ui()

    def init_ui(self):
        layout = QGridLayout()

        # Excel File Dropdown + Input
        self.excel_combo = QComboBox()
        self.excel_combo.addItem("Select a predefined Excel folder...")
        for name in self.network_folders:
            self.excel_combo.addItem(name)
        self.excel_combo.currentTextChanged.connect(self.handle_excel_folder_selection)
        layout.addWidget(QLabel("Quick Excel Folder:"), 0, 0)
        layout.addWidget(self.excel_combo, 0, 1, 1, 2)

        self.file_input = QLineEdit()
        self.browse_file_btn = QPushButton("Browse Excel...")
        self.browse_file_btn.clicked.connect(self.select_excel_file)
        layout.addWidget(QLabel("Excel File:"), 1, 0)
        layout.addWidget(self.file_input, 1, 1)
        layout.addWidget(self.browse_file_btn, 1, 2)

        # Output Folder Dropdown + Input
        self.folder_combo = QComboBox()
        self.folder_combo.addItem("Select a predefined folder...")
        for name in self.network_folders:
            self.folder_combo.addItem(name)
        self.folder_combo.currentTextChanged.connect(self.handle_folder_selection)
        layout.addWidget(QLabel("Quick Network Folder:"), 2, 0)
        layout.addWidget(self.folder_combo, 2, 1, 1, 2)

        self.folder_input = QLineEdit()
        self.browse_folder_btn = QPushButton("Browse Folder...")
        self.browse_folder_btn.clicked.connect(self.select_output_folder)
        layout.addWidget(QLabel("Save JSON To:"), 3, 0)
        layout.addWidget(self.folder_input, 3, 1)
        layout.addWidget(self.browse_folder_btn, 3, 2)

        # Keywords Input
        self.keyword_input = QLineEdit()
        layout.addWidget(QLabel("Keywords (comma-separated):"), 4, 0)
        layout.addWidget(self.keyword_input, 4, 1, 1, 2)

        # Start/End Cell Inputs
        self.start_cell_input = QLineEdit("A1")
        self.end_cell_input = QLineEdit("T58")
        layout.addWidget(QLabel("Start Cell (e.g., A1):"), 5, 0)
        layout.addWidget(self.start_cell_input, 5, 1, 1, 2)
        layout.addWidget(QLabel("End Cell (e.g., T58):"), 6, 0)
        layout.addWidget(self.end_cell_input, 6, 1, 1, 2)

        # Scan Button
        self.scan_btn = QPushButton("Scan Excel")
        self.scan_btn.clicked.connect(self.run_scan)
        layout.addWidget(self.scan_btn, 7, 0, 1, 3)

        # Output Area
        self.output_box = QTextEdit()
        self.output_box.setReadOnly(True)
        layout.addWidget(QLabel("Results:"), 8, 0)
        layout.addWidget(self.output_box, 9, 0, 1, 3)

        self.setLayout(layout)

    def load_network_folders(self, config_path):
        try:
            with open(config_path, "r") as f:
                return json.load(f)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load network folders: {e}")
            return {}

    def handle_folder_selection(self, selection):
        if selection in self.network_folders:
            self.folder_input.setText(self.network_folders[selection])

    def handle_excel_folder_selection(self, selection):
        if selection in self.network_folders:
            self.file_input.setText(self.network_folders[selection] + os.sep)

    def select_excel_file(self):
        start_dir = self.file_input.text().strip() or ""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Excel File", start_dir, "Excel Files (*.xlsx *.xls)"
        )
        if file_path:
            self.file_input.setText(file_path)

    def select_output_folder(self):
        start_dir = self.folder_input.text().strip() or ""
        folder_path = QFileDialog.getExistingDirectory(self, "Select Output Folder", start_dir)
        if folder_path:
            self.folder_input.setText(folder_path)

    def get_cell_coords(self, cell_str):
        col_str = ''.join(filter(str.isalpha, cell_str))
        row_str = ''.join(filter(str.isdigit, cell_str))
        if not col_str or not row_str:
            raise ValueError(f"Invalid cell format: {cell_str}")
        col_idx = column_index_from_string(col_str) - 1
        row_idx = int(row_str) - 1
        return col_idx, row_idx

    def run_scan(self):
        file_path = self.file_input.text().strip()
        keyword_str = self.keyword_input.text().strip()
        folder_path = self.folder_input.text().strip() or os.path.dirname(file_path)

        if not file_path or not Path(file_path).exists():
            QMessageBox.critical(self, "Error", "Please select a valid Excel file.")
            return

        if not keyword_str:
            QMessageBox.critical(self, "Error", "Please enter at least one keyword.")
            return

        try:
            start_col_idx, start_row_idx = self.get_cell_coords(self.start_cell_input.text().upper())
            end_col_idx, end_row_idx = self.get_cell_coords(self.end_cell_input.text().upper())
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Invalid cell format: {e}")
            return

        keywords = [k.strip() for k in keyword_str.split(",")]

        try:
            result = self.scan_excel_with_pandas(
                file_path, keywords, start_col_idx, end_col_idx, start_row_idx, end_row_idx
            )
            formatted_result = json.dumps(result, indent=4)
            self.output_box.setPlainText(formatted_result)

            stem = Path(file_path).stem
            output_path = Path(folder_path) / f"{stem}.json"
            output_path.parent.mkdir(parents=True, exist_ok=True)

            with open(output_path, "w") as f:
                json.dump(result, f, indent=4)

            QMessageBox.information(self, "Success", f"Scan complete.\nFile saved to:\n{output_path}")

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to process file: {e}")

    def scan_excel_with_pandas(self, file_path, keywords, start_col=0, end_col=19, start_row=0, end_row=58):
        df = pd.read_excel(file_path, header=None)
        result = {}
        for row_idx, row in df.iterrows():
            if not (start_row <= row_idx <= end_row):
                continue
            for col_idx, value in enumerate(row):
                if not (start_col <= col_idx <= end_col):
                    continue
                cell_value = str(value).strip() if pd.notna(value) else ""
                if not cell_value:
                    continue
                matched_keywords = [k for k in keywords if k.lower() in cell_value.lower()]
                if matched_keywords:
                    cell_address = f"{chr(65 + col_idx)}{row_idx + 1}"
                    result[cell_address] = {k: cell_value for k in matched_keywords}
        return result


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelScannerGUI()
    window.show()
    sys.exit(app.exec())
