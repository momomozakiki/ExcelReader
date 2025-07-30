import sys
import json
import pandas as pd
import os
from pathlib import Path
from PyQt6.QtWidgets import (
    QApplication, QWidget, QPushButton, QLineEdit, QLabel, QTextEdit,
    QFileDialog, QMessageBox, QGridLayout, QComboBox, QInputDialog
)
from openpyxl.utils import column_index_from_string

DEFAULTS_PATH = "folder_paths.json"

# Have problem to add new default path

class ExcelScannerGUI(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Keyword Scanner")
        self.resize(700, 550)

        self.default_folders = self.load_default_folders(DEFAULTS_PATH)

        self.init_ui()
        self.set_default_selection()

    def load_default_folders(self, config_path):
        try:
            with open(config_path, "r") as f:
                return json.load(f)
        except Exception as e:
            QMessageBox.warning(self, "Warning", f"Could not load default folders: {e}")
            return {"Folders": {}, "Default_Open": "", "Default_Save": ""}

    def save_default_folders(self):
        try:
            with open(DEFAULTS_PATH, "w") as f:
                json.dump(self.default_folders, f, indent=4)
        except Exception as e:
            QMessageBox.warning(self, "Warning", f"Could not save default folders: {e}")

    def init_ui(self):
        layout = QGridLayout()

        # Open Folder Dropdown
        self.open_combo = QComboBox()
        self.open_combo.addItem("Select a folder...")
        for name, path in self.default_folders.get("Folders", {}).items():
            self.open_combo.addItem(f"{name} ({path})", userData=name)
        self.open_combo.currentIndexChanged.connect(self.handle_open_selection)
        layout.addWidget(QLabel("Default Open Folder:"), 0, 0)
        layout.addWidget(self.open_combo, 0, 1, 1, 2)

        # Save Folder Dropdown
        self.save_combo = QComboBox()
        self.save_combo.addItem("Select a folder...")
        for name, path in self.default_folders.get("Folders", {}).items():
            self.save_combo.addItem(f"{name} ({path})", userData=name)
        self.save_combo.currentIndexChanged.connect(self.handle_save_selection)
        layout.addWidget(QLabel("Default Save Folder:"), 1, 0)
        layout.addWidget(self.save_combo, 1, 1, 1, 2)

        # Excel File Path
        self.file_input = QLineEdit()
        self.browse_file_btn = QPushButton("Browse Excel...")
        self.browse_file_btn.clicked.connect(self.select_excel_file)
        print(self.browse_file_btn)
        layout.addWidget(QLabel("Excel File:"), 2, 0)
        layout.addWidget(self.file_input, 2, 1)
        layout.addWidget(self.browse_file_btn, 2, 2)

        # Save Folder Path
        self.folder_input = QLineEdit()
        self.browse_folder_btn = QPushButton("Browse Folder...")
        self.browse_folder_btn.clicked.connect(self.select_output_folder)
        print(self.browse_folder_btn)
        layout.addWidget(QLabel("Save JSON To:"), 3, 0)
        layout.addWidget(self.folder_input, 3, 1)
        layout.addWidget(self.browse_folder_btn, 3, 2)

        # Set Default Button
        self.set_default_btn = QPushButton("Set As Default")
        self.set_default_btn.clicked.connect(self.set_selected_as_default)
        layout.addWidget(self.set_default_btn, 4, 0, 1, 3)

        # Add Current Folder to Defaults
        self.add_folder_btn = QPushButton("Add Folder to Defaults")
        self.add_folder_btn.clicked.connect(self.add_folder_to_defaults)
        layout.addWidget(self.add_folder_btn, 5, 0, 1, 3)

        # Keywords Input
        self.keyword_input = QLineEdit()
        layout.addWidget(QLabel("Keywords (comma-separated):"), 6, 0)
        layout.addWidget(self.keyword_input, 6, 1, 1, 2)

        # Start/End Cell Inputs
        self.start_cell_input = QLineEdit("A1")
        self.end_cell_input = QLineEdit("T58")
        layout.addWidget(QLabel("Start Cell (e.g., A1):"), 7, 0)
        layout.addWidget(self.start_cell_input, 7, 1, 1, 2)
        layout.addWidget(QLabel("End Cell (e.g., T58):"), 8, 0)
        layout.addWidget(self.end_cell_input, 8, 1, 1, 2)

        # Scan Button
        self.scan_btn = QPushButton("Scan Excel")
        self.scan_btn.clicked.connect(self.run_scan)
        layout.addWidget(self.scan_btn, 9, 0, 1, 3)

        # Output Area
        self.output_box = QTextEdit()
        self.output_box.setReadOnly(True)
        layout.addWidget(QLabel("Results:"), 10, 0)
        layout.addWidget(self.output_box, 11, 0, 1, 3)

        self.setLayout(layout)

    def set_default_selection(self):
        open_key = self.default_folders.get("Default_Open", "")
        save_key = self.default_folders.get("Default_Save", "")
        if open_key:
            index = self.open_combo.findData(open_key)
            if index != -1:
                self.open_combo.setCurrentIndex(index)
        if save_key:
            index = self.save_combo.findData(save_key)
            if index != -1:
                self.save_combo.setCurrentIndex(index)

    def set_selected_as_default(self):
        open_key = self.open_combo.currentData()
        save_key = self.save_combo.currentData()
        if open_key:
            self.default_folders["Default_Open"] = open_key
        if save_key:
            self.default_folders["Default_Save"] = save_key
        self.save_default_folders()
        QMessageBox.information(self, "Saved", "Default folder settings updated.")

    def handle_open_selection(self):
        key = self.open_combo.currentData()
        path = self.default_folders.get("Folders", {}).get(key, "")
        if path:
            self.file_input.setText(path + os.sep)

    def handle_save_selection(self):
        key = self.save_combo.currentData()
        path = self.default_folders.get("Folders", {}).get(key, "")
        if path:
            self.folder_input.setText(path)

    def add_folder_to_defaults(self):
        folder_path = QFileDialog.getExistingDirectory(self, "Select Folder to Add")
        if not folder_path:
            return

        name, ok = QInputDialog.getText(self, "Folder Name", "Enter name for the folder:")
        if not ok or not name.strip():
            return
        name = name.strip()

        if name in self.default_folders["Folders"]:
            QMessageBox.warning(self, "Exists", f"Folder name '{name}' already exists.")
            return

        self.default_folders["Folders"][name] = folder_path
        self.save_default_folders()
        self.open_combo.addItem(f"{name} ({folder_path})", userData=name)
        self.save_combo.addItem(f"{name} ({folder_path})", userData=name)

    def select_excel_file(self):
        start_dir = self.file_input.text().strip() or ""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Excel File", start_dir, "Excel Files (*.xlsx *.xls)"
        )
        print(f"selected file path : {file_path}")
        if file_path:
            self.file_input.setText(file_path)

    def select_output_folder(self):
        start_dir = self.folder_input.text().strip() or ""
        folder_path = QFileDialog.getExistingDirectory(self, "Select Output Folder", start_dir)
        print(f"selected output folder path : {folder_path}")
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

        keywords = [k.strip() for k in keyword_str.split(",") if k.strip()]

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
