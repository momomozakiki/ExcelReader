import sys
import json
import pandas as pd
import os
from pathlib import Path
from PyQt6.QtWidgets import (
    QApplication, QWidget, QPushButton, QLineEdit,
    QLabel, QTextEdit, QFileDialog, QMessageBox, QGridLayout
)
from openpyxl.utils import column_index_from_string

# Local network drive
sales_dir = r'\\192.168.0.105\Sales Doc'
service_dir = r'\\192.168.0.105\Service'
scan_dir = r'\\192.168.0.105\Scan'
apps_dir = r'\\192.168.0.105\Apps'

class ExcelScannerGUI(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Keyword Scanner")
        self.resize(700, 550)

        self.init_ui()

    def init_ui(self):
        layout = QGridLayout()

        # File Selection
        self.file_input = QLineEdit()
        self.browse_file_btn = QPushButton("Browse Excel...")
        self.browse_file_btn.clicked.connect(self.select_excel_file)

        layout.addWidget(QLabel("Excel File:"), 0, 0)
        layout.addWidget(self.file_input, 0, 1)
        layout.addWidget(self.browse_file_btn, 0, 2)

        # Output Folder Selection
        self.folder_input = QLineEdit()
        self.browse_folder_btn = QPushButton("Browse Folder...")
        self.browse_folder_btn.clicked.connect(self.select_output_folder)

        layout.addWidget(QLabel("Save JSON To:"), 1, 0)
        layout.addWidget(self.folder_input, 1, 1)
        layout.addWidget(self.browse_folder_btn, 1, 2)

        # Keywords Input
        self.keyword_input = QLineEdit()
        layout.addWidget(QLabel("Keywords (comma-separated):"), 2, 0)
        layout.addWidget(self.keyword_input, 2, 1, 1, 2)

        # Start Cell
        self.start_cell_input = QLineEdit("A1")
        layout.addWidget(QLabel("Start Cell (e.g., A1):"), 3, 0)
        layout.addWidget(self.start_cell_input, 3, 1, 1, 2)

        # End Cell
        self.end_cell_input = QLineEdit("T58")
        layout.addWidget(QLabel("End Cell (e.g., T58):"), 4, 0)
        layout.addWidget(self.end_cell_input, 4, 1, 1, 2)

        # Scan Button
        self.scan_btn = QPushButton("Scan Excel")
        self.scan_btn.clicked.connect(self.run_scan)
        layout.addWidget(self.scan_btn, 5, 0, 1, 3)

        # Output Area
        self.output_box = QTextEdit()
        self.output_box.setReadOnly(True)
        layout.addWidget(QLabel("Results:"), 6, 0)
        layout.addWidget(self.output_box, 7, 0, 1, 3)

        self.setLayout(layout)

    def select_excel_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Excel File", "", "Excel Files (*.xlsx)"
        )
        if file_path:
            self.file_input.setText(file_path)

    def select_output_folder(self):
        folder_path = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if folder_path:
            self.folder_input.setText(folder_path)

    def get_cell_coords(self, cell_str):
        """Convert Excel-style cell (e.g., 'A1') to (col_idx, row_idx)"""
        col_str = ''.join(filter(str.isalpha, cell_str))
        row_str = ''.join(filter(str.isdigit, cell_str))

        if not col_str or not row_str:
            raise ValueError(f"Invalid cell format: {cell_str}")

        col_idx = column_index_from_string(col_str) - 1  # 0-based
        row_idx = int(row_str) - 1  # 0-based
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

            # Prepare output path
            stem = Path(file_path).stem
            output_path = Path(folder_path) / f"{stem}.json"

            # Create folder if not exists
            output_path.parent.mkdir(parents=True, exist_ok=True)

            # Save to JSON
            with open(output_path, "w") as f:
                json.dump(result, f, indent=4)

            QMessageBox.information(self, "Success", f"Scan complete.\nFile saved to:\n{output_path}")

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to process file: {e}")

    def scan_excel_with_pandas(self, file_path, keywords, start_col=0, end_col=19, start_row=0, end_row=58):
        df = pd.read_excel(file_path, header=None)
        result = {}

        for row_idx, row in df.iterrows():
            if not (start_row <= row_idx < end_row):
                continue

            for col_idx, value in enumerate(row):
                if not (start_col <= col_idx <= end_col):
                    continue

                cell_value = str(value).strip() if pd.notna(value) else ""
                if not cell_value:
                    continue

                matched_keywords = [
                    keyword for keyword in keywords
                    if keyword.lower() in cell_value.lower()
                ]

                if matched_keywords:
                    column_letter = chr(65 + col_idx)
                    cell_address = f"{column_letter}{row_idx + 1}"
                    result[cell_address] = {
                        keyword: cell_value for keyword in matched_keywords
                    }

        return result


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelScannerGUI()
    window.show()
    sys.exit(app.exec())