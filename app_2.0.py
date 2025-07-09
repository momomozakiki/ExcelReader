import sys
import json
import pandas as pd
from pathlib import Path
from PyQt6.QtWidgets import (
    QApplication, QWidget, QPushButton, QLineEdit,
    QLabel, QTextEdit, QFileDialog, QGridLayout, QMessageBox
)
from openpyxl.utils import column_index_from_string


class ExcelScannerGUI(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Keyword Scanner")
        self.resize(700, 500)

        self.init_ui()

    def init_ui(self):
        layout = QGridLayout()

        # File Selection
        self.file_input = QLineEdit()
        self.browse_btn = QPushButton("Browse...")
        self.browse_btn.clicked.connect(self.select_file)

        layout.addWidget(QLabel("Excel File:"), 0, 0)
        layout.addWidget(self.file_input, 0, 1)
        layout.addWidget(self.browse_btn, 0, 2)

        # Keywords Input
        self.keyword_input = QLineEdit()
        layout.addWidget(QLabel("Keywords (comma-separated):"), 1, 0)
        layout.addWidget(self.keyword_input, 1, 1, 1, 2)

        # Start Cell
        self.start_cell_input = QLineEdit("A1")
        layout.addWidget(QLabel("Start Cell (e.g., A1):"), 2, 0)
        layout.addWidget(self.start_cell_input, 2, 1, 1, 2)

        # End Cell
        self.end_cell_input = QLineEdit("T58")
        layout.addWidget(QLabel("End Cell (e.g., T58):"), 3, 0)
        layout.addWidget(self.end_cell_input, 3, 1, 1, 2)

        # Scan Button
        self.scan_btn = QPushButton("Scan Excel")
        self.scan_btn.clicked.connect(self.run_scan)
        layout.addWidget(self.scan_btn, 4, 0, 1, 3)

        # Output Area
        self.output_box = QTextEdit()
        self.output_box.setReadOnly(True)
        layout.addWidget(QLabel("Results:"), 5, 0)
        layout.addWidget(self.output_box, 6, 0, 1, 3)

        self.setLayout(layout)

    def select_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Excel File", "", "Excel Files (*.xlsx)"
        )
        if file_path:
            self.file_input.setText(file_path)

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

            # Save to JSON
            stem = Path(file_path).stem
            with open(f"{stem}.json", "w") as f:
                json.dump(result, f, indent=4)

            QMessageBox.information(self, "Success", f"Scan complete. Result saved to '{stem}.json'")

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