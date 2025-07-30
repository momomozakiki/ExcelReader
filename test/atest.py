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

print(QWidget.__)
