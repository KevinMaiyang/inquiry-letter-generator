# main.py
import sys
from PyQt6.QtWidgets import QApplication

# 显式导入（帮助 PyInstaller 识别）
import gui
import template_manager
import excel_generator
import pdf_generator
import utils

from gui import ExcelToInquiryLetter, DarkTheme

def main():
    app = QApplication(sys.argv)
    DarkTheme.apply(app)
    window = ExcelToInquiryLetter()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()