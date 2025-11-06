# main.py
import sys
from PyQt6.QtWidgets import QApplication
from gui import ExcelToInquiryLetter, DarkTheme

def main():
    app = QApplication(sys.argv)
    DarkTheme.apply(app)
    window = ExcelToInquiryLetter()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()