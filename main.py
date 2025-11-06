# main.py
import sys
import os
from PyQt6.QtWidgets import QApplication

# 关键：添加项目根目录到Python路径
ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

# 绝对导入
from core.gui import ExcelToInquiryLetter, DarkTheme

def main():
    app = QApplication(sys.argv)
    DarkTheme.apply(app)
    window = ExcelToInquiryLetter()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()