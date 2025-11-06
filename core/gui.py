# core/gui.py
import sys
import os
import pandas as pd
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QFileDialog, QLabel, QLineEdit, QMessageBox, QInputDialog, QDateEdit, QMenuBar, QMenu
)
from PyQt6.QtCore import Qt, QDate
from PyQt6.QtGui import QPalette, QColor, QIcon

# 添加根目录到路径（确保模块能找到）
ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

# 绝对导入
from core.template_manager import TemplateManager
from generators.excel_generator import generate_excel
from generators.pdf_generator import generate_pdfs
from core.utils import get_user_template_path, get_season_from_date

class DarkTheme:
    @staticmethod
    def apply(app):
        palette = QPalette()
        palette.setColor(QPalette.ColorRole.Window, QColor(45, 45, 48))
        palette.setColor(QPalette.ColorRole.WindowText, QColor(240, 240, 240))
        palette.setColor(QPalette.ColorRole.Base, QColor(30, 30, 30))
        palette.setColor(QPalette.ColorRole.AlternateBase, QColor(45, 45, 48))
        palette.setColor(QPalette.ColorRole.ToolTipBase, QColor(45, 45, 48))
        palette.setColor(QPalette.ColorRole.ToolTipText, QColor(240, 240, 240))
        palette.setColor(QPalette.ColorRole.Text, QColor(240, 240, 240))
        palette.setColor(QPalette.ColorRole.Button, QColor(68, 68, 68))
        palette.setColor(QPalette.ColorRole.ButtonText, QColor(240, 240, 240))
        palette.setColor(QPalette.ColorRole.BrightText, Qt.GlobalColor.red)
        palette.setColor(QPalette.ColorRole.Link, QColor(42, 130, 218))
        palette.setColor(QPalette.ColorRole.Highlight, QColor(42, 130, 218))
        palette.setColor(QPalette.ColorRole.HighlightedText, QColor(35, 35, 35))
        app.setPalette(palette)
        app.setStyle("Fusion")

# ========== 保留：普通按钮（用于“选择文件”等） ==========
class StyledButton(QPushButton):
    def __init__(self, text, parent=None):
        super().__init__(text, parent)
        self.setStyleSheet("""
            QPushButton {
                background-color: #4a4a4a;
                color: white;
                border: 1px solid #5a5a5a;
                border-radius: 5px;
                padding: 8px 16px;
                font-weight: bold;
                min-height: 20px;
            }
            QPushButton:hover {
                background-color: #5a5a5a;
                border: 1px solid #6a6a6a;
            }
            QPushButton:pressed {
                background-color: #3a3a3a;
            }
            QPushButton:disabled {
                background-color: #353535;
                color: #7a7a7a;
            }
        """)


# ========== 新增：带图标+渐变色的按钮（用于“导出Excel/PDF”） ==========
class IconButton(QPushButton):
    def __init__(self, text, icon_path, gradient_color_start, gradient_color_end, parent=None):
        super().__init__(text, parent)
        if os.path.exists(icon_path):
            self.setIcon(QIcon(icon_path))
        self.setStyleSheet(f"""
            QPushButton {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 {gradient_color_start},
                    stop:1 {gradient_color_end});
                color: white;
                border: none;
                border-radius: 6px;
                padding: 8px 16px 8px 40px;
                font-weight: bold;
                text-align: left;
                min-height: 32px;
                icon-size: 20px;
            }}
            QPushButton:hover {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 {gradient_color_start},
                    stop:1 #444444);
            }}
            QPushButton:pressed {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #222222,
                    stop:1 {gradient_color_end});
            }}
            QPushButton:disabled {{
                background: #555555;
                color: #aaaaaa;
            }}
        """)

class IconButton(QPushButton):
    def __init__(self, text, icon_path, gradient_color_start, gradient_color_end, parent=None):
        super().__init__(text, parent)
        if os.path.exists(icon_path):
            self.setIcon(QIcon(icon_path))
        self.setStyleSheet(f"""
            QPushButton {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 {gradient_color_start},
                    stop:1 {gradient_color_end});
                color: white;
                border: none;
                border-radius: 6px;
                padding: 8px 16px 8px 40px;
                font-weight: bold;
                text-align: left;
                min-height: 32px;
                icon-size: 20px;
            }}
            QPushButton:hover {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 {gradient_color_start},
                    stop:1 #444444);
            }}
            QPushButton:pressed {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #222222,
                    stop:1 {gradient_color_end});
            }}
            QPushButton:disabled {{
                background: #555555;
                color: #aaaaaa;
            }}
        """)


class StyledLineEdit(QLineEdit):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setStyleSheet("""
            QLineEdit {
                background-color: #252525;
                color: white;
                border: 1px solid #5a5a5a;
                border-radius: 3px;
                padding: 6px 8px;
                selection-background-color: #3daee9;
            }
            QLineEdit:focus {
                border: 1px solid #3daee9;
            }
            QLineEdit:disabled {
                background-color: #353535;
                color: #7a7a7a;
            }
        """)


class StyledLabel(QLabel):
    def __init__(self, text, parent=None, is_section=False):
        super().__init__(text, parent)
        if is_section:
            self.setStyleSheet("""
                QLabel {
                    color: #3daee9;
                    font-weight: bold;
                    font-size: 12px;
                    padding: 5px 0px;
                }
            """)
        else:
            self.setStyleSheet("""
                QLabel {
                    color: #e0e0e0;
                    padding: 2px 0px;
                }
            """)


class ExcelToInquiryLetter(QWidget):
    def __init__(self):
        super().__init__()
        self.tm = TemplateManager()
        self.input_path = ""
        self.selected_sheet = None
        self.sheet_label = None
        self.init_ui()
        self.load_template_fields()

    def init_ui(self):
        self.setWindowTitle("询证函生成器V1.0--by KevinMai")
        self.resize(700, 480)

        menubar = QMenuBar(self)
        help_menu = menubar.addMenu("帮助(&H)")
        help_action = help_menu.addAction("查看帮助文档(&D)")
        help_action.triggered.connect(self.show_help)

        self.setStyleSheet("QWidget { background-color: #2b2b2b; color: #e0e0e0; }")

        layout = QVBoxLayout()
        layout.setSpacing(10)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setMenuBar(menubar)

        self.label_input = StyledLabel("请选择询证函台账文件：未选择")
        self.btn_browse = StyledButton("选择询证函台账文件")
        self.btn_browse.clicked.connect(self.browse_input)
        layout.addWidget(self.label_input)

        self.sheet_label = StyledLabel("工作表：未选择")
        layout.addWidget(self.sheet_label)
        layout.addWidget(self.btn_browse)

        separator = QLabel("─" * 50)
        separator.setStyleSheet("color: #5a5a5a; padding: 10px 0px;")
        separator.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(separator)

        layout.addWidget(StyledLabel("请编辑以下模板字段（将应用到所有询证函）", is_section=True))

        addr_layout = QHBoxLayout()
        addr_layout.addWidget(StyledLabel("回函地址："))
        self.addr_edit = StyledLineEdit()
        addr_layout.addWidget(self.addr_edit)
        layout.addLayout(addr_layout)

        contact_layout = QHBoxLayout()
        contact_layout.addWidget(StyledLabel("联系人："))
        self.contact_edit = StyledLineEdit()
        contact_layout.addWidget(self.contact_edit)
        layout.addLayout(contact_layout)

        h1 = QHBoxLayout()
        h1.addWidget(StyledLabel("电话："))
        self.phone_edit = StyledLineEdit()
        h1.addWidget(self.phone_edit)
        h1.addWidget(StyledLabel("邮箱："))
        self.email_edit = StyledLineEdit()
        h1.addWidget(self.email_edit)
        h1.setStretchFactor(self.phone_edit, 1)
        h1.setStretchFactor(self.email_edit, 1)
        layout.addLayout(h1)

        issuer_layout = QHBoxLayout()
        issuer_layout.addWidget(StyledLabel("发函单位："))
        self.issuer_edit = StyledLineEdit()
        issuer_layout.addWidget(self.issuer_edit)
        layout.addLayout(issuer_layout)

        date_layout = QHBoxLayout()
        date_layout.addWidget(StyledLabel("发函日期："))
        self.date_edit = QDateEdit()
        self.date_edit.setDisplayFormat("yyyy.M.d")
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setStyleSheet("""
            QDateEdit {
                background-color: #252525;
                color: white;
                border: 1px solid #5a5a5a;
                border-radius: 3px;
                padding: 6px 8px;
            }
            QDateEdit:focus {
                border: 1px solid #3daee9;
            }
        """)
        date_layout.addWidget(self.date_edit)
        layout.addLayout(date_layout)

        layout.addStretch(1)

        # ========== 新增：导出按钮上方的分割线 ==========
        export_separator = QLabel("─" * 50)
        export_separator.setStyleSheet("color: #5a5a5a; padding: 10px 0px;")
        export_separator.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(export_separator)

        # ========== 新增：带图标、渐变色的导出按钮 ==========
        btn_layout = QHBoxLayout()

        # 获取图标路径 - 使用 resource_path（推荐）
        from core.utils import resource_path
        excel_icon = resource_path("assets/excel.png")
        pdf_icon = resource_path("assets/pdf.png")

        self.btn_process = IconButton(
            "导出Excel",
            excel_icon,
            gradient_color_start="#4caf50",   # 绿色起始
            gradient_color_end="#2e7d32"      # 绿色结束
        )
        self.btn_process.clicked.connect(self.process)

        self.btn_process_pdf = IconButton(
            "导出PDF",
            pdf_icon,
            gradient_color_start="#e91e63",   # 粉色起始
            gradient_color_end="#ad1457"      # 粉色结束
        )
        self.btn_process_pdf.clicked.connect(self.process_pdf)

        btn_layout.addWidget(self.btn_process)
        btn_layout.addWidget(self.btn_process_pdf)
        layout.addLayout(btn_layout)

        self.setLayout(layout)

    def show_help(self):
        help_text = """
        <h2>询证函生成器 V1.0 — 使用帮助</h2>
        <p><b>作者：</b>KevinMai</p>
        <p>本程序用于从 Excel 台账文件批量生成询证函。</p>
        <h3>📌 使用步骤</h3>
        <ol>
          <li>选择台账文件（需含：工作表名称、编号、函证单位等列）</li>
          <li>编辑模板字段</li>
          <li>点击“导出Excel”或“导出PDF”</li>
        </ol>
        """
        help_dialog = QMessageBox(self)
        help_dialog.setWindowTitle("帮助文档")
        help_dialog.setText(help_text)
        help_dialog.setTextFormat(Qt.TextFormat.RichText)
        help_dialog.setStandardButtons(QMessageBox.StandardButton.Ok)
        help_dialog.exec()

    def load_template_fields(self):
        fields = self.tm.load_fields()
        self.addr_edit.setText(fields['address'])
        self.contact_edit.setText(fields['contact'])
        self.phone_edit.setText(fields['phone'])
        self.email_edit.setText(fields['email'])
        self.issuer_edit.setText(fields['issuer'])
        try:
            y, m, d = map(int, fields['date'].split('.'))
            self.date_edit.setDate(QDate(y, m, d))
        except:
            self.date_edit.setDate(QDate.currentDate())

        self.addr_edit.setPlaceholderText("例如：四川省成都市金牛区...")
        self.contact_edit.setPlaceholderText("例如：李四")
        self.phone_edit.setPlaceholderText("例如：13588888888")
        self.email_edit.setPlaceholderText("例如：999999999@QQ.COM")
        self.issuer_edit.setPlaceholderText("例如：XXXXX有限责任公司")

    def browse_input(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "选择原始 Excel 文件", "", "Excel Files (*.xlsx)")
        if not file_path:
            return

        try:
            excel_file = pd.ExcelFile(file_path)
            sheet_names = excel_file.sheet_names
            if not sheet_names:
                raise ValueError("Excel 文件中没有工作表！")
            elif len(sheet_names) == 1:
                selected_sheet = sheet_names[0]
            else:
                selected_sheet, ok = QInputDialog.getItem(self, "选择工作表", "请选择要处理的工作表：", sheet_names, 0, False)
                if not ok:
                    return

            self.input_path = file_path
            self.selected_sheet = selected_sheet
            self.label_input.setText(f"已选择：{os.path.basename(file_path)}")
            self.label_input.setStyleSheet("color: #7ecb7e; font-weight: bold;")
            self.sheet_label.setText(f"工作表：{selected_sheet}")
            self.sheet_label.setStyleSheet("color: #7ecb7e; font-weight: bold;")

        except Exception as e:
            QMessageBox.critical(self, "错误", f"无法读取 Excel 文件：\n{str(e)}")

    def _prepare_data(self):
        df = pd.read_excel(self.input_path, sheet_name=self.selected_sheet, dtype=str)
        required_cols = [
            "工作表名称", "编号", "函证单位", "工程项目",
            "应收帐款（已开票末付款）", "长期应收款（质量保金）", "合计"
        ]
        missing = [col for col in required_cols if col not in df.columns]
        if missing:
            raise ValueError(f"缺少列：{', '.join(missing)}")

        df = df.fillna("")
        date = self.date_edit.date().toString("yyyy.M.d")
        season = get_season_from_date(date)

        data_list = []
        for _, row in df.iterrows():
            data_list.append({
                'sheet_name': str(row["工作表名称"]),
                'number': row["编号"],
                'unit': row["函证单位"],
                'project': row["工程项目"] or "",
                'receivable': row["应收帐款（已开票末付款）"] or "0.00",
                'long_term': row["长期应收款（质量保金）"] or "0.00",
                'total': row["合计"] or "0.00",
                'address': self.addr_edit.text().strip(),
                'contact': self.contact_edit.text().strip(),
                'phone': self.phone_edit.text().strip(),
                'email': self.email_edit.text().strip(),
                'issuer': self.issuer_edit.text().strip(),
                'date': date,
                'season': season
            })
        return data_list

    def process(self):
        if not self.input_path or not self.selected_sheet:
            QMessageBox.warning(self, "错误", "请先选择原始 Excel 文件及工作表！")
            return

        try:
            data_list = self._prepare_data()
            if not all([d['address'], d['contact'], d['phone'], d['email'], d['issuer'], d['date']] for d in data_list[:1]):
                QMessageBox.warning(self, "警告", "请填写所有模板字段！")
                return

            output_path, _ = QFileDialog.getSaveFileName(self, "保存询证函文件", "询证函.xlsx", "Excel Files (*.xlsx)")
            if not output_path:
                return
            if not output_path.endswith('.xlsx'):
                output_path += '.xlsx'

            user_template = get_user_template_path()
            generate_excel(data_list, user_template, output_path)

            # 保存模板
            fields = {
                'address': self.addr_edit.text().strip(),
                'contact': self.contact_edit.text().strip(),
                'phone': self.phone_edit.text().strip(),
                'email': self.email_edit.text().strip(),
                'issuer': self.issuer_edit.text().strip(),
                'date': self.date_edit.date().toString("yyyy.M.d")
            }
            self.tm.save_fields(fields)

            QMessageBox.information(self, "成功", f"询证函已生成：\n{output_path}")

        except Exception as e:
            QMessageBox.critical(self, "错误", f"处理失败：\n{str(e)}")

    def process_pdf(self):
        if not self.input_path or not self.selected_sheet:
            QMessageBox.warning(self, "错误", "请先选择原始 Excel 文件及工作表！")
            return

        try:
            data_list = self._prepare_data()
            if not all([d['address'], d['contact'], d['phone'], d['email'], d['issuer'], d['date']] for d in data_list[:1]):
                QMessageBox.warning(self, "警告", "请填写所有模板字段！")
                return

            base_dir = QFileDialog.getExistingDirectory(self, "选择PDF保存文件夹", "")
            if not base_dir:
                return

            pdf_dir = os.path.join(base_dir, "pdf")
            generate_pdfs(data_list, pdf_dir)

            # 保存模板
            fields = {
                'address': self.addr_edit.text().strip(),
                'contact': self.contact_edit.text().strip(),
                'phone': self.phone_edit.text().strip(),
                'email': self.email_edit.text().strip(),
                'issuer': self.issuer_edit.text().strip(),
                'date': self.date_edit.date().toString("yyyy.M.d")
            }
            self.tm.save_fields(fields)

            QMessageBox.information(self, "成功", f"PDF询证函已生成：\n{pdf_dir}")

        except Exception as e:
            QMessageBox.critical(self, "错误", f"PDF生成失败：\n{str(e)}")