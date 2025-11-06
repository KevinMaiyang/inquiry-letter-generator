# pdf_generator.py
import os
from fpdf import FPDF

class InquiryPDF(FPDF):
    def __init__(self):
        super().__init__()
        font_dir = os.path.join(os.path.dirname(__file__), "fonts")
        self.add_font("AlibabaPuHuiTi-M", "", os.path.join(font_dir, "AlibabaPuHuiTi-3-65-Medium.ttf"), uni=True)
        self.add_font("AlibabaPuHuiTi-L", "", os.path.join(font_dir, "AlibabaPuHuiTi-3-45-Light.ttf"), uni=True)
    def header(self):
        self.set_font("AlibabaPuHuiTi-M", "", 16)
        self.cell(0, 10, "对 账 函", border=0, align="C")
        self.ln(12)

def format_currency(value):
    try:
        return f"{float(value):,.2f}"
    except (ValueError, TypeError):
        return "0.00"

def generate_single_pdf_content(pdf, data):
    """将单份询证函内容绘制到给定的 PDF 对象中（供单页或合并使用）"""
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)

    # 抬头单位 + 冒号
    pdf.set_font("AlibabaPuHuiTi-M", size=12)
    pdf.cell(0, 10, f"{data['unit']}：")
    pdf.ln(10)

    # 正文（首行缩进：两个全角空格）
    pdf.set_font("AlibabaPuHuiTi-L", size=12)
    text = f"\u3000\u3000我公司承担的{data['project']}项目，已完成了合同约定的相应工作，我公司核算截止到该项目{data['season']}计价，债权记录截止到{data['date']}尚有下表列示数据未收到，请贵公司核对，如与贵单位记录相符，请在本函下端“信息证明无误”处签章证明；如有不符，请在“信息不符”处列明不符金额。"
    pdf.multi_cell(0, 6, text)
    pdf.ln(5)

    # 回函信息
    pdf.set_font("AlibabaPuHuiTi-M", size=12)
    pdf.cell(0, 10, "回函请直接寄至：")
    pdf.ln(8)
    pdf.set_font("AlibabaPuHuiTi-L", size=12)
    pdf.cell(0, 10, f"回函地址：{data['address']}    联系人：{data['contact']}")
    pdf.ln(6)
    pdf.cell(0, 10, f"电话：{data['phone']}    邮箱：{data['email']}")
    pdf.ln(10)

    # 表格标题
    pdf.set_font("AlibabaPuHuiTi-M", size=12)
    pdf.cell(0, 10, "本单位与贵单位的往来账项列示如下：")
    pdf.ln(0)
    pdf.cell(0, 10, "单位：元", align="R")
    pdf.ln(10)

    # 表格
    col_widths = [60, 30, 35, 35, 30]
    headers = ["本单位账户", "截止日期", "贵单位欠", "欠贵单位", "备 注"]
    for i, header in enumerate(headers):
        pdf.cell(col_widths[i], 10, header, border=1, align="C")
    pdf.ln()

    pdf.set_font("AlibabaPuHuiTi-L", size=11)
    receivable = format_currency(data['receivable'])
    long_term = format_currency(data['long_term'])
    total = format_currency(data['total'])

    pdf.cell(col_widths[0], 10, "应收帐款（已开票末付款）", border=1)
    pdf.cell(col_widths[1], 10, data['date'], border=1, align="C")
    pdf.cell(col_widths[2], 10, receivable, border=1, align="R")
    pdf.cell(col_widths[3], 10, "", border=1)
    pdf.cell(col_widths[4], 10, "", border=1)
    pdf.ln()

    pdf.cell(col_widths[0], 10, "长期应收款（质量保金）", border=1)
    pdf.cell(col_widths[1], 10, data['date'], border=1, align="C")
    pdf.cell(col_widths[2], 10, long_term, border=1, align="R")
    pdf.cell(col_widths[3], 10, "", border=1)
    pdf.cell(col_widths[4], 10, "", border=1)
    pdf.ln()

    pdf.set_font("AlibabaPuHuiTi-M", size=12)
    pdf.cell(col_widths[0], 10, "合计", border=1)
    pdf.cell(col_widths[1], 10, "", border=1)
    pdf.cell(col_widths[2], 10, total, border=1, align="R")
    pdf.cell(col_widths[3], 10, "", border=1)
    pdf.cell(col_widths[4], 10, "", border=1)
    pdf.ln(30)

    # 落款（靠右）
    pdf.cell(0, 10, data['issuer'], align="R")
    pdf.ln(8)
    pdf.cell(0, 10, data['date'], align="R")
    pdf.ln(15)

    # ========== 结论部分（保持你现有逻辑）==========
    pdf.set_font("AlibabaPuHuiTi-L", size=12)
    left_width = 95
    right_width = 95
    total_width = left_width + right_width
    left_lines = [
        "1.信息证明无误",
        "(盖章)          ",
        "年       月       日",
        "经办人：          "
    ]
    right_lines = [
        "2.信息不符，请列明不符的详细情况",
        "(盖章)"          ,
        "年       月       日",
        "经办人：          "
    ]
    line_height = 10
    box_height = len(left_lines) * line_height
    pdf.cell(left_width, box_height, border=1, ln=False)
    pdf.cell(right_width, box_height, border=1, ln=True)
    pdf.set_xy(pdf.l_margin, pdf.get_y() - box_height)
    for i, line in enumerate(left_lines):
        pdf.set_x(pdf.l_margin)
        if i == 0:
            pdf.cell(left_width, line_height, line, border=0, ln=False)
        else:
            pdf.cell(left_width, line_height, line, border=0, align="R", ln=False)
        pdf.set_x(pdf.l_margin + left_width)
        pdf.cell(right_width, line_height, right_lines[i], border=0, align="R" if i > 0 else "L", ln=True)

    remark_text = f"3.备注：（如果截止{data['date']}日后情况有变化，请在此列明最新情况）"
    remark_lines = [
        remark_text,
        "(盖章)"          ,
        "年       月       日",
        "经办人：          "
    ]
    remark_box_height = len(remark_lines) * line_height
    pdf.cell(total_width, remark_box_height, border=1, ln=True)
    pdf.set_xy(pdf.l_margin, pdf.get_y() - remark_box_height)
    for i, line in enumerate(remark_lines):
        if i == 0:
            pdf.cell(total_width, line_height, line, border=0, ln=True)
        else:
            pdf.cell(total_width, line_height, line, border=0, align="R", ln=True)

def generate_pdfs(data_list, output_dir):
    os.makedirs(output_dir, exist_ok=True)

    # 1. 生成单个 PDF
    for data in data_list:
        safe_name = "".join(c for c in data['sheet_name'] if c.isalnum() or c in (' ', '-', '_')).rstrip()
        if not safe_name:
            safe_name = "询证函"
        output_path = os.path.join(output_dir, f"{safe_name}.pdf")
        pdf = InquiryPDF()
        generate_single_pdf_content(pdf, data)
        pdf.output(output_path)

    # 2. 生成合并 PDF
    merged_pdf = InquiryPDF()
    for data in data_list:
        generate_single_pdf_content(merged_pdf, data)
    merged_output = os.path.join(output_dir, "询证函-合并.pdf")
    merged_pdf.output(merged_output)