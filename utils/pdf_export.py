"""
PDF导出模块
"""
import io
import os
from datetime import datetime

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT


def _find_chinese_font():
    """查找系统中可用的中文字体"""
    paths = [
        "C:/Windows/Fonts/simhei.ttf",
        "C:/Windows/Fonts/msyh.ttc",
        "C:/Windows/Fonts/msyhbd.ttc",
        "C:/Windows/Fonts/simsun.ttc",
        "C:/Windows/Fonts/simkai.ttf",
    ]
    for p in paths:
        if os.path.exists(p):
            return p
    return None


def export_pdf(results: list, summary: dict, period: str, anomalies: list = None) -> bytes:
    """导出抵消分录PDF报告"""
    output = io.BytesIO()
    doc = SimpleDocTemplate(output, pagesize=A4, topMargin=20 * mm, bottomMargin=20 * mm)

    # 注册中文字体
    font_name = "Helvetica"
    font_path = _find_chinese_font()
    if font_path:
        try:
            pdfmetrics.registerFont(TTFont("CJK", font_path))
            font_name = "CJK"
        except Exception:
            pass

    # 创建带中文字体的样式
    styles = getSampleStyleSheet()
    for key in ["Normal", "Title", "Heading1", "Heading2", "Heading3"]:
        if key in styles:
            styles[key].fontName = font_name

    elements = []

    # 标题
    elements.append(Paragraph("内部存货交易未实现利润抵消报告", styles["Title"]))
    elements.append(Paragraph(
        f"会计期间：{period} | 生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M')}",
        styles["Normal"]
    ))
    elements.append(Spacer(1, 10 * mm))

    # 汇总表
    summary_data = [
        ["指标", "数值"],
        ["内部交易总笔数", str(summary["total_trades"])],
        ["需抵消笔数", str(summary["eliminated"])],
        ["无需抵消", str(summary["skipped"])],
        ["未识别/异常", str(summary.get("anomaly_count", 0))],
        ["抵消营业收入合计", f"{summary['total_revenue_offset']:,.2f}"],
        ["冲减存货合计", f"{summary['total_inventory_offset']:,.2f}"],
        ["确认递延所得税资产", f"{summary['total_dta']:,.2f}"],
        ["合并净利润净影响", f"{summary['total_net_profit_impact']:,.2f}"],
    ]
    t = Table(summary_data)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1a2333")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, -1), font_name),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("ALIGN", (1, 1), (1, -1), "RIGHT"),
    ]))
    elements.append(t)
    elements.append(Spacer(1, 10 * mm))

    # 异常提示
    if anomalies:
        elements.append(Paragraph("<b>异常/警告项：</b>", styles["Normal"]))
        for a in anomalies:
            elements.append(Paragraph(f"  * {a}", styles["Normal"]))
        elements.append(Spacer(1, 5 * mm))

    # 分录明细
    for r in results:
        # 简化的分录标题
        label = r.get("trade_label", "")
        elements.append(Paragraph(
            f"<b>{r['elim_no']}</b> | {label} | {r['seller']}->{r['buyer']} | "
            f"批次{r['batch_no']} | 未实现利润：{r['unrealized_profit']:,.2f}",
            styles["Normal"]
        ))
        for e in r["entries"]:
            if e["dr"]:
                elements.append(Paragraph(
                    f"    借：{e['dr']}    {e['amount']:,.2f}",
                    styles["Normal"]
                ))
            if e["cr"]:
                elements.append(Paragraph(
                    f"        贷：{e['cr']}    {e['amount']:,.2f}",
                    styles["Normal"]
                ))
        elements.append(Spacer(1, 3 * mm))

    doc.build(elements)
    return output.getvalue()
