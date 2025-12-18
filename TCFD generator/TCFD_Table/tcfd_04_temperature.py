"""
TCFD 表格引擎 04 - 溫升風險
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pathlib import Path

OUTPUT_DIR = Path(__file__).parent.parent / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

TABLE_TITLE = "TCFD 溫升風險分析"
TYPE_NAME = "溫升風險"
RISK_TYPES = ['升溫1.5°C情境', '升溫2°C以上情境']


def create_table(csv_lines, industry="企業", filename=None, output_dir=None):
    """從 CSV 生成 TCFD PPTX"""
    if output_dir is None:
        output_dir = OUTPUT_DIR
    else:
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)
    
    prs = Presentation()
    # 16:9 寬螢幕
    # A4 橫向: 11.69" x 8.27"
    prs.slide_width = Inches(11.69)
    prs.slide_height = Inches(8.27)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # 表格尺寸與置中計算
    slide_w = 11.69  # A4 橫向
    table_w = 10.0   # 表格寬度 ≈ 25cm  # 表格寬度
    left_margin = (slide_w - table_w) / 2  # 置中
    
    # 標題
    title = slide.shapes.add_textbox(Inches(left_margin), Inches(0.2), Inches(table_w), Inches(0.5))
    p = title.text_frame.paragraphs[0]
    p.text = f"{TABLE_TITLE} - {industry}"
    p.font.size = Pt(20)
    p.font.bold = True
    p.alignment = PP_ALIGN.CENTER
    
    # 表格 (置中)
    rows = 2 + len(csv_lines)
    tbl = slide.shapes.add_table(rows, 6, Inches(left_margin), Inches(0.85), Inches(table_w), Inches(6.2)).table
    
    # 欄寬 (總寬 10.0" ≈ 25cm)
    tbl.columns[0].width = Inches(1.1)
    tbl.columns[1].width = Inches(1.4)
    tbl.columns[2].width = Inches(0.9)
    tbl.columns[3].width = Inches(1.5)
    tbl.columns[4].width = Inches(2.5)
    tbl.columns[5].width = Inches(2.6)
    
    # 表頭列高
    tbl.rows[0].height = Inches(0.56)
    tbl.rows[1].height = Inches(0.56)
    
    # Row 0: 分割表頭
    tbl.cell(0, 0).merge(tbl.cell(0, 2))
    left_cell = tbl.cell(0, 0)
    left_cell.fill.solid()
    left_cell.fill.fore_color.rgb = RGBColor(0x2F, 0x52, 0x33)
    left_cell.text = "氣候相關風險"
    left_cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
    left_cell.text_frame.paragraphs[0].font.bold = True
    left_cell.text_frame.paragraphs[0].font.size = Pt(11)
    left_cell.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    
    tbl.cell(0, 3).merge(tbl.cell(0, 5))
    right_cell = tbl.cell(0, 3)
    right_cell.fill.solid()
    right_cell.fill.fore_color.rgb = RGBColor(0x80, 0x80, 0x80)
    right_cell.text = "財務影響"
    right_cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
    right_cell.text_frame.paragraphs[0].font.bold = True
    right_cell.text_frame.paragraphs[0].font.size = Pt(11)
    right_cell.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    
    # Row 1: 欄位標題
    headers = ['類型', '氣候風險', '期間', '風險描述', '潛在影響', '因應措施']
    for c, h in enumerate(headers):
        cell = tbl.cell(1, c)
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(0x80, 0x80, 0x80)
        cell.text = h
        cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
        cell.text_frame.paragraphs[0].font.bold = True
        cell.text_frame.paragraphs[0].font.size = Pt(10)
    
    # 資料行
    for i, line in enumerate(csv_lines):
        r = i + 2
        parts = [p.strip() for p in line.split('|||')]
        
        cell = tbl.cell(r, 0)
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(0x8B, 0x9D, 0x83)
        if i == 0:
            cell.text = TYPE_NAME
            cell.text_frame.paragraphs[0].font.bold = True
            cell.text_frame.paragraphs[0].font.size = Pt(11)
        
        risk = RISK_TYPES[i] if i < len(RISK_TYPES) else f"風險{i+1}"
        tbl.cell(r, 1).text = risk
        tbl.cell(r, 1).text_frame.paragraphs[0].font.size = Pt(11)
        
        tbl.cell(r, 2).text = "長期"
        tbl.cell(r, 2).text_frame.paragraphs[0].font.size = Pt(11)
        
        if len(parts) >= 1:
            _set_bullet_text(tbl.cell(r, 3), parts[0])
        if len(parts) >= 2:
            _set_bullet_text(tbl.cell(r, 4), parts[1])
        if len(parts) >= 3:
            _set_bullet_text(tbl.cell(r, 5), parts[2])
    
    # 儲存
    if filename is None:
        from datetime import datetime
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"TCFD_04_溫升風險_{industry}_{timestamp}.pptx"
    
    filepath = output_dir / filename
    prs.save(str(filepath))
    return filepath


def _set_bullet_text(cell, text):
    tf = cell.text_frame
    tf.clear()
    points = [p.strip() for p in text.split(';') if p.strip()]
    for idx, point in enumerate(points):
        if idx == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = f"• {point}"
        p.font.size = Pt(11)
        p.alignment = PP_ALIGN.LEFT

