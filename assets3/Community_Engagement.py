"""
esg_table_S5.py
-----------------------------------
Reusable ESG Table Component (S5: Community Engagement & Social Investment)
Author: Chi Hsing Wu / ChatGPT
Version: 1.0
-----------------------------------
Purpose:
    Generate a 14 cm width ReportLab Table (for A4 landscape ESG reports)
    with pre-filled values and word-wrapping.
Usage:
    from esg_table_S5 import ESGTableS5

    s5 = ESGTableS5()
    table = s5.create_table()
    doc.build([table])
"""

from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle, Paragraph
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm


class ESGTableS5:
    def __init__(self):
        self.cell_style = ParagraphStyle(
            name="ESGCell",
            fontName="Helvetica",
            fontSize=8.5,
            leading=10,
            alignment=0,  # left
        )

    def create_table(self):
        """Return a ready-to-use ReportLab Table for S5 Community Engagement."""

        data = [
            ["Program / Area", "Type of Initiative", "KPI or Outcome",
             "2024 Result", "2025 Target", "Remarks"],
            ["STEM Education",
             Paragraph("School partnership & internship program", self.cell_style),
             Paragraph("Students reached", self.cell_style),
             "180", "250",
             Paragraph("Collaborative education initiative", self.cell_style)],
            ["Environmental Restoration",
             Paragraph("Recycling & community cleanup", self.cell_style),
             Paragraph("Waste collected (kg)", self.cell_style),
             "600", "800",
             Paragraph("ISO 26000 6.8.3 alignment", self.cell_style)],
            ["Employee Volunteering",
             Paragraph("Volunteer hours contributed", self.cell_style),
             Paragraph("Total hours", self.cell_style),
             "230", "260",
             Paragraph("Annual corporate volunteering program", self.cell_style)],
            ["Local Supplier Development",
             Paragraph("Capacity-building workshops", self.cell_style),
             Paragraph("SMEs trained", self.cell_style),
             "12", "15",
             Paragraph("Sustainable procurement program", self.cell_style)],
            ["Community Health",
             Paragraph("Health awareness workshop", self.cell_style),
             Paragraph("Participants reached", self.cell_style),
             "180", "250",
             Paragraph("In partnership with local clinic / NGO", self.cell_style)]
        ]

        col_widths = [2.6 * cm, 3.8 * cm, 2.4 * cm, 1.2 * cm, 1.4 * cm, 2.6 * cm]

        table = Table(data, colWidths=col_widths)
        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E0E0E0")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("ALIGN", (1, 1), (1, -1), "LEFT"),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, 0), 9.5),
            ("FONTSIZE", (0, 1), (-1, -1), 8.5),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
            ("BACKGROUND", (0, 1), (-1, -1), colors.whitesmoke),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ]))
        return table
