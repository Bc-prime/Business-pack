import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
from datetime import datetime

# ============================================
# CONFIGURACIÓN — El cliente solo edita esto
# ============================================
ARCHIVO = "productos.xlsx"
NOMBRE_NEGOCIO = "Mi Negocio"
EMAIL_NEGOCIO = "contacto@minegocio.com"
TEL_NEGOCIO = "+54 11 1234-5678"
VALIDEZ_DIAS = 15
# ============================================

# 1. Leer productos seleccionados
df = pd.read_excel(ARCHIVO)
print("✅ Products loaded successfully")

# 2. Calcular totales
df["Subtotal"] = df["Quantity"] * df["Unit Price"]
subtotal = df["Subtotal"].sum()
iva = (subtotal * 0.21).round(2)
total = (subtotal + iva).round(2)
fecha = datetime.now().strftime("%B %d, %Y")
numero_presupuesto = datetime.now().strftime("%Y%m%d%H%M")

# 3. Resumen terminal
print("\n" + "="*50)
print(f"     📄 QUOTE — {NOMBRE_NEGOCIO}")
print("="*50)
print(f"\n  Quote #: {numero_presupuesto}")
print(f"  Date: {fecha}")
print(f"\n  Products:")
for _, row in df.iterrows():
    print(f"    {row['Product']} x{row['Quantity']} — ${row['Subtotal']:,.2f}")
print(f"\n  Subtotal: ${subtotal:,.2f}")
print(f"  IVA 21%:  ${iva:,.2f}")
print(f"  TOTAL:    ${total:,.2f}")
print("="*50)

# 4. Generar PDF
doc = SimpleDocTemplate(
    f"quote_{numero_presupuesto}.pdf",
    pagesize=A4,
    rightMargin=15*mm,
    leftMargin=15*mm,
    topMargin=12*mm,
    bottomMargin=12*mm
)

# Colores
DARK = colors.HexColor("#1E293B")
ACCENT = colors.HexColor("#3B82F6")
LIGHT = colors.HexColor("#F1F5F9")
WHITE = colors.white
GREEN = colors.HexColor("#16A34A")

styles = getSampleStyleSheet()
story = []

# Estilos
title_style = ParagraphStyle("title", fontSize=22, textColor=WHITE, fontName="Helvetica-Bold", alignment=TA_LEFT, leading=26)
subtitle_style = ParagraphStyle("subtitle", fontSize=10, textColor=colors.HexColor("#93C5FD"), fontName="Helvetica", alignment=TA_LEFT)
body_style = ParagraphStyle("body", fontSize=9, textColor=DARK, fontName="Helvetica", alignment=TA_LEFT, leading=14)
bold_style = ParagraphStyle("bold", fontSize=9, textColor=DARK, fontName="Helvetica-Bold", alignment=TA_LEFT)
right_style = ParagraphStyle("right", fontSize=9, textColor=DARK, fontName="Helvetica", alignment=TA_RIGHT)
total_style = ParagraphStyle("total", fontSize=12, textColor=WHITE, fontName="Helvetica-Bold", alignment=TA_RIGHT)

# Header
header_data = [[
    Table([
        [Paragraph(NOMBRE_NEGOCIO, title_style)],
        [Paragraph("Professional Quote", subtitle_style)],
        [Spacer(1, 8)],
        [Paragraph(EMAIL_NEGOCIO, subtitle_style)],
        [Paragraph(TEL_NEGOCIO, subtitle_style)],
    ], colWidths=[100*mm]),
    Table([
        [Paragraph("QUOTE", ParagraphStyle("q", fontSize=28, textColor=colors.HexColor("#3B82F6"), fontName="Helvetica-Bold", alignment=TA_RIGHT))],
        [Paragraph(f"# {numero_presupuesto}", ParagraphStyle("qn", fontSize=9, textColor=colors.HexColor("#93C5FD"), fontName="Helvetica", alignment=TA_RIGHT))],
        [Spacer(1, 8)],
        [Paragraph(f"Date: {fecha}", ParagraphStyle("qd", fontSize=9, textColor=WHITE, fontName="Helvetica", alignment=TA_RIGHT))],
        [Paragraph(f"Valid for: {VALIDEZ_DIAS} days", ParagraphStyle("qv", fontSize=9, textColor=WHITE, fontName="Helvetica", alignment=TA_RIGHT))],
    ], colWidths=[70*mm]),
]]

header_table = Table(header_data, colWidths=[100*mm, 75*mm])
header_table.setStyle(TableStyle([
    ("BACKGROUND", (0, 0), (-1, -1), DARK),
    ("PADDING", (0, 0), (-1, -1), 14),
    ("VALIGN", (0, 0), (-1, -1), "TOP"),
]))
story.append(header_table)
story.append(Spacer(1, 10))

# Tabla de productos
product_headers = [["#", "Product", "Description", "Qty", "Unit Price", "Subtotal"]]
product_data = []
for i, (_, row) in enumerate(df.iterrows(), 1):
    product_data.append([
        str(i),
        str(row["Product"]),
        str(row.get("Description", "")),
        str(int(row["Quantity"])),
        f"${row['Unit Price']:,.2f}",
        f"${row['Subtotal']:,.2f}",
    ])

table_data = product_headers + product_data
product_table = Table(table_data, colWidths=[10*mm, 35*mm, 55*mm, 12*mm, 22*mm, 22*mm])
product_table.setStyle(TableStyle([
    ("BACKGROUND", (0, 0), (-1, 0), ACCENT),
    ("TEXTCOLOR", (0, 0), (-1, 0), WHITE),
    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
    ("FONTSIZE", (0, 0), (-1, -1), 9),
    ("ALIGN", (0, 0), (-1, -1), "CENTER"),
    ("ALIGN", (1, 1), (2, -1), "LEFT"),
    ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.HexColor("#F8FAFC"), WHITE]),
    ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#E2E8F0")),
    ("PADDING", (0, 0), (-1, -1), 6),
    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
]))
story.append(product_table)
story.append(Spacer(1, 8))

# Totales
totals_data = [
    ["", "", "", "", "Subtotal:", f"${subtotal:,.2f}"],
    ["", "", "", "", "IVA 21%:", f"${iva:,.2f}"],
]
totals_table = Table(totals_data, colWidths=[10*mm, 35*mm, 55*mm, 12*mm, 22*mm, 22*mm])
totals_table.setStyle(TableStyle([
    ("ALIGN", (4, 0), (-1, -1), "RIGHT"),
    ("FONTSIZE", (0, 0), (-1, -1), 9),
    ("PADDING", (0, 0), (-1, -1), 4),
]))
story.append(totals_table)

# Total final
total_data = [["", "", "", "", "TOTAL:", f"${total:,.2f}"]]
total_table = Table(total_data, colWidths=[10*mm, 35*mm, 55*mm, 12*mm, 22*mm, 22*mm])
total_table.setStyle(TableStyle([
    ("BACKGROUND", (4, 0), (-1, -1), DARK),
    ("TEXTCOLOR", (4, 0), (-1, -1), WHITE),
    ("FONTNAME", (4, 0), (-1, -1), "Helvetica-Bold"),
    ("FONTSIZE", (4, 0), (-1, -1), 11),
    ("ALIGN", (4, 0), (-1, -1), "RIGHT"),
    ("PADDING", (0, 0), (-1, -1), 6),
]))
story.append(total_table)
story.append(Spacer(1, 15))

# Nota al pie
story.append(HRFlowable(width="100%", thickness=1, color=ACCENT))
story.append(Spacer(1, 6))
story.append(Paragraph(
    f"Thank you for your interest in {NOMBRE_NEGOCIO}. This quote is valid for {VALIDEZ_DIAS} days from the date of issue. "
    f"For any questions please contact us at {EMAIL_NEGOCIO} or {TEL_NEGOCIO}.",
    ParagraphStyle("footer", fontSize=8, textColor=colors.HexColor("#64748B"), fontName="Helvetica", alignment=TA_CENTER)
))

doc.build(story)
print(f"\n📄 Quote saved as quote_{numero_presupuesto}.pdf")