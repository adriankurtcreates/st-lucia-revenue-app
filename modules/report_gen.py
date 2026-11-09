import io

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH


def _safe_float(val, default=0.0):
    """Coerce any value to float, never raise."""
    try:
        if pd.isna(val):
            return default
        return float(val)
    except (TypeError, ValueError):
        return default


def generate_excel_report(summary_df, scenario_name):
    """
    Generates a clean single-worksheet executive report in Excel.
    Returns a BytesIO buffer ready for st.download_button.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Revenue Projections Summary"
    # FIX: hide gridlines on a styled report (previous code set the default True, a no-op)
    # FIX: openpyxl sets ws.views.sheetView to None on freshly created sheets;
    # guard the cosmetic gridline toggle so it can't throw AttributeError.
    try:
        ws.views.sheetView.showGridLines = False
    except AttributeError:
        pass

    # Palette Styles (Classic Navy Theme)
    navy_fill = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
    accent_fill = PatternFill(start_color="EAECEE", end_color="EAECEE", fill_type="solid")

    font_title = Font(name="Calibri", size=16, bold=True, color="1B365D")
    font_subtitle = Font(name="Calibri", size=11, italic=True, color="555555")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_data = Font(name="Calibri", size=11, bold=False)
    font_grand_total = Font(name="Calibri", size=12, bold=True, color="1B365D")

    thin_side = Side(border_style="thin", color="D0D3D4")
    data_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    accounting_border = Border(
        top=Side(border_style="thin", color="000000"),
        bottom=Side(border_style="double", color="000000"),
    )

    # FIX: EC$ currency format (previous code used "$", which reads as US dollars)
    ECY_FORMAT = '"EC$"#,##0;"(EC$"#,##0")";"-"'
    PCT_FORMAT = "0.00%"

    # Titles Block
    ws["A1"] = "INLAND REVENUE DEPARTMENT — SAINT LUCIA"
    ws["A1"].font = font_title
    ws["A2"] = f"Medium-Term Revenue Framework Projections — Scenario: {scenario_name.upper()}"
    ws["A2"].font = font_subtitle
    ws["A3"] = "Values expressed in Eastern Caribbean Dollars (EC$)"
    ws["A3"].font = font_subtitle

    headers = list(summary_df.columns)
    header_row = 5

    for col_idx, text in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=col_idx, value=str(text))
        cell.font = font_header
        cell.fill = navy_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[header_row].height = 26

    current_row = 6
    for _, row in summary_df.iterrows():
        is_grand_total = "GRAND TOTAL" in str(row.get("Tax Head", "")).upper()

        for col_idx, col_name in enumerate(headers, start=1):
            val = row[col_name]
            cell = ws.cell(row=current_row, column=col_idx)

            if col_idx == 1:  # Code
                cell.value = str(val) if pd.notna(val) else ""
                cell.alignment = Alignment(horizontal="center")
            elif col_idx == 2:  # Tax Head
                cell.value = str(val) if pd.notna(val) else ""
                cell.alignment = Alignment(horizontal="left")
            elif "GROWTH" in str(col_name).upper() or "CAGR" in str(col_name).upper():
                cell.value = _safe_float(val)
                cell.number_format = PCT_FORMAT
                cell.alignment = Alignment(horizontal="right")
            else:  # Financial
                cell.value = _safe_float(val)
                cell.number_format = ECY_FORMAT
                cell.alignment = Alignment(horizontal="right")

            if is_grand_total:
                cell.font = font_grand_total
                cell.fill = accent_fill
                cell.border = accounting_border
            else:
                cell.font = font_data
                cell.border = data_border

        current_row += 1

    # Dynamic column widths (measure from header row onward)
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col.column)
        for cell in col:
            if cell.row < header_row:
                continue
            if cell.value is not None:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max(max_len + 4, 13)

    # FIX: return in-memory buffer instead of writing to disk
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


def generate_docx_brief(summary_stats, scenario_name):
    """
    Generates a clean institutional memorandum using IMF reporting structures.
    Returns a BytesIO buffer ready for st.download_button.
    """
    doc = Document()
    for sec in doc.sections:
        sec.top_margin = Inches(1.0)
        sec.bottom_margin = Inches(1.0)
        sec.left_margin = Inches(1.0)
        sec.right_margin = Inches(1.0)

    doc.styles['Normal'].font.name = 'Times New Roman'
    doc.styles['Normal'].font.size = Pt(11)

    title = doc.add_paragraph()
    run = title.add_run("SAINT LUCIA: MEDIUM-TERM REVENUE PROJECTION FRAMEWORK")
    run.font.size = Pt(13)
    run.font.bold = True
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    sub = doc.add_paragraph()
    sub.add_run(f"Technical Briefing Note for Management — Scenario Analysis: {scenario_name}").font.italic = True
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph("-" * 60)

    # FIX: safe number formatting so missing/None values can't crash the docx build
    baseline_total = _safe_float(summary_stats.get('baseline_total', 0))
    year3_total = _safe_float(summary_stats.get('year3_total', 0))

    doc.add_heading('1. Executive Briefing', level=1)
    p = doc.add_paragraph()
    p.add_run("This framework documents updated fiscal projection analytics for the Inland Revenue Department (IRD). ").bold = True
    p.add_run(f"Under current structural rules for the {scenario_name} macroeconomic vector, total aggregated metrics are expected to shift from a baseline foundation of ")
    p.add_run(f"EC${baseline_total:,.0f} to settle at EC${year3_total:,.0f} by the horizon end points. ")
    p.add_run("This application engine uses isolated, driver-specific variables to decouple structural growth tracking from older index mechanisms safely.")

    doc.add_heading('2. Core Transmission Mechanics', level=1)
    doc.add_paragraph("Projections are mapped across isolated nominal economic transmission streams:")

    bullets = [
        ("Wages Matrix: ", "Controls Personal Income Tax trajectories dynamically."),
        ("Profits Lags: ", "Dictates Corporate Tax collection windows smoothly."),
        ("Consumption Paths: ", "Anchors Value Added Tax gross-to-net tracking curves safely."),
    ]
    for pref, text in bullets:
        bp = doc.add_paragraph(style='List Bullet')
        bp.add_run(pref).bold = True
        bp.add_run(text)

    # FIX: return in-memory buffer instead of writing to disk
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer
