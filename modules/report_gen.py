import os
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

def generate_excel_report(summary_df, scenario_name, output_path="IRD_Revenue_Projections_Report.xlsx"):
    """
    Generates a clean single-worksheet executive report in Excel.
    Safely handles numbers and strings to completely eliminate TypeErrors.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Revenue Projections Summary"
    ws.views.sheetView.showGridLines = True

    # Palette Styles (Classic Navy Theme)
    navy_fill = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
    subtotal_fill = PatternFill(start_color="F4F6F9", end_color="F4F6F9", fill_type="solid")
    accent_fill = PatternFill(start_color="EAECEE", end_color="EAECEE", fill_type="solid")
    
    font_title = Font(name="Calibri", size=16, bold=True, color="1B365D")
    font_subtitle = Font(name="Calibri", size=11, italic=True, color="555555")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_data = Font(name="Calibri", size=11, bold=False)
    font_subtotal = Font(name="Calibri", size=11, bold=True)
    font_grand_total = Font(name="Calibri", size=12, bold=True, color="1B365D")

    # Borders
    thin_side = Side(border_style="thin", color="D0D3D4")
    data_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    accounting_border = Border(top=Side(border_style="thin", color="000000"), bottom=Side(border_style="double", color="000000"))

    # Titles Block
    ws["A1"] = "INLAND REVENUE DEPARTMENT — SAINT LUCIA"
    ws["A1"].font = font_title
    ws["A2"] = f"Medium-Term Revenue Framework Projections — Scenario: {scenario_name.upper()}"
    ws["A2"].font = font_subtitle
    ws["A3"] = "Values expressed in Eastern Caribbean Dollars (EC$)"
    ws["A3"].font = font_subtitle

    # Re-map Columns dynamically to prevent structural indexing breaks
    headers = list(summary_df.columns)
    header_row = 5
    
    for col_idx, text in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=col_idx, value=str(text))
        cell.font = font_header
        cell.fill = navy_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[header_row].height = 26

    # Safe Row Processing Data Injection Loop
    current_row = 6
    for _, row in summary_df.iterrows():
        row_cells = []
        is_grand_total = "GRAND TOTAL" in str(row.get("Tax Head", "")).upper()
        
        for col_idx, col_name in enumerate(headers, start=1):
            val = row[col_name]
            cell = ws.cell(row=current_row, column=col_idx)
            
            # Safe Value Casting Block
            if col_idx == 1: # Code column
                cell.value = str(val) if pd.notna(val) else ""
                cell.alignment = Alignment(horizontal="center")
            elif col_idx == 2: # Tax Head Description column
                cell.value = str(val) if pd.notna(val) else ""
                cell.alignment = Alignment(horizontal="left")
            elif "GROWTH" in str(col_name).upper() or "CAGR" in str(col_name).upper():
                try:
                    cell.value = float(val) if pd.notna(val) else 0.0
                except:
                    cell.value = 0.0
                cell.number_format = "0.00%"
                cell.alignment = Alignment(horizontal="right")
            else: # Financial fields
                try:
                    cell.value = float(val) if pd.notna(val) else 0.0
                except:
                    cell.value = str(val)
                cell.number_format = "$#,##0;($#,##0);\"-\""
                cell.alignment = Alignment(horizontal="right")

            # Apply fonts and grid structural lines safely
            if is_grand_total:
                cell.font = font_grand_total
                cell.fill = accent_fill
                cell.border = accounting_border
            else:
                cell.font = font_data
                cell.border = data_border
                
        current_row += 1

    # Safe dynamic column resizing framework
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col.column)
        for cell in col:
            if cell.row < 5: continue
            if cell.value is not None:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max(max_len + 4, 13)

    wb.save(output_path)
    return output_path


def generate_docx_brief(summary_stats, scenario_name, output_path="IRD_Revenue_Framework_Brief.docx"):
    """
    Generates a clean institutional memorandum using IMF reporting structures.
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

    doc.add_heading('1. Executive Briefing', level=1)
    p = doc.add_paragraph()
    p.add_run("This framework documents updated fiscal projection analytics for the Inland Revenue Department (IRD). ").bold = True
    p.add_run(f"Under current structural rules for the {scenario_name} macroeconomic vector, total aggregated metrics are expected to shift from a baseline foundation of ")
    p.add_run(f"EC${summary_stats.get('baseline_total', 0):,.0f} to settle at EC${summary_stats.get('year3_total', 0):,.0f} by the horizon end points. ")
    p.add_run("This application engine uses isolated, driver-specific variables to decouple structural growth tracking from older index mechanisms safely.")

    doc.add_heading('2. Core Transmission Mechanics', level=1)
    doc.add_paragraph("Projections are mapped across isolated nominal economic transmission streams:")
    
    bullets = [
        ("Wages Matrix: ", "Controls Personal Income Tax trajectories dynamically."),
        ("Profits Lags: ", "Dictates Corporate Tax collection windows smoothly."),
        ("Consumption Paths: ", "Anchors Value Added Tax gross-to-net tracking curves safely.")
    ]
    for pref, text in bullets:
        bp = doc.add_paragraph(style='List Bullet')
        bp.add_run(pref).bold = True
        bp.add_run(text)

    doc.save(output_path)
    return output_path
