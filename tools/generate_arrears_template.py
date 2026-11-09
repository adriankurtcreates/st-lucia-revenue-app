"""
Generates Arrears_Stock_Template.xlsx — the fixed-schema data-entry workbook
that the arrears supervisor fills in and uploads back to the Streamlit app.

Run once locally:
    python tools/generate_arrears_template.py

Output:
    Arrears_Stock_Template.xlsx (in the current working directory)

Schema version: 1.0
Sheets: Instructions, Arrears_Stock_Flow
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

SCHEMA_VERSION = "1.0"
OUTPUT_FILE = "Arrears_Stock_Template.xlsx"

# ---------------------------------------------------------------------------
# Pre-populated arrears codes (confirmed from FY2025/26 upload).
# Input cells left blank for the supervisor to fill.
# ---------------------------------------------------------------------------
CONFIRMED_CODES = [
    # (Code, Description, Underlying Current Code)
    ("3111 002", "Arrears on Personal Income Tax", "3111 001"),
    ("3112 002", "Arrears on Corporate Income Tax", "3112 001"),
    ("3113 002", "Arrears on Withholding Income Tax", "3113 001"),
]

# Blank formula-bearing rows for the remaining arrears heads.
# The supervisor fills Code + Description + inputs; formulas are pre-wired.
GUIDE_CODES = [
    # (suggested Code, suggested Description, underlying current code)
    ("3131 002", "Arrears on Property Tax", "3131 001"),
    ("3141 002", "Arrears on VAT", "3141 001"),
    ("3141 004", "Arrears on Hotel Accommodation Tax", "3141 004"),
    ("3141 005", "Arrears on Domestic Sales Tax on Travel", "3141 005"),
    ("3141 006", "Arrears on Health and Citizen Security Levy", "3141 006"),
    ("3142 005", "Arrears on Insurance Premium Tax", "3142 005"),
    ("3142 006", "Arrears on Stamp Duty - Inland Revenue", "3142 006"),
    ("3143 010", "Arrears on Telecommunication Class Licence", "3143 010"),
    ("3151 009", "Arrears on Passenger Facility Fee", "3151 009"),
    ("", "", ""),  # spare row
]

# ---------------------------------------------------------------------------
# Styling
# ---------------------------------------------------------------------------
NAVY_FILL = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
INPUT_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")   # yellow
FORMULA_FILL = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid") # grey
GUIDE_FILL = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")   # light grey
HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
TITLE_FONT = Font(name="Calibri", size=14, bold=True, color="1B365D")
SUBTITLE_FONT = Font(name="Calibri", size=10, italic=True, color="555555")
BODY_FONT = Font(name="Calibri", size=11)
NOTE_FONT = Font(name="Calibri", size=10, italic=True, color="404040")

THIN = Side(border_style="thin", color="D0D3D4")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _style_header_row(ws, row, n_cols):
    for c in range(1, n_cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = HEADER_FONT
        cell.fill = NAVY_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
    ws.row_dimensions[row].height = 42


def build_instructions(wb):
    ws = wb.create_sheet("Instructions")
    ws.sheet_view.showGridLines = False

    ws["A1"] = "ARR EARS STOCK DATA-ENTRY TEMPLATE"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = "Saint Lucia Inland Revenue Department — Medium-Term Revenue Framework"
    ws["A2"].font = SUBTITLE_FONT

    ws["A4"] = "Schema Version:"
    ws["A4"].font = BODY_FONT
    ws["B4"] = SCHEMA_VERSION
    ws["B4"].font = Font(name="Calibri", size=11, bold=True)
    ws["C4"] = "← Do not change. The app checks this before accepting an upload."
    ws["C4"].font = NOTE_FONT

    notes = [
        ("How to use this workbook", True),
        ("1. Go to the 'Arrears_Stock_Flow' sheet.", False),
        ("2. Enter values ONLY in the yellow cells. Grey cells are formulas and are ignored by the app.", False),
        ("3. Fill in a row for every arrears code your register reports. Pre-populated rows show the three confirmed codes; blank rows are for the remaining arrears heads.", False),
        ("4. Leave 'Opening Stock' blank if unknown — the line will be flagged Provisional in the app output.", False),
        ("5. Save this workbook as .xlsx and upload it via the app's 'Arrears & Amnesty Workbook' sidebar uploader.", False),
        ("", False),
        ("What the app does with this file", True),
        ("The app reads ONLY the yellow input cells (columns A–H) and recomputes the arrears tank in Python.", False),
        ("The grey formula cells are a visual preview for you; the app does NOT use them. If a preview disagrees with the app, the app wins.", False),
        ("", False),
        ("Column reference", True),
        ("Code — Tax head code (matches the current-year upload).", False),
        ("Description — Tax head name.", False),
        ("Opening Stock — Arrears outstanding at the start of FY2025/26 (from your register).", False),
        ("Base Collection Rate — Historical recovery rate (e.g., 0.20 = 20%).", False),
        ("Write-off Rate — Historical write-off rate (e.g., 0.02 = 2%).", False),
        ("Accrual Rate — New arrears as a fraction of current-year collections of the underlying tax head.", False),
        ("Rate Adjustment — Policy tweak to the collection rate (default 0).", False),
        ("Underlying Current Code — The current-year code whose collections drive new arrears accrual.", False),
        ("", False),
        ("Excluded heads", True),
        ("Collection Fees - Towns & Villages (3422 008) is NOT an arrears head. It is a fee charged on property tax arrears collections and flows automatically through the baseline upload. Do not add it here.", False),
        ("", False),
        ("Amnesty", True),
        ("Amnesties are not modelled in this template. Amnesty-induced payments are captured in the amnesty year's actual collections upload, not in the arrears tank. If an amnesty is announced, run the model again with that year's actuals to measure its effect.", False),
    ]

    row = 6
    for text, is_bold in notes:
        cell = ws.cell(row=row, column=1, value=text)
        cell.font = Font(name="Calibri", size=11, bold=is_bold) if is_bold else NOTE_FONT
        if text == "":
            row += 1
        else:
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
            row += 1

    ws.column_dimensions["A"].width = 110
    ws.column_dimensions["B"].width = 8
    ws.column_dimensions["C"].width = 60


def build_stock_sheet(wb):
    ws = wb.create_sheet("Arrears_Stock_Flow")
    ws.sheet_view.showGridLines = False

    # Title block
    ws["A1"] = "ARREARS STOCK FLOW — FY2025/26 ONWARD"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = ("Enter values in yellow cells only. Grey cells are formulas for your preview; "
                "the app reads only yellow cells and recomputes in Python.")
    ws["A2"].font = SUBTITLE_FONT
    ws.merge_cells("A2:N2")

    # Column layout:
    # A Code | B Description | C Opening Stock | D Base Collection | E Write-off
    # F Accrual | G Rate Adjustment | H Underlying Current Code
    # I Collections Y1 | J Write-offs Y1 | K New Arrears Y1 | L Closing Y1
    # M Closing Y2 | N Closing Y3
    headers = [
        "Code",
        "Description",
        "Opening Stock (EC$)",
        "Base Collection Rate",
        "Write-off Rate",
        "Accrual Rate",
        "Rate Adjustment",
        "Underlying Current Code",
        "Collections Y1 (preview)",
        "Write-offs Y1 (preview)",
        "New Arrears Y1 (preview)",
        "Closing Y1 (preview)",
        "Closing Y2 (preview)",
        "Closing Y3 (preview)",
    ]
    header_row = 4
    for col_idx, text in enumerate(headers, start=1):
        ws.cell(row=header_row, column=col_idx, value=text)
    _style_header_row(ws, header_row, len(headers))

    # Data rows start at row 5
    data_start = 5
    all_rows = CONFIRMED_CODES + GUIDE_CODES  # 3 confirmed + 10 guide = 13 rows

    for i, (code, desc, underlying) in enumerate(all_rows):
        r = data_start + i

        # Yellow input cells (A–H)
        for c in range(1, 9):
            cell = ws.cell(row=r, column=c)
            cell.fill = INPUT_FILL
            cell.border = BORDER
            cell.font = BODY_FONT
            cell.alignment = Alignment(horizontal="left" if c == 2 else "center",
                                       vertical="center")

        # Pre-fill known codes + descriptions + underlying links
        if code:
            ws.cell(row=r, column=1, value=code)
            ws.cell(row=r, column=2, value=desc)
            ws.cell(row=r, column=8, value=underlying)

        # Pre-fill default rates for the three confirmed codes only
        if code in ("3111 002", "3112 002", "3113 002"):
            ws.cell(row=r, column=4, value=0.20)
            ws.cell(row=r, column=5, value=0.02)
            ws.cell(row=r, column=6, value=0.22888)
        elif code:
            # Guide rows: leave rates blank for the supervisor
            pass
        # spare row: fully blank

        # Grey preview formula cells (I–N)
        for c in range(9, 15):
            cell = ws.cell(row=r, column=c)
            cell.fill = FORMULA_FILL
            cell.border = BORDER
            cell.font = Font(name="Calibri", size=11, color="808080")
            cell.alignment = Alignment(horizontal="right", vertical="center")

        # Preview formulas (these are only for visual aid; the app ignores them)
        # I: Collections Y1 = Opening * (Base + Rate Adj)
        ws.cell(row=r, column=9).value = (
            f'=IF(OR($C{r}="",$D{r}=""),"",$C{r}*($D{r}+IF($G{r}="",0,$G{r})))'
        )
        # J: Write-offs Y1 = Opening * Write-off Rate
        ws.cell(row=r, column=10).value = (
            f'=IF(OR($C{r}="",$E{r}=""),"",$C{r}*$E{r})'
        )
        # K: New Arrears Y1 — needs current-year collections; show a note instead
        #    because this value comes from the OTHER upload at runtime.
        ws.cell(row=r, column=11).value = (
            f'=IF(OR($C{r}="",$F{r}=""),"","[needs Upload 1]")'
        )
        # L: Closing Y1 (approx, without new arrears from upload 1)
        ws.cell(row=r, column=12).value = (
            f'=IF($C{r}="","",MAX(0,$C{r}-IF(ISNUMBER($I{r}),$I{r},0)'
            f'-IF(ISNUMBER($J{r}),$J{r},0)))'
        )
        # M, N: Closing Y2/Y3 — rolling approximation, again without Upload 1 accruals
        ws.cell(row=r, column=13).value = (
            f'=IF($L{r}="","",MAX(0,$L{r}-$L{r}*IF($D{r}="",0,$D{r})'
            f'-$L{r}*IF($E{r}="",0,$E{r})))'
        )
        ws.cell(row=r, column=14).value = (
            f'=IF($M{r}="","",MAX(0,$M{r}-$M{r}*IF($D{r}="",0,$D{r})'
            f'-$M{r}*IF($E{r}="",0,$E{r})))'
        )

        # Percentage formats on rate columns
        for c in (4, 5, 6, 7):
            ws.cell(row=r, column=c).number_format = "0.00%"
        # Currency formats on $ columns
        for c in (3, 9, 10, 11, 12, 13, 14):
            ws.cell(row=r, column=c).number_format = '"EC$"#,##0'

    # Column widths
    widths = [12, 42, 18, 14, 14, 14, 14, 18, 18, 18, 20, 18, 18, 18]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.row_dimensions[header_row].height = 42

    # Freeze panes below header so supervisor can scroll with Code column visible
    ws.freeze_panes = "C5"

    # Add a note cell below the data
    note_row = data_start + len(all_rows) + 2
    ws.cell(row=note_row, column=1,
            value=("Note: 'Closing Y1/Y2/Y3 (preview)' columns are simplified — they omit "
                   "the new-arrears accrual term, which depends on the current-year upload. "
                   "The app computes the full stock-flow in Python using both uploads."))
    ws.cell(row=note_row, column=1).font = NOTE_FONT
    ws.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=14)


def main():
    wb = Workbook()
    # Remove default sheet
    wb.remove(wb.active)

    build_instructions(wb)
    build_stock_sheet(wb)

    wb.save(OUTPUT_FILE)
    print(f"Wrote {OUTPUT_FILE} (schema v{SCHEMA_VERSION})")
    print("Sheets:", wb.sheetnames)


if __name__ == "__main__":
    main()
