import xlwings as xw
import pandas as pd
import shutil
import io
from datetime import datetime

# ─────────────────────────────────────────────────────────────────────────────
# Module-level state — shared across all Robot Framework tasks in this session
# ─────────────────────────────────────────────────────────────────────────────
_app = None
_wb = None
_output_file = None


def _read_source_file(source_file, skiprows=1):
    """
    Core data ingestion engine equipped with an auto-detection fallback architecture.
    Sequentially attempts pyxlsb, openpyxl, xlrd, and finally a custom low-memory 
    HTML string parser to handle masked-HTML ERP extracts without Pandas MemoryErrors.
    
    Args:
        source_file (str): The absolute path to the raw input data.
        skiprows (int): Number of top non-data rows to skip before the header.
        
    Returns:
        pd.DataFrame: A fully extracted dataframe ready for Excel insertion.
    """
    ext = source_file.lower().rsplit('.', 1)[-1]
    print(f"  Reading: {source_file}")

    if ext == 'xlsb':
        try:
            df = pd.read_excel(source_file, sheet_name=0, header=None, skiprows=skiprows, engine='pyxlsb')
            print(f"  Format: .xlsb (pyxlsb). Rows: {len(df)}")
            return df
        except Exception as e:
            print(f"  pyxlsb failed ({e}) — trying HTML fallback...")

    elif ext in ('xlsx', 'xlsm'):
        try:
            df = pd.read_excel(source_file, sheet_name=0, header=None, skiprows=skiprows, engine='openpyxl')
            print(f"  Format: .xlsx (openpyxl). Rows: {len(df)}")
            return df
        except Exception as e:
            print(f"  openpyxl failed ({e}) — trying HTML fallback...")

    elif ext == 'xls':
        try:
            xl = pd.ExcelFile(source_file, engine='xlrd')
            sheet = xl.sheet_names[0]
            df = pd.read_excel(source_file, sheet_name=sheet, header=None, skiprows=skiprows, engine='xlrd')
            print(f"  Format: real .xls (sheet: '{sheet}'). Rows: {len(df)}")
            return df
        except Exception as e:
            print(f"  xlrd failed ({e}) — trying HTML fallback...")

    # HTML / MHTML fallback — works for any extension disguised as HTML
    print("  Trying HTML/MHTML reader with encoding detection...")
    with open(source_file, 'rb') as f:
        raw = f.read()
    for enc in ['utf-8', 'latin-1', 'cp1252', 'utf-16']:
        try:
            html_str = raw.decode(enc)
            
            # Fast custom HTML table parser for huge files (1GB+)
            # Avoids pd.read_html memory crashes
            import re
            trs = html_str.split('<tr')
            if len(trs) > 1:
                data = []
                empty_top_rows = 0
                tag_cleaner = re.compile(r'<[^>]+>')
                for tr in trs[1:]:
                    if not tr.strip(): continue
                    tds = tr.split('<td')
                    if len(tds) > 1:
                        row = []
                        for td in tds[1:]:
                            end_idx = td.find('</td>')
                            cell_html = td[:end_idx] if end_idx != -1 else td
                            tag_close_idx = cell_html.find('>')
                            if tag_close_idx != -1:
                                cell_html = cell_html[tag_close_idx+1:]
                            text = tag_cleaner.sub('', cell_html).strip()
                            row.append(text)
                        if any(row):
                            data.append(row)
                        elif len(data) == 0:
                            empty_top_rows += 1
                
                if len(data) > 1:
                    # Our HTML parser automatically strips completely empty rows.
                    # Therefore, data[0] is ALWAYS the true header. We ignore 'skiprows' for HTML.
                    df = pd.DataFrame(data[1:], columns=data[0])
                    print(f"  Format: Huge HTML/MHTML (fast string split, encoding: {enc}). Rows: {len(df)} (Auto-skipped {empty_top_rows} blank top rows)")
                    return df
        except Exception as parse_e:
            print(f"  Fast HTML parse failed for {enc}: {parse_e}")
            continue

    raise ValueError(f"Cannot read file in any supported format: {source_file}")


def _write_data(ws, source_file, skiprows=1):
    """
    Internal Helper: Computes dynamic template mapping, clears obsolete data, 
    and writes fresh dataset cleanly into the respective Excel layout.
    
    Args:
        ws (xw.Sheet): The target XLWings Worksheet object.
        source_file (str): The raw export file mapped to this sheet.
        skiprows (int): How many rows to bypass during data ingestion.
        
    Returns:
        tuple: Coordinate boundaries (last_row, col1, col2, col3, col4, start_row, header_row) 
               required for accurately appending subsequent calculation formulas.
    """
    df = _read_source_file(source_file, skiprows)

    # 🛑 DYNAMIC HEADER DETECTION 🛑
    # Instead of hardcoding start_row=4, the robot dynamically scans column A of the dashboard
    # to find EXACTLY where the template has its headers placed!
    first_col_name = str(df.columns[0]).strip()
    template_column_A = ws.range("A1:A20").value
    
    header_row = 1
    start_row = 2
    for i, cell_val in enumerate(template_column_A):
        if str(cell_val).strip() == first_col_name:
            header_row = i + 1
            start_row = header_row + 1
            break
            
    print(f"  Dynamic Mapping -> Template Headers detected on Row {header_row}. Data starts on Row {start_row}.")

    print(f"  Clearing old data from row {start_row} downwards...")
    ws.range(f"{start_row}:100000").clear_contents()

    print(f"  Writing {len(df)} rows starting at A{start_row}...")
    ws.range(f"A{start_row}").options(index=False, header=False).value = df

    last_row = len(df) + start_row - 1
    # Use df column count — not used_range — to avoid doubling columns on re-runs
    last_col_num = len(df.columns)
    col1 = xw.utils.col_name(last_col_num + 1)
    col2 = xw.utils.col_name(last_col_num + 2)
    col3 = xw.utils.col_name(last_col_num + 3)
    col4 = xw.utils.col_name(last_col_num + 4)
    print(f"  Data cols: {last_col_num} → new cols at: {col1}, {col2}, {col3}, {col4}")
    return last_row, col1, col2, col3, col4, start_row, header_row


# ─────────────────────────────────────────────────────────────────────────────
# Public keywords — called by Robot Framework tasks
# Each function has its OWN formulas — change them independently
# ─────────────────────────────────────────────────────────────────────────────

def initialize_dashboard(dashboard_file):
    """
    [Robot Framework Suite Setup Task] 
    Clones the Master Dashboard template to preserve the original, generates 
    a timestamped output duplicate, and initializes a quiet XLWings COM connection.
    """
    global _app, _wb, _output_file

    timestamp = datetime.now().strftime("%d%b%Y_%H%M%S")
    dashboard_path = dashboard_file.replace('.xlsb', '')
    _output_file = f"{dashboard_path}_{timestamp}.xlsb"

    print(f"Creating output file: {_output_file}")
    shutil.copy2(dashboard_file, _output_file)

    _app = xw.App(visible=False)
    _app.screen_updating = False
    _app.display_alerts = False
    _app.enable_events = False
    _app.calculation = 'manual'

    _wb = _app.books.open(_output_file)
    print(f"Dashboard opened. Sheets: {[s.name for s in _wb.sheets]}")
    print(f"Output file: {_output_file}")
    return _output_file


def update_issuance(issuance_file, sheet_name="Issuance"):
    """
    [Robot Keyword] Executes the Issuance reconciliation phase.
    Ingests issuance data, aligns logically to the template, and inserts exactly 
    4 calculation columns. Formulas herein are isolated to Issuance workflows.
    """
    global _wb
    print(f"\n{'='*50}\n  Updating '{sheet_name}' sheet\n{'='*50}")
    ws = _wb.sheets[sheet_name]
    last_row, col1, col2, col3, col4, start_row, header_row = _write_data(ws, issuance_file, skiprows=3)

    # ── Issuance formulas — change only here ──────────────────────────────
    ws.range(f"{col1}{header_row}").value = "Year"
    ws.range(f"{col2}{header_row}").value = "Aging"
    ws.range(f"{col3}{header_row}").value = "Item cate 01"
    ws.range(f"{col4}{header_row}").value = "Item Cate 02"
    
    if last_row >= start_row:
        ws.range(f"{col1}{start_row}:{col1}{last_row}").formula = "=ROUND(((TODAY()-V2)/365),0)"
        ws.range(f"{col2}{start_row}:{col2}{last_row}").formula = '=IF(TODAY()-V2<=365,"Within 1 Year","More than 1 Year")'
        ws.range(f"{col3}{start_row}:{col3}{last_row}").formula = "=VLOOKUP($A2,'Item Category'!$A:$E,4,0)"
        ws.range(f"{col4}{start_row}:{col4}{last_row}").formula = "=VLOOKUP($A2,'Item Category'!$A:$E,5,0)"
    # ──────────────────────────────────────────────────────────────────────
    print(f"  Issuance sheet updated.")


def update_inventory(aging_file, sheet_name="Inventory"):
    """
    [Robot Keyword] Executes the Inventory reconciliation phase. 
    Applies custom Inventory-specific calculations and age-profiling columns.
    """
    global _wb
    print(f"\n{'='*50}\n  Updating '{sheet_name}' sheet\n{'='*50}")
    ws = _wb.sheets[sheet_name]
    last_row, col1, col2, col3, col4, start_row, header_row = _write_data(ws, aging_file, skiprows=1)

    # ── Inventory formulas — change only here ─────────────────────────────
    ws.range(f"{col1}{header_row}").value = "Year"
    ws.range(f"{col2}{header_row}").value = "Aging"
    ws.range(f"{col3}{header_row}").value = "Item cate 01"
    ws.range(f"{col4}{header_row}").value = "Item Cate 02"
    
    if last_row >= start_row:
        ws.range(f"{col1}{start_row}:{col1}{last_row}").formula = "=ROUND(((TODAY()-V2)/365),0)"
        ws.range(f"{col2}{start_row}:{col2}{last_row}").formula = '=IF(TODAY()-V2<=365,"Within 1 Year","More than 1 Year")'
        ws.range(f"{col3}{start_row}:{col3}{last_row}").formula = "=VLOOKUP($A2,'Item Category'!$A:$E,4,0)"
        ws.range(f"{col4}{start_row}:{col4}{last_row}").formula = "=VLOOKUP($A2,'Item Category'!$A:$E,5,0)"
    # ──────────────────────────────────────────────────────────────────────
    print(f"  Inventory sheet updated.")


def update_rnr(rnr_file, sheet_name="R&R"):
    """
    [Robot Keyword] Executes the R&R (Returns and Repairs) reconciliation phase.
    Contains isolated Business Logic formulas strictly applied to the R&R dataset.
    """
    global _wb
    print(f"\n{'='*50}\n  Updating '{sheet_name}' sheet\n{'='*50}")
    ws = _wb.sheets[sheet_name]
    last_row, col1, col2, col3, col4, start_row, header_row = _write_data(ws, rnr_file, skiprows=1)

    # ── R&R formulas — change only here ───────────────────────────────────
    ws.range(f"{col1}{header_row}").value = "Year"
    ws.range(f"{col2}{header_row}").value = "Aging"
    ws.range(f"{col3}{header_row}").value = "Item cate 01"
    ws.range(f"{col4}{header_row}").value = "Item Cate 02"
    
    if last_row >= start_row:
        # NOTE: If R&R formulas use different cell refs than V2, change them below based on start_row
        ws.range(f"{col1}{start_row}:{col1}{last_row}").formula = f"=ROUND(((TODAY()-V{start_row})/365),0)"
        ws.range(f"{col2}{start_row}:{col2}{last_row}").formula = f'=IF(TODAY()-V{start_row}<=365,"Within 1 Year","More than 1 Year")'
        ws.range(f"{col3}{start_row}:{col3}{last_row}").formula = f"=VLOOKUP($A{start_row},'Item Category'!$A:$E,4,0)"
        ws.range(f"{col4}{start_row}:{col4}{last_row}").formula = f"=VLOOKUP($A{start_row},'Item Category'!$A:$E,5,0)"
    # ──────────────────────────────────────────────────────────────────────
    print(f"  R&R sheet updated.")


def update_receipt(receipt_file, sheet_name="Receipt"):
    """
    [Robot Keyword] Executes the Receipt reconciliation phase.
    Maps daily Receipts extract against the Master Dashboard utilizing 
    isolated Receipt-specific dynamic formula columns.
    """
    global _wb
    print(f"\n{'='*50}\n  Updating '{sheet_name}' sheet\n{'='*50}")
    ws = _wb.sheets[sheet_name]
    last_row, col1, col2, col3, col4, start_row, header_row = _write_data(ws, receipt_file, skiprows=1)

    # ── Receipt formulas — change only here ───────────────────────────────
    ws.range(f"{col1}{header_row}").value = "Year"
    ws.range(f"{col2}{header_row}").value = "Aging"
    ws.range(f"{col3}{header_row}").value = "Item cate 01"
    ws.range(f"{col4}{header_row}").value = "Item Cate 02"
    
    if last_row >= start_row:
        ws.range(f"{col1}{start_row}:{col1}{last_row}").formula = "=ROUND(((TODAY()-V2)/365),0)"
        ws.range(f"{col2}{start_row}:{col2}{last_row}").formula = '=IF(TODAY()-V2<=365,"Within 1 Year","More than 1 Year")'
        ws.range(f"{col3}{start_row}:{col3}{last_row}").formula = "=VLOOKUP($A2,'Item Category'!$A:$E,4,0)"
        ws.range(f"{col4}{start_row}:{col4}{last_row}").formula = "=VLOOKUP($A2,'Item Category'!$A:$E,5,0)"
    # ──────────────────────────────────────────────────────────────────────
    print(f"  Receipt sheet updated.")


def finalize_dashboard():
    """
    [Robot Framework Suite Teardown] 
    Triggers global workbook calculations, forcefully saves the final state, 
    and safely purges all Excel background workers via `App.Quit()`.
    """
    global _app, _wb, _output_file
    try:
        print(f"\n{'='*50}")
        print("  Finalizing Dashboard")
        print(f"{'='*50}")
        print("  Calculating workbook...")
        _app.calculate()
        print("  Saving...")
        _wb.save()
        print(f"  Saved to: {_output_file}")
    except Exception as e:
        print(f"ERROR in finalize: {e}")
        raise
    finally:
        if _wb:
            _wb.close()
        if _app:
            _app.quit()
    print("\nAll 4 sheets updated successfully!")
