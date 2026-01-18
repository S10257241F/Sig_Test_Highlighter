# data_processor.py
from pathlib import Path
import pandas as pd
import openpyxl
import html
from typing import Optional, Set, Tuple, Dict
import re

# correct import for column index conversion
from openpyxl.utils.cell import column_index_from_string

# ---------- simple preview ----------
def preview_file(path: Path, nrows: int = 10):
    """
    Return small pandas.DataFrame preview for display.
    Handles Excel and CSV.
    """
    path = Path(path)
    if not path.exists():
        return None
    try:
        if path.suffix.lower() in [".xlsx", ".xls"]:
            df = pd.read_excel(path, engine="openpyxl")
        elif path.suffix.lower() == ".csv":
            df = pd.read_csv(path)
        else:
            return pd.DataFrame({"info": [f"Unsupported file type: {path.suffix}"]})
        return df.head(nrows)
    except Exception as e:
        return pd.DataFrame({"error": [str(e)]})

# ---------- styled preview using openpyxl fills ----------
def _openpyxl_color_to_hex(openpyxl_color) -> Optional[str]:
    """
    Convert openpyxl Color (fgColor.rgb or similar) to HEX string like #RRGGBB.
    Returns None if color not found.
    """
    if openpyxl_color is None:
        return None
    rgb = getattr(openpyxl_color, "rgb", None)
    if rgb:
        rgb = rgb.upper()
        if len(rgb) == 8 and rgb.startswith("FF"):
            rgb = rgb[2:]
        if len(rgb) == 6:
            return f"#{rgb}"
    return None

def preview_file_with_styles(path: Path, nrows: int = 10, max_cols: int = 20) -> Optional[str]:
    """
    Return HTML table (string) that shows first nrows with background fills from Excel.
    - path: Path to .xlsx file
    - nrows: number of rows to display
    - max_cols: maximum number of columns to render (safety)
    Returns HTML string or None if file can't be read / unsupported.
    """
    path = Path(path)
    if not path.exists():
        return None
    suffix = path.suffix.lower()
    if suffix not in [".xlsx", ".xlsm", ".xltx", ".xltm", ".xls"]:
        # styles not supported for CSV
        return None

    # read values via pandas
    try:
        try:
            df = pd.read_excel(path, engine="openpyxl", header=0)
            header_present = True
        except Exception:
            df = pd.read_excel(path, engine="openpyxl", header=None)
            header_present = False
    except Exception as e:
        return f"<pre>Could not read Excel file: {html.escape(str(e))}</pre>"

    # limit rows and cols
    df_preview = df.head(nrows).iloc[:, :max_cols]

    # openpyxl workbook to read fills
    try:
        wb = openpyxl.load_workbook(filename=str(path), read_only=False, data_only=True)
        ws = wb.active
    except Exception as e:
        return f"<pre>Could not open workbook for styles: {html.escape(str(e))}</pre>"

    # CSS + table build
    css = """
    <style>
    table.excel-preview { border-collapse: collapse; font-family: Arial, Helvetica, sans-serif; width: 100%; }
    table.excel-preview th, table.excel-preview td { border: 1px solid #ddd; padding: 6px 8px; text-align: left; vertical-align: top; }
    table.excel-preview th { background:#f6f6f6; font-weight:600; }
    </style>
    """

    html_rows = []

    # Build header row
    header_cells = []
    for col_idx, col_name in enumerate(df_preview.columns, start=1):
        header_cells.append(f"<th>{html.escape(str(col_name))}</th>")
    html_rows.append("<tr>" + "".join(header_cells) + "</tr>")

    # Data rows
    # If header_present True -> pandas row 0 maps to Excel row 2 (Excel header at row1).
    # If header_present False -> pandas row 0 maps to Excel row 1.
    for ridx, (_, row_series) in enumerate(df_preview.iterrows()):
        cells_html = []
        excel_row_num = (2 + ridx) if header_present else (1 + ridx)
        for cidx, col_name in enumerate(df_preview.columns, start=1):
            try:
                cell_value = row_series.iloc[cidx - 1]
            except Exception:
                cell_value = row_series[cidx - 1]
            text = "" if pd.isna(cell_value) else str(cell_value)
            # Get openpyxl cell to inspect fill
            try:
                wb_cell = ws.cell(row=excel_row_num, column=cidx)
                fill = getattr(wb_cell, "fill", None)
                bg = None
                if fill:
                    fg = getattr(fill, "fgColor", None)
                    bg = _openpyxl_color_to_hex(fg) if fg is not None else None
                style_attr = f" style='background:{bg};'" if (bg is not None) else ""
            except Exception:
                style_attr = ""
            cells_html.append(f"<td{style_attr}>{html.escape(text)}</td>")
        html_rows.append("<tr>" + "".join(cells_html) + "</tr>")

    table_html = f"{css}<table class='excel-preview'>{''.join(html_rows)}</table>"
    return table_html

# ---------- helpers for CSV-based combined highlight (green/red) ----------

def _parse_cell_address_token(tok: str) -> Optional[Tuple[int, int]]:
    """
    Parse an address token in A1 style like 'B12' and return (excel_row, excel_col) 1-based.
    Returns None if cannot parse.
    """
    tok = str(tok).strip()
    if not tok:
        return None
    m = re.match(r"^([A-Za-z]+)(\d+)$", tok)
    if m:
        col_letters = m.group(1)
        row_num = int(m.group(2))
        try:
            col_idx = column_index_from_string(col_letters.upper())
            return (row_num, col_idx)
        except Exception:
            return None
    return None

def _parse_highlight_csv(path: Path) -> Dict[str, Set[Tuple[int,int]]]:
    """
    Parse a highlighted-cells CSV into a dict keyed by optional 'table' or 'sheet' or filename,
    mapping to a set of (excel_row, excel_col) pairs (1-based).
    """
    path = Path(path)
    if not path.exists():
        return {}

    try:
        df = pd.read_csv(path, dtype=str)
    except Exception:
        try:
            df = pd.read_csv(path, engine="python", dtype=str)
        except Exception:
            return {}

    cols = {c.lower(): c for c in df.columns}

    table_key_col = None
    for candidate in ("table", "sheet", "file", "table_name"):
        if candidate in cols:
            table_key_col = cols[candidate]
            break

    results: Dict[str, Set[Tuple[int,int]]] = {}

    def add_coord(key: str, row: int, col: int):
        results.setdefault(key, set()).add((row, col))

    for _, row in df.iterrows():
        key = "__ALL__" if table_key_col is None else (row.get(table_key_col) or "__ALL__")
        # Try direct address cell
        if "cell" in cols:
            tok = row.get(cols["cell"])
            parsed = _parse_cell_address_token(tok) if pd.notna(tok) else None
            if parsed:
                add_coord(key, parsed[0], parsed[1])
                continue
        if "address" in cols:
            tok = row.get(cols["address"])
            parsed = _parse_cell_address_token(tok) if pd.notna(tok) else None
            if parsed:
                add_coord(key, parsed[0], parsed[1])
                continue
        # Try numeric row/col (various names)
        r = None
        c = None
        for rc in ("row", "r", "excel_row", "row_index"):
            if rc in cols:
                val = row.get(cols[rc])
                if pd.notna(val):
                    try:
                        r = int(float(val))
                        break
                    except Exception:
                        r = None
        for cc in ("col", "c", "column", "col_index"):
            if cc in cols:
                val = row.get(cols[cc])
                if pd.notna(val):
                    try:
                        c = int(float(val))
                        break
                    except Exception:
                        c = None
        if r is not None and c is not None:
            if r >= 1 and c >= 1:
                add_coord(key, r, c)
            else:
                add_coord(key, r + 1, c + 1)
            continue
        # If generic coordinate-like value present, try parse it
        for possible in df.columns:
            val = row.get(possible)
            if pd.notna(val):
                parsed = _parse_cell_address_token(val)
                if parsed:
                    add_coord(key, parsed[0], parsed[1])
                    break

    return results

def _excel_coords_to_df_indices(coords: Set[Tuple[int,int]], header_present: bool) -> Set[Tuple[int,int]]:
    """
    Map excel coordinates (row, col) 1-based to dataframe (r_idx, c_idx) 0-based.
    """
    out = set()
    for erow, ecol in coords:
        if header_present:
            df_r = erow - 2
        else:
            df_r = erow - 1
        df_c = ecol - 1
        if df_r >= 0 and df_c >= 0:
            out.add((df_r, df_c))
    return out

def build_combined_highlight_html(
    excel_path: Path,
    highlighted_csv: Optional[Path],
    siggreen_csv: Optional[Path],
    nrows: int = 10,
    max_cols: int = 20
) -> Optional[str]:
    """
    Build HTML table showing values from the excel_path and applying red/green highlights
    based on companion CSVs.
    """
    p = Path(excel_path)
    if not p.exists():
        return None
    # read dataframe with header detection
    try:
        try:
            df = pd.read_excel(p, engine="openpyxl", header=0)
            header_present = True
        except Exception:
            df = pd.read_excel(p, engine="openpyxl", header=None)
            header_present = False
    except Exception as e:
        return f"<pre>Could not read Excel file: {html.escape(str(e))}</pre>"

    df_preview = df.head(nrows).iloc[:, :max_cols]

    # parse companion CSVs
    green_coords = set()
    red_coords = set()

    base_key_guess = p.stem
    base_key_guess = re.sub(r'(_Green|_highlighted|_SigTestTable|_SigTestTable_Green|_highlighted_cells|_cells)$', '', base_key_guess, flags=re.IGNORECASE)

    def extract_coords_from_csv(csv_path: Optional[Path], collect_set: Set[Tuple[int,int]]):
        if csv_path is None or not Path(csv_path).exists():
            return
        parsed = _parse_highlight_csv(Path(csv_path))
        if not parsed:
            return
        if base_key_guess in parsed:
            coords = parsed[base_key_guess]
        elif "__ALL__" in parsed:
            coords = parsed["__ALL__"]
        else:
            if len(parsed) == 1:
                coords = next(iter(parsed.values()))
            else:
                coords = set().union(*parsed.values())
        collect_set.update(coords)

    extract_coords_from_csv(highlighted_csv, red_coords)
    extract_coords_from_csv(siggreen_csv, green_coords)

    green_df_coords = _excel_coords_to_df_indices(green_coords, header_present)
    red_df_coords = _excel_coords_to_df_indices(red_coords, header_present)

    # Build HTML
    css = """
    <style>
    table.excel-preview { border-collapse: collapse; font-family: Arial, Helvetica, sans-serif; width: 100%; }
    table.excel-preview th, table.excel-preview td { border: 1px solid #ddd; padding: 6px 8px; text-align: left; vertical-align: top; }
    table.excel-preview th { background:#f6f6f6; font-weight:600; }
    </style>
    """
    html_rows = []

    # header
    header_cells = []
    for col_name in df_preview.columns:
        header_cells.append(f"<th>{html.escape(str(col_name))}</th>")
    html_rows.append("<tr>" + "".join(header_cells) + "</tr>")

    # data rows
    for r_idx in range(len(df_preview)):
        row_series = df_preview.iloc[r_idx]
        cells_html = []
        for c_idx in range(len(df_preview.columns)):
            val = row_series.iloc[c_idx] if hasattr(row_series, "iloc") else row_series[c_idx]
            text = "" if pd.isna(val) else str(val)
            style = ""
            if (r_idx, c_idx) in green_df_coords:
                style = "background:#C6EFCE;"
            elif (r_idx, c_idx) in red_df_coords:
                style = "background:#FFC7CE;"
            cells_html.append(f"<td style='{style}'>{html.escape(text)}</td>")
        html_rows.append("<tr>" + "".join(cells_html) + "</tr>")

    return css + "<table class='excel-preview'>" + "".join(html_rows) + "</table>"
