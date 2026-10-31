"""Excel writers (xlsxwriter, values only — no formulas)."""

from __future__ import annotations

import io
import re
from datetime import date, datetime

import pandas as pd
import xlsxwriter

from buybox import labels as L
from buybox.core import Params, result_columns_present

HEADER_BLUE = "#1F4E78"
HEADER_COMPARE = "#7F6000"
FILL_WIN = "#C6EFCE"
FILL_LOSS = "#FFC7CE"
FILL_NOT_EVAL = "#FFEB9C"
COMPARE_COLUMNS = ("Comment_Prev_File", "BuyBox_Change")
PRICE_COLUMNS = ("Selling_Price", "BuyBox_Min_Price", "Target_Price", "RRP_Price")
PCT_COLUMNS = ("Gap_to_BuyBox", "Required_Cut_%")

SELLER_PCT = ("Win_Rate", "Comp_Win_Rate", "Prev_Comp_Win_Rate", "Median_Loss_Gap", "Share")
SELLER_INT = ("Variants", "Not_Evaluated", "Evaluated", "Wins", "Comp_Variants", "Comp_Wins", "Gained",
              "Lost", "Single_Seller_Variants", "No_BuyBox", "Near_Miss")

QUICKWIN_COLUMNS: dict[str, str] = {
    "DKPC": "DKPC",
    "DKP": "DKP",
    "Product_Name": "Product_Name",
    "Weight": "Weight",
    "Selling_Price": "Selling_Price",
    "BuyBox_Min_Price": "BuyBox_Min_Price",
    "Gap_to_BuyBox": "Gap",
    "Target_Price": "Target_Price",
    "Required_Cut_%": "Required_Cut_%",
    "Margin_Flag": "Margin_Flag",
}


def safe_filename(text: str, max_len: int = 60) -> str:
    """Make a string safe to use inside a file name (keeps Persian letters)."""
    cleaned = re.sub(r'[\\/:*?"<>|\r\n\t]+', "_", str(text)).strip(" ._")
    return (cleaned or "file")[:max_len]


def _values(df: pd.DataFrame) -> list[list[object]]:
    """Frame -> list of rows with NA mapped to None and datetimes to ISO strings."""
    out = df.copy()
    for c in out.columns:
        if pd.api.types.is_datetime64_any_dtype(out[c]):
            out[c] = out[c].dt.strftime("%Y-%m-%d %H:%M").where(out[c].notna(), None)
    obj = out.astype(object)
    return obj.where(pd.notna(obj), None).to_numpy().tolist()


def _write_rows(ws: xlsxwriter.worksheet.Worksheet, first_row: int, rows: list[list[object]],
                fmt: object | None = None) -> None:
    for i, r in enumerate(rows):
        for j, v in enumerate(r):
            if v is None:
                continue
            if isinstance(v, float) and v != v:  # NaN
                continue
            if isinstance(v, (datetime, date)):
                v = v.isoformat()
            if fmt is None:
                ws.write(first_row + i, j, v)
            else:
                ws.write(first_row + i, j, v, fmt)


# ---------------------------------------------------------------------------
# Main workbook
# ---------------------------------------------------------------------------
def build_main_workbook(
    raw: pd.DataFrame,
    rows: pd.DataFrame,
    seller_table: pd.DataFrame,
    kam_merged: pd.DataFrame | None,
    quality: pd.DataFrame,
    params: Params,
    input_name: str,
    has_prev: bool,
    prev_label: str = "",
) -> bytes:
    """Build the main analysis workbook and return its bytes."""
    buf = io.BytesIO()
    wb = xlsxwriter.Workbook(buf, {"in_memory": True, "nan_inf_to_errors": True})
    f = _formats(wb)
    _seller_sheet(wb, f, seller_table, params, input_name, has_prev, prev_label)
    _rows_sheet(wb, f, raw, rows, has_prev)
    _kam_sheet(wb, f, kam_merged)
    _quality_sheet(wb, f, quality)
    wb.close()
    return buf.getvalue()


def _formats(wb: xlsxwriter.Workbook) -> dict[str, object]:
    base = {"font_name": "Tahoma", "font_size": 10}
    return {
        "title": wb.add_format({**base, "bold": True, "font_size": 14, "font_color": HEADER_BLUE}),
        "label": wb.add_format({**base, "bold": True}),
        "text": wb.add_format(base),
        "head": wb.add_format({**base, "bold": True, "font_color": "#FFFFFF", "bg_color": HEADER_BLUE,
                               "text_wrap": True, "valign": "vcenter", "align": "center", "border": 1}),
        "head_cmp": wb.add_format({**base, "bold": True, "font_color": "#FFFFFF", "bg_color": HEADER_COMPARE,
                                   "text_wrap": True, "valign": "vcenter", "align": "center", "border": 1}),
        "int": wb.add_format({**base, "num_format": "#,##0"}),
        "num": wb.add_format({**base, "num_format": "#,##0.####"}),
        "pct": wb.add_format({**base, "num_format": "0.0%"}),
        "pp": wb.add_format({**base, "num_format": '+0.0" pp";-0.0" pp";0.0" pp"'}),
        "ratio": wb.add_format({**base, "num_format": "0.00"}),
        "t_text": wb.add_format({**base, "bold": True, "top": 2, "bg_color": "#D9E1F2"}),
        "t_int": wb.add_format({**base, "bold": True, "top": 2, "bg_color": "#D9E1F2", "num_format": "#,##0"}),
        "t_pct": wb.add_format({**base, "bold": True, "top": 2, "bg_color": "#D9E1F2", "num_format": "0.0%"}),
        "t_pp": wb.add_format({**base, "bold": True, "top": 2, "bg_color": "#D9E1F2",
                               "num_format": '+0.0" pp";-0.0" pp";0.0" pp"'}),
        "green": wb.add_format({"bg_color": FILL_WIN, "font_color": "#006100"}),
        "red": wb.add_format({"bg_color": FILL_LOSS, "font_color": "#9C0006"}),
        "amber": wb.add_format({"bg_color": FILL_NOT_EVAL, "font_color": "#9C5700"}),
        "note": wb.add_format({**base, "font_color": "#404040"}),
    }


def _seller_sheet(wb, f, table: pd.DataFrame, params: Params, input_name: str, has_prev: bool,
                  prev_label: str) -> None:
    ws = wb.add_worksheet("BuyBox_Seller_Report")
    ws.right_to_left()
    cols = [c for c in L.SELLER_COLUMNS if c in table.columns]
    if not has_prev:
        cols = [c for c in cols if c not in L.PREV_ONLY_COLUMNS]

    ws.write(0, 0, L.EXCEL_TITLE, f["title"])
    ws.write(1, 0, "فایل ورودی:", f["label"])
    ws.write(1, 1, input_name, f["text"])
    ws.write(1, 3, "تاریخ تهیه:", f["label"])
    ws.write(1, 4, datetime.now().strftime("%Y-%m-%d %H:%M"), f["text"])
    if has_prev:
        ws.write(1, 6, "دوره قبل:", f["label"])
        ws.write(1, 7, prev_label, f["text"])
    ws.write(2, 0, "آستانه باخت نزدیک:", f["label"])
    ws.write_number(2, 1, params.near_miss_threshold, f["pct"])
    ws.write(2, 3, "زیر قیمت زدن (ریال):", f["label"])
    ws.write_number(2, 4, params.undercut, f["int"])
    ws.write(2, 6, "حداکثر کاهش قابل قبول:", f["label"])
    ws.write_number(2, 7, params.max_acceptable_cut, f["pct"])

    hr = 4
    for j, c in enumerate(cols):
        ws.write(hr, j, L.SELLER_COLUMNS[c], f["head"])
    ws.set_row(hr, 42)

    data = table[~table["Is_Total"]]
    total = table[table["Is_Total"]]
    r = hr + 1
    for rec in data[cols].itertuples(index=False):
        _write_seller_row(ws, f, r, cols, rec, total_row=False)
        r += 1
    last_data_row = r - 1
    for rec in total[cols].itertuples(index=False):
        _write_seller_row(ws, f, r, cols, rec, total_row=True)
        r += 1

    if last_data_row >= hr + 1:
        if "Comp_Win_Rate" in cols:
            j = cols.index("Comp_Win_Rate")
            ws.conditional_format(hr + 1, j, last_data_row, j, {
                "type": "3_color_scale", "min_type": "num", "min_value": 0, "mid_type": "num",
                "mid_value": 0.45, "max_type": "num", "max_value": 1,
                "min_color": "#F8696B", "mid_color": "#FFEB84", "max_color": "#63BE7B"})
        if "Comp_Win_Rate_Change" in cols:
            j = cols.index("Comp_Win_Rate_Change")
            ws.conditional_format(hr + 1, j, last_data_row, j,
                                  {"type": "cell", "criteria": ">=", "value": 5, "format": f["green"]})
            ws.conditional_format(hr + 1, j, last_data_row, j,
                                  {"type": "cell", "criteria": "<=", "value": -5, "format": f["red"]})

    ws.freeze_panes(hr + 1, 2)
    ws.autofilter(hr, 0, max(last_data_row, hr), len(cols) - 1)
    widths = {"Seller_Name": 28, "KAM": 22, "Main_Brand": 16, "Seller_ID": 11}
    for j, c in enumerate(cols):
        ws.set_column(j, j, widths.get(c, 13))

    r += 2
    for line in L.EXCEL_NOTES:
        ws.write(r, 0, line, f["label"] if line.endswith(":") else f["note"])
        r += 1


def _write_seller_row(ws, f, r: int, cols: list[str], rec: tuple, total_row: bool) -> None:
    p = "t_" if total_row else ""
    for j, (c, v) in enumerate(zip(cols, rec)):
        missing = v is None or v is pd.NA or (isinstance(v, float) and v != v)
        if c in SELLER_PCT:
            if missing:
                ws.write_string(r, j, "-", f[p + "text"] if total_row else f["text"])
            else:
                ws.write_number(r, j, float(v), f[p + "pct"] if total_row else f["pct"])
        elif c == "Comp_Win_Rate_Change":
            if missing:
                ws.write_string(r, j, "-", f[p + "text"] if total_row else f["text"])
            else:
                ws.write_number(r, j, float(v), f[p + "pp"] if total_row else f["pp"])
        elif c in SELLER_INT:
            ws.write_number(r, j, 0 if missing else int(v), f[p + "int"] if total_row else f["int"])
        elif c == "Seller_ID":
            if missing:
                ws.write_blank(r, j, None, f[p + "text"] if total_row else f["text"])
            else:
                ws.write(r, j, int(v) if isinstance(v, (int,)) or str(v).isdigit() else str(v),
                         f["t_text"] if total_row else f["text"])
        else:
            ws.write_string(r, j, "" if missing else str(v), f[p + "text"] if total_row else f["text"])


def sheet1_frame(raw: pd.DataFrame, rows: pd.DataFrame, has_prev: bool) -> tuple[pd.DataFrame, int]:
    """Original columns (original order) followed by computed columns.

    Returns the frame and the index of the first comparison column (or -1).
    Original columns whose names clash with computed ones get an ``_input`` suffix.
    """
    computed = result_columns_present(rows)
    if has_prev:
        computed += [c for c in COMPARE_COLUMNS if c in rows.columns]
    orig = raw.copy()
    orig.columns = [f"{c}_input" if c in computed else c for c in orig.columns]
    out = pd.concat([orig.reset_index(drop=True), rows[computed].reset_index(drop=True)], axis=1)
    first_cmp = len(orig.columns) + computed.index("Comment_Prev_File") if "Comment_Prev_File" in computed else -1
    return out, first_cmp


def _rows_sheet(wb, f, raw: pd.DataFrame, rows: pd.DataFrame, has_prev: bool) -> None:
    ws = wb.add_worksheet("Sheet1")
    df, first_cmp = sheet1_frame(raw, rows, has_prev)
    for j, c in enumerate(df.columns):
        is_cmp = first_cmp >= 0 and j >= first_cmp
        ws.write(0, j, str(c), f["head_cmp"] if is_cmp else f["head"])
        if c in PRICE_COLUMNS:
            ws.set_column(j, j, 15, f["int"])
        elif c in PCT_COLUMNS:
            ws.set_column(j, j, 12, f["pct"])
        elif c == "Cut_vs_Wage":
            ws.set_column(j, j, 12, f["ratio"])
        elif c in ("Comment", "Competition_Type", "BuyBox_Change", "Comment_Prev_File", "BuyBox_Winner_Seller"):
            ws.set_column(j, j, 24)
        elif c == "Product_Name":
            ws.set_column(j, j, 40)
        else:
            ws.set_column(j, j, 13)
    ws.set_row(0, 30)
    _write_rows(ws, 1, _values(df))
    n = len(df)
    ws.freeze_panes(1, 0)
    ws.autofilter(0, 0, max(n, 1), len(df.columns) - 1)
    if n and "Comment" in df.columns:
        j = list(df.columns).index("Comment")
        ws.conditional_format(1, j, n, j, {"type": "cell", "criteria": "==", "value": f'"{L.COMMENT_WIN}"',
                                           "format": f["green"]})
        ws.conditional_format(1, j, n, j, {"type": "cell", "criteria": "==", "value": f'"{L.COMMENT_LOSS}"',
                                           "format": f["red"]})
        ws.conditional_format(1, j, n, j, {"type": "text", "criteria": "containing",
                                           "value": "قابل ارزیابی نیست", "format": f["amber"]})


def _kam_sheet(wb, f, kam_merged: pd.DataFrame | None) -> None:
    ws = wb.add_worksheet("KAM_LIST")
    ws.right_to_left()
    headers = ["Seller Id", "Seller Name", "Category Level1 Name Fa", "KAM", "تعداد ردیف در لیست اصلی", "یادداشت"]
    for j, h in enumerate(headers):
        ws.write(0, j, h, f["head"])
    if kam_merged is None or kam_merged.empty:
        ws.write(1, 0, "لیست KAM بارگذاری نشده است.", f["text"])
        return
    cols = ["Seller_ID", "Seller_Name", "Category", "KAM", "Rows", "KAM_Note"]
    _write_rows(ws, 1, _values(kam_merged[cols]))
    ws.set_column(0, 0, 12)
    ws.set_column(1, 3, 30)
    ws.set_column(4, 5, 18)
    ws.freeze_panes(1, 0)
    ws.autofilter(0, 0, len(kam_merged), len(cols) - 1)


def _quality_sheet(wb, f, quality: pd.DataFrame) -> None:
    ws = wb.add_worksheet("Data_Quality")
    ws.right_to_left()
    for j, h in enumerate(quality.columns):
        ws.write(0, j, h, f["head"])
    _write_rows(ws, 1, _values(quality), f["text"])
    ws.set_column(0, 0, 60)
    ws.set_column(1, 2, 14)
    ws.set_column(3, 3, 40)


# ---------------------------------------------------------------------------
# Quick-win workbook (one seller)
# ---------------------------------------------------------------------------
def quickwin_filename(seller_name: str, when: date | None = None) -> str:
    """``QuickWin_<SellerName>_<date>.xlsx``."""
    when = when or date.today()
    return f"QuickWin_{safe_filename(seller_name)}_{when.strftime('%Y%m%d')}.xlsx"


def build_quickwin_workbook(near_miss_rows: pd.DataFrame, seller_name: str, params: Params) -> bytes:
    """Excel with one seller's near-miss rows, to send to the seller."""
    buf = io.BytesIO()
    wb = xlsxwriter.Workbook(buf, {"in_memory": True})
    f = _formats(wb)
    ws = wb.add_worksheet("QuickWin")
    ws.right_to_left()
    cols = [c for c in QUICKWIN_COLUMNS if c in near_miss_rows.columns]
    for j, c in enumerate(cols):
        ws.write(0, j, QUICKWIN_COLUMNS[c], f["head"])
        if c in ("Selling_Price", "BuyBox_Min_Price", "Target_Price"):
            ws.set_column(j, j, 16, f["int"])
        elif c in ("Gap_to_BuyBox", "Required_Cut_%"):
            ws.set_column(j, j, 14, f["pct"])
        elif c == "Product_Name":
            ws.set_column(j, j, 45)
        elif c == "Margin_Flag":
            ws.set_column(j, j, 14)
        else:
            ws.set_column(j, j, 13)
    data = near_miss_rows[cols].sort_values("Gap_to_BuyBox") if "Gap_to_BuyBox" in cols else near_miss_rows[cols]
    _write_rows(ws, 1, _values(data))
    ws.freeze_panes(1, 0)
    ws.autofilter(0, 0, max(len(data), 1), len(cols) - 1)
    notes = wb.add_worksheet("توضیحات")
    notes.right_to_left()
    notes.set_column(0, 0, 90)
    notes.write(0, 0, f"فروشنده: {seller_name}", f["label"])
    notes.write(1, 0, f"تاریخ: {date.today().isoformat()}", f["text"])
    notes.write(2, 0, f"قیمت هدف = قیمت بای باکس − {params.undercut:,.0f} ریال.", f["note"])
    notes.write(3, 0, "Gap = قیمت شما / قیمت بای باکس − ۱ ؛ Required_Cut_% = ۱ − قیمت هدف / قیمت شما", f["note"])
    notes.write(4, 0, "قانون تساوی الگوریتم واقعی بای باکس مشخص نیست؛ هم‌قیمت شدن ممکن است کافی نباشد.", f["note"])
    wb.close()
    return buf.getvalue()
