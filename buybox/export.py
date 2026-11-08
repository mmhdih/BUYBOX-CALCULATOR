"""Standard Excel exports (xlsxwriter, values only — no formulas).

``BuyBox_<input>.xlsx`` keeps the specified structure (seller report, Sheet1 with the original
columns + computed columns, KAM_LIST, Data_Quality) and adds formatting for readability:
colour key, header tooltips, banded rows, data bars and a column guide sheet.
"""

from __future__ import annotations

import io
import re
from datetime import date, datetime

import pandas as pd
import xlsxwriter

from buybox import labels as L
from buybox.core import Params, result_columns_present
from buybox.xlstyle import (
    AMBER_BG, BIG_WORKBOOK_ROWS, BROWN, COLOR_KEY, GREEN_BG, MAX_SHEET_ROWS, NAVY, NAVY_LIGHT, NUM_FORMATS, RED_BG,
    ROW_KINDS, Sheet, SheetWriter, Styles, Table, clean, frame_rows, iter_frame_rows, workbook_options,
)

HEADER_BLUE = NAVY
HEADER_COMPARE = BROWN
FILL_WIN = GREEN_BG
FILL_LOSS = RED_BG
FILL_NOT_EVAL = AMBER_BG
BAND = "#F3F6FA"
COMPARE_COLUMNS = ("Comment_Prev_File", "BuyBox_Change")

SELLER_PCT = ("Win_Rate", "Comp_Win_Rate", "Prev_Comp_Win_Rate", "Median_Loss_Gap", "Share")
SELLER_INT = ("Variants", "Not_Evaluated", "Evaluated", "Wins", "Comp_Variants", "Comp_Wins", "Gained",
              "Lost", "Single_Seller_Variants", "No_BuyBox", "Near_Miss")

#: QuickWin columns -> English name required by the spec (shown together with a Persian title).
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
QUICKWIN_FA: dict[str, str] = {
    "DKPC": "شناسه تنوع", "DKP": "شناسه کالا", "Product_Name": "نام کالا", "Weight": "وزن (گرم)",
    "Selling_Price": "قیمت فعلی شما", "BuyBox_Min_Price": "قیمت بای باکس", "Gap_to_BuyBox": "فاصله تا بای باکس",
    "Target_Price": "قیمت پیشنهادی", "Required_Cut_%": "کاهش قیمت لازم", "Margin_Flag": "هشدار حاشیه",
}


def safe_filename(text: str, max_len: int = 60) -> str:
    """Make a string safe to use inside a file name (keeps Persian letters)."""
    cleaned = re.sub(r'[\\/:*?"<>|\r\n\t]+', "_", str(text)).strip(" ._")
    return (cleaned or "file")[:max_len]


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
    big = len(rows) > BIG_WORKBOOK_ROWS
    wb = xlsxwriter.Workbook(buf, workbook_options(len(rows)))
    wb.set_properties({"title": L.EXCEL_TITLE, "comments": "Buy Box Analyzer"})
    st = Styles(wb)
    _seller_sheet(wb, st, seller_table, params, input_name, has_prev, prev_label)
    _rows_sheets(wb, st, raw, rows, has_prev, banding=not big)
    _kam_sheet(wb, st, kam_merged)
    _quality_sheet(wb, st, quality)
    _guide_sheet(wb, st, rows)
    wb.close()
    return buf.getvalue()


def _seller_sheet(wb, st: Styles, table: pd.DataFrame, params: Params, input_name: str, has_prev: bool,
                  prev_label: str) -> None:
    ws = wb.add_worksheet("BuyBox_Seller_Report")
    ws.right_to_left()
    ws.hide_gridlines(2)
    cols = [c for c in L.SELLER_COLUMNS if c in table.columns]
    if not has_prev:
        cols = [c for c in cols if c not in L.PREV_ONLY_COLUMNS]
    span = len(cols)

    # row 0: title band, row 1: file/date, row 2: parameters, row 3: colour key, row 4: header
    ws.set_row(0, 30)
    ws.merge_range(0, 0, 0, span - 1, L.EXCEL_TITLE, st.title())
    label = st.get(bold=True, font_color=NAVY, align="right", bg_color=NAVY_LIGHT, border=1,
                   border_color="#FFFFFF")
    value = st.get(align="right", bg_color="#FFFFFF", border=1, border_color=NAVY_LIGHT)
    ws.write(1, 0, "فایل ورودی:", label)
    ws.merge_range(1, 1, 1, 2, input_name, value)
    ws.write(1, 3, "تاریخ تهیه:", label)
    ws.write(1, 4, datetime.now().strftime("%Y-%m-%d %H:%M"), value)
    if has_prev and span > 7:
        ws.write(1, 6, "دوره قبل:", label)
        ws.write(1, 7, prev_label, value)
    ws.write(2, 0, "آستانه باخت نزدیک:", label)
    ws.write_number(2, 1, params.near_miss_threshold, st.get(num_format="0.0%", align="right", bold=True,
                                                             border=1, border_color=NAVY_LIGHT))
    ws.write(2, 3, "زیر قیمت زدن (ریال):", label)
    ws.write_number(2, 4, params.undercut, st.get(num_format="#,##0", align="right", bold=True, border=1,
                                                  border_color=NAVY_LIGHT))
    if span > 7:
        ws.write(2, 6, "حداکثر کاهش قابل قبول:", label)
        ws.write_number(2, 7, params.max_acceptable_cut, st.get(num_format="0.0%", align="right", bold=True,
                                                                border=1, border_color=NAVY_LIGHT))
    ws.set_row(3, 28)
    ws.merge_range(3, 0, 3, span - 1, "رنگ وین‌ریت رقابتی: طیف قرمز (پایین) ← زرد ← سبز (بالا). "
                   "تغییر: سبز = حداقل ۵+ واحد درصد، قرمز = حداقل ۵− واحد درصد. "
                   "میله آبی = اندازه نسبی. روی سرستون‌ها مکث کنید تا توضیح نمایش داده شود.",
                   st.get(italic=True, font_color="#595959", align="right", text_wrap=True))

    hr = 4
    ws.set_row(hr, 44)
    for j, c in enumerate(cols):
        ws.write(hr, j, L.SELLER_COLUMNS[c], st.header())
        tip = L.SELLER_COLUMN_HELP.get(c)
        if tip:
            ws.write_comment(hr, j, tip, {"x_scale": 2.2, "y_scale": 1.1, "font_name": "Tahoma"})

    data = table[~table["Is_Total"]]
    total = table[table["Is_Total"]]
    r = hr + 1
    for i, rec in enumerate(data[cols].itertuples(index=False)):
        _write_seller_row(ws, st, r, cols, rec, total_row=False, band=bool(i % 2))
        r += 1
    last_data_row = r - 1
    for rec in total[cols].itertuples(index=False):
        _write_seller_row(ws, st, r, cols, rec, total_row=True, band=False)
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
                                  {"type": "cell", "criteria": ">=", "value": 5, "format": st.green()})
            ws.conditional_format(hr + 1, j, last_data_row, j,
                                  {"type": "cell", "criteria": "<=", "value": -5, "format": st.red()})
        for c in ("Variants", "Comp_Variants", "Near_Miss"):
            if c in cols:
                j = cols.index(c)
                ws.conditional_format(hr + 1, j, last_data_row, j, {
                    "type": "data_bar", "bar_color": "#9DC3E6", "bar_solid": True, "bar_direction": "right"})

    ws.freeze_panes(hr + 1, 2)
    ws.autofilter(hr, 0, max(last_data_row, hr), len(cols) - 1)
    widths = {"Seller_Name": 28, "KAM": 22, "Main_Brand": 16, "Seller_ID": 11}
    for j, c in enumerate(cols):
        ws.set_column(j, j, widths.get(c, 13))

    r += 2
    box = st.get(align="right", text_wrap=True, bg_color="#F7F7F7", font_color="#404040", left=1, right=1,
                 left_color="#D9D9D9", right_color="#D9D9D9")
    head = st.get(bold=True, align="right", bg_color=NAVY_LIGHT, font_color=NAVY, top=1, top_color="#D9D9D9")
    for k, line in enumerate(L.EXCEL_NOTES):
        ws.merge_range(r, 0, r, span - 1, line, head if k == 0 else box)
        r += 1


def _write_seller_row(ws, st: Styles, r: int, cols: list[str], rec: tuple, total_row: bool, band: bool) -> None:
    extra = {"bg_color": BAND} if band else {}
    for j, (c, v) in enumerate(zip(cols, rec)):
        missing = v is None or v is pd.NA or (isinstance(v, float) and v != v)
        if c in SELLER_PCT:
            kind = "pct"
        elif c == "Comp_Win_Rate_Change":
            kind = "pp"
        elif c in SELLER_INT:
            kind = "int"
        elif c == "Seller_ID":
            kind = "id"
        else:
            kind = "text"
        fmt = st.cell(kind, total=total_row, **({} if total_row else extra))
        if kind in ("pct", "pp") and missing:
            ws.write_string(r, j, "-", st.cell("text", total=total_row, align="right",
                                               **({} if total_row else extra)))
        elif kind == "int":
            ws.write_number(r, j, 0 if missing else int(v), fmt)
        elif kind == "pp":
            ws.write_number(r, j, round(float(v), 1) + 0.0, fmt)
        elif kind == "pct":
            ws.write_number(r, j, float(v), fmt)
        elif missing:
            ws.write_blank(r, j, None, fmt)
        elif kind == "id":
            ws.write(r, j, int(v) if str(v).isdigit() else str(v), fmt)
        else:
            ws.write_string(r, j, str(v), fmt)


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


def _rows_sheets(wb, st: Styles, raw: pd.DataFrame, rows: pd.DataFrame, has_prev: bool, banding: bool) -> None:
    """Sheet1 (split into Sheet1, Sheet1 (2) ... beyond Excel's row limit)."""
    df, first_cmp = sheet1_frame(raw, rows, has_prev)
    n_parts = max(1, -(-len(df) // MAX_SHEET_ROWS))
    for k in range(n_parts):
        part = df.iloc[k * MAX_SHEET_ROWS:(k + 1) * MAX_SHEET_ROWS]
        _rows_sheet(wb, st, part, first_cmp, "Sheet1" if k == 0 else f"Sheet1 ({k + 1})", banding)


def _rows_sheet(wb, st: Styles, df: pd.DataFrame, first_cmp: int, name: str, banding: bool) -> None:
    ws = wb.add_worksheet(name)
    cols = list(df.columns)
    ws.set_row(0, 32)
    for j, c in enumerate(cols):
        is_cmp = first_cmp >= 0 and j >= first_cmp
        ws.write(0, j, str(c), st.header(HEADER_COMPARE if is_cmp else HEADER_BLUE))
        if c in L.ROW_COLUMNS:
            ws.write_comment(0, j, f"{L.ROW_COLUMNS[c][0]}: {L.ROW_COLUMNS[c][1]}",
                             {"x_scale": 2.4, "y_scale": 1.1, "font_name": "Tahoma"})
        kind = ROW_KINDS.get(str(c), "text")
        width = {"money": 15, "pct": 12, "ratio": 11, "weight": 9, "id": 12}.get(kind, 13)
        if c in ("Comment", "Competition_Type", "BuyBox_Change", "Comment_Prev_File", "BuyBox_Winner_Seller"):
            width = 24
        elif c == "Product_Name":
            width = 42
        elif c == "Margin_Flag":
            width = 13
        ws.set_column(j, j, width, st.cell(kind) if kind != "text" else None)
    n = len(df)
    for i, row in enumerate(iter_frame_rows(df), start=1):
        for j, v in enumerate(row):
            if v is not None:
                ws.write(i, j, v)
    ws.freeze_panes(1, 1)
    ws.autofilter(0, 0, max(n, 1), len(cols) - 1)
    if not n:
        return

    def rule(col: str, text: str, fmt) -> None:
        if col in cols:
            j = cols.index(col)
            ws.conditional_format(1, j, n, j, {"type": "cell", "criteria": "==", "value": f'"{text}"',
                                               "format": fmt})

    rule("Comment", L.COMMENT_WIN, st.green())
    rule("Comment", L.COMMENT_LOSS, st.red())
    if "Comment" in cols:
        j = cols.index("Comment")
        ws.conditional_format(1, j, n, j, {"type": "text", "criteria": "containing",
                                           "value": "قابل ارزیابی نیست", "format": st.amber()})
    rule("Comment_Prev_File", L.COMMENT_WIN, st.green())
    rule("Comment_Prev_File", L.COMMENT_LOSS, st.red())
    rule("BuyBox_Change", L.CHANGE_GAINED, st.green())
    rule("BuyBox_Change", L.CHANGE_LOST, st.red())
    rule("BuyBox_Change", L.CHANGE_NEW, st.blue())
    rule("Margin_Flag", L.MARGIN_FLAG, st.red())
    rule("Competition_Type", L.COMP_COMPETITIVE, st.blue())
    if banding:  # lowest priority so the colour rules above win
        ws.conditional_format(1, 0, n, len(cols) - 1, {"type": "formula", "criteria": "=MOD(ROW(),2)=0",
                                                       "format": st.get(bg_color=BAND)})


def _kam_sheet(wb, st: Styles, kam_merged: pd.DataFrame | None) -> None:
    ws = wb.add_worksheet("KAM_LIST")
    ws.right_to_left()
    headers = ["Seller Id", "Seller Name", "Category Level1 Name Fa", "KAM", "تعداد ردیف در لیست اصلی", "یادداشت"]
    ws.set_row(0, 30)
    for j, h in enumerate(headers):
        ws.write(0, j, h, st.header())
    if kam_merged is None or kam_merged.empty:
        ws.write(1, 0, "لیست KAM بارگذاری نشده است.", st.note())
        ws.set_column(0, 0, 40)
        return
    cols = ["Seller_ID", "Seller_Name", "Category", "KAM", "Rows", "KAM_Note"]
    kinds = ["id", "text", "text", "text", "int", "text"]
    for i, row in enumerate(frame_rows(kam_merged[cols]), start=1):
        band = {"bg_color": BAND} if i % 2 == 0 else {}
        for j, v in enumerate(row):
            fmt = st.cell(kinds[j], **band)
            if v is None:
                ws.write_blank(i, j, None, fmt)
            else:
                ws.write(i, j, v, fmt)
    n = len(kam_merged)
    ws.conditional_format(1, 0, n, len(cols) - 1, {
        "type": "formula", "criteria": f'=$F2="{L.KAM_DUPLICATE_FLAG}"', "format": st.amber()})
    ws.set_column(0, 0, 12)
    ws.set_column(1, 3, 30)
    ws.set_column(4, 5, 20)
    ws.freeze_panes(1, 0)
    ws.autofilter(0, 0, n, len(cols) - 1)


def _quality_sheet(wb, st: Styles, quality: pd.DataFrame) -> None:
    ws = wb.add_worksheet("Data_Quality")
    ws.right_to_left()
    ws.set_row(0, 26)
    for j, h in enumerate(quality.columns):
        ws.write(0, j, h, st.header())
    for i, row in enumerate(frame_rows(quality), start=1):
        band = {"bg_color": BAND} if i % 2 == 0 else {}
        for j, v in enumerate(row):
            ws.write(i, j, "" if v is None else v, st.get(align="right", text_wrap=True, **band))
    n = len(quality)
    if "وضعیت" in quality.columns and n:
        j = list(quality.columns).index("وضعیت")
        for text, fmt in (("OK", st.green()), ("هشدار", st.amber()), ("اطلاع", st.blue())):
            ws.conditional_format(1, j, n, j, {"type": "cell", "criteria": "==", "value": f'"{text}"',
                                               "format": fmt})
    ws.set_column(0, 0, 60)
    ws.set_column(1, 2, 14)
    ws.set_column(3, 3, 44)
    ws.freeze_panes(1, 0)


def _guide_sheet(wb, st: Styles, rows: pd.DataFrame) -> None:
    present = [c for c in L.ROW_COLUMNS if c in rows.columns or c in ("Comment_Prev_File", "BuyBox_Change")]
    guide = pd.DataFrame([(c, L.ROW_COLUMNS[c][0], L.ROW_COLUMNS[c][1]) for c in present],
                         columns=["ستون در Sheet1", "عنوان فارسی", "توضیح"])
    sellers = pd.DataFrame([(L.SELLER_COLUMNS[k], v) for k, v in L.SELLER_COLUMN_HELP.items()],
                           columns=["ستون گزارش فروشندگان", "توضیح"])
    SheetWriter(wb, st).write(Sheet(
        name="راهنما", title="راهنمای ستون‌ها و روش محاسبه",
        tables=[Table(title="ستون‌های Sheet1", df=guide), Table(title="ستون‌های گزارش فروشندگان", df=sellers)],
        notes=[COLOR_KEY, *L.EXCEL_NOTES],
    ))


# ---------------------------------------------------------------------------
# Quick-win workbook (one seller)
# ---------------------------------------------------------------------------
def quickwin_filename(seller_name: str, when: date | None = None) -> str:
    """``QuickWin_<SellerName>_<date>.xlsx``."""
    when = when or date.today()
    return f"QuickWin_{safe_filename(seller_name)}_{when.strftime('%Y%m%d')}.xlsx"


def quickwin_header(col: str) -> str:
    """Persian title + the English column name, in one header cell."""
    return f"{QUICKWIN_FA[col]}\n{QUICKWIN_COLUMNS[col]}"


def build_quickwin_workbook(near_miss_rows: pd.DataFrame, seller_name: str, params: Params) -> bytes:
    """Excel with one seller's near-miss rows, to send to the seller."""
    buf = io.BytesIO()
    wb = xlsxwriter.Workbook(buf, {"in_memory": True, "nan_inf_to_errors": True})
    st = Styles(wb)
    cols = [c for c in QUICKWIN_COLUMNS if c in near_miss_rows.columns]
    data = near_miss_rows[cols].sort_values("Gap_to_BuyBox") if "Gap_to_BuyBox" in cols else near_miss_rows[cols]
    thr = params.near_miss_threshold * 100
    writer = SheetWriter(wb, st)
    writer.write(Sheet(
        name="QuickWin",
        title=f"پیشنهاد قیمت برای گرفتن بای باکس — {seller_name}",
        subtitle=(f"تاریخ: {date.today().isoformat()}  |  {len(data):,} تنوع که با کاهش قیمت حداکثر "
                  f"{thr:.1f}٪ به بای باکس می‌رسند"),
        tables=[Table(
            title="تنوع‌ها (مرتب از کمترین فاصله)", df=data,
            headers={c: quickwin_header(c) for c in cols},
            tips={c: L.ROW_COLUMNS[c][1] for c in cols if c in L.ROW_COLUMNS},
            text_rules={"Margin_Flag": {L.MARGIN_FLAG: "red"}},
        )],
        notes=[
            f"قیمت پیشنهادی = قیمت بای باکس − {params.undercut:,.0f} ریال.",
            "فاصله تا بای باکس = قیمت فعلی ÷ قیمت بای باکس − ۱  |  کاهش قیمت لازم = ۱ − قیمت پیشنهادی ÷ قیمت فعلی.",
            "قانون تساوی الگوریتم واقعی بای باکس مشخص نیست؛ هم‌قیمت شدن ممکن است کافی نباشد. "
            "بای باکس فقط بر اساس قیمت سنجیده شده است.",
        ],
    ))
    notes = wb.add_worksheet("توضیحات")
    notes.right_to_left()
    notes.set_column(0, 0, 90)
    lines = [f"فروشنده: {seller_name}", f"تاریخ: {date.today().isoformat()}",
             f"آستانه باخت نزدیک: {thr:.1f}٪", f"زیر قیمت زدن: {params.undercut:,.0f} ریال"]
    for i, line in enumerate(lines):
        notes.write(i, 0, line, st.get(align="right", bold=i == 0))
    wb.close()
    return buf.getvalue()


# kept for backwards compatibility with older callers/tests
_values = frame_rows
__all__ = ["build_main_workbook", "build_quickwin_workbook", "quickwin_filename", "safe_filename", "sheet1_frame",
           "clean", "NUM_FORMATS"]
