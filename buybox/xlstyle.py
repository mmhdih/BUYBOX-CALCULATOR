"""Shared Excel styling for every workbook the app writes (xlsxwriter, values only).

Gives all exports the same look: Tahoma font, navy headers, banded Excel tables with
filter buttons, Persian headers with explanatory tooltips, consistent number formats and
the red / amber / green colour key used in the app.
"""

from __future__ import annotations

import math
import re
from dataclasses import dataclass, field
from datetime import date, datetime

import pandas as pd

FONT = "Tahoma"
NAVY = "#1F4E78"
NAVY_LIGHT = "#D9E1F2"
SLATE = "#44546A"
BROWN = "#7F6000"
GREY = "#595959"
GREEN_BG, GREEN_FG = "#C6EFCE", "#006100"
RED_BG, RED_FG = "#FFC7CE", "#9C0006"
AMBER_BG, AMBER_FG = "#FFEB9C", "#9C5700"
BLUE_BG, BLUE_FG = "#DDEBF7", "#1F4E78"

#: Excel's hard limit is 1,048,576 rows per sheet (one is the header).
MAX_SHEET_ROWS = 1_000_000
#: Tables longer than this are streamed (plain rows + autofilter) instead of built as Excel tables.
STREAM_ROWS = 100_000
#: Workbooks with more data rows than this use xlsxwriter's constant-memory (streaming) mode.
BIG_WORKBOOK_ROWS = 200_000


def workbook_options(n_rows: int) -> dict:
    """xlsxwriter options: in memory for normal files, streaming (low RAM) for very large ones."""
    base = {"nan_inf_to_errors": True, "strings_to_numbers": False, "strings_to_urls": False}
    return {**base, "constant_memory": True} if n_rows > BIG_WORKBOOK_ROWS else {**base, "in_memory": True}

NUM_FORMATS: dict[str, str] = {
    "int": "#,##0",
    "money": "#,##0",
    "pct": "0.0%",
    "pp": '+0.0" pp";-0.0" pp";0.0" pp"',
    "weight": "0.####",
    "ratio": "0.00",
    "id": "0",
    "text": "General",
}

#: Column-name based defaults for row-level data.
ROW_KINDS: dict[str, str] = {
    "DKPC": "id", "DKP": "id", "Seller_ID": "id", "Group_ID": "id",
    "Selling_Price": "money", "BuyBox_Min_Price": "money", "Target_Price": "money", "RRP_Price": "money",
    "Weight": "weight", "Gap_to_BuyBox": "pct", "Required_Cut_%": "pct", "Cut_vs_Wage": "ratio",
    "Other_Sellers_Variants": "int", "Same_Seller_Variants": "int",
}


class Styles:
    """Memoized xlsxwriter formats."""

    def __init__(self, wb) -> None:
        self.wb = wb
        self._cache: dict[tuple, object] = {}

    def get(self, **props) -> object:
        key = tuple(sorted(props.items()))
        if key not in self._cache:
            self._cache[key] = self.wb.add_format({"font_name": FONT, "font_size": 10, "valign": "vcenter",
                                                   **props})
        return self._cache[key]

    # common formats ------------------------------------------------------
    def title(self):
        return self.get(bold=True, font_size=16, font_color="#FFFFFF", bg_color=NAVY, align="right",
                        indent=1)

    def subtitle(self):
        return self.get(italic=True, font_color=GREY, align="right")

    def section(self):
        return self.get(bold=True, font_size=12, font_color=NAVY, bottom=2, bottom_color=NAVY, align="right")

    def note(self):
        return self.get(font_color=GREY, italic=True, align="right", text_wrap=True, valign="top")

    def header(self, color: str = NAVY):
        return self.get(bold=True, font_color="#FFFFFF", bg_color=color, text_wrap=True, align="center",
                        border=1, border_color="#BFBFBF")

    def cell(self, kind: str = "text", total: bool = False, **extra):
        props = {"num_format": NUM_FORMATS.get(kind, "General")}
        if kind in ("int", "money", "pct", "pp", "weight", "ratio", "id"):
            props["align"] = "right"
        if total:
            props.update(bold=True, bg_color=NAVY_LIGHT, top=2, top_color=NAVY)
        props.update(extra)
        return self.get(**props)

    def green(self):
        return self.get(bg_color=GREEN_BG, font_color=GREEN_FG)

    def red(self):
        return self.get(bg_color=RED_BG, font_color=RED_FG)

    def amber(self):
        return self.get(bg_color=AMBER_BG, font_color=AMBER_FG)

    def blue(self):
        return self.get(bg_color=BLUE_BG, font_color=BLUE_FG)


_NUM_RE = re.compile(r"(?<![\w\-−.,٫/:])([+\-−]?[0-9][0-9,٬]*(?:[.٫][0-9]+)?(?:\s?(?:%|٪|pp))?)(?![\w\-/:])")


def bidi_text(text: str) -> str:
    """Anchor numbers (with % / pp) inside Persian text with LRM marks so Excel keeps their order."""
    return _NUM_RE.sub("\u200e\\1\u200e", str(text))


def round_pp(v: object) -> object:
    """Percentage-point change rounded to 1 decimal, without a '-0.0'."""
    v = clean(v)
    if isinstance(v, (int, float)):
        return round(float(v), 1) + 0.0
    return v


def clean(v: object) -> object:
    """Python value safe for xlsxwriter (NA/NaN/inf -> None, dates -> text)."""
    if v is None or v is pd.NA or v is pd.NaT:
        return None
    if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
        return None
    if isinstance(v, (pd.Timestamp, datetime)):
        return v.strftime("%Y-%m-%d %H:%M")
    if isinstance(v, date):
        return v.isoformat()
    if hasattr(v, "item"):  # numpy scalar
        try:
            return clean(v.item())
        except (ValueError, AttributeError):
            return str(v)
    return v


def frame_rows(df: pd.DataFrame) -> list[list[object]]:
    """Frame -> list of rows of clean Python values."""
    out = df.copy()
    for c in out.columns:
        if pd.api.types.is_datetime64_any_dtype(out[c]):
            out[c] = out[c].dt.strftime("%Y-%m-%d %H:%M")
    obj = out.astype(object)
    return [[clean(v) for v in row] for row in obj.where(pd.notna(obj), None).to_numpy().tolist()]


def iter_frame_rows(df: pd.DataFrame, chunk: int = 50_000):
    """Yield clean rows chunk by chunk (keeps memory flat for millions of rows)."""
    for start in range(0, len(df), chunk):
        yield from frame_rows(df.iloc[start:start + chunk])


def guess_width(header: str, values: list[object], kind: str) -> float:
    sample = [v for v in values[:300] if v is not None]
    if kind in ("int", "money"):
        longest = max((len(f"{v:,.0f}") for v in sample if isinstance(v, (int, float))), default=4)
    elif kind in ("pct", "pp", "ratio", "weight"):
        longest = 8
    else:
        longest = max((len(str(v)) for v in sample), default=4)
    head = max(len(part) for part in str(header).split("\n"))
    return float(min(max(longest + 2, head * 0.9 + 3, 8), 48))


@dataclass
class Table:
    """A data block written as an Excel table (filter buttons, banded rows)."""

    title: str
    df: pd.DataFrame
    kinds: dict[str, str] = field(default_factory=dict)       # column -> number kind
    headers: dict[str, str] = field(default_factory=dict)     # column -> displayed header
    tips: dict[str, str] = field(default_factory=dict)        # column -> tooltip (by column key)
    rate_cols: tuple[str, ...] = ()                           # red/amber/green 30%/60%
    change_cols: tuple[str, ...] = ()                         # green >= +5pp, red <= -5pp
    text_rules: dict[str, dict[str, str]] = field(default_factory=dict)  # column -> {text: colour}
    total_last: bool = False
    note: str = ""
    dash_for_missing: tuple[str, ...] = ()                    # columns that show "-" when empty
    chart: dict | None = None                                 # {"category", "value", "title"} -> bar chart


@dataclass
class Sheet:
    """One worksheet: title, optional KPI list and bullets, then tables."""

    name: str
    title: str
    subtitle: str = ""
    kpis: list[tuple[str, object, str]] = field(default_factory=list)   # (label, value, kind)
    bullets: list[str] = field(default_factory=list)
    tables: list[Table] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)
    rtl: bool = True


def _unique_headers(names: list[str]) -> list[str]:
    seen: dict[str, int] = {}
    out = []
    for n in names:
        n = str(n) if str(n).strip() else "-"
        if n in seen:
            seen[n] += 1
            n = f"{n} ({seen[n]})"
        else:
            seen[n] = 1
        out.append(n)
    return out


class SheetWriter:
    """Lays out :class:`Sheet` objects vertically on worksheets."""

    def __init__(self, wb, styles: Styles) -> None:
        self.wb = wb
        self.st = styles
        self._table_no = 0
        self._sheet_names: set[str] = set()
        self._col_formats: dict[int, object] = {}
        self._has_autofilter: dict[str, bool] = {}

    def _sheet_name(self, name: str) -> str:
        base = "".join(ch for ch in name if ch not in '[]:*?/\\')[:31] or "Sheet"
        n, i = base, 2
        while n.lower() in self._sheet_names:
            suffix = f" ({i})"
            n = base[: 31 - len(suffix)] + suffix
            i += 1
        self._sheet_names.add(n.lower())
        return n

    def write(self, sheet: Sheet) -> None:
        # split very large single-table sheets across several worksheets
        if len(sheet.tables) == 1 and len(sheet.tables[0].df) > MAX_SHEET_ROWS:
            t = sheet.tables[0]
            parts = range(0, len(t.df), MAX_SHEET_ROWS)
            for k, start in enumerate(parts, 1):
                part = Table(**{**t.__dict__, "df": t.df.iloc[start:start + MAX_SHEET_ROWS],
                                "title": f"{t.title} — بخش {k} از {len(parts)}"})
                self._write(Sheet(**{**sheet.__dict__, "name": f"{sheet.name} {k}", "tables": [part]}))
            return
        self._write(sheet)

    def _write(self, sheet: Sheet) -> None:
        ws = self.wb.add_worksheet(self._sheet_name(sheet.name))
        if sheet.rtl:
            ws.right_to_left()
        ws.hide_gridlines(2)
        widths: dict[int, float] = {0: 34.0, 1: 18.0}
        self._col_formats = {}
        span = max([len(t.df.columns) for t in sheet.tables] + [4])
        r = 0
        # set_row() always before writing a row: required by streaming (constant_memory) workbooks
        ws.set_row(r, 30)
        ws.merge_range(r, 0, r, span - 1, sheet.title, self.st.title())
        r += 1
        if sheet.subtitle:
            ws.set_row(r, 32)
            ws.merge_range(r, 0, r, span - 1, bidi_text(sheet.subtitle),
                           self.st.get(italic=True, font_color=GREY, align="right", text_wrap=True))
            r += 1
        r += 1
        if sheet.kpis:
            ws.write(r, 0, "شاخص", self.st.header())
            ws.write(r, 1, "مقدار", self.st.header())
            r += 1
            for i, (label, value, kind) in enumerate(sheet.kpis):
                band = {"bg_color": "#F3F6FA"} if i % 2 else {}
                ws.write(r, 0, label, self.st.get(align="right", border=1, border_color="#D9D9D9", **band))
                v = round_pp(value) if kind == "pp" else clean(value)
                fmt = self.st.cell(kind, bold=True, font_size=11, border=1, border_color="#D9D9D9", **band)
                if v is None:
                    ws.write(r, 1, "-", fmt)
                else:
                    ws.write(r, 1, v, fmt)
                r += 1
            r += 1
        if sheet.bullets:
            ws.write(r, 0, "نکات کلیدی", self.st.section())
            r += 1
            for b in sheet.bullets:
                ws.set_row(r, 30)
                ws.merge_range(r, 0, r, span - 1, f"• {bidi_text(b)}", self.st.get(align="right", text_wrap=True))
                r += 1
            r += 1
        single = len(sheet.tables) == 1 and not sheet.kpis and not sheet.bullets
        for t in sheet.tables:
            r = self._write_table(ws, r, t, widths, freeze=single)
            r += 2
        for line in sheet.notes:
            ws.set_row(r, 30 if len(line) > 90 else 18)
            ws.merge_range(r, 0, r, span - 1, bidi_text(line), self.st.note())
            r += 1
        for c, w in widths.items():
            ws.set_column(c, c, w, self._col_formats.get(c))

    def _write_table(self, ws, r: int, t: Table, widths: dict[int, float], freeze: bool) -> int:
        ws.write(r, 0, t.title, self.st.section())
        r += 1
        if t.note:
            ws.write(r, 0, t.note, self.st.note())
            r += 1
        if t.df.empty:
            ws.write(r, 0, "— بدون ردیف —", self.st.note())
            return r + 1
        cols = list(t.df.columns)
        headers = _unique_headers([t.headers.get(c, str(c)) for c in cols])
        # Excel tables (add_table) are not available in streaming workbooks
        if len(t.df) > STREAM_ROWS or getattr(self.wb, "constant_memory", False):
            return self._stream_table(ws, r, t, cols, headers, widths, freeze)
        data = frame_rows(t.df)
        for row in data:
            for j, c in enumerate(cols):
                if row[j] is None and c in t.dash_for_missing:
                    row[j] = "-"
        kinds = [t.kinds.get(c, ROW_KINDS.get(str(c), "text")) for c in cols]
        for j, k in enumerate(kinds):
            if k == "pp":
                for row in data:
                    row[j] = round_pp(row[j])
        self._table_no += 1
        columns = []
        for j, (c, h) in enumerate(zip(cols, headers)):
            columns.append({"header": h, "format": self.st.cell(kinds[j]),
                            "header_format": self.st.header()})
            widths[j] = max(widths.get(j, 0), guess_width(h, [row[j] for row in data], kinds[j]))
        first, last = r, r + len(data)
        ws.set_row(first, 36)
        ws.add_table(first, 0, last, len(cols) - 1, {
            "data": data, "columns": columns, "name": f"Table_{self._table_no}",
            "style": "Table Style Light 9", "autofilter": True,
        })
        for j, c in enumerate(cols):
            tip = t.tips.get(c)
            if tip:
                ws.write_comment(first, j, tip, {"x_scale": 2.2, "y_scale": 1.1, "font_name": FONT})
        if t.total_last and data:
            for j, v in enumerate(data[-1]):
                fmt = self.st.cell(kinds[j], total=True)
                if v is None:
                    ws.write_blank(last, j, None, fmt)
                else:
                    ws.write(last, j, v, fmt)
        self._body_rules(ws, t, cols, first, last - 1 if t.total_last else last)
        if freeze:
            ws.freeze_panes(first + 1, 0)
        return last + 1

    def _body_rules(self, ws, t: Table, cols: list, first: int, body_last: int) -> None:
        """Colour rules (rate, change, text) on the data rows of a table."""
        if body_last >= first + 1:
            for c in t.rate_cols:
                if c in cols:
                    j = cols.index(c)
                    self._rate_rules(ws, first + 1, body_last, j)
            for c in t.change_cols:
                if c in cols:
                    j = cols.index(c)
                    ws.conditional_format(first + 1, j, body_last, j, {"type": "cell", "criteria": ">=",
                                                                        "value": 5, "format": self.st.green()})
                    ws.conditional_format(first + 1, j, body_last, j, {"type": "cell", "criteria": "<=",
                                                                        "value": -5, "format": self.st.red()})
            for c, rules in t.text_rules.items():
                if c in cols:
                    j = cols.index(c)
                    for text, colour in rules.items():
                        ws.conditional_format(first + 1, j, body_last, j, {
                            "type": "cell", "criteria": "==", "value": f'"{text}"',
                            "format": getattr(self.st, colour)()})

    def _stream_table(self, ws, r: int, t: Table, cols: list, headers: list[str], widths: dict[int, float],
                      freeze: bool) -> int:
        """Very large table: header + rows streamed in chunks, column formats, autofilter (low memory)."""
        kinds = [t.kinds.get(c, ROW_KINDS.get(str(c), "text")) for c in cols]
        first = r
        ws.set_row(first, 36)
        sample = frame_rows(t.df.iloc[:300])
        for j, (c, h) in enumerate(zip(cols, headers)):
            ws.write(first, j, h, self.st.header())
            if t.tips.get(c):
                ws.write_comment(first, j, t.tips[c], {"x_scale": 2.2, "y_scale": 1.1, "font_name": FONT})
            if kinds[j] != "text":
                self._col_formats[j] = self.st.cell(kinds[j])
            widths[j] = max(widths.get(j, 0), guess_width(h, [row[j] for row in sample], kinds[j]))
        pp_cols = [j for j, k in enumerate(kinds) if k == "pp"]
        dash = [j for j, c in enumerate(cols) if c in t.dash_for_missing]
        n = len(t.df)
        small = n <= STREAM_ROWS
        cell_fmt = [self.st.cell(k) for k in kinds]
        total_fmt = [self.st.cell(k, total=True) for k in kinds]
        r = first
        for i, row in enumerate(iter_frame_rows(t.df)):
            r += 1
            is_total = t.total_last and i == n - 1
            for j in pp_cols:
                row[j] = round_pp(row[j])
            for j, v in enumerate(row):
                fmt = total_fmt[j] if is_total else (cell_fmt[j] if small else self._col_formats.get(j))
                if v is None:
                    if j in dash:
                        ws.write_string(r, j, "-", fmt)
                    elif is_total or small:
                        ws.write_blank(r, j, None, fmt)
                    continue
                ws.write(r, j, v, fmt)
        last = r
        if not self._has_autofilter.get(ws.get_name()):  # one autofilter per worksheet
            ws.autofilter(first, 0, last, len(cols) - 1)
            self._has_autofilter[ws.get_name()] = True
        self._body_rules(ws, t, cols, first, last - 1 if t.total_last else last)
        if small and last > first:  # banded rows, lowest priority
            ws.conditional_format(first + 1, 0, last - 1 if t.total_last else last, len(cols) - 1, {
                "type": "formula", "criteria": "=MOD(ROW(),2)=0", "format": self.st.get(bg_color="#F3F6FA")})
        if freeze:
            ws.freeze_panes(first + 1, 0)
        return last + 1

    def _rate_rules(self, ws, r0: int, r1: int, j: int) -> None:
        """Same colour key as the app: red < 30%, amber 30–60%, green > 60% (text '-' untouched)."""
        ws.conditional_format(r0, j, r1, j, {"type": "cell", "criteria": "between", "minimum": 0,
                                             "maximum": 0.2999999, "format": self.st.red()})
        ws.conditional_format(r0, j, r1, j, {"type": "cell", "criteria": "between", "minimum": 0.3,
                                             "maximum": 0.6, "format": self.st.amber()})
        ws.conditional_format(r0, j, r1, j, {"type": "cell", "criteria": "between", "minimum": 0.6000001,
                                             "maximum": 1, "format": self.st.green()})


COLOR_KEY = ("رنگ‌ها: وین‌ریت رقابتی قرمز = کمتر از ۳۰٪، زرد = ۳۰ تا ۶۰٪، سبز = بیشتر از ۶۰٪. "
             "تغییر: سبز = حداقل ۵+ واحد درصد، قرمز = حداقل ۵− واحد درصد. "
             "روی سرستون‌ها مکث کنید تا توضیح هر ستون نمایش داده شود.")
