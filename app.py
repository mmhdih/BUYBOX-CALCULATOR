"""Streamlit UI for the Buy Box Analyzer (UI only; logic lives in the ``buybox`` package)."""

from __future__ import annotations

import dataclasses
import hashlib
import json
import os
import re
import sys
import threading
from datetime import date
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent
if str(APP_DIR) not in sys.path:  # needed for the portable (embedded Python) build
    sys.path.insert(0, str(APP_DIR))

import pandas as pd  # noqa: E402
import plotly.express as px  # noqa: E402
import streamlit as st  # noqa: E402

from buybox import compare, history  # noqa: E402
from buybox import labels as L  # noqa: E402
from buybox.core import (  # noqa: E402
    Params,
    market_summary,
    rate_by_dimension,
    seller_metrics,
    single_seller_dkps,
)
from buybox.export import (  # noqa: E402
    build_main_workbook,
    build_quickwin_workbook,
    quickwin_filename,
    safe_filename,
)
from buybox.io import (  # noqa: E402
    OPTIONAL_COLUMNS,
    REQUIRED_COLUMNS,
    InputFileError,
    MappingResult,
    detect_columns,
    merge_kam,
    read_kam_list,
    read_table_bytes,
)
from buybox.pipeline import Analysis, analyze, attach_comparison, filter_mask, kam_options  # noqa: E402
from buybox.quality import STATUS_OK, STATUS_WARN, checks_to_frame, run_checks  # noqa: E402
from buybox.report import ReportContext, build_report, report_filename  # noqa: E402

DATA_DIR = APP_DIR / "data"
RUNS_DIR = DATA_DIR / "runs"
SETTINGS_FILE = DATA_DIR / "settings.json"
KAM_SAVED = DATA_DIR / "kam_list_last.xlsx"
UPLOAD_TYPES = ["xlsx", "csv"]
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

RED, AMBER, GREEN = "#D64545", "#E8A33D", "#3A9D5D"
BG_RED, BG_AMBER, BG_GREEN = "#FFC7CE", "#FFEB9C", "#C6EFCE"

CANON_LABELS = {
    "DKPC": "DKPC (شناسه تنوع)",
    "DKP": "DKP (شناسه کالا)",
    "Seller_ID": "Seller_ID (شناسه فروشنده)",
    "Weight": "Weight (وزن - گرم)",
    "Selling_Price": "Selling_Price (قیمت فروش - ریال)",
    "Seller_Name": "Seller_Name (نام فروشنده)",
    "Product_Name": "Product_Name (نام کالا)",
    "Brand_Name_En": "Brand_Name_En (برند)",
    "Leaf_Category_Name": "Leaf_Category_Name (دسته‌بندی)",
    "True_Wage": "True_Wage (اجرت واقعی، برای هشدار حاشیه)",
    "Actual_BuyBox": L.MAPPING_ACTUAL_BB,
}

ROW_VIEW_COLUMNS = ["DKPC", "DKP", "Product_Name", "Weight", "Selling_Price", "BuyBox_Min_Price",
                    "BuyBox_Winner_Seller", "Gap_to_BuyBox", "Target_Price", "Required_Cut_%", "Margin_Flag"]


# ---------------------------------------------------------------------------
# Settings (remembered on disk, next to the app — portable friendly)
# ---------------------------------------------------------------------------
def load_settings() -> dict:
    try:
        return json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}


def save_settings(settings: dict) -> None:
    try:
        DATA_DIR.mkdir(parents=True, exist_ok=True)
        SETTINGS_FILE.write_text(json.dumps(settings, ensure_ascii=False, indent=1), encoding="utf-8")
    except OSError:
        st.toast("ذخیره تنظیمات ممکن نشد (دسترسی نوشتن روی پوشه data وجود ندارد).")


# ---------------------------------------------------------------------------
# Cached computation (keyed on file bytes + parameters)
# ---------------------------------------------------------------------------
# Large files (up to 500 MB) are cached by content hash: the bytes themselves are passed in
# underscore-prefixed arguments, which Streamlit does not hash, and results are kept with
# st.cache_resource (shared, not copied on every rerun). Cached frames are never mutated.
@st.cache_resource(show_spinner="در حال خواندن فایل…", max_entries=4)
def cached_read(key: str, name: str, _src) -> pd.DataFrame:
    """_src: an uploaded file (bytes read only on a cache miss) or bytes."""
    return read_table_bytes(_src.getvalue() if hasattr(_src, "getvalue") else _src, name)


@st.cache_data(show_spinner="در حال خواندن لیست KAM…", max_entries=4)
def cached_kam(key: str, name: str, _data: bytes) -> tuple[pd.DataFrame, list[str]]:
    kam, warnings = read_kam_list(_data, name)
    return merge_kam(kam), warnings


@st.cache_resource(show_spinner="در حال محاسبه بای باکس…", max_entries=4)
def cached_analysis(key: str, name: str, mapping_items: tuple, kam_key: str, params: tuple[float, float, float],
                    _src, _kam: tuple[str, bytes] | None) -> Analysis:
    raw = cached_read(key, name, _src)
    kam = cached_kam(kam_key, _kam[0], _kam[1])[0] if _kam else None
    return analyze(raw, dict(mapping_items), kam, Params(*params), name)


@st.cache_resource(show_spinner="در حال بارگذاری اجرای ذخیره‌شده…", max_entries=4)
def cached_run(path: str, mtime: float) -> pd.DataFrame:  # mtime only busts the cache
    return history.load_run(history.RunInfo(Path(path), {}))


def upload_key(f) -> str:
    """Content hash of an uploaded file, computed once per upload (not on every rerun)."""
    memo = st.session_state.setdefault("_upload_hashes", {})
    fid = getattr(f, "file_id", None) or f"{f.name}:{f.size}"
    if fid not in memo:
        memo[fid] = file_hash(f.getvalue())
    return memo[fid]


def file_hash(data: bytes) -> str:
    return hashlib.sha1(data).hexdigest()


# ---------------------------------------------------------------------------
# Formatting helpers
# ---------------------------------------------------------------------------
def fmt_int(v: object) -> str:
    return "-" if v is None or pd.isna(v) else f"{float(v):,.0f}"


def fmt_pct(v: object) -> str:
    return "-" if v is None or pd.isna(v) else f"{float(v) * 100:.1f}%"


def fmt_pp(v: object) -> str:
    if v is None or pd.isna(v):
        return "-"
    return "0.0 pp" if abs(float(v)) < 0.05 else f"{float(v):+.1f} pp"


def delta_args(pp: object) -> dict:
    """st.metric delta kwargs for a percentage-point change (neutral colour when ~0)."""
    if pp is None or pd.isna(pp):
        return {}
    if abs(float(pp)) < 0.05:
        return {"delta": fmt_pp(pp), "delta_color": "off", "delta_arrow": "off"}
    return {"delta": fmt_pp(pp), "delta_color": "normal"}


def fmt_weight(v: object) -> str:
    return "-" if v is None or pd.isna(v) else f"{float(v):g}"


_NUM_RE = re.compile(r"(?<![\w\-−.,٫/:])([+\-−]?[0-9][0-9,٬]*(?:[.٫][0-9]+)?(?:\s?(?:%|٪|pp))?)(?![\w\-/:])")


def bidi(text: str) -> str:
    """Wrap numbers (incl. % / pp) in LTR isolates so they read correctly inside Persian text."""
    return _NUM_RE.sub("\u2066\\1\u2069", text)


def rate_bg(v: object) -> str:
    if v is None or pd.isna(v):
        return ""
    return f"background-color: {BG_RED if v < 0.3 else BG_AMBER if v <= 0.6 else BG_GREEN}"


def change_bg(v: object) -> str:
    if v is None or pd.isna(v):
        return ""
    return f"background-color: {BG_GREEN}" if v >= 5 else f"background-color: {BG_RED}" if v <= -5 else ""


def _fmt_id(v: object) -> str:
    if v is None or pd.isna(v):
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def _fmt_auto(v: object) -> tuple[str, bool]:
    """Format an unspecified value; returns (text, is_number)."""
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return "-", True
    if isinstance(v, bool):
        return ("✓" if v else ""), False
    if isinstance(v, (int,)) or (hasattr(v, "dtype") and pd.api.types.is_integer_dtype(getattr(v, "dtype", None))):
        return f"{int(v):,}", True
    if isinstance(v, float):
        return (f"{v:,.0f}" if v.is_integer() else f"{v:,.4g}"), True
    return str(v), False


def rtl_table(df: pd.DataFrame, pct=(), ints=(), pp=(), weight=(), ratio=(), rate_cols=(), change_cols=(),
              ids=(), height: int | None = None, max_rows: int = 2000, total_mask: pd.Series | None = None,
              wrap_cols=()) -> None:
    """Render a read-only right-to-left HTML table.

    Columns start on the right, the header and the first column are sticky, numbers are
    isolated left-to-right (so % and − sit correctly) and cells can be colour-coded.
    """
    import html as _html

    if df.empty:
        st.caption("— بدون ردیف —")
        return
    shown = df.head(max_rows)
    tot = (total_mask.reindex(shown.index).fillna(False).to_numpy() if total_mask is not None
           else [False] * len(shown))
    fixed: dict[str, object] = {}
    for c in shown.columns:
        if c in pct:
            fixed[c] = fmt_pct
        elif c in ints:
            fixed[c] = fmt_int
        elif c in pp:
            fixed[c] = fmt_pp
        elif c in weight:
            fixed[c] = fmt_weight
        elif c in ratio:
            fixed[c] = lambda v: "-" if v is None or pd.isna(v) else f"{v:.2f}"
        elif c in ids:
            fixed[c] = _fmt_id
    cols_html: list[list[str]] = []
    for c in shown.columns:
        vals = shown[c].tolist()
        f = fixed.get(c)
        bg = rate_bg if c in rate_cols else change_bg if c in change_cols else None
        cells = []
        for v in vals:
            if f is not None:
                text, num = f(v), True
            else:
                text, num = _fmt_auto(v)
            cls = "n" if num else ("w" if c in wrap_cols else "t")
            style = ""
            if bg is not None:
                css = bg(v)
                style = f' style="{css}"' if css else ""
            esc = _html.escape(text)
            title = f' title="{esc}"' if cls == "t" and len(text) > 40 else ""
            cells.append(f'<td class="{cls}"{style}{title}>{esc}</td>')
        cols_html.append(cells)
    head = "".join(f"<th>{_html.escape(str(c))}</th>" for c in shown.columns)
    body = []
    for i in range(len(shown)):
        row_cls = ' class="tot"' if tot[i] else ""
        body.append(f"<tr{row_cls}>" + "".join(col[i] for col in cols_html) + "</tr>")
    h = f"max-height:{height}px;" if height else ""
    st.markdown(
        f'<div class="bb-tw" style="{h}"><table class="bb-t"><thead><tr>{head}</tr></thead>'
        f'<tbody>{"".join(body)}</tbody></table></div>',
        unsafe_allow_html=True,
    )
    if len(df) > max_rows:
        st.caption(bidi(f"فقط {fmt_int(max_rows)} ردیف اول از {fmt_int(len(df))} ردیف نمایش داده شد؛ "
                        "همه ردیف‌ها در فایل اکسل خروجی هست."))


def row_table(rows: pd.DataFrame, extra: tuple[str, ...] = (), height: int | None = None) -> None:
    cols = [c for c in (*ROW_VIEW_COLUMNS, *extra) if c in rows.columns]
    rtl_table(rows[cols], pct=("Gap_to_BuyBox", "Required_Cut_%"),
              ints=("Selling_Price", "BuyBox_Min_Price", "Target_Price"), weight=("Weight",),
              ids=("DKPC", "DKP", "Seller_ID"), height=height)


def inject_css() -> None:
    st.markdown(
        """
<style>
@import url('https://fonts.googleapis.com/css2?family=Vazirmatn:wght@300;400;600;800&display=swap');
@font-face { font-family: 'VazirmatnLocal'; font-weight: 100 900; font-display: swap;
  src: url('app/static/fonts/vazirmatn-arabic.woff2') format('woff2');
  unicode-range: U+0600-06FF, U+0750-077F, U+200C-200E, U+2010-2011, U+204F, U+FB50-FDFF, U+FE70-FEFC; }
@font-face { font-family: 'VazirmatnLocal'; font-weight: 100 900; font-display: swap;
  src: url('app/static/fonts/vazirmatn-latin.woff2') format('woff2');
  unicode-range: U+0000-00FF, U+0131, U+0152-0153, U+02BB-02BC, U+02C6, U+02DA, U+02DC, U+2000-206F, U+20AC; }
html, body, .stApp, .stMarkdown, p, li, h1, h2, h3, h4, h5, h6, label, input, textarea, select,
button, .stTabs [data-baseweb="tab"], [data-testid="stMetricLabel"], [data-testid="stMetricValue"],
[data-testid="stCaptionContainer"], [data-testid="stAlert"] {
  font-family: 'Vazirmatn', 'VazirmatnLocal', Tahoma, sans-serif !important;
}
[data-testid="stIconMaterial"], .material-symbols-rounded, span[class*="material"] {
  font-family: 'Material Symbols Rounded' !important;
}
/* ---------- right-to-left layout ---------- */
.stApp, [data-testid="stSidebar"], [data-testid="stMain"], [data-testid="stBottom"] { direction: rtl; }
/* Streamlit sets text-align:left on these; force right */
.stApp h1, .stApp h2, .stApp h3, .stApp h4, .stApp h5, .stApp h6,
.stApp [data-testid="stHeadingWithActionElements"],
.stApp [data-testid="stMarkdownContainer"], .stApp [data-testid="stMarkdownContainer"] p,
.stApp [data-testid="stMarkdownContainer"] li, .stApp [data-testid="stCaptionContainer"],
.stApp [data-testid="stCaptionContainer"] p, .stApp label, .stApp [data-testid="stWidgetLabel"],
.stApp [data-testid="stWidgetLabel"] p, .stApp [data-testid="stAlertContainer"] p,
.stApp [data-testid="stExpander"] summary, .stApp [data-testid="stMetricLabel"],
.stApp [data-testid="stMetricLabel"] p, .stApp [data-testid="stTooltipContent"] {
  text-align: right !important; direction: rtl;
}
.stApp [data-testid="stHeadingWithActionElements"] > * { justify-content: flex-start; }
.stApp [data-testid="stMarkdownContainer"] ul, .stApp [data-testid="stMarkdownContainer"] ol {
  padding-right: 1.3rem !important; padding-left: 0 !important; margin-right: 0 !important;
}
/* metric value + delta: numbers LTR but placed on the right */
.stApp [data-testid="stMetricValue"], .stApp [data-testid="stMetricValue"] > div {
  direction: ltr; text-align: right !important; justify-content: flex-end !important;
}
.stApp [data-testid="stMetric"] { text-align: right; }
.stApp [data-testid="stMetricDelta"] { direction: ltr; justify-content: flex-end; margin-left: auto; width: fit-content; }
.stApp [data-testid="stMetricDelta"], .stApp [data-testid="stMetricDelta"] * { direction: ltr !important; }
/* file uploader: Persian texts instead of Streamlit's English ones */
.stApp [data-testid="stFileUploaderDropzoneInstructions"] * { display: none !important; }
.stApp [data-testid="stFileUploaderDropzoneInstructions"]::after {
  content: "xlsx یا csv — حداکثر ۵۰۰ مگابایت"; font-size: 0.8rem; color: #6B7280; }
.stApp [data-testid="stFileUploaderDropzone"] button[data-testid="stBaseButton-secondary"] > * { display: none !important; }
.stApp [data-testid="stFileUploaderDropzone"] button[data-testid="stBaseButton-secondary"]::after {
  content: "انتخاب فایل"; font-size: 0.9rem; }
/* sliders and number boxes keep LTR mechanics (min on the left, typed digits LTR) */
.stApp [data-testid="stSlider"] [data-baseweb="slider"], .stApp [data-testid="stSliderThumbValue"],
.stApp [data-testid="stSliderTickBar"] { direction: ltr; }
.stApp [data-testid="stNumberInput"] input { direction: ltr; text-align: right; }
.stApp [data-testid="stNumberInputContainer"] { direction: ltr; }
.js-plotly-plot, code, pre { direction: ltr; text-align: left; }
.stTabs [data-baseweb="tab-list"] { gap: 0.25rem; flex-wrap: wrap; }
div[data-testid="stExpander"] summary p { font-weight: 600; }
/* ---------- RTL tables ---------- */
.bb-tw { overflow: auto; direction: rtl; border: 1px solid #E3E6EC; border-radius: 8px; margin: 0.25rem 0 0.75rem; }
.bb-t { border-collapse: separate; border-spacing: 0; width: 100%; font-size: 0.84rem; direction: rtl; }
.bb-t th { position: sticky; top: 0; z-index: 2; background: #1F4E78; color: #fff; font-weight: 600;
  padding: 7px 10px; text-align: right; white-space: nowrap; }
.bb-t td { padding: 5px 10px; border-bottom: 1px solid #EEF0F4; text-align: right; white-space: nowrap;
  background: #fff; }
.bb-t tbody tr:nth-child(even) td { background: #F8F9FB; }
.bb-t tbody tr:hover td { background: #EEF4FB; }
.bb-t td.n { direction: ltr; unicode-bidi: isolate; text-align: right; font-variant-numeric: tabular-nums; }
.bb-t td.t { max-width: 340px; overflow: hidden; text-overflow: ellipsis; }
.bb-t td.w { white-space: normal; min-width: 160px; }
.bb-t th:first-child, .bb-t td:first-child { position: sticky; right: 0; z-index: 1; box-shadow: -1px 0 0 #E3E6EC; }
.bb-t th:first-child { z-index: 3; }
.bb-t tr.tot td { font-weight: 700; background: #D9E1F2 !important; border-top: 2px solid #1F4E78; }
.bb-chip { display:inline-block; padding:0.1rem 0.55rem; border-radius:999px; font-size:0.8rem; margin-left:0.3rem; }
.bb-ok { background:#C6EFCE; color:#006100; } .bb-warn { background:#FFEB9C; color:#9C5700; }
.bb-info { background:#DDEBF7; color:#1F4E78; }
</style>
""",
        unsafe_allow_html=True,
    )


# ---------------------------------------------------------------------------
# Column mapping UI
# ---------------------------------------------------------------------------
def mapping_ui(columns: list[str], detected: MappingResult, key_prefix: str, title: str) -> dict | None:
    """Selectboxes to map canonical columns; returns the mapping or None if incomplete."""
    sig = hashlib.md5("|".join(columns).encode()).hexdigest()[:8]
    opts = [L.MAPPING_NONE, *columns]
    mapping: dict[str, str | None] = {}
    with st.expander(f"{L.MAPPING_TITLE} — {title}", expanded=not detected.ok):
        if not detected.ok:
            st.warning(L.MAPPING_HELP)
        grid = st.columns(3)
        for i, canon in enumerate((*REQUIRED_COLUMNS, *OPTIONAL_COLUMNS)):
            cur = detected.mapping.get(canon)
            label = CANON_LABELS[canon] + (" *" if canon in REQUIRED_COLUMNS else "")
            sel = grid[i % 3].selectbox(label, opts, index=opts.index(cur) if cur in opts else 0,
                                        key=f"map_{key_prefix}_{canon}_{sig}")
            mapping[canon] = None if sel == L.MAPPING_NONE else sel
        st.caption("علامت \\* یعنی ستون الزامی")
    missing = [c for c in REQUIRED_COLUMNS if not mapping.get(c)]
    chosen = [v for v in mapping.values() if v]
    if len(chosen) != len(set(chosen)):
        st.warning("یک ستون برای بیش از یک مورد انتخاب شده است.")
    if missing:
        st.error("ستون‌های الزامی تعیین نشده‌اند: " + "، ".join(missing))
        return None
    return mapping


# ---------------------------------------------------------------------------
# Sections
# ---------------------------------------------------------------------------
def kpi_cards(s: dict, mk: dict | None) -> None:
    c = st.columns(6)
    c[0].metric(L.KPI_EVALUATED, fmt_int(s["evaluated"]))
    c[1].metric(L.KPI_COMPETITIVE, fmt_int(s["comp_rows"]))
    c[2].metric(L.KPI_COMP_WIN_RATE, fmt_pct(s["comp_win_rate"]),
                **(delta_args(mk["comp_win_rate_change_pp"]) if mk is not None else {}))
    c[3].metric(L.KPI_NEAR_MISS, fmt_int(s["near_miss"]))
    c[4].metric(L.KPI_GAINED, fmt_int(mk["gained"]) if mk else "-")
    c[5].metric(L.KPI_LOST, fmt_int(mk["lost"]) if mk else "-")


def seller_bar_chart(table: pd.DataFrame) -> None:
    t = table[~table["Is_Total"] & (table["Comp_Variants"] > 0)]
    t = t.sort_values("Comp_Variants", ascending=False).head(20).copy()
    if t.empty:
        st.info("هیچ تنوع رقابتی در داده فیلترشده وجود ندارد.")
        return
    t["band"] = pd.cut(t["Comp_Win_Rate"], [-0.01, 0.2999999, 0.6, 1.01], labels=["<30%", "30–60%", ">60%"])
    t["label"] = t["Seller_Name"].astype(str) + " (" + t["Comp_Variants"].map(fmt_int) + ")"
    t = t.iloc[::-1]
    fig = px.bar(t, x="Comp_Win_Rate", y="label", orientation="h", color="band",
                 color_discrete_map={"<30%": RED, "30–60%": AMBER, ">60%": GREEN},
                 text=t["Comp_Win_Rate"].map(fmt_pct),
                 category_orders={"band": ["<30%", "30–60%", ">60%"]})
    # right-to-left: bars grow from the right, seller names on the right axis
    fig.update_layout(height=max(360, 26 * len(t) + 80), margin=dict(l=10, r=10, t=40, b=10),
                      xaxis=dict(tickformat=".0%", range=[1.12, 0], title="وین‌ریت رقابتی"),
                      yaxis=dict(title="", side="right", categoryorder="array", categoryarray=list(t["label"])),
                      legend=dict(orientation="h", x=1, xanchor="right", y=1.02, yanchor="bottom",
                                  title_text="", traceorder="normal"),
                      font=dict(family="Vazirmatn, VazirmatnLocal, Tahoma"))
    fig.update_traces(textposition="outside", cliponaxis=False)
    st.plotly_chart(fig, width="stretch")
    st.caption("۲۰ فروشنده با بیشترین تنوع در رقابت (عدد داخل پرانتز). قرمز < ۳۰٪، کهربایی ۳۰ تا ۶۰٪، سبز > ۶۰٪.")


def summary_bullets(s: dict, table: pd.DataFrame, mk: dict | None, changes: pd.DataFrame | None,
                    seller_info: pd.DataFrame, params: Params) -> list[str]:
    """3–5 bullets built strictly from the computed numbers."""
    b: list[str] = []
    line = (f"وین‌ریت رقابتی {fmt_pct(s['comp_win_rate'])} است ({fmt_int(s['comp_wins'])} بای باکس از "
            f"{fmt_int(s['comp_rows'])} تنوع در رقابت).")
    if mk is not None and not pd.isna(mk["prev_comp_win_rate"]):
        line += (f" دوره قبل {fmt_pct(mk['prev_comp_win_rate'])} بود؛ تغییر "
                 f"{fmt_pp(mk['comp_win_rate_change_pp'])}.")
    b.append(line)
    if mk is not None and changes is not None and not changes.empty:
        b.append(f"{fmt_int(mk['gained'])} بای باکس گرفته شد و {fmt_int(mk['lost'])} از دست رفت؛ "
                 f"{fmt_int(mk['new'])} تنوع جدید و {fmt_int(mk['removed'])} تنوع حذف‌شده.")
        winners, losers = compare.net_movers(changes, seller_info, top=1)
        if not winners.empty:
            w = winners.iloc[0]
            b.append(f"بیشترین رشد خالص: {w.Seller_Name} — گرفته {fmt_int(w.Gained)}، از دست رفته "
                     f"{fmt_int(w.Lost)}، خالص {int(w.Net):+d}.")
        if not losers.empty:
            lo = losers.iloc[0]
            b.append(f"بیشترین افت خالص: {lo.Seller_Name} — گرفته {fmt_int(lo.Gained)}، از دست رفته "
                     f"{fmt_int(lo.Lost)}، خالص {int(lo.Net):+d}.")
    data = table[~table["Is_Total"]]
    if not data.empty and data["Near_Miss"].max() > 0:
        top = data.sort_values("Near_Miss", ascending=False).iloc[0]
        b.append(f"بیشترین باخت نزدیک (فاصله ≤ {params.near_miss_threshold * 100:.1f}٪): "
                 f"{top.Seller_Name} با {fmt_int(top.Near_Miss)} تنوع — از تب Quick Win فایل بفرستید.")
    if len(b) < 5 and s["margin_flags"]:
        b.append(f"{fmt_int(s['margin_flags'])} تنوع بازنده برای رسیدن به قیمت هدف بیش از "
                 f"{params.max_acceptable_cut * 100:.1f}٪ کاهش قیمت لازم دارند (هشدار حاشیه).")
    if len(b) < 5 and not pd.isna(s["single_seller_group_pct"]):
        b.append(f"{fmt_pct(s['single_seller_group_pct'])} از گروه‌های DKP + وزن فقط یک فروشنده دارند.")
    return b[:5]


def trend_section(runs: list[history.RunInfo], sellers: list[object]) -> None:
    st.subheader("روند")
    tr = history.trend_frame(runs, sellers or None)
    if len(tr) < 1:
        st.caption("هنوز اجرای ذخیره‌شده‌ای برای نمایش روند وجود ندارد.")
        return
    title = "وین‌ریت رقابتی تجمیعی" + (" (فروشندگان انتخاب‌شده)" if sellers else " (کل بازار)")
    fig = px.line(tr, x="created_at", y="comp_win_rate", markers=True, hover_data=["run", "comp_rows"])
    fig.update_layout(height=300, margin=dict(l=10, r=10, t=40, b=10), title=title,
                      yaxis=dict(tickformat=".0%", title=""), xaxis=dict(title=""),
                      font=dict(family="Vazirmatn, VazirmatnLocal, Tahoma"))
    st.plotly_chart(fig, width="stretch")
    st.caption("روند بر اساس اجراهای ذخیره‌شده است (فیلتر KAM/دسته/برند روی روند اعمال نمی‌شود؛ فیلتر فروشنده اعمال می‌شود).")


def quality_section(checks, compact: bool) -> None:
    icons = {STATUS_OK: ("bb-ok", "✓"), STATUS_WARN: ("bb-warn", "!"), "info": ("bb-info", "i")}
    for chk in checks:
        cls, ic = icons[chk.status]
        note = f" — <small>{chk.note}</small>" if chk.note else ""
        line = f"<span class='bb-chip {cls}'>{ic}</span> {chk.label}: <b dir='ltr'>{chk.display}</b>{note}"
        if compact or chk.rows is None or chk.rows.empty:
            st.markdown(line, unsafe_allow_html=True)
        else:
            st.markdown(line, unsafe_allow_html=True)
            with st.expander(f"مشاهده ردیف‌ها — {chk.label}"):
                rtl_table(chk.rows.head(5000), ints=("Selling_Price", "BuyBox_Min_Price"),
                                    weight=("Weight",), ids=("DKPC", "DKP", "Seller_ID", "Group_ID"))


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main() -> None:
    st.set_page_config(layout="wide", page_title=L.APP_TITLE, page_icon="🪙")
    inject_css()
    settings = load_settings()
    runs = history.list_runs(RUNS_DIR)

    sb = st.sidebar
    sb.markdown(f"### {L.APP_TITLE}")
    sb.caption(L.APP_SUBTITLE)

    # ---- Files ---------------------------------------------------------------
    sb.subheader(L.SIDEBAR_FILES)
    cur_file = sb.file_uploader(L.UPLOAD_CURRENT, type=UPLOAD_TYPES, key="up_cur")
    kam_file = sb.file_uploader(L.UPLOAD_KAM, type=UPLOAD_TYPES, key="up_kam")
    kam_bytes: bytes | None = None
    kam_name = ""  # empty name -> format detected from content (remembered list may be xlsx or csv)
    if kam_file is not None:
        kam_bytes = kam_file.getvalue()
        kam_name = kam_file.name
        try:
            if not KAM_SAVED.exists() or KAM_SAVED.read_bytes() != kam_bytes:
                DATA_DIR.mkdir(parents=True, exist_ok=True)
                KAM_SAVED.write_bytes(kam_bytes)
        except OSError:
            sb.warning("ذخیره لیست KAM روی دیسک ممکن نشد.")
    elif KAM_SAVED.exists():
        kam_bytes = KAM_SAVED.read_bytes()
        sb.caption(L.KAM_SAVED_INFO)
        if sb.button(L.KAM_SAVED_FORGET, key="forget_kam"):
            KAM_SAVED.unlink(missing_ok=True)
            st.rerun()

    cur_hash = upload_key(cur_file) if cur_file is not None else ""
    usable_runs = [r for r in runs if r.meta.get("file_hash") != cur_hash]
    prev_options = [L.PREV_NONE, L.PREV_FILE, L.PREV_RUN]
    prev_kind = sb.radio(L.PREV_SOURCE, prev_options, index=2 if usable_runs else 0, key="prev_kind")
    prev_file = None
    prev_run: history.RunInfo | None = None
    if prev_kind == L.PREV_FILE:
        prev_file = sb.file_uploader(L.UPLOAD_PREV, type=UPLOAD_TYPES, key="up_prev")
    elif prev_kind == L.PREV_RUN:
        if usable_runs:
            by_path = {str(r.parquet): r for r in usable_runs}
            picked = sb.selectbox(L.SELECT_RUN, list(by_path), format_func=lambda p: by_path[p].label,
                                  key="sel_run")
            prev_run = by_path.get(picked)
        else:
            sb.caption("هنوز اجرای ذخیره‌شده‌ای (غیر از فایل فعلی) وجود ندارد.")

    # ---- Parameters ----------------------------------------------------------
    sb.subheader(L.SIDEBAR_PARAMS)
    thr = sb.slider(L.PARAM_NEAR_MISS, 0.1, 10.0, 1.0, 0.1, key="p_thr")
    undercut = sb.number_input(L.PARAM_UNDERCUT, min_value=0, value=0, step=10_000, key="p_under",
                               help=L.PARAM_UNDERCUT_HELP)
    sb.caption(L.PARAM_UNDERCUT_HELP)
    max_cut = sb.number_input(L.PARAM_MAX_CUT, min_value=0.0, max_value=100.0, value=1.0, step=0.1,
                              format="%.1f", key="p_maxcut")
    params = Params(near_miss_threshold=thr / 100, undercut=float(undercut), max_acceptable_cut=max_cut / 100)

    filters_box = sb.container()

    with sb.expander(L.SIDEBAR_SETTINGS):
        my_kam = st.text_input(L.SETTINGS_MY_KAM, value=settings.get("my_kam", ""), key="set_kam")
        autosave = st.checkbox(L.SETTINGS_AUTOSAVE, value=settings.get("autosave", True), key="set_auto")
        if my_kam != settings.get("my_kam", "") or autosave != settings.get("autosave", True):
            settings.update(my_kam=my_kam, autosave=autosave)
            save_settings(settings)
        st.caption(f"محل ذخیره داده‌ها: {DATA_DIR}")
        if runs:
            st.markdown("**تاریخچه اجراها**")
            run_by_path = {str(r.parquet): r for r in runs}
            to_del = st.selectbox("حذف اجرا", ["", *run_by_path],
                                  format_func=lambda p: "—" if not p else run_by_path[p].label, key="del_run")
            if to_del and st.button("حذف اجرای انتخاب‌شده", key="del_btn"):
                history.delete_run(run_by_path[to_del])
                st.rerun()

    if os.environ.get("BUYBOX_LAUNCHER"):  # started by BuyBox.exe / launch.py
        if sb.button(L.SHUTDOWN, key="shutdown", icon=":material/power_settings_new:", width="stretch"):
            sb.success(L.SHUTDOWN_DONE)
            st.success(L.SHUTDOWN_DONE)
            threading.Timer(1.5, os._exit, args=(0,)).start()
            st.stop()

    # ---- Header --------------------------------------------------------------
    st.title(L.APP_TITLE)
    if cur_file is None:
        st.info(L.INFO_UPLOAD_FIRST)
        st.caption("داده‌ها فقط روی همین کامپیوتر پردازش می‌شوند و به هیچ سرویس بیرونی ارسال نمی‌شوند.")
        if runs:
            trend_section(runs, [])
        return

    # ---- Read + map current file --------------------------------------------
    try:
        raw = cached_read(cur_hash, cur_file.name, cur_file)
    except InputFileError as exc:
        st.error(str(exc))
        return
    cur_name = cur_file.name
    mapping = mapping_ui([str(c) for c in raw.columns], detect_columns(list(raw.columns)), "cur", cur_name)
    if mapping is None:
        return

    kam_merged = None
    kam_key = file_hash(kam_bytes) if kam_bytes else ""
    if kam_bytes:
        try:
            kam_merged, kam_warnings = cached_kam(kam_key, kam_name, kam_bytes)
            for w in kam_warnings:
                st.warning(w)
        except InputFileError as exc:
            st.error("لیست KAM: " + str(exc))
            kam_bytes = None

    params_t = (params.near_miss_threshold, params.undercut, params.max_acceptable_cut)
    try:
        analysis = cached_analysis(cur_hash, cur_name, tuple(sorted(mapping.items())), kam_key, params_t,
                                   cur_file, (kam_name, kam_bytes) if kam_bytes else None)
        # cached object is shared: work on a shallow copy so adding comparison columns never leaks
        analysis = dataclasses.replace(analysis, rows=analysis.rows.copy(deep=False))
    except Exception as exc:  # noqa: BLE001
        st.error(f"خطا در پردازش فایل: {exc}")
        return

    # ---- Previous period ----------------------------------------------------
    comparison = None
    prev_columns: list[str] | None = None
    if prev_file is not None:
        try:
            prev_key = upload_key(prev_file)
            prev_raw = cached_read(prev_key, prev_file.name, prev_file)
            pmap = mapping_ui([str(c) for c in prev_raw.columns], detect_columns(list(prev_raw.columns)),
                              "prev", prev_file.name)
            if pmap is not None:
                prev_an = cached_analysis(prev_key, prev_file.name, tuple(sorted(pmap.items())), kam_key,
                                          params_t, prev_file,
                                          (kam_name, kam_bytes) if kam_bytes else None)
                prev_columns = prev_an.original_columns
                comparison = attach_comparison(analysis, prev_an.rows, prev_file.name, prev_columns)
        except InputFileError as exc:
            st.error("فایل دوره قبل: " + str(exc))
    elif prev_run is not None:
        prev_rows = cached_run(str(prev_run.parquet), prev_run.parquet.stat().st_mtime)
        prev_columns = prev_run.meta.get("columns")
        comparison = attach_comparison(analysis, prev_rows, prev_run.label, prev_columns)
    has_prev = comparison is not None

    # ---- Save run -----------------------------------------------------------
    if autosave and history.find_run(RUNS_DIR, cur_hash) is None:
        try:
            history.save_run(analysis.rows, params, cur_name, RUNS_DIR, file_hash=cur_hash,
                             original_columns=analysis.original_columns)
            st.toast("این اجرا در تاریخچه ذخیره شد.")
            runs = history.list_runs(RUNS_DIR)
        except OSError as exc:
            st.warning(f"ذخیره اجرا ممکن نشد: {exc}")

    rows_all = analysis.rows
    info = analysis.seller_info

    # ---- Filters ------------------------------------------------------------
    with filters_box:
        st.subheader(L.SIDEBAR_FILTERS)
        k_opts = kam_options(info)
        if "f_kam" not in st.session_state:
            st.session_state["f_kam"] = [my_kam] if my_kam and my_kam in k_opts else []
        st.session_state["f_kam"] = [k for k in st.session_state["f_kam"] if k in k_opts]
        f_kam = st.multiselect(L.FILTER_KAM, k_opts, key="f_kam")
        f_cat = f_brand = []
        if "Leaf_Category_Name" in rows_all.columns:
            cats = sorted(rows_all["Leaf_Category_Name"].dropna().unique().tolist())
            st.session_state["f_cat"] = [c for c in st.session_state.get("f_cat", []) if c in cats]
            f_cat = st.multiselect(L.FILTER_CATEGORY, cats, key="f_cat")
        if "Brand_Name_En" in rows_all.columns:
            brands = sorted(rows_all["Brand_Name_En"].dropna().unique().tolist())
            st.session_state["f_brand"] = [b for b in st.session_state.get("f_brand", []) if b in brands]
            f_brand = st.multiselect(L.FILTER_BRAND, brands, key="f_brand")
        seller_ids = info.index.tolist()
        st.session_state["f_seller"] = [s for s in st.session_state.get("f_seller", []) if s in seller_ids]
        f_seller = st.multiselect(L.FILTER_SELLER, seller_ids, key="f_seller",
                                  format_func=lambda s: f"{info.at[s, 'Seller_Name']} ({s})")

    mask = filter_mask(rows_all, info, f_kam, f_cat, f_brand, f_seller)
    rows = rows_all[mask]
    filtered = not bool(mask.all())
    if filtered:
        st.caption(f"فیلتر فعال: {fmt_int(len(rows))} از {fmt_int(len(rows_all))} تنوع.")
    if rows.empty:
        st.warning("با فیلترهای انتخاب‌شده هیچ ردیفی باقی نمی‌ماند.")
        return

    s = market_summary(rows)
    mk = changes = prev_comp = prev_f = None
    if has_prev:
        prev_rows_all = comparison.prev_rows
        prev_info_mask = filter_mask(prev_rows_all, info, f_kam, f_cat, f_brand, f_seller)
        prev_f = prev_rows_all[prev_info_mask]
        removed_f = compare.removed_rows(rows_all, prev_f)
        mk = compare.market_changes(rows, prev_f, removed_f)
        changes = compare.seller_changes(rows)
        prev_comp = compare.prev_comp_by_seller(prev_f)
    table = seller_metrics(rows, info, prev_comp, changes)

    # ---- Data quality first ---------------------------------------------------
    checks = run_checks(rows_all, analysis.duplicate_rows, info, kam_merged, analysis.original_columns,
                        prev_columns)
    n_warn = sum(c.status == STATUS_WARN for c in checks)
    with st.expander(f"کیفیت داده — {n_warn} هشدار" if n_warn else "کیفیت داده — بدون هشدار",
                     expanded=bool(n_warn) and not st.session_state.get("dq_seen")):
        quality_section(checks, compact=True)
        st.caption("جزئیات و فهرست ردیف‌ها در تب «کیفیت داده».")
    st.session_state["dq_seen"] = True

    bullets = summary_bullets(s, table, mk, changes, info, params)
    ctx = ReportContext(
        input_name=cur_name, params=params, rows=rows, seller_info=info, table=table, summary=s,
        bullets=bullets, checks=checks, kam_merged=kam_merged,
        prev_label=comparison.prev_label if has_prev else "", prev_rows=prev_f,
        removed=removed_f if has_prev else None, mk=mk, changes=changes,
        filters={"KAM": f_kam, "دسته‌بندی": f_cat, "برند": f_brand,
                 "فروشنده": [str(info.at[x, "Seller_Name"]) for x in f_seller]},
    )
    export_bar(ctx, analysis, rows_all, info, kam_merged, checks, comparison, params, cur_name, filtered)

    tabs = st.tabs([L.TAB_SUMMARY, L.TAB_SELLERS, L.TAB_NEAR_MISS, L.TAB_CHANGES, L.TAB_CATEGORY, L.TAB_QUALITY])

    # ---- Tab 1: summary -------------------------------------------------------
    with tabs[0]:
        tab_export(ctx, "summary")
        st.info(bidi(L.INFO_COMPETITIVE_KPI.format(single_pct=fmt_pct(s["single_seller_group_pct"]),
                                              overall=fmt_pct(s["win_rate"]),
                                              competitive=fmt_pct(s["comp_win_rate"]))))
        kpi_cards(s, mk)
        left, right = st.columns([3, 2])
        with left:
            seller_bar_chart(table)
        with right:
            st.subheader("نکات کلیدی")
            for bullet in bullets:
                st.markdown(f"- {bidi(bullet)}")
            st.caption(L.INFO_PRICE_ONLY)
        trend_section(runs, f_seller)

    # ---- Tab 2: sellers -------------------------------------------------------
    with tabs[1]:
        tab_export(ctx, "sellers")
        show_cols = [c for c in L.SELLER_COLUMNS if c in table.columns]
        disp = table[show_cols].rename(columns=L.SELLER_COLUMNS)
        sort_opts = {L.SELLER_COLUMNS[c]: c for c in show_cols if c not in ("Seller_ID", "Seller_Name", "KAM", "Main_Brand")}
        sc1, sc2, _ = st.columns([2, 1, 3])
        sort_by = sc1.selectbox("مرتب‌سازی بر اساس", list(sort_opts), index=0, key="seller_sort")
        ascending = sc2.radio("ترتیب", ["نزولی", "صعودی"], horizontal=True, key="seller_order") == "صعودی"
        body = table[~table["Is_Total"]].sort_values(sort_opts[sort_by], ascending=ascending, na_position="last")
        ordered = pd.concat([body, table[table["Is_Total"]]])
        disp_cols = ["Seller_Name", "Seller_ID"] + [c for c in show_cols if c not in ("Seller_ID", "Seller_Name")]
        disp = ordered[disp_cols].rename(columns=L.SELLER_COLUMNS)
        st.caption("سطر آخر «جمع کل» تجمیعی است (مجموع بردها ÷ مجموع تنوع‌ها، نه میانگین درصدها). "
                   "برای جزئیات هر فروشنده، آن را از فهرست زیر جدول انتخاب کنید.")
        rtl_table(
            disp,
            pct=[L.SELLER_COLUMNS[c] for c in ("Win_Rate", "Comp_Win_Rate", "Prev_Comp_Win_Rate",
                                                 "Median_Loss_Gap", "Share")],
            ints=[L.SELLER_COLUMNS[c] for c in ("Variants", "Not_Evaluated", "Evaluated", "Wins",
                                                  "Comp_Variants", "Comp_Wins", "Gained", "Lost",
                                                  "Single_Seller_Variants", "No_BuyBox", "Near_Miss")],
            pp=[L.SELLER_COLUMNS["Comp_Win_Rate_Change"]],
            rate_cols=[L.SELLER_COLUMNS["Comp_Win_Rate"]],
            change_cols=[L.SELLER_COLUMNS["Comp_Win_Rate_Change"]],
            ids=["Seller_ID"], height=460, total_mask=ordered["Is_Total"].reset_index(drop=True)
            .set_axis(disp.index), max_rows=5000,
        )
        data_ids = body["Seller_ID"].tolist()
        if st.session_state.get("drill") not in data_ids:
            st.session_state.pop("drill", None)
        sid = st.selectbox("فروشنده برای جزئیات", data_ids, key="drill",
                           format_func=lambda x: f"{info.at[x, 'Seller_Name']} ({x})")
        if sid is not None:
            seller_drilldown(rows, table, sid, info, params, mk is not None)

    # ---- Tab 3: near misses -----------------------------------------------------
    with tabs[2]:
        tab_export(ctx, "near_miss")
        nm = rows[rows["Near_Miss"]]
        st.caption(f"باخت نزدیک: تنوع بازنده‌ای که فاصله قیمتش تا بای باکس بیشتر از صفر و حداکثر "
                   f"{thr:.1f}٪ است. {L.PARAM_UNDERCUT_HELP}")
        if nm.empty:
            st.info("در داده فیلترشده باخت نزدیکی وجود ندارد.")
        else:
            counts = (nm.groupby("Seller_ID").agg(n_rows=("DKPC", "size"),
                                                   n_margin=("Margin_Flag", lambda x: int((x != "").sum())))
                      .sort_values("n_rows", ascending=False))
            st.markdown(f"**{fmt_int(len(nm))} باخت نزدیک در {fmt_int(len(counts))} فروشنده**")
            for seller_id, rec in counts.iterrows():
                name = info.at[seller_id, "Seller_Name"] if seller_id in info.index else str(seller_id)
                srows = nm[nm["Seller_ID"] == seller_id].sort_values("Gap_to_BuyBox")
                n_margin = int(rec["n_margin"])
                with st.expander(bidi(f"{name} — {fmt_int(rec['n_rows'])} تنوع" +
                                      (f" (هشدار حاشیه: {n_margin})" if n_margin else ""))):
                    row_table(srows, height=380)

                    def _qw(r=srows, n=name) -> bytes:
                        return build_quickwin_workbook(r, n, params)

                    st.download_button(f"{L.EXPORT_QUICKWIN} — {name}", data=_qw, mime=XLSX_MIME,
                                       file_name=quickwin_filename(name), on_click="ignore",
                                       key=f"qw_{seller_id}")

    # ---- Tab 4: changes -----------------------------------------------------------
    with tabs[3]:
        tab_export(ctx, "changes")
        if not has_prev:
            st.info(L.INFO_NO_PREV)
        else:
            changes_tab(rows, prev_f, comparison, mk, changes, info, rows_all)

    # ---- Tab 5: category & brand -----------------------------------------------------
    with tabs[4]:
        tab_export(ctx, "category")
        category_tab(rows)

    # ---- Tab 6: data quality ---------------------------------------------------------
    with tabs[5]:
        tab_export(ctx, "quality")
        quality_section(checks, compact=False)


def export_bar(ctx: ReportContext, analysis: Analysis, rows_all: pd.DataFrame, info: pd.DataFrame,
               kam_merged, checks, comparison, params: Params, cur_name: str, filtered: bool) -> None:
    """Top export bar: full styled report (all tabs) + the standard BuyBox_<file>.xlsx."""
    has_prev = comparison is not None
    with st.container(border=True):
        c1, c2 = st.columns([1, 1])
        with c1:
            st.download_button(L.EXPORT_ALL, data=lambda: build_report(ctx), mime=XLSX_MIME, type="primary",
                               file_name=report_filename(cur_name), on_click="ignore", key="dl_all",
                               width="stretch")
            st.caption(L.EXPORT_ALL_HELP + (" (فیلترهای فعلی اعمال شده‌اند.)" if filtered else ""))
        with c2:
            def _std() -> bytes:
                full_table = seller_metrics(
                    rows_all, info,
                    compare.prev_comp_by_seller(comparison.prev_rows) if has_prev else None,
                    compare.seller_changes(rows_all) if has_prev else None)
                return build_main_workbook(analysis.raw, rows_all, full_table, kam_merged, checks_to_frame(checks),
                                           params, cur_name, has_prev, comparison.prev_label if has_prev else "")

            st.download_button(L.EXPORT_MAIN, data=_std, mime=XLSX_MIME, on_click="ignore", key="dl_main",
                               file_name=f"BuyBox_{safe_filename(Path(cur_name).stem, 120)}.xlsx", width="stretch")
            st.caption(L.EXPORT_MAIN_HELP)
        if len(rows_all) > 300_000:
            st.caption(bidi(f"فایل بزرگ است ({fmt_int(len(rows_all))} ردیف)؛ ساخت اکسل ممکن است چند دقیقه طول بکشد. "
                            "ردیف‌های بیش از ۱٬۰۰۰٬۰۰۰ در چند شیت پشت سر هم نوشته می‌شوند."))


def tab_export(ctx: ReportContext, tab: str) -> None:
    """Small per-tab download button (left side of the tab header row)."""
    _, right = st.columns([4, 1])
    right.download_button(L.EXPORT_TAB, data=lambda: build_report(ctx, [tab]), mime=XLSX_MIME,
                          file_name=report_filename(ctx.input_name, tab), on_click="ignore",
                          key=f"dl_tab_{tab}", icon=":material/download:", width="stretch")


def seller_drilldown(rows: pd.DataFrame, table: pd.DataFrame, sid: object, info: pd.DataFrame,
                     params: Params, has_prev: bool) -> None:
    rec = table[(~table["Is_Total"]) & (table["Seller_ID"] == sid)].iloc[0]
    srows = rows[rows["Seller_ID"] == sid]
    st.markdown(f"#### {rec.Seller_Name}  ·  KAM: {rec.KAM}")
    c = st.columns(6)
    c[0].metric("تنوع فعال", fmt_int(rec.Variants))
    c[1].metric("تنوع در رقابت", fmt_int(rec.Comp_Variants))
    c[2].metric("وین‌ریت رقابتی", fmt_pct(rec.Comp_Win_Rate),
                **(delta_args(rec.Comp_Win_Rate_Change) if has_prev and "Comp_Win_Rate_Change" in rec else {}))
    c[3].metric("باخت نزدیک", fmt_int(rec.Near_Miss))
    c[4].metric("بای باکس گرفته", fmt_int(rec.Gained) if has_prev else "-")
    c[5].metric("بای باکس از دست رفته", fmt_int(rec.Lost) if has_prev else "-")

    l, r = st.columns([1, 2])
    with l:
        st.markdown("**باخت به چه کسی؟ (۵ رقیب اول در تنوع‌های رقابتی)**")
        ltw = compare.lost_to_whom(rows, sid)
        if ltw.empty:
            st.caption("باخت رقابتی ندارد.")
        else:
            rtl_table(ltw.rename(columns={"BuyBox_Winner_Seller": "برنده بای باکس", "Count": "تعداد"}),
                                ints=("تعداد",))
    with r:
        nm = srows[srows["Near_Miss"]].sort_values("Gap_to_BuyBox")
        st.markdown(f"**باخت‌های نزدیک ({fmt_int(len(nm))})**")
        if nm.empty:
            st.caption("باخت نزدیکی ندارد.")
        else:
            row_table(nm, height=250)
    lost = srows[srows["Is_Loss"]].sort_values("Gap_to_BuyBox")
    st.markdown(f"**همه بای باکس‌های ازدست‌رفته ({fmt_int(len(lost))})** — "
                "ستون Margin_Flag نشان می‌دهد کاهش قیمت لازم از حد قابل قبول بیشتر است.")
    if not lost.empty:
        extra = ("Competition_Type", "BuyBox_Change") if "BuyBox_Change" in lost.columns else ("Competition_Type",)
        row_table(lost, extra, height=320)


def changes_tab(rows, prev_f, comparison, mk, changes, info, rows_all) -> None:
    st.caption(f"مقایسه با: {comparison.prev_label} (تطبیق بر اساس DKPC)")
    c = st.columns(6)
    c[0].metric(L.KPI_GAINED, fmt_int(mk["gained"]))
    c[1].metric(L.KPI_LOST, fmt_int(mk["lost"]))
    c[2].metric(L.KPI_NEW, fmt_int(mk["new"]))
    c[3].metric(L.KPI_REMOVED, fmt_int(mk["removed"]))
    c[4].metric("وین‌ریت رقابتی دوره قبل", fmt_pct(mk["prev_comp_win_rate"]))
    c[5].metric("وین‌ریت رقابتی فعلی", fmt_pct(mk["cur_comp_win_rate"]), **delta_args(mk["comp_win_rate_change_pp"]))

    winners, losers = compare.net_movers(changes, info)
    a, b = st.columns(2)
    rename = {"Seller_Name": "فروشنده", "Gained": "گرفته", "Lost": "از دست رفته", "Net": "خالص"}
    with a:
        st.markdown("**۵ برنده خالص اول**")
        rtl_table(winners.rename(columns=rename), ints=("گرفته", "از دست رفته", "خالص"),
                            ids=("Seller_ID",))
    with b:
        st.markdown("**۵ بازنده خالص اول**")
        rtl_table(losers.rename(columns=rename), ints=("گرفته", "از دست رفته", "خالص"),
                            ids=("Seller_ID",))

    st.markdown("**تغییرات به تفکیک فروشنده**")
    t = compare.seller_change_table(rows, prev_f, changes, info)
    rtl_table(t.rename(columns={"Seller_Name": "فروشنده", "Prev_Rate": "وین‌ریت رقابتی قبل",
                                          "Cur_Rate": "وین‌ریت رقابتی فعلی", "Change_pp": "تغییر",
                                          "Gained": "گرفته", "Lost": "از دست رفته", "Net": "خالص",
                                          "New": "تنوع جدید"}),
                        pct=("وین‌ریت رقابتی قبل", "وین‌ریت رقابتی فعلی"), pp=("تغییر",),
                        ints=("گرفته", "از دست رفته", "خالص", "تنوع جدید"), change_cols=("تغییر",),
                        ids=("Seller_ID",), height=380)

    st.markdown("**ردیف‌های تغییرکرده**")
    kinds = [L.CHANGE_GAINED, L.CHANGE_LOST, L.CHANGE_NEW, L.CHANGE_NO_WEIGHT, L.CHANGE_SAME, L.CHANGE_BAD_PRICE]
    pick = st.multiselect("نوع تغییر", kinds, default=[L.CHANGE_GAINED, L.CHANGE_LOST], key="chg_kinds")
    ch_rows = rows[rows["BuyBox_Change"].isin(pick)]
    cols = [c for c in ("DKPC", "DKP", "Seller_Name", "Product_Name", "Weight", "Selling_Price", "BuyBox_Min_Price",
                        "BuyBox_Winner_Seller", "Comment_Prev_File", "Comment", "BuyBox_Change") if c in ch_rows.columns]
    st.caption(f"{fmt_int(len(ch_rows))} ردیف (حداکثر ۵٬۰۰۰ ردیف نمایش داده می‌شود)")
    rtl_table(ch_rows[cols].head(5000), ints=("Selling_Price", "BuyBox_Min_Price"), weight=("Weight",),
                        ids=("DKPC", "DKP"), height=360)

    removed = compare.removed_rows(rows_all, prev_f)
    st.markdown(f"**تنوع‌های حذف‌شده (در فایل قبل بودند، در فایل فعلی نیستند): {fmt_int(len(removed))}**")
    if not removed.empty:
        rtl_table(removed.head(5000), ints=("Selling_Price",), weight=("Weight",),
                            ids=("DKPC", "DKP", "Seller_ID"), height=300)


def category_tab(rows: pd.DataFrame) -> None:
    rename = {"Variants": "تعداد تنوع", "Evaluated": "ارزیابی‌شده", "Comp_Variants": "تنوع در رقابت",
              "Comp_Wins": "بای باکس در رقابت", "Comp_Win_Rate": "وین‌ریت رقابتی",
              "Median_Loss_Gap": "میانه فاصله باخت‌ها", "Near_Miss": "باخت نزدیک", "Sellers": "تعداد فروشنده"}
    ints = ("تعداد تنوع", "ارزیابی‌شده", "تنوع در رقابت", "بای باکس در رقابت", "باخت نزدیک", "تعداد فروشنده")
    a, b = st.columns(2)
    for box, col, title in ((a, "Leaf_Category_Name", "به تفکیک دسته‌بندی"), (b, "Brand_Name_En", "به تفکیک برند")):
        with box:
            st.markdown(f"**{title}**")
            t = rate_by_dimension(rows, col)
            if t.empty:
                st.caption(f"ستون {col} در فایل نیست.")
                continue
            rtl_table(t.rename(columns=rename), pct=("وین‌ریت رقابتی", "میانه فاصله باخت‌ها"),
                                ints=ints, rate_cols=("وین‌ریت رقابتی",), height=360)
    st.markdown("**گروه‌های تک‌فروشنده — DKPهای پرتنوع (کاندید جذب فروشنده دوم)**")
    t = single_seller_dkps(rows, top=50)
    if t.empty:
        st.caption("گروه تک‌فروشنده‌ای وجود ندارد.")
    else:
        rtl_table(t.rename(columns={"Variants": "تنوع تک‌فروشنده", "Weights": "تعداد وزن",
                                              "Sellers": "فروشنده(ها)"}),
                            ints=("تنوع تک‌فروشنده", "تعداد وزن"), ids=("DKP",), height=380)


if __name__ == "__main__":
    main()
