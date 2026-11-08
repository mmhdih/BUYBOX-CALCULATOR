"""Styled Persian Excel report covering every tab of the app (or a single tab).

Every sheet reflects the filters active in the app (KAM, category, brand, seller); the
filters are written on each sheet. The standard ``BuyBox_<file>.xlsx`` export
(:mod:`buybox.export`) stays unfiltered and in the specified format.
"""

from __future__ import annotations

import io
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import pandas as pd
import xlsxwriter

from buybox import compare
from buybox import labels as L
from buybox.core import Params, rate_by_dimension, result_columns_present, single_seller_dkps
from buybox.export import safe_filename
from buybox.quality import QualityCheck, checks_to_frame
from buybox.xlstyle import COLOR_KEY, Sheet, SheetWriter, Styles, Table, workbook_options

#: Tab keys in app order, with their Persian names (used for per-tab file names).
TABS: dict[str, str] = {
    "summary": L.TAB_SUMMARY,
    "sellers": L.TAB_SELLERS,
    "near_miss": "باخت‌های نزدیک",
    "changes": L.TAB_CHANGES,
    "category": L.TAB_CATEGORY,
    "quality": L.TAB_QUALITY,
    "rows": "همه ردیف‌ها",
}
ALL_TABS: tuple[str, ...] = tuple(TABS)


@dataclass
class ReportContext:
    """Everything the report needs, already filtered the way the app shows it."""

    input_name: str
    params: Params
    rows: pd.DataFrame
    seller_info: pd.DataFrame
    table: pd.DataFrame
    summary: dict
    bullets: list[str] = field(default_factory=list)
    checks: list[QualityCheck] = field(default_factory=list)
    kam_merged: pd.DataFrame | None = None
    prev_label: str = ""
    prev_rows: pd.DataFrame | None = None
    removed: pd.DataFrame | None = None
    mk: dict | None = None
    changes: pd.DataFrame | None = None
    filters: dict[str, list[str]] = field(default_factory=dict)
    created: datetime = field(default_factory=datetime.now)

    @property
    def has_prev(self) -> bool:
        return self.mk is not None and self.prev_rows is not None


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------
def _subtitle(ctx: ReportContext) -> str:
    parts = [f"فایل: {ctx.input_name}", f"تاریخ تهیه: {ctx.created:%Y-%m-%d %H:%M}"]
    if ctx.prev_label:
        parts.append(f"دوره قبل: {ctx.prev_label}")
    active = {k: v for k, v in ctx.filters.items() if v}
    parts.append("فیلتر: " + ("؛ ".join(f"{k}: {'، '.join(map(str, v))}" for k, v in active.items())
                              if active else "ندارد (همه داده‌ها)"))
    return "  |  ".join(parts)


def _row_table(title: str, rows: pd.DataFrame, cols: list[str], note: str = "") -> Table:
    cols = [c for c in cols if c in rows.columns]
    df = rows[cols].copy()
    if "Near_Miss" in df.columns:
        df["Near_Miss"] = df["Near_Miss"].map({True: "بله", False: ""})
    return Table(
        title=title, df=df, note=note,
        headers={c: L.ROW_COLUMNS.get(c, (c, ""))[0] for c in cols},
        tips={c: L.ROW_COLUMNS[c][1] for c in cols if c in L.ROW_COLUMNS},
        text_rules={
            "Comment": {L.COMMENT_WIN: "green", L.COMMENT_LOSS: "red", L.COMMENT_NO_WEIGHT: "amber",
                        L.COMMENT_BAD_PRICE: "amber", L.COMMENT_MISSING_ID: "amber"},
            "BuyBox_Change": {L.CHANGE_GAINED: "green", L.CHANGE_LOST: "red", L.CHANGE_NEW: "blue",
                              L.CHANGE_NO_WEIGHT: "amber"},
            "Margin_Flag": {L.MARGIN_FLAG: "red"},
            "Competition_Type": {L.COMP_COMPETITIVE: "blue"},
        },
    )


LOSS_COLUMNS = ["Seller_Name", "KAM", "DKPC", "DKP", "Product_Name", "Weight", "Selling_Price",
                "BuyBox_Min_Price", "BuyBox_Winner_Seller", "Gap_to_BuyBox", "Target_Price", "Required_Cut_%",
                "Margin_Flag", "Competition_Type", "Near_Miss", "BuyBox_Change"]


# ---------------------------------------------------------------------------
# sheets
# ---------------------------------------------------------------------------
def _summary(ctx: ReportContext) -> list[Sheet]:
    s = ctx.summary
    kpis: list[tuple[str, object, str]] = [
        ("تعداد کل تنوع‌ها (DKPC)", s["rows"], "int"),
        ("تنوع ارزیابی‌شده", s["evaluated"], "int"),
        ("تنوع خارج از ارزیابی (بدون وزن / داده ناقص)", s["not_evaluated"], "int"),
        ("گروه‌های DKP + وزن", s["groups"], "int"),
        ("سهم گروه‌های تک‌فروشنده", s["single_seller_group_pct"], "pct"),
        ("تعداد بای باکس", s["wins"], "int"),
        ("وین‌ریت کل", s["win_rate"], "pct"),
        ("تنوع در رقابت", s["comp_rows"], "int"),
        ("بای باکس در رقابت", s["comp_wins"], "int"),
        ("وین‌ریت رقابتی (شاخص اصلی)", s["comp_win_rate"], "pct"),
        ("باخت نزدیک", s["near_miss"], "int"),
        ("بازنده با هشدار حاشیه", s["margin_flags"], "int"),
        ("گروه‌های دارای تساوی", s["tie_groups"], "int"),
    ]
    if ctx.mk is not None:
        mk = ctx.mk
        kpis += [
            ("وین‌ریت رقابتی دوره قبل", mk["prev_comp_win_rate"], "pct"),
            ("تغییر وین‌ریت رقابتی", mk["comp_win_rate_change_pp"], "pp"),
            ("بای باکس گرفته", mk["gained"], "int"),
            ("بای باکس از دست رفته", mk["lost"], "int"),
            ("تنوع جدید", mk["new"], "int"),
            ("تنوع حذف‌شده", mk["removed"], "int"),
        ]
    p = ctx.params
    notes = [
        f"پارامترها: آستانه باخت نزدیک {p.near_miss_threshold * 100:.1f}٪ | زیر قیمت زدن {p.undercut:,.0f} ریال | "
        f"حداکثر کاهش قابل قبول {p.max_acceptable_cut * 100:.1f}٪",
        COLOR_KEY,
        *L.EXCEL_NOTES,
    ]
    top = ctx.table[~ctx.table["Is_Total"] & (ctx.table["Comp_Variants"] > 0)]
    top = top.sort_values("Comp_Variants", ascending=False).head(20)
    top_t = Table(
        title="۲۰ فروشنده با بیشترین تنوع در رقابت",
        df=top[["Seller_Name", "Comp_Variants", "Comp_Wins", "Comp_Win_Rate"]],
        headers={c: L.SELLER_COLUMNS[c] for c in ("Seller_Name", "Comp_Variants", "Comp_Wins", "Comp_Win_Rate")},
        kinds={"Comp_Variants": "int", "Comp_Wins": "int", "Comp_Win_Rate": "pct"},
        rate_cols=("Comp_Win_Rate",),
        tips={"Comp_Win_Rate": L.SELLER_COLUMN_HELP["Comp_Win_Rate"]},
        chart={"category": "Seller_Name", "value": "Comp_Win_Rate", "title": "وین‌ریت رقابتی"},
    )
    return [Sheet(name="خلاصه", title=L.EXCEL_TITLE, subtitle=_subtitle(ctx), kpis=kpis,
                  bullets=ctx.bullets, tables=[top_t], notes=notes)]


def seller_table(ctx: ReportContext) -> Table:
    cols = [c for c in L.SELLER_COLUMNS if c in ctx.table.columns]
    if not ctx.has_prev:
        cols = [c for c in cols if c not in L.PREV_ONLY_COLUMNS]
    kinds = {c: "int" for c in ("Variants", "Not_Evaluated", "Evaluated", "Wins", "Comp_Variants", "Comp_Wins",
                                "Gained", "Lost", "Single_Seller_Variants", "No_BuyBox", "Near_Miss")}
    kinds.update({c: "pct" for c in ("Win_Rate", "Comp_Win_Rate", "Prev_Comp_Win_Rate", "Median_Loss_Gap",
                                     "Share")})
    kinds.update({"Comp_Win_Rate_Change": "pp", "Seller_ID": "id"})
    return Table(
        title="گزارش فروشندگان (سطر آخر = جمع کل تجمیعی)",
        df=ctx.table[cols], kinds=kinds, headers=dict(L.SELLER_COLUMNS),
        tips=dict(L.SELLER_COLUMN_HELP), rate_cols=("Comp_Win_Rate",), change_cols=("Comp_Win_Rate_Change",),
        total_last=True, dash_for_missing=("Comp_Win_Rate", "Prev_Comp_Win_Rate", "Comp_Win_Rate_Change"),
        note="وین‌ریت رقابتی جمع کل = مجموع بای باکس در رقابت ÷ مجموع تنوع در رقابت (نه میانگین درصدها).",
    )


def _sellers(ctx: ReportContext) -> list[Sheet]:
    sheets = [Sheet(name="فروشندگان", title="فروشندگان", subtitle=_subtitle(ctx), tables=[seller_table(ctx)],
                    notes=[COLOR_KEY])]
    ltw = compare.lost_to_whom_all(ctx.rows, ctx.seller_info)
    sheets.append(Sheet(
        name="باخت به چه کسی", title="باخت به چه کسی — ۵ رقیب اول هر فروشنده در تنوع‌های رقابتی",
        subtitle=_subtitle(ctx),
        tables=[Table(title="برندگان بای باکس در برابر هر فروشنده", df=ltw,
                      headers={"Seller_Name": "فروشنده", "BuyBox_Winner_Seller": "برنده بای باکس",
                               "Count": "تعداد باخت", "Share": "سهم از باخت‌های رقابتی"},
                      kinds={"Count": "int", "Share": "pct", "Seller_ID": "id"})],
    ))
    losses = ctx.rows[ctx.rows["Is_Loss"].astype(bool)].sort_values(["Seller_Name", "Gap_to_BuyBox"])
    sheets.append(Sheet(
        name="بای باکس‌های از دست رفته", title="همه تنوع‌های بدون بای باکس", subtitle=_subtitle(ctx),
        tables=[_row_table(f"{len(losses):,} تنوع بازنده — مرتب بر اساس فروشنده و فاصله", losses, LOSS_COLUMNS,
                           note="هشدار حاشیه یعنی کاهش قیمت لازم از حد قابل قبول بیشتر است.")],
    ))
    return sheets


def _near_miss(ctx: ReportContext) -> list[Sheet]:
    nm = ctx.rows[ctx.rows["Near_Miss"].astype(bool)]
    per = pd.DataFrame(columns=["Seller_Name", "KAM", "n", "flags", "median"])
    if not nm.empty:
        per = (nm.groupby("Seller_ID")
               .agg(n=("DKPC", "size"), flags=("Margin_Flag", lambda x: int((x != "").sum())),
                    median=("Gap_to_BuyBox", "median"))
               .sort_values("n", ascending=False))
        per.insert(0, "Seller_Name", ctx.seller_info["Seller_Name"].reindex(per.index))
        per.insert(1, "KAM", ctx.seller_info["KAM"].reindex(per.index))
        per = per.reset_index()
    thr = ctx.params.near_miss_threshold * 100
    return [Sheet(
        name="باخت‌های نزدیک", title=f"باخت‌های نزدیک (Quick Win) — فاصله بزرگ‌تر از صفر و حداکثر {thr:.1f}٪",
        subtitle=_subtitle(ctx),
        tables=[
            Table(title="خلاصه به تفکیک فروشنده", df=per,
                  headers={"Seller_Name": "فروشنده", "n": "تعداد باخت نزدیک", "flags": "دارای هشدار حاشیه",
                           "median": "میانه فاصله"},
                  kinds={"n": "int", "flags": "int", "median": "pct", "Seller_ID": "id"}),
            _row_table(f"همه باخت‌های نزدیک ({len(nm):,})",
                       nm.sort_values(["Seller_Name", "Gap_to_BuyBox"]), LOSS_COLUMNS,
                       note=L.PARAM_UNDERCUT_HELP),
        ],
    )]


def _changes(ctx: ReportContext) -> list[Sheet]:
    if not ctx.has_prev:
        return [Sheet(name="تغییرات", title=L.TAB_CHANGES, subtitle=_subtitle(ctx), notes=[L.INFO_NO_PREV])]
    mk = ctx.mk
    kpis = [("بای باکس گرفته", mk["gained"], "int"), ("بای باکس از دست رفته", mk["lost"], "int"),
            ("تنوع جدید", mk["new"], "int"), ("تنوع حذف‌شده", mk["removed"], "int"),
            ("وین‌ریت رقابتی دوره قبل", mk["prev_comp_win_rate"], "pct"),
            ("وین‌ریت رقابتی فعلی", mk["cur_comp_win_rate"], "pct"),
            ("تغییر", mk["comp_win_rate_change_pp"], "pp")]
    winners, losers = compare.net_movers(ctx.changes, ctx.seller_info)
    mv_headers = {"Seller_Name": "فروشنده", "Gained": "گرفته", "Lost": "از دست رفته", "Net": "خالص"}
    mv_kinds = {"Gained": "int", "Lost": "int", "Net": "int", "Seller_ID": "id"}
    sct = compare.seller_change_table(ctx.rows, ctx.prev_rows, ctx.changes, ctx.seller_info)
    changed = ctx.rows[ctx.rows["BuyBox_Change"] != L.CHANGE_SAME].sort_values(["BuyBox_Change", "Seller_Name"])
    cols = ["BuyBox_Change", "Seller_Name", "DKPC", "DKP", "Product_Name", "Weight", "Selling_Price",
            "BuyBox_Min_Price", "BuyBox_Winner_Seller", "Comment_Prev_File", "Comment"]
    removed = ctx.removed if ctx.removed is not None else pd.DataFrame()
    return [Sheet(
        name="تغییرات", title=f"{L.TAB_CHANGES} — مقایسه با {ctx.prev_label}", subtitle=_subtitle(ctx),
        kpis=kpis,
        tables=[
            Table(title="۵ برنده خالص اول", df=winners, headers=mv_headers, kinds=mv_kinds),
            Table(title="۵ بازنده خالص اول", df=losers, headers=mv_headers, kinds=mv_kinds),
            Table(title="تغییرات به تفکیک فروشنده", df=sct,
                  headers={"Seller_Name": "فروشنده", "Prev_Rate": "وین‌ریت رقابتی قبل",
                           "Cur_Rate": "وین‌ریت رقابتی فعلی", "Change_pp": "تغییر", "Gained": "گرفته",
                           "Lost": "از دست رفته", "Net": "خالص", "New": "تنوع جدید"},
                  kinds={"Prev_Rate": "pct", "Cur_Rate": "pct", "Change_pp": "pp", "Gained": "int", "Lost": "int",
                         "Net": "int", "New": "int", "Seller_ID": "id"},
                  rate_cols=("Cur_Rate",), change_cols=("Change_pp",),
                  dash_for_missing=("Prev_Rate", "Cur_Rate", "Change_pp")),
            _row_table(f"ردیف‌های تغییرکرده ({len(changed):,}) — بدون «بدون تغییر»", changed, cols),
            _row_table(f"تنوع‌های حذف‌شده ({len(removed):,}) — در فایل قبل بودند، در فایل فعلی نیستند",
                       removed, ["DKPC", "DKP", "Seller_ID", "Seller_Name", "Product_Name", "Weight",
                                 "Selling_Price", "Comment"]),
        ],
        notes=[COLOR_KEY],
    )]


def _category(ctx: ReportContext) -> list[Sheet]:
    headers = {"Variants": "تعداد تنوع", "Evaluated": "ارزیابی‌شده", "Comp_Variants": "تنوع در رقابت",
               "Comp_Wins": "بای باکس در رقابت", "Comp_Win_Rate": "وین‌ریت رقابتی",
               "Median_Loss_Gap": "میانه فاصله باخت‌ها", "Near_Miss": "باخت نزدیک", "Sellers": "تعداد فروشنده",
               "Leaf_Category_Name": "دسته‌بندی", "Brand_Name_En": "برند"}
    kinds = {c: "int" for c in ("Variants", "Evaluated", "Comp_Variants", "Comp_Wins", "Near_Miss", "Sellers")}
    kinds.update({"Comp_Win_Rate": "pct", "Median_Loss_Gap": "pct"})
    tables = []
    for col, title in (("Leaf_Category_Name", "به تفکیک دسته‌بندی"), ("Brand_Name_En", "به تفکیک برند")):
        t = rate_by_dimension(ctx.rows, col)
        if not t.empty:
            tables.append(Table(title=title, df=t, headers=headers, kinds=kinds, rate_cols=("Comp_Win_Rate",),
                                dash_for_missing=("Comp_Win_Rate", "Median_Loss_Gap")))
    single = single_seller_dkps(ctx.rows, top=200)
    tables.append(Table(title="گروه‌های تک‌فروشنده — DKPهای پرتنوع (کاندید جذب فروشنده دوم)", df=single,
                        headers={"Variants": "تنوع تک‌فروشنده", "Weights": "تعداد وزن", "Sellers": "فروشنده(ها)",
                                 "Product_Name": "نام کالا"},
                        kinds={"Variants": "int", "Weights": "int", "DKP": "id"}))
    return [Sheet(name="دسته‌بندی و برند", title=L.TAB_CATEGORY, subtitle=_subtitle(ctx), tables=tables,
                  notes=[COLOR_KEY])]


def _quality(ctx: ReportContext) -> list[Sheet]:
    q = checks_to_frame(ctx.checks)
    tables = [Table(title="چک‌لیست کیفیت داده", df=q,
                    text_rules={"وضعیت": {"OK": "green", "هشدار": "amber", "اطلاع": "blue"}})]
    for chk in ctx.checks:
        if chk.rows is not None and not chk.rows.empty:
            tables.append(_row_table(f"{chk.label} — {chk.display}", chk.rows.head(MAX_DETAIL_ROWS),
                                     list(chk.rows.columns)))
    if ctx.kam_merged is not None and not ctx.kam_merged.empty:
        tables.append(Table(title="لیست KAM استفاده‌شده (Seller Id تکراری ادغام شده)", df=ctx.kam_merged,
                            headers={"Seller_Name": "Seller Name", "Category": "Category Level1 Name Fa",
                                     "Rows": "تعداد ردیف در لیست اصلی", "KAM_Note": "یادداشت"},
                            kinds={"Rows": "int", "Seller_ID": "id"},
                            text_rules={"KAM_Note": {L.KAM_DUPLICATE_FLAG: "amber"}}))
    return [Sheet(name="کیفیت داده", title=L.TAB_QUALITY, subtitle=_subtitle(ctx), tables=tables)]


MAX_DETAIL_ROWS = 100_000


def _rows(ctx: ReportContext) -> list[Sheet]:
    base = ["DKPC", "DKP", "Seller_ID", "Seller_Name", "KAM", "Product_Name", "Brand_Name_En",
            "Leaf_Category_Name", "Weight", "Selling_Price"]
    cols = base + result_columns_present(ctx.rows) + ["Near_Miss", "Comment_Prev_File", "BuyBox_Change"]
    return [Sheet(name="همه ردیف‌ها", title="همه ردیف‌ها با نتیجه بای باکس", subtitle=_subtitle(ctx),
                  tables=[_row_table(f"{len(ctx.rows):,} ردیف", ctx.rows, cols)])]


def _guide(ctx: ReportContext) -> list[Sheet]:
    row_guide = pd.DataFrame([(k, fa, desc) for k, (fa, desc) in L.ROW_COLUMNS.items()],
                             columns=["نام ستون (انگلیسی)", "عنوان فارسی", "توضیح"])
    seller_guide = pd.DataFrame([(L.SELLER_COLUMNS[k], desc) for k, desc in L.SELLER_COLUMN_HELP.items()],
                                columns=["ستون", "توضیح"])
    return [Sheet(name="راهنمای ستون‌ها", title="راهنمای ستون‌ها و روش محاسبه", tables=[
        Table(title="ستون‌های ردیفی (هر DKPC)", df=row_guide),
        Table(title="ستون‌های گزارش فروشندگان", df=seller_guide),
    ], notes=[COLOR_KEY, *L.EXCEL_NOTES])]


SHEET_BUILDERS = {
    "summary": [_summary],
    "sellers": [_sellers],
    "near_miss": [_near_miss],
    "changes": [_changes],
    "category": [_category],
    "quality": [_quality],
    "rows": [_rows],
}


class _ChartWriter(SheetWriter):
    """SheetWriter that also draws a coloured bar chart next to tables that ask for one."""

    def _write_table(self, ws, r, t, widths, freeze):  # noqa: ANN001
        first = r + 1 + (1 if t.note else 0)
        end = super()._write_table(ws, r, t, widths, freeze)
        spec = t.chart
        if spec and not t.df.empty:
            cols = list(t.df.columns)
            ci, vi = cols.index(spec["category"]), cols.index(spec["value"])
            n = len(t.df)
            sheet = ws.get_name()
            points = []
            for v in t.df[spec["value"]].tolist():
                colour = "#D64545" if v < 0.3 else "#E8A33D" if v <= 0.6 else "#3A9D5D"
                points.append({"fill": {"color": colour}})
            chart = self.wb.add_chart({"type": "bar"})
            chart.add_series({
                "categories": [sheet, first + 1, ci, first + n, ci],
                "values": [sheet, first + 1, vi, first + n, vi],
                "points": points,
                "data_labels": {"value": True, "num_format": "0.0%"},
            })
            chart.set_title({"name": spec["title"], "name_font": {"name": "Tahoma", "size": 11}})
            chart.set_x_axis({"num_format": "0%", "min": 0, "max": 1, "reverse": True,
                              "num_font": {"name": "Tahoma"}})
            chart.set_y_axis({"reverse": True, "num_font": {"name": "Tahoma"}})
            chart.set_legend({"none": True})
            chart.set_size({"width": 620, "height": max(320, 24 * n + 80)})
            ws.insert_chart(first, len(cols) + 1, chart)
        return end


def build_report(ctx: ReportContext, tabs: tuple[str, ...] | list[str] = ALL_TABS) -> bytes:
    """Write the selected tabs (default: all) plus the column guide; return the .xlsx bytes."""
    buf = io.BytesIO()
    n = len(ctx.rows) if "rows" in tabs else 0
    wb = xlsxwriter.Workbook(buf, workbook_options(n))
    wb.set_properties({"title": L.EXCEL_TITLE, "comments": "Buy Box Analyzer"})
    writer = _ChartWriter(wb, Styles(wb))
    for tab in tabs:
        for build in SHEET_BUILDERS[tab]:
            for sheet in build(ctx):
                writer.write(sheet)
    for sheet in _guide(ctx):
        writer.write(sheet)
    wb.close()
    return buf.getvalue()


def report_filename(input_name: str, tab: str | None = None, when: datetime | None = None) -> str:
    """``BuyBox_Report_<input>_<date>.xlsx`` (all tabs) or ``..._<tab>_<date>.xlsx``."""
    when = when or datetime.now()
    stem = safe_filename(Path(input_name).stem, 80)
    part = f"_{safe_filename(TABS[tab], 40)}" if tab else ""
    return f"BuyBox_Report_{stem}{part}_{when:%Y%m%d_%H%M}.xlsx"
