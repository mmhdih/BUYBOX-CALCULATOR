import io

import openpyxl
import pandas as pd
import pytest

from buybox import compare
from buybox import labels as L
from buybox.core import Params, market_summary, seller_metrics
from buybox.io import merge_kam
from buybox.pipeline import attach_comparison
from buybox.quality import run_checks
from buybox.report import ALL_TABS, ReportContext, build_report, report_filename
from tests.conftest import run
from tests.test_compare import CUR, PREV


def _ctx(with_prev: bool = True) -> ReportContext:
    cur, prev = run(CUR), run(PREV)
    kam = merge_kam(pd.DataFrame({"Seller_ID": pd.array([10, 10], dtype="Int64"), "Seller_Name": ["A", "A2"],
                                  "Category": ["x", "x"], "KAM": ["k1", "k2"]}))
    kw = {}
    pc = ch = None
    if with_prev:
        comp = attach_comparison(cur, prev.rows, "prev.xlsx")
        ch = compare.seller_changes(cur.rows)
        pc = compare.prev_comp_by_seller(prev.rows)
        kw = dict(prev_label="prev.xlsx", prev_rows=prev.rows, removed=comp.removed,
                  mk=compare.market_changes(cur.rows, prev.rows, comp.removed), changes=ch)
    return ReportContext(
        input_name="cur.xlsx", params=Params(), rows=cur.rows, seller_info=cur.seller_info,
        table=seller_metrics(cur.rows, cur.seller_info, pc, ch), summary=market_summary(cur.rows),
        bullets=["نکته ۱"], checks=run_checks(cur.rows, cur.duplicate_rows, cur.seller_info, kam,
                                               cur.original_columns), kam_merged=kam,
        filters={"KAM": ["k1"]}, **kw)


def _load(data: bytes):
    return openpyxl.load_workbook(io.BytesIO(data))


def test_full_report_has_a_sheet_per_tab():
    wb = _load(build_report(_ctx()))
    assert wb.sheetnames == ["خلاصه", "فروشندگان", "باخت به چه کسی", "بای باکس‌های از دست رفته",
                             "باخت‌های نزدیک", "تغییرات", "دسته‌بندی و برند", "کیفیت داده", "همه ردیف‌ها",
                             "راهنمای ستون‌ها"]
    for ws in wb.worksheets:
        assert ws.sheet_view.rightToLeft
        for row in ws.iter_rows():
            for c in row:
                assert not (isinstance(c.value, str) and c.value.startswith("="))  # values only


def test_summary_sheet_contents():
    wb = _load(build_report(_ctx()))
    ws = wb["خلاصه"]
    values = {r[0]: r[1] for r in ws.iter_rows(values_only=True) if r[0]}
    assert values["وین‌ریت رقابتی (شاخص اصلی)"] == pytest.approx(0.5)
    assert values["بای باکس گرفته"] == 1 and values["بای باکس از دست رفته"] == 1
    text = " ".join(str(v) for row in ws.iter_rows(values_only=True) for v in row if v)
    assert "فیلتر: KAM: k1" in text and "• نکته ۱" in text
    assert ws._charts  # bar chart of competitive win rate


def test_seller_sheet_persian_headers_total_and_tooltips():
    wb = _load(build_report(_ctx(), ["sellers"]))
    ws = wb["فروشندگان"]
    (table,) = ws.tables.values()
    first = ws[table.ref.split(":")[0]].row
    header = [c.value for c in ws[first] if c.value]
    assert header[:3] == ["Seller_ID", "Seller Name", "KAM"]
    assert L.SELLER_COLUMNS["Comp_Win_Rate"] in header
    j = header.index(L.SELLER_COLUMNS["Comp_Win_Rate"]) + 1
    assert ws.cell(first, j).comment is not None
    last = ws[table.ref.split(":")[1]].row
    assert ws.cell(last, 2).value == L.TOTAL_ROW_LABEL and ws.cell(last, 2).font.b


def test_single_tab_export_and_without_prev():
    wb = _load(build_report(_ctx(with_prev=False), ["changes"]))
    assert wb.sheetnames == ["تغییرات", "راهنمای ستون‌ها"]
    text = " ".join(str(v) for row in wb["تغییرات"].iter_rows(values_only=True) for v in row if v)
    assert L.INFO_NO_PREV in text


def test_each_tab_builds_on_its_own():
    ctx = _ctx()
    for tab in ALL_TABS:
        assert _load(build_report(ctx, [tab])).sheetnames[-1] == "راهنمای ستون‌ها"


def test_report_filename():
    assert report_filename("اسورتمنت شمش.xlsx").startswith("BuyBox_Report_اسورتمنت شمش_")
    assert "_فروشندگان_" in report_filename("a.csv", "sellers")


def test_streaming_mode_for_huge_files(monkeypatch):
    """Force the constant-memory / streamed-table paths with tiny thresholds."""
    import buybox.export as ex
    import buybox.xlstyle as xs
    monkeypatch.setattr(xs, "BIG_WORKBOOK_ROWS", 2)
    monkeypatch.setattr(xs, "STREAM_ROWS", 3)
    monkeypatch.setattr(ex, "BIG_WORKBOOK_ROWS", 2)
    ctx = _ctx()
    wb = _load(build_report(ctx))
    ws = wb["همه ردیف‌ها"]
    assert not ws.tables  # streamed, not an Excel table
    values = [c.value for row in ws.iter_rows() for c in row]
    header_row = next(i for i, row in enumerate(ws.iter_rows(values_only=True), 1) if row[0] == "DKPC")
    assert ws.auto_filter.ref.startswith(f"A{header_row}")
    dkpcs = [ws.cell(r, 1).value for r in range(header_row + 1, ws.max_row + 1) if ws.cell(r, 1).value]
    assert dkpcs == ctx.rows["DKPC"].tolist()
    assert L.COMMENT_WIN in values
    # small tables in a streaming workbook must still carry their data (no Excel table objects there)
    sellers = wb["فروشندگان"]
    names = [c.value for row in sellers.iter_rows() for c in row]
    assert L.TOTAL_ROW_LABEL in names and "A" in names
    summary = {r[0]: r[1] for r in wb["خلاصه"].iter_rows(values_only=True) if r[0]}
    assert summary["وین‌ریت رقابتی (شاخص اصلی)"] == pytest.approx(0.5)

    from buybox.quality import checks_to_frame
    data = ex.build_main_workbook(ctx.rows[["DKPC", "DKP", "Seller_ID", "Weight", "Selling_Price"]], ctx.rows,
                                  ctx.table, None, checks_to_frame(ctx.checks), Params(), "x", True, "p")
    std = _load(data)
    assert [c.value for c in std["Sheet1"][2]][0] == ctx.rows["DKPC"].iloc[0]
    assert std["Sheet1"].max_row == len(ctx.rows) + 1
