import io

import openpyxl
import pandas as pd

from buybox import compare
from buybox import labels as L
from buybox.core import Params, seller_metrics
from buybox.export import build_main_workbook, build_quickwin_workbook, quickwin_filename
from buybox.io import merge_kam, read_excel_bytes
from buybox.pipeline import attach_comparison
from buybox.quality import checks_to_frame, run_checks
from tests.conftest import make_raw, run
from tests.test_compare import CUR, PREV


def _build(has_prev=True):
    cur = run(CUR)
    prev = run(PREV)
    kam = merge_kam(pd.DataFrame({"Seller_ID": pd.array([10, 10], dtype="Int64"), "Seller_Name": ["A", "A2"],
                                  "Category": ["x", "x"], "KAM": ["k1", "k2"]}))
    pc = ch = None
    if has_prev:
        comp = attach_comparison(cur, prev.rows, "prev")
        pc, ch = compare.prev_comp_by_seller(prev.rows), compare.seller_changes(cur.rows)
    table = seller_metrics(cur.rows, cur.seller_info, pc, ch)
    q = checks_to_frame(run_checks(cur.rows, cur.duplicate_rows, cur.seller_info, kam, cur.original_columns,
                                   prev.original_columns if has_prev else None))
    data = build_main_workbook(cur.raw, cur.rows, table, kam, q, Params(), "cur.xlsx", has_prev, "prev")
    return cur, openpyxl.load_workbook(io.BytesIO(data))


def test_workbook_structure_and_values():
    cur, wb = _build()
    assert wb.sheetnames == ["BuyBox_Seller_Report", "Sheet1", "KAM_LIST", "Data_Quality", "راهنما"]
    assert wb["BuyBox_Seller_Report"].sheet_view.rightToLeft
    ws = wb["Sheet1"]
    header = [c.value for c in ws[1]]
    original = list(make_raw(CUR).columns)
    assert header[: len(original)] == original
    assert header[len(original):] == [
        "BuyBox_Min_Price", "Other_Sellers_Variants", "Same_Seller_Variants", "Gap_to_BuyBox",
        "Competition_Type", "Comment", "BuyBox_Winner_Seller", "Target_Price", "Required_Cut_%",
        "Margin_Flag", "Comment_Prev_File", "BuyBox_Change"]
    assert ws.max_row == len(CUR) + 1
    assert ws.freeze_panes == "B2"  # header row + DKPC column
    assert ws.cell(1, header.index("Gap_to_BuyBox") + 1).comment is not None  # explanatory tooltip
    assert ws["A1"].fill.fgColor.rgb.endswith("1F4E78")
    assert ws.cell(1, header.index("BuyBox_Change") + 1).fill.fgColor.rgb.endswith("7F6000")
    # values, not formulas
    for row in ws.iter_rows(min_row=2):
        for c in row:
            assert not (isinstance(c.value, str) and c.value.startswith("="))
    comments = [r[header.index("Comment")] for r in ws.iter_rows(min_row=2, values_only=True)]
    assert comments == list(cur.rows["Comment"])


def test_seller_report_total_row():
    _, wb = _build()
    ws = wb["BuyBox_Seller_Report"]
    names = [ws.cell(r, 2).value for r in range(1, ws.max_row + 1)]
    assert L.TOTAL_ROW_LABEL in names
    r = names.index(L.TOTAL_ROW_LABEL) + 1
    assert ws.cell(r, 2).font.b
    assert ws["A1"].value == L.EXCEL_TITLE
    assert ws["B3"].value == 0.01  # threshold parameter


def test_workbook_without_prev_has_no_compare_columns():
    _, wb = _build(has_prev=False)
    header = [c.value for c in wb["Sheet1"][1]]
    assert "BuyBox_Change" not in header
    heads = [c.value for c in wb["BuyBox_Seller_Report"][5]]
    assert L.SELLER_COLUMNS["Prev_Comp_Win_Rate"] not in heads


def test_quickwin_workbook():
    a = run([(1, 1, 10, 1, 100, "A"), (2, 1, 20, 1, 101, "B/C")])
    nm = a.rows[a.rows["Near_Miss"]]
    data = build_quickwin_workbook(nm, "B/C", Params())
    wb = openpyxl.load_workbook(io.BytesIO(data))
    ws = wb["QuickWin"]
    assert ws.sheet_view.rightToLeft
    (table,) = ws.tables.values()  # an Excel table with filter buttons
    first = ws[table.ref.split(":")[0]]
    header = [c.value for c in ws[first.row]]
    header = [h for h in header if h]
    # Persian title + the English name required by the spec, in each header cell
    assert [h.split("\n")[1] for h in header] == ["DKPC", "DKP", "Weight", "Selling_Price", "BuyBox_Min_Price",
                                                  "Gap", "Target_Price", "Required_Cut_%", "Margin_Flag"]
    assert header[0] == "شناسه تنوع\nDKPC"
    assert ws.cell(first.row + 1, 1).value == 2
    name = quickwin_filename("B/C")
    assert name.startswith("QuickWin_B_C_") and name.endswith(".xlsx")


def test_main_workbook_splits_rows_beyond_sheet_limit(monkeypatch):
    import buybox.export as ex
    monkeypatch.setattr(ex, "MAX_SHEET_ROWS", 4)
    cur, wb = _build()
    assert [n for n in wb.sheetnames if n.startswith("Sheet1")] == ["Sheet1", "Sheet1 (2)", "Sheet1 (3)"]
    total = sum(wb[n].max_row - 1 for n in wb.sheetnames if n.startswith("Sheet1"))
    assert total == len(cur.rows)


def test_roundtrip_through_excel_reader():
    raw = make_raw(CUR)
    buf = io.BytesIO()
    raw.to_excel(buf, index=False)
    back = read_excel_bytes(buf.getvalue())
    assert list(back.columns) == list(raw.columns) and len(back) == len(raw)
