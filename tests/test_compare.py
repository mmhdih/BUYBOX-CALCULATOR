import pandas as pd
import pytest

from buybox import compare
from buybox import labels as L
from buybox.core import Params, seller_metrics
from buybox.history import list_runs, load_run, save_run, trend_frame
from buybox.pipeline import attach_comparison
from tests.conftest import run

PREV = [
    (1, 1, 10, 1, 100, "A"), (2, 1, 20, 1, 110, "B"),   # A wins
    (3, 2, 10, 1, 300, "A"), (4, 2, 20, 1, 200, "B"),   # B wins
    (5, 3, 10, 1, 500, "A"),                             # A alone, wins
    (6, 4, 20, 1, 700, "B"), (7, 4, 10, 1, 800, "A"),    # B wins
    (8, 5, 10, None, 100, "A"),                          # no weight before
    (9, 6, 20, 1, 50, "B"),                              # removed later
]
CUR = [
    (1, 1, 10, 1, 120, "A"), (2, 1, 20, 1, 110, "B"),   # A lost, B gained
    (3, 2, 10, 1, 300, "A"), (4, 2, 20, 1, 200, "B"),   # unchanged
    (5, 3, 10, None, 500, "A"),                          # weight now blank
    (6, 4, 20, 1, 700, "B"), (7, 4, 10, 1, 800, "A"),    # unchanged
    (8, 5, 10, 1, 100, "A"), (10, 5, 20, 1, 90, "B"),   # 8: was not evaluated, now loses; 10 new
]


def _setup():
    prev = run(PREV)
    cur = run(CUR)
    comp = attach_comparison(cur, prev.rows, "prev")
    return cur, prev, comp


def test_all_change_labels_and_removed():
    cur, prev, comp = _setup()
    r = cur.rows.set_index("DKPC")
    assert r.loc[1, "BuyBox_Change"] == L.CHANGE_LOST
    assert r.loc[2, "BuyBox_Change"] == L.CHANGE_GAINED
    assert r.loc[3, "BuyBox_Change"] == L.CHANGE_SAME
    assert r.loc[4, "BuyBox_Change"] == L.CHANGE_SAME
    assert r.loc[5, "BuyBox_Change"] == L.CHANGE_NO_WEIGHT
    assert r.loc[8, "BuyBox_Change"] == L.CHANGE_SAME  # had no buy box, still none (not "lost")
    assert r.loc[10, "BuyBox_Change"] == L.CHANGE_NEW
    assert r.loc[10, "Comment_Prev_File"] == ""
    assert r.loc[1, "Comment_Prev_File"] == L.COMMENT_WIN
    assert r.loc[8, "Comment_Prev_File"] == L.COMMENT_NO_WEIGHT
    assert list(comp.removed["DKPC"]) == [9]


def test_market_and_seller_changes():
    cur, prev, comp = _setup()
    m = compare.market_changes(cur.rows, prev.rows, comp.removed)
    assert (m["gained"], m["lost"], m["new"], m["removed"]) == (1, 1, 1, 1)
    # prev competitive: rows 1,2,3,4,6,7 -> wins 1,4,6 => 3/6 ; cur: 1,2,3,4,6,7,8,10 -> wins 2,4,6,10 => 4/8
    assert m["prev_comp_win_rate"] == pytest.approx(0.5)
    assert m["cur_comp_win_rate"] == pytest.approx(0.5)
    ch = compare.seller_changes(cur.rows)
    assert ch.loc[10, "Lost"] == 1 and ch.loc[20, "Gained"] == 1
    winners, losers = compare.net_movers(ch, cur.seller_info)
    assert list(winners["Seller_Name"]) == ["B"] and list(losers["Seller_Name"]) == ["A"]
    t = seller_metrics(cur.rows, cur.seller_info, compare.prev_comp_by_seller(prev.rows), ch)
    a = t.set_index("Seller_Name").loc["A"]
    # A prev: comp rows 1,3,7 -> wins 1 => 1/3 ; cur: comp rows 1,3,7,8 -> 0 wins
    assert a["Prev_Comp_Win_Rate"] == pytest.approx(1 / 3)
    assert a["Comp_Win_Rate_Change"] == pytest.approx(-100 / 3)
    assert a["Gained"] == 0 and a["Lost"] == 1
    total = t[t["Is_Total"]].iloc[0]
    assert total["Gained"] == 1 and total["Lost"] == 1
    assert total["Prev_Comp_Win_Rate"] == pytest.approx(0.5)


def test_lost_to_whom():
    cur, _, _ = _setup()
    t = compare.lost_to_whom(cur.rows, 10)
    assert dict(zip(t["BuyBox_Winner_Seller"], t["Count"])) == {"B": 4}


def test_string_vs_int_dkpc_matching():
    prev = run([("1", 1, 10, 1, 100, "A")])
    cur = run([(1, 1, 10, 1, 100, "A")])
    attach_comparison(cur, prev.rows, "p")
    assert cur.rows.iloc[0]["BuyBox_Change"] == L.CHANGE_SAME


def test_history_roundtrip(tmp_path):
    cur, prev, _ = _setup()
    info = save_run(prev.rows, Params(), "اسورتمنت شمش.xlsx", tmp_path, file_hash="abc")
    assert info.parquet.name.endswith("_اسورتمنت شمش.parquet")
    runs = list_runs(tmp_path)
    assert len(runs) == 1
    loaded = load_run(runs[0])
    cur2 = run(CUR)
    attach_comparison(cur2, loaded, "run")
    pd.testing.assert_series_equal(cur2.rows["BuyBox_Change"], cur.rows["BuyBox_Change"])
    tr = trend_frame(runs)
    assert tr.iloc[0]["comp_win_rate"] == pytest.approx(0.5)
    tr_a = trend_frame(runs, [10])
    assert tr_a.iloc[0]["comp_win_rate"] == pytest.approx(1 / 3)
