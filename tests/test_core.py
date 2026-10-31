import time

import numpy as np
import pandas as pd
import pytest

from buybox import labels as L
from buybox.core import Params, market_summary, pooled_rate, seller_metrics
from buybox.io import merge_kam, normalize_name, detect_columns
from tests.conftest import run


def _row(a, dkpc):
    return a.rows.set_index("DKPC").loc[dkpc]


def test_simple_win_loss():
    a = run([(1, 100, 10, 1.0, 1000, "A"), (2, 100, 20, 1.0, 1100, "B")])
    w, l = _row(a, 1), _row(a, 2)
    assert w.Comment == L.COMMENT_WIN and l.Comment == L.COMMENT_LOSS
    assert w.BuyBox_Min_Price == 1000 and l.BuyBox_Min_Price == 1000
    assert w.Gap_to_BuyBox == 0
    assert l.Gap_to_BuyBox == pytest.approx(0.1)
    assert l.BuyBox_Winner_Seller == "A" and w.BuyBox_Winner_Seller == ""
    assert w.Competition_Type == l.Competition_Type == L.COMP_COMPETITIVE
    assert w.Other_Sellers_Variants == 1 and w.Same_Seller_Variants == 1
    assert l.Target_Price == 1000
    assert l["Required_Cut_%"] == pytest.approx(1 - 1000 / 1100)
    assert l.Margin_Flag == L.MARGIN_FLAG  # 9.1% > 1%
    assert not l.Near_Miss  # 10% > 1%


def test_tie_between_two_sellers_both_win():
    a = run([(1, 100, 10, 2.0, 500, "A"), (2, 100, 20, 2.0, 500, "B"), (3, 100, 30, 2.0, 600, "C")])
    assert _row(a, 1).Comment == L.COMMENT_WIN
    assert _row(a, 2).Comment == L.COMMENT_WIN
    assert _row(a, 3).BuyBox_Winner_Seller == "A / B"
    s = market_summary(a.rows)
    assert s["tie_groups"] == 1 and s["wins"] == 2 and s["groups"] == 1


def test_same_seller_two_dkpcs_only_cheaper_wins():
    a = run([(1, 100, 10, 1.0, 1000, "A"), (2, 100, 10, 1.0, 1050, "A")])
    cheap, dear = _row(a, 1), _row(a, 2)
    assert cheap.Comment == L.COMMENT_WIN and dear.Comment == L.COMMENT_LOSS
    assert cheap.Competition_Type == dear.Competition_Type == L.COMP_SINGLE_DUP
    assert dear.Same_Seller_Variants == 2 and dear.Other_Sellers_Variants == 0
    assert not dear.Is_Competitive


def test_blank_weight_excluded_from_denominators():
    a = run([
        (1, 100, 10, 1.0, 1000, "A"), (2, 100, 20, 1.0, 1100, "B"),
        (3, 100, 10, None, 10, "A"), (4, 100, 20, np.nan, 99999, "B"),
    ])
    r3 = _row(a, 3)
    assert r3.Comment == L.COMMENT_NO_WEIGHT and r3.Competition_Type == L.COMP_NO_WEIGHT
    assert not r3.Is_Evaluated
    # blank-weight rows must not be grouped by DKP alone (row 3's price would otherwise win)
    assert _row(a, 1).Comment == L.COMMENT_WIN
    t = seller_metrics(a.rows, a.seller_info).set_index("Seller_Name")
    assert t.loc["A", "Evaluated"] == 1 and t.loc["A", "Not_Evaluated"] == 1 and t.loc["A", "Variants"] == 2
    assert t.loc["A", "Win_Rate"] == 1.0 and t.loc["B", "Win_Rate"] == 0.0
    s = market_summary(a.rows)
    assert s["evaluated"] == 2 and s["no_weight"] == 2


def test_invalid_price_excluded():
    a = run([(1, 100, 10, 1.0, 0, "A"), (2, 100, 20, 1.0, 1100, "B")])
    assert not _row(a, 1).Is_Evaluated
    assert _row(a, 1).Comment == L.COMMENT_BAD_PRICE
    assert _row(a, 2).Comment == L.COMMENT_WIN
    assert _row(a, 2).Competition_Type == L.COMP_SINGLE


def test_single_seller_group():
    a = run([(1, 100, 10, 5.0, 1000, "A")])
    r = _row(a, 1)
    assert r.Comment == L.COMMENT_WIN and r.Competition_Type == L.COMP_SINGLE
    t = seller_metrics(a.rows, a.seller_info)
    assert np.isnan(t.iloc[0]["Comp_Win_Rate"])  # displayed as "-"
    assert t.iloc[0]["Single_Seller_Variants"] == 1


def test_near_miss_boundary_inclusive():
    a = run([
        (1, 100, 10, 1.0, 100_000_000, "A"),
        (2, 100, 20, 1.0, 101_000_000, "B"),   # exactly 1% -> near miss
        (3, 100, 30, 1.0, 101_000_001, "C"),   # just above -> not near miss
    ], Params(near_miss_threshold=0.01))
    assert _row(a, 2).Near_Miss
    assert not _row(a, 3).Near_Miss
    assert _row(a, 2).Margin_Flag == ""  # required cut 0.99% <= 1%
    b = run([(1, 7, 10, 1.0, 100, "A"), (2, 7, 20, 1.0, 101, "B")])  # 101/100-1 = 0.010000000000000009
    assert _row(b, 2).Near_Miss


def test_float_weight_mismatch():
    a = run([(1, 100, 10, 0.30000000004, 1000, "A"), (2, 100, 20, 0.3, 1100, "B")])
    assert _row(a, 1).Group_ID == _row(a, 2).Group_ID
    assert _row(a, 2).Comment == L.COMMENT_LOSS
    assert _row(a, 2).Competition_Type == L.COMP_COMPETITIVE


def test_different_weights_are_different_groups():
    a = run([(1, 100, 10, 1.0, 1000, "A"), (2, 100, 20, 2.0, 1100, "B")])
    assert _row(a, 1).Comment == _row(a, 2).Comment == L.COMMENT_WIN


def test_undercut_and_margin_guard():
    a = run([(1, 1, 10, 1.0, 1_000_000, "A"), (2, 1, 20, 1.0, 1_005_000, "B")],
            Params(near_miss_threshold=0.01, undercut=1000, max_acceptable_cut=0.005))
    r = _row(a, 2)
    assert r.Target_Price == 999_000
    assert r["Required_Cut_%"] == pytest.approx(1 - 999_000 / 1_005_000)
    assert r.Margin_Flag == L.MARGIN_FLAG


def test_cut_vs_wage():
    raw = pd.DataFrame({"DKPC": [1, 2], "DKP": [1, 1], "Seller_ID": [10, 20], "Weight": [1, 1],
                        "Selling_Price": [100, 110], "True_Wage": [0.1, 0.2]})
    from buybox.pipeline import analyze
    a = analyze(raw, detect_columns(list(raw.columns)).mapping, None, Params(), "x")
    r = a.rows.set_index("DKPC").loc[2]
    assert r.Cut_vs_Wage == pytest.approx((1 - 100 / 110) / 0.2)


def test_column_matching_case_space_underscore():
    m = detect_columns(["dkpc", "D K P", "seller id", "WEIGHT", "Selling Price", "seller_name"])
    assert m.ok
    assert m.mapping["Seller_ID"] == "seller id" and m.mapping["Selling_Price"] == "Selling Price"
    assert normalize_name(" Seller_ Id ") == "sellerid"
    m2 = detect_columns(["DKPC", "DKP", "Seller_ID", "Selling_Price"])
    assert m2.missing_required == ["Weight"]


def test_kam_duplicate_merge():
    kam = pd.DataFrame({
        "Seller_ID": pd.array([2016514, 2016514, 5, 6], dtype="Int64"),
        "Seller_Name": ["نام اول", "نام دوم", "X", "Y"],
        "Category": ["طلا", "طلا", "طلا", "طلا"],
        "KAM": ["علی", "مریم", "علی", "رضا"],
    })
    m = merge_kam(kam).set_index("Seller_ID")
    assert len(m) == 3
    assert m.loc[2016514, "Seller_Name"] == "نام اول / نام دوم"
    assert m.loc[2016514, "KAM"] == "علی / مریم"
    assert m.loc[2016514, "KAM_Note"] == L.KAM_DUPLICATE_FLAG
    assert m.loc[5, "KAM_Note"] == ""
    a = run([(1, 1, 2016514, 1, 10), (2, 1, 99, 1, 11), (3, 1, 1, 1, 12)], kam=merge_kam(kam))
    info = a.seller_info
    assert info.loc[2016514, "Seller_Name"] == "نام اول / نام دوم"  # no name in assortment -> KAM list
    assert info.loc[99, "KAM"] == L.KAM_NOT_IN_LIST
    assert info.loc[1, "Seller_Name"] == L.DIGIKALA_RETAIL
    from buybox.pipeline import filter_mask, kam_options
    assert kam_options(info) == sorted(["علی", "مریم", L.KAM_NOT_IN_LIST])
    assert filter_mask(a.rows, info, kams=["مریم"]).sum() == 1


def test_pooled_vs_averaged_win_rate():
    # Seller A: 1 competitive row, wins. Seller B: 3 competitive rows, wins 0.
    a = run([
        (1, 1, 10, 1, 100, "A"), (2, 1, 20, 1, 200, "B"),
        (3, 2, 20, 1, 300, "B"), (4, 2, 30, 1, 100, "C"),
        (5, 3, 20, 1, 300, "B"), (6, 3, 30, 1, 100, "C"),
    ])
    t = seller_metrics(a.rows, a.seller_info)
    per = t[~t["Is_Total"]].set_index("Seller_Name")
    total = t[t["Is_Total"]].iloc[0]
    assert per.loc["A", "Comp_Win_Rate"] == 1.0
    assert per.loc["B", "Comp_Win_Rate"] == 0.0
    assert per.loc["C", "Comp_Win_Rate"] == 1.0
    assert total["Comp_Win_Rate"] == pytest.approx(3 / 6)
    averaged = per["Comp_Win_Rate"].mean()
    assert averaged == pytest.approx(2 / 3) and total["Comp_Win_Rate"] != pytest.approx(averaged)
    assert t.iloc[0]["Seller_Name"] == "B"  # sorted by variants desc
    assert total["Seller_Name"] == L.TOTAL_ROW_LABEL and total["Variants"] == 6
    assert pooled_rate(0, 0) != pooled_rate(0, 0)  # NaN


def test_duplicate_dkpc_keeps_first():
    a = run([(1, 1, 10, 1, 100, "A"), (1, 1, 20, 1, 50, "B"), (2, 1, 30, 1, 120, "C")])
    assert len(a.rows) == 2 and len(a.duplicate_rows) == 1
    assert _row(a, 1).Comment == L.COMMENT_WIN


def test_performance_25k_rows():
    rng = np.random.default_rng(0)
    n = 25_000
    dkp = rng.integers(0, 6000, n)
    raw = pd.DataFrame({
        "DKPC": np.arange(n), "DKP": dkp, "Seller_ID": rng.integers(1, 400, n),
        "Weight": rng.choice([0.1, 0.5, 1, 2.5, 5, 10, np.nan], n),
        "Selling_Price": rng.integers(10_000_000, 900_000_000, n),
        "Brand_Name_En": rng.choice(["A", "B", "C"], n),
        "Leaf_Category_Name": rng.choice(["شمش طلا", "سکه پارسیان (گرمی)"], n),
    })
    from buybox.pipeline import analyze
    from buybox.quality import run_checks
    t0 = time.perf_counter()
    a = analyze(raw, detect_columns(list(raw.columns)).mapping, None, Params(), "perf")
    seller_metrics(a.rows, a.seller_info)
    market_summary(a.rows)
    run_checks(a.rows, a.duplicate_rows, a.seller_info, None, a.original_columns)
    assert time.perf_counter() - t0 < 5
