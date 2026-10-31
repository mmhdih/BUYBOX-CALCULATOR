"""Regression test against the real exports (files are NOT committed).

Place the files in ``tests/fixtures/`` to enable:
  - ``اسورتمنت شمش 14050802.xlsx`` (current)
  - ``اسورتمنت شمش.xlsx`` (previous)
"""

import pytest

from buybox import compare
from buybox.core import Params, comp_by_seller, market_summary
from buybox.io import detect_columns, read_excel_bytes
from buybox.pipeline import analyze, attach_comparison
from tests.conftest import FIXTURES

CUR_FILE = FIXTURES / "اسورتمنت شمش 14050802.xlsx"
PREV_FILE = FIXTURES / "اسورتمنت شمش.xlsx"


def _load(path):
    raw = read_excel_bytes(path.read_bytes())
    m = detect_columns(list(raw.columns))
    assert m.ok, m.missing_required
    return analyze(raw, m.mapping, None, Params(near_miss_threshold=0.01), path.name)


@pytest.fixture(scope="module")
def cur():
    if not CUR_FILE.exists():
        pytest.skip(f"fixture missing: {CUR_FILE.name}")
    return _load(CUR_FILE)


@pytest.fixture(scope="module")
def prev():
    if not PREV_FILE.exists():
        pytest.skip(f"fixture missing: {PREV_FILE.name}")
    return _load(PREV_FILE)


def test_current_file(cur):
    s = market_summary(cur.rows)
    assert s["rows"] == 21_461
    assert s["no_weight"] == 161
    assert s["groups"] == 18_217
    assert s["wins"] == 18_218
    assert s["comp_rows"] == 4_789
    assert s["comp_wins"] == 1_711
    assert round(s["comp_win_rate"] * 100, 1) == 35.7
    assert s["losses"] == 3_082
    assert s["near_miss"] == 1_263
    assert s["tie_groups"] == 1


def test_previous_file(prev):
    s = market_summary(prev.rows)
    assert s["rows"] == 20_808
    assert s["wins"] == 18_013
    assert s["comp_rows"] == 4_419
    assert s["comp_wins"] == 1_628
    assert round(s["comp_win_rate"] * 100, 1) == 36.8


def test_comparison(cur, prev):
    comp = attach_comparison(cur, prev.rows, prev.input_name)
    m = compare.market_changes(cur.rows, prev.rows, comp.removed)
    assert (m["gained"], m["lost"], m["new"], m["removed"]) == (630, 708, 1_240, 587)


@pytest.mark.parametrize("seller_id, comp_rows, comp_wins", [
    (1737909, 748, 379),   # سکه شیخ
    (795353, 386, 96),     # مدوپد
    (1, 67, 50),           # دیجی‌کالا
])
def test_seller_spot_checks(cur, seller_id, comp_rows, comp_wins):
    t = comp_by_seller(cur.rows)
    assert int(t.loc[seller_id, "Comp_Variants"]) == comp_rows
    assert int(t.loc[seller_id, "Comp_Wins"]) == comp_wins
