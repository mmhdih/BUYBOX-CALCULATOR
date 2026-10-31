import sys
from pathlib import Path

import pandas as pd
import pytest

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from buybox.core import Params  # noqa: E402
from buybox.io import detect_columns  # noqa: E402
from buybox.pipeline import analyze  # noqa: E402

FIXTURES = ROOT / "tests" / "fixtures"


def make_raw(records: list[tuple]) -> pd.DataFrame:
    """records: (DKPC, DKP, Seller_ID, Weight, Selling_Price[, Seller_Name])"""
    cols = ["DKPC", "DKP", "Seller_ID", "Weight", "Selling_Price", "Seller_Name"]
    rows = [r + (None,) * (6 - len(r)) for r in records]
    return pd.DataFrame(rows, columns=cols)


def run(records: list[tuple], params: Params | None = None, kam=None):
    raw = make_raw(records)
    mapping = detect_columns(list(raw.columns)).mapping
    return analyze(raw, mapping, kam, params or Params(), "test.xlsx")


@pytest.fixture
def analyze_records():
    return run
