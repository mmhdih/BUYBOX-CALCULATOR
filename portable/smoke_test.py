"""Smoke test run with the *embedded* Python of a built portable package.

    <package>\\python\\python.exe portable\\smoke_test.py <package>\\app

Exercises every heavy dependency (pandas, numpy, pyarrow, calamine, openpyxl, xlsxwriter,
plotly, streamlit import) through the real pipeline.
"""

from __future__ import annotations

import io
import sys
import tempfile
from pathlib import Path

app_dir = Path(sys.argv[1]).resolve()
sys.path.insert(0, str(app_dir))

import pandas as pd  # noqa: E402
import plotly.express  # noqa: E402,F401
import streamlit  # noqa: E402,F401

from buybox import labels as L  # noqa: E402
from buybox.core import Params, market_summary, seller_metrics  # noqa: E402
from buybox.export import build_main_workbook, build_quickwin_workbook  # noqa: E402
from buybox.history import list_runs, load_run, save_run  # noqa: E402
from buybox.io import detect_columns, read_excel_bytes  # noqa: E402
from buybox.pipeline import analyze, attach_comparison  # noqa: E402
from buybox.quality import checks_to_frame, run_checks  # noqa: E402

raw = pd.DataFrame({
    "DKPC": [1, 2, 3, 4], "DKP": [10, 10, 11, 11], "Seller_ID": [1, 5, 5, 6],
    "Weight": [1.0, 1.0, 0.5, None], "Selling_Price": [100, 101, 50, 60],
    "Product_Name": ["شمش ۱ گرمی", "شمش ۱ گرمی", "سکه", "سکه"],
})
buf = io.BytesIO()
raw.to_excel(buf, index=False)
raw2 = read_excel_bytes(buf.getvalue())  # calamine
a = analyze(raw2, detect_columns(list(raw2.columns)).mapping, None, Params(), "smoke.xlsx")
s = market_summary(a.rows)
assert s["wins"] == 2 and s["near_miss"] == 1 and s["no_weight"] == 1, s
assert a.seller_info.loc[1, "Seller_Name"] == L.DIGIKALA_RETAIL
with tempfile.TemporaryDirectory() as tmp:
    run = save_run(a.rows, Params(), "smoke.xlsx", Path(tmp), file_hash="x")  # pyarrow
    prev = load_run(list_runs(Path(tmp))[0])
    attach_comparison(a, prev, run.label)
table = seller_metrics(a.rows, a.seller_info)
q = checks_to_frame(run_checks(a.rows, a.duplicate_rows, a.seller_info, None, a.original_columns))
xlsx = build_main_workbook(a.raw, a.rows, table, None, q, Params(), "smoke.xlsx", True, "prev")
qw = build_quickwin_workbook(a.rows[a.rows["Near_Miss"]], "test", Params())
import openpyxl  # noqa: E402

assert openpyxl.load_workbook(io.BytesIO(xlsx)).sheetnames[0] == "BuyBox_Seller_Report"
assert len(qw) > 1000
print("SMOKE TEST OK", sys.version.split()[0], "pandas", pd.__version__, "streamlit", streamlit.__version__)
