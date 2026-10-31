"""Generate synthetic sample files (fake data) for trying the app.

Usage:  python scripts/make_sample_data.py [out_dir]
Writes: sample_assortment_prev.xlsx, sample_assortment_cur.xlsx, sample_kam_list.xlsx
"""

from __future__ import annotations

import sys
from pathlib import Path

import numpy as np
import pandas as pd


def make(n_dkp: int = 1500, seed: int = 7) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    rng = np.random.default_rng(seed)
    sellers = pd.DataFrame({
        "Seller_ID": [1] + list(rng.choice(np.arange(100_000, 2_500_000), 59, replace=False)),
    })
    sellers["Seller_Name"] = ["دیجی‌کالا"] + [f"فروشنده نمونه {i}" for i in range(1, 60)]
    brands = ["Parsian", "Melli Gold", "Zar", "Shams", "Sekeh Shiekh", "Aram"]
    cats = ["شمش طلا", "سکه پارسیان (گرمی)", "پک شمش و پلاک طلا"]
    weights = [0.05, 0.1, 0.2, 0.25, 0.5, 1.0, 1.5, 2.0, 2.5, 5.0, 10.0]
    recs = []
    dkpc = 50_000_000
    seller_weights = rng.dirichlet(np.ones(len(sellers)) * 0.6)
    for dkp in range(10_000_000, 10_000_000 + n_dkp):
        cat = rng.choice(cats)
        brand = rng.choice(brands)
        for w in rng.choice(weights, rng.integers(1, 4), replace=False):
            base = 72_000_000 * w * (1 + rng.uniform(0.04, 0.12))
            k = 1 if rng.random() < 0.85 else rng.integers(2, 5)
            for s in rng.choice(len(sellers), k, p=seller_weights):
                dkpc += 1
                price = round(base * (1 + rng.normal(0, 0.012)), -3)
                weight = np.nan if rng.random() < 0.008 else w
                recs.append({
                    "DKPC": dkpc, "DKP": dkp, "Product_Name": f"{cat} {brand} {w:g} گرمی",
                    "Brand_Name_En": brand, "Leaf_Category_Name": cat,
                    "Seller_ID": int(sellers.Seller_ID[s]), "Seller_Name": sellers.Seller_Name[s],
                    "Weight": weight, "Selling_Price": int(price), "RRP_Price": int(price * 1.02),
                    "True_Wage": round(rng.uniform(0.03, 0.12), 3),
                })
    prev = pd.DataFrame(recs)
    cur = prev.sample(frac=0.97, random_state=seed).copy()
    moved = cur.sample(frac=0.25, random_state=seed + 1).index
    cur.loc[moved, "Selling_Price"] = (cur.loc[moved, "Selling_Price"] * rng.normal(1, 0.01, len(moved))).round(-3)
    new = prev.sample(frac=0.05, random_state=seed + 2).copy()
    new["DKPC"] = np.arange(90_000_000, 90_000_000 + len(new))
    cur = pd.concat([cur, new]).sort_values(["DKP", "DKPC"]).reset_index(drop=True)
    cur["Selling_Price"] = cur["Selling_Price"].astype("int64")
    prev = prev.drop(columns=["Seller_Name"])  # older schema had no Seller_Name

    kams = ["علی رضایی", "مریم احمدی", "سارا محمدی"]
    kam = sellers.iloc[1:50].copy()
    kam["Category Level1 Name Fa"] = "طلا"
    kam["KAM"] = rng.choice(kams, len(kam))
    dup = kam.iloc[[0]].copy()
    dup["Seller_Name"] = dup["Seller_Name"] + " (نام دوم)"
    dup["KAM"] = "سارا محمدی" if kam.iloc[0]["KAM"] != "سارا محمدی" else "علی رضایی"
    kam = pd.concat([kam, dup]).rename(columns={"Seller_ID": "Seller Id", "Seller_Name": "Seller Name"})
    return prev, cur, kam


def main() -> None:
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("sample_data")
    out.mkdir(parents=True, exist_ok=True)
    prev, cur, kam = make()
    prev.to_excel(out / "sample_assortment_prev.xlsx", index=False)
    cur.to_excel(out / "sample_assortment_cur.xlsx", index=False)
    with pd.ExcelWriter(out / "sample_kam_list.xlsx") as xw:
        kam.to_excel(xw, sheet_name="KA Selection", index=False)
    print(f"wrote {len(prev)} / {len(cur)} rows and {len(kam)} KAM rows to {out}")


if __name__ == "__main__":
    main()
