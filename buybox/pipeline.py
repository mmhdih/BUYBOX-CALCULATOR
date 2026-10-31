"""Orchestration of the full analysis (no Streamlit): input frame -> analysis bundle."""

from __future__ import annotations

from dataclasses import dataclass, field

import pandas as pd

from buybox import compare
from buybox.core import Params, compute_buybox
from buybox.io import build_seller_info, build_work_frame, id_key, parse_actual_buybox, split_kams

#: Columns persisted per DKPC in saved runs (enough to act as a "previous file").
HISTORY_COLUMNS: tuple[str, ...] = (
    "DKPC", "DKP", "Seller_ID", "Seller_Name", "KAM", "Weight", "Selling_Price", "Product_Name",
    "Brand_Name_En", "Leaf_Category_Name", "BuyBox_Min_Price", "Gap_to_BuyBox", "Competition_Type",
    "Comment", "BuyBox_Winner_Seller", "Is_Evaluated", "Is_Win", "Is_Loss", "Is_Competitive",
    "Near_Miss",
)


@dataclass
class Analysis:
    """Everything derived from one assortment file."""

    raw: pd.DataFrame                 #: original columns, duplicate DKPCs removed
    rows: pd.DataFrame                #: canonical columns + buy-box result + names
    seller_info: pd.DataFrame         #: indexed by Seller_ID
    duplicate_rows: pd.DataFrame      #: dropped duplicate-DKPC rows (original columns)
    mapping: dict[str, str | None]
    params: Params
    input_name: str = ""
    original_columns: list[str] = field(default_factory=list)


def analyze(
    raw: pd.DataFrame,
    mapping: dict[str, str | None],
    kam_merged: pd.DataFrame | None,
    params: Params,
    input_name: str = "",
) -> Analysis:
    """Run the full per-row analysis on an already-read assortment frame."""
    original_columns = [str(c) for c in raw.columns]
    raw = raw.copy()
    raw.columns = original_columns
    dkpc_col = mapping["DKPC"]
    key = id_key(raw[dkpc_col])
    dup_mask = key.notna() & key.duplicated(keep="first")
    duplicates = raw[dup_mask]
    raw = raw[~dup_mask].reset_index(drop=True)

    work = build_work_frame(raw, mapping)
    seller_info = build_seller_info(work, kam_merged)
    result = compute_buybox(work, seller_info["Seller_Name"], params)

    rows = pd.concat([work.drop(columns=["Seller_Name"], errors="ignore"), result], axis=1)
    rows["Seller_Name"] = rows["Seller_ID"].map(seller_info["Seller_Name"]).fillna(rows["Seller_ID"].astype(str))
    rows["KAM"] = rows["Seller_ID"].map(seller_info["KAM"])
    if "Actual_BuyBox" in rows.columns:
        rows["Actual_BuyBox"] = parse_actual_buybox(rows["Actual_BuyBox"])
    return Analysis(raw=raw, rows=rows, seller_info=seller_info, duplicate_rows=duplicates,
                    mapping=dict(mapping), params=params, input_name=input_name,
                    original_columns=original_columns)


@dataclass
class Comparison:
    """Comparison of an analysis with a previous period."""

    prev_rows: pd.DataFrame
    removed: pd.DataFrame
    prev_label: str
    prev_columns: list[str] | None = None


def attach_comparison(analysis: Analysis, prev_rows: pd.DataFrame, prev_label: str,
                      prev_columns: list[str] | None = None) -> Comparison:
    """Add ``Comment_Prev_File``/``BuyBox_Change`` to ``analysis.rows`` and return the comparison."""
    ch = compare.row_changes(analysis.rows, prev_rows)
    analysis.rows["Comment_Prev_File"] = ch["Comment_Prev_File"]
    analysis.rows["BuyBox_Change"] = ch["BuyBox_Change"]
    removed = compare.removed_rows(analysis.rows, prev_rows)
    return Comparison(prev_rows=prev_rows, removed=removed, prev_label=prev_label, prev_columns=prev_columns)


def history_frame(rows: pd.DataFrame) -> pd.DataFrame:
    """Subset of rows persisted for a saved run."""
    cols = [c for c in HISTORY_COLUMNS if c in rows.columns]
    return rows[cols].reset_index(drop=True)


# ---------------------------------------------------------------------------
# Filters
# ---------------------------------------------------------------------------
def kam_options(seller_info: pd.DataFrame) -> list[str]:
    """Distinct individual KAM names (joined KAM values are split)."""
    names: set[str] = set()
    for v in seller_info["KAM"].dropna().unique():
        names.update(split_kams(v))
    return sorted(names)


def sellers_for_kams(seller_info: pd.DataFrame, kams: list[str]) -> pd.Index:
    """Seller IDs assigned (possibly jointly) to any of the given KAMs."""
    wanted = set(kams)
    hit = seller_info["KAM"].map(lambda v: bool(wanted.intersection(split_kams(v))))
    return seller_info.index[hit.astype(bool).to_numpy()]


def filter_mask(
    rows: pd.DataFrame,
    seller_info: pd.DataFrame,
    kams: list[str] | None = None,
    categories: list[str] | None = None,
    brands: list[str] | None = None,
    sellers: list[object] | None = None,
) -> pd.Series:
    """Boolean mask of rows matching all non-empty filters."""
    m = pd.Series(True, index=rows.index)
    if kams:
        m &= rows["Seller_ID"].isin(sellers_for_kams(seller_info, kams))
    if categories and "Leaf_Category_Name" in rows.columns:
        m &= rows["Leaf_Category_Name"].isin(categories)
    if brands and "Brand_Name_En" in rows.columns:
        m &= rows["Brand_Name_En"].isin(brands)
    if sellers:
        m &= rows["Seller_ID"].isin(sellers)
    return m.fillna(False).astype(bool)
