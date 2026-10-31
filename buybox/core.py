"""Core buy-box logic (pure pandas, no Streamlit).

The buy box of a ``(DKP, Weight)`` group is the lowest ``Selling_Price`` in the group.
Everything here is vectorized; no row-wise ``apply``.
"""

from __future__ import annotations

from dataclasses import asdict, dataclass

import numpy as np
import pandas as pd

from buybox import labels as L

#: Tolerance for "<= threshold" comparisons on floating point ratios.
EPS = 1e-9
WEIGHT_DECIMALS = 4

#: Computed per-row output columns, in export order.
RESULT_COLUMNS: tuple[str, ...] = (
    "BuyBox_Min_Price",
    "Other_Sellers_Variants",
    "Same_Seller_Variants",
    "Gap_to_BuyBox",
    "Competition_Type",
    "Comment",
    "BuyBox_Winner_Seller",
    "Target_Price",
    "Required_Cut_%",
    "Cut_vs_Wage",
    "Margin_Flag",
)


@dataclass(frozen=True)
class Params:
    """User-tunable analysis parameters (fractions, not percents; undercut in Rial)."""

    near_miss_threshold: float = 0.01
    undercut: float = 0.0
    max_acceptable_cut: float = 0.01

    def as_dict(self) -> dict[str, float]:
        return asdict(self)


def compute_buybox(
    work: pd.DataFrame,
    seller_names: pd.Series | None = None,
    params: Params | None = None,
) -> pd.DataFrame:
    """Compute per-row buy-box status for a canonical working frame.

    Args:
        work: frame with at least ``DKPC, DKP, Seller_ID, Weight, Selling_Price``
            (``True_Wage`` optional).
        seller_names: ``Seller_ID -> display name`` used for ``BuyBox_Winner_Seller``.
        params: analysis parameters.

    Returns:
        Frame indexed like ``work`` holding :data:`RESULT_COLUMNS` plus internal boolean
        helper columns (``Is_Evaluated``, ``Is_Win``, ``Is_Loss``, ``Is_Competitive``,
        ``Near_Miss``, ``Tie_Group``, ``Same_Seller_Dup``, ``Group_ID``, ``Weight_Key``).
    """
    params = params or Params()
    n = len(work)
    idx = work.index
    price = work["Selling_Price"].astype("float64")
    weight = work["Weight"].astype("float64")

    weight_ok = weight.notna()
    price_ok = price.notna() & (price > 0)
    ids_ok = work["DKP"].notna() & work["Seller_ID"].notna()
    evaluated = weight_ok & price_ok & ids_ok

    out = pd.DataFrame(index=idx)
    out["Weight_Key"] = weight.round(WEIGHT_DECIMALS)

    min_price = pd.Series(np.nan, index=idx)
    other = pd.Series(pd.NA, index=idx, dtype="Int64")
    same = pd.Series(pd.NA, index=idx, dtype="Int64")
    group_id = pd.Series(-1, index=idx, dtype="int64")
    n_winners = pd.Series(0, index=idx, dtype="int64")
    winner_names = pd.Series("", index=idx, dtype=object)

    if evaluated.any():
        ev = pd.DataFrame(
            {
                "DKP": work.loc[evaluated, "DKP"],
                "Seller_ID": work.loc[evaluated, "Seller_ID"],
                "W": out.loc[evaluated, "Weight_Key"],
                "P": price[evaluated],
            }
        )
        gid = ev.groupby(["DKP", "W"], sort=False).ngroup()
        ev["G"] = gid
        g = ev.groupby("G", sort=False)["P"]
        ev_min = g.transform("min")
        ev_size = g.transform("size")
        ev_same = ev.groupby(["G", "Seller_ID"], sort=False)["P"].transform("size")
        ev_win = ev["P"] == ev_min
        ev_nwin = ev_win.groupby(ev["G"]).transform("sum")

        min_price[evaluated] = ev_min
        same[evaluated] = ev_same.astype("int64")
        other[evaluated] = (ev_size - ev_same).astype("int64")
        group_id[evaluated] = gid.astype("int64")
        n_winners[evaluated] = ev_nwin.astype("int64")

        # Winner seller name(s) per group, ties joined with " / ".
        winners = ev.loc[ev_win, ["G", "Seller_ID"]].drop_duplicates()
        names = winners["Seller_ID"].map(seller_names) if seller_names is not None else None
        if names is None:
            names = winners["Seller_ID"].astype(str)
        winners = winners.assign(N=names.fillna(winners["Seller_ID"].astype(str)).astype(str))
        multi = winners["G"].duplicated(keep=False)
        by_group = winners.loc[~multi].set_index("G")["N"]
        if multi.any():
            joined = winners.loc[multi].sort_values(["G", "N"]).groupby("G")["N"].agg(L.JOIN_SEP.join)
            by_group = pd.concat([by_group, joined])
        winner_names[evaluated] = ev["G"].map(by_group).to_numpy()

    is_win = evaluated & (price == min_price)
    is_loss = evaluated & ~is_win
    competitive = evaluated & (other.fillna(0) > 0)
    same_dup = evaluated & ~competitive & (same.fillna(0) > 1)

    gap = (price / min_price - 1).where(evaluated)
    gap = gap.where(~is_win, 0.0)

    comment = np.select(
        [is_win, is_loss, ~weight_ok, ~price_ok],
        [L.COMMENT_WIN, L.COMMENT_LOSS, L.COMMENT_NO_WEIGHT, L.COMMENT_BAD_PRICE],
        default=L.COMMENT_MISSING_ID,
    )
    comp_type = np.select(
        [competitive, same_dup, evaluated, ~weight_ok, ~price_ok],
        [L.COMP_COMPETITIVE, L.COMP_SINGLE_DUP, L.COMP_SINGLE, L.COMP_NO_WEIGHT, L.COMP_BAD_PRICE],
        default=L.COMP_MISSING_ID,
    )

    near_miss = is_loss & (gap > 0) & (gap <= params.near_miss_threshold + EPS)
    target = (min_price - params.undercut).where(is_loss)
    req_cut = (1 - target / price).where(is_loss)
    margin = np.where(is_loss & (req_cut > params.max_acceptable_cut + EPS), L.MARGIN_FLAG, "")

    out["BuyBox_Min_Price"] = min_price
    out["Other_Sellers_Variants"] = other
    out["Same_Seller_Variants"] = same
    out["Gap_to_BuyBox"] = gap
    out["Competition_Type"] = comp_type
    out["Comment"] = comment
    out["BuyBox_Winner_Seller"] = winner_names.where(is_loss, "")
    out["Target_Price"] = target
    out["Required_Cut_%"] = req_cut
    if "True_Wage" in work.columns:
        wage = work["True_Wage"].astype("float64")
        out["Cut_vs_Wage"] = (req_cut / wage.where(wage > 0)).where(is_loss)
    out["Margin_Flag"] = margin

    out["Is_Evaluated"] = evaluated
    out["Is_Win"] = is_win
    out["Is_Loss"] = is_loss
    out["Is_Competitive"] = competitive
    out["Near_Miss"] = near_miss
    out["Tie_Group"] = evaluated & (n_winners > 1)
    out["Same_Seller_Dup"] = evaluated & (same.fillna(0) > 1)
    out["Group_ID"] = group_id
    assert len(out) == n
    return out


def result_columns_present(rows: pd.DataFrame) -> list[str]:
    """The subset of :data:`RESULT_COLUMNS` present in ``rows`` (Cut_vs_Wage is optional)."""
    return [c for c in RESULT_COLUMNS if c in rows.columns]


# ---------------------------------------------------------------------------
# Aggregations
# ---------------------------------------------------------------------------
def pooled_rate(wins: float, total: float) -> float:
    """wins / total, NaN when total is 0."""
    return float(wins) / float(total) if total else float("nan")


def comp_counts(rows: pd.DataFrame) -> tuple[int, int]:
    """(competitive rows, competitive wins)."""
    comp = rows["Is_Competitive"].astype(bool)
    return int(comp.sum()), int((comp & rows["Is_Win"].astype(bool)).sum())


def market_summary(rows: pd.DataFrame) -> dict[str, float]:
    """Market-level KPIs for a (possibly filtered) set of rows."""
    ev = rows["Is_Evaluated"].astype(bool)
    comp, comp_wins = comp_counts(rows)
    wins = int(rows["Is_Win"].sum())
    groups = rows.loc[ev, "Group_ID"]
    n_groups = int(groups.nunique())
    single_groups = int(rows.loc[ev & (rows["Other_Sellers_Variants"].fillna(0) == 0), "Group_ID"].nunique())
    return {
        "rows": int(len(rows)),
        "unique_dkp": int(rows["DKP"].nunique()),
        "unique_sellers": int(rows["Seller_ID"].nunique()),
        "not_evaluated": int((~ev).sum()),
        "no_weight": int(rows["Weight"].isna().sum()),
        "evaluated": int(ev.sum()),
        "groups": n_groups,
        "single_seller_groups": single_groups,
        "single_seller_group_pct": pooled_rate(single_groups, n_groups),
        "wins": wins,
        "losses": int(rows["Is_Loss"].sum()),
        "win_rate": pooled_rate(wins, int(ev.sum())),
        "comp_rows": comp,
        "comp_wins": comp_wins,
        "comp_win_rate": pooled_rate(comp_wins, comp),
        "near_miss": int(rows["Near_Miss"].sum()),
        "tie_groups": int(rows.loc[rows["Tie_Group"], "Group_ID"].nunique()),
        "same_seller_dup_groups": int(rows.loc[rows["Same_Seller_Dup"], "Group_ID"].nunique()),
        "margin_flags": int((rows["Margin_Flag"] == L.MARGIN_FLAG).sum()),
    }


def _main_brand(rows: pd.DataFrame) -> pd.Series:
    if "Brand_Name_En" not in rows.columns:
        return pd.Series(dtype=object)
    b = rows.dropna(subset=["Brand_Name_En"])
    if b.empty:
        return pd.Series(dtype=object)
    counts = b.groupby(["Seller_ID", "Brand_Name_En"], observed=True).size().reset_index(name="n")
    counts = counts.sort_values(["Seller_ID", "n"], ascending=[True, False]).drop_duplicates("Seller_ID")
    return counts.set_index("Seller_ID")["Brand_Name_En"]


def seller_metrics(
    rows: pd.DataFrame,
    seller_info: pd.DataFrame,
    prev_comp: pd.DataFrame | None = None,
    changes: pd.DataFrame | None = None,
) -> pd.DataFrame:
    """Seller-level table (one row per Seller_ID) plus a pooled total row.

    Args:
        rows: per-DKPC rows (work columns + :func:`compute_buybox` output).
        seller_info: frame indexed by Seller_ID with ``Seller_Name`` and ``KAM``.
        prev_comp: optional frame indexed by Seller_ID with ``Comp_Variants``/``Comp_Wins``
            of the previous period.
        changes: optional frame indexed by Seller_ID with ``Gained``/``Lost`` counts.

    Returns:
        Frame with the internal keys of :data:`labels.SELLER_COLUMNS` and an ``Is_Total``
        flag; the last row is the total. Sorted by number of variants, descending.
    """
    ev = rows["Is_Evaluated"].astype(bool)
    win = rows["Is_Win"].astype(bool)
    comp = rows["Is_Competitive"].astype(bool)
    base = pd.DataFrame(
        {
            "Seller_ID": rows["Seller_ID"],
            "Variants": 1,
            "Evaluated": ev.astype(int),
            "Wins": win.astype(int),
            "Comp_Variants": comp.astype(int),
            "Comp_Wins": (comp & win).astype(int),
            "Near_Miss": rows["Near_Miss"].astype(int),
        }
    )
    t = base.groupby("Seller_ID").sum()
    t["Not_Evaluated"] = t["Variants"] - t["Evaluated"]
    t["Single_Seller_Variants"] = t["Evaluated"] - t["Comp_Variants"]
    t["No_BuyBox"] = t["Evaluated"] - t["Wins"]
    t["Win_Rate"] = t["Wins"] / t["Evaluated"].where(t["Evaluated"] > 0)
    t["Comp_Win_Rate"] = t["Comp_Wins"] / t["Comp_Variants"].where(t["Comp_Variants"] > 0)
    losses = rows.loc[rows["Is_Loss"].astype(bool)]
    t["Median_Loss_Gap"] = losses.groupby("Seller_ID")["Gap_to_BuyBox"].median().reindex(t.index)
    total_rows = len(rows)
    t["Share"] = t["Variants"] / total_rows if total_rows else np.nan
    t["Main_Brand"] = _main_brand(rows).reindex(t.index)
    t["Seller_Name"] = seller_info["Seller_Name"].reindex(t.index).fillna(pd.Series(t.index.astype(str), index=t.index))
    t["KAM"] = seller_info["KAM"].reindex(t.index).fillna(L.KAM_NOT_IN_LIST)

    has_prev = prev_comp is not None
    if has_prev:
        pc = prev_comp.reindex(t.index)
        t["Prev_Comp_Win_Rate"] = pc["Comp_Wins"] / pc["Comp_Variants"].where(pc["Comp_Variants"] > 0)
        t["Comp_Win_Rate_Change"] = (t["Comp_Win_Rate"] - t["Prev_Comp_Win_Rate"]) * 100
        ch = changes.reindex(t.index) if changes is not None else pd.DataFrame(index=t.index)
        t["Gained"] = ch.get("Gained", pd.Series(0, index=t.index)).fillna(0).astype(int)
        t["Lost"] = ch.get("Lost", pd.Series(0, index=t.index)).fillna(0).astype(int)

    t = t.sort_values(["Variants", "Comp_Variants"], ascending=False).reset_index()
    t["Is_Total"] = False

    tot: dict[str, object] = {
        "Seller_ID": pd.NA,
        "Seller_Name": L.TOTAL_ROW_LABEL,
        "KAM": "",
        "Main_Brand": "",
        "Is_Total": True,
    }
    for c in ("Variants", "Evaluated", "Wins", "Comp_Variants", "Comp_Wins", "Near_Miss",
              "Not_Evaluated", "Single_Seller_Variants", "No_BuyBox"):
        tot[c] = int(t[c].sum())
    tot["Win_Rate"] = pooled_rate(tot["Wins"], tot["Evaluated"])  # type: ignore[arg-type]
    tot["Comp_Win_Rate"] = pooled_rate(tot["Comp_Wins"], tot["Comp_Variants"])  # type: ignore[arg-type]
    tot["Median_Loss_Gap"] = float(losses["Gap_to_BuyBox"].median()) if len(losses) else np.nan
    tot["Share"] = 1.0 if total_rows else np.nan
    if has_prev:
        pc_all = prev_comp
        tot["Prev_Comp_Win_Rate"] = pooled_rate(pc_all["Comp_Wins"].sum(), pc_all["Comp_Variants"].sum())
        tot["Comp_Win_Rate_Change"] = (tot["Comp_Win_Rate"] - tot["Prev_Comp_Win_Rate"]) * 100  # type: ignore[operator]
        tot["Gained"] = int(t["Gained"].sum())
        tot["Lost"] = int(t["Lost"].sum())
    t = pd.concat([t, pd.DataFrame([tot])], ignore_index=True)

    cols = [c for c in L.SELLER_COLUMNS if c in t.columns]
    return t[cols + ["Is_Total"]]


def comp_by_seller(rows: pd.DataFrame) -> pd.DataFrame:
    """Competitive rows / wins per Seller_ID (used for previous-period rates and history)."""
    comp = rows["Is_Competitive"].astype(bool)
    win = rows["Is_Win"].astype(bool)
    return (
        pd.DataFrame({"Seller_ID": rows["Seller_ID"], "Comp_Variants": comp.astype(int),
                      "Comp_Wins": (comp & win).astype(int)})
        .groupby("Seller_ID")
        .sum()
    )


def rate_by_dimension(rows: pd.DataFrame, column: str) -> pd.DataFrame:
    """Competitive win rate and median loss gap by a dimension (category, brand ...)."""
    if column not in rows.columns:
        return pd.DataFrame()
    r = rows.assign(_dim=rows[column].astype("string").fillna("(نامشخص)"))
    comp = r["Is_Competitive"].astype(bool)
    win = r["Is_Win"].astype(bool)
    base = pd.DataFrame(
        {
            "_dim": r["_dim"],
            "Variants": 1,
            "Evaluated": r["Is_Evaluated"].astype(int),
            "Comp_Variants": comp.astype(int),
            "Comp_Wins": (comp & win).astype(int),
            "Near_Miss": r["Near_Miss"].astype(int),
        }
    )
    t = base.groupby("_dim").sum()
    t["Sellers"] = r.groupby("_dim")["Seller_ID"].nunique()
    t["Comp_Win_Rate"] = t["Comp_Wins"] / t["Comp_Variants"].where(t["Comp_Variants"] > 0)
    loss = r[r["Is_Loss"].astype(bool)]
    t["Median_Loss_Gap"] = loss.groupby("_dim")["Gap_to_BuyBox"].median().reindex(t.index)
    t.index.name = column
    return t.sort_values("Variants", ascending=False).reset_index()


def single_seller_dkps(rows: pd.DataFrame, top: int = 50) -> pd.DataFrame:
    """Top DKPs (by variant count) whose evaluated variants sit in single-seller groups.

    These are candidates for recruiting a second seller.
    """
    ev = rows["Is_Evaluated"].astype(bool)
    single = rows[ev & ~rows["Is_Competitive"].astype(bool)]
    if single.empty:
        return pd.DataFrame()
    agg: dict[str, tuple[str, object]] = {
        "Variants": ("DKPC", "size"),
        "Weights": ("Weight_Key", "nunique"),
        "Sellers": ("Seller_Name", lambda s: L.JOIN_SEP.join(sorted(set(s.astype(str))))),
    }
    if "Product_Name" in single.columns:
        agg["Product_Name"] = ("Product_Name", "first")
    t = single.groupby("DKP").agg(**agg).sort_values("Variants", ascending=False).head(top).reset_index()
    return t
