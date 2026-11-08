"""Comparison of the current assortment with a previous one (matched by DKPC)."""

from __future__ import annotations

import numpy as np
import pandas as pd

from buybox import labels as L
from buybox.core import comp_by_seller, comp_counts, pooled_rate
from buybox.io import id_key

#: Columns kept for the removed-DKPC table when available.
REMOVED_COLUMNS = ("DKPC", "DKP", "Seller_ID", "Seller_Name", "Product_Name", "Weight",
                   "Selling_Price", "Comment")


def row_changes(cur: pd.DataFrame, prev: pd.DataFrame) -> pd.DataFrame:
    """Per-row comparison columns for the current rows.

    Returns a frame indexed like ``cur`` with ``Comment_Prev_File`` (blank when the DKPC is
    new) and ``BuyBox_Change``.

    Label precedence: new DKPC -> blank current weight -> other non-evaluable row ->
    unchanged comment -> gained -> lost. A row that was not evaluable before and loses now
    did not *lose* a buy box it held, so it is reported as unchanged.
    """
    cur_key = id_key(cur["DKPC"])
    prev_key = id_key(prev["DKPC"])
    prev_comment = pd.Series(prev["Comment"].to_numpy(), index=prev_key.to_numpy())
    prev_comment = prev_comment[~prev_comment.index.duplicated()]

    comment_prev = cur_key.map(prev_comment)
    is_new = comment_prev.isna().to_numpy()
    cur_comment = cur["Comment"].to_numpy()
    prev_c = comment_prev.fillna("").to_numpy()

    change = np.select(
        [
            is_new,
            cur["Weight"].isna().to_numpy(),
            ~cur["Is_Evaluated"].astype(bool).to_numpy(),
            cur_comment == prev_c,
            (cur_comment == L.COMMENT_WIN) & (prev_c != L.COMMENT_WIN),
            (cur_comment == L.COMMENT_LOSS) & (prev_c == L.COMMENT_WIN),
        ],
        [L.CHANGE_NEW, L.CHANGE_NO_WEIGHT, L.CHANGE_BAD_PRICE, L.CHANGE_SAME, L.CHANGE_GAINED,
         L.CHANGE_LOST],
        default=L.CHANGE_SAME,
    )
    return pd.DataFrame(
        {"Comment_Prev_File": comment_prev.fillna("").astype(str).to_numpy(), "BuyBox_Change": change},
        index=cur.index,
    )


def removed_rows(cur: pd.DataFrame, prev: pd.DataFrame) -> pd.DataFrame:
    """Rows of the previous file whose DKPC is not in the current file."""
    cur_keys = set(id_key(cur["DKPC"]).dropna())
    mask = ~id_key(prev["DKPC"]).isin(cur_keys)
    cols = [c for c in REMOVED_COLUMNS if c in prev.columns]
    return prev.loc[mask, cols].reset_index(drop=True)


def seller_changes(cur: pd.DataFrame) -> pd.DataFrame:
    """Gained / lost / net counts per Seller_ID (``cur`` must hold ``BuyBox_Change``)."""
    ch = cur["BuyBox_Change"]
    t = pd.DataFrame(
        {
            "Seller_ID": cur["Seller_ID"],
            "Gained": (ch == L.CHANGE_GAINED).astype(int),
            "Lost": (ch == L.CHANGE_LOST).astype(int),
            "New": (ch == L.CHANGE_NEW).astype(int),
        }
    ).groupby("Seller_ID").sum()
    t["Net"] = t["Gained"] - t["Lost"]
    return t


def prev_comp_by_seller(prev: pd.DataFrame) -> pd.DataFrame:
    """Previous-period competitive counts per seller."""
    return comp_by_seller(prev)


def market_changes(cur: pd.DataFrame, prev: pd.DataFrame, removed: pd.DataFrame) -> dict[str, float]:
    """Market-level comparison KPIs."""
    ch = cur["BuyBox_Change"]
    c_comp, c_wins = comp_counts(cur)
    p_comp, p_wins = comp_counts(prev)
    cur_rate = pooled_rate(c_wins, c_comp)
    prev_rate = pooled_rate(p_wins, p_comp)
    return {
        "gained": int((ch == L.CHANGE_GAINED).sum()),
        "lost": int((ch == L.CHANGE_LOST).sum()),
        "new": int((ch == L.CHANGE_NEW).sum()),
        "removed": int(len(removed)),
        "unchanged": int((ch == L.CHANGE_SAME).sum()),
        "prev_rows": int(len(prev)),
        "prev_comp_rows": p_comp,
        "prev_comp_wins": p_wins,
        "prev_comp_win_rate": prev_rate,
        "cur_comp_win_rate": cur_rate,
        "comp_win_rate_change_pp": (cur_rate - prev_rate) * 100,
    }


def net_movers(changes: pd.DataFrame, seller_info: pd.DataFrame, top: int = 5) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Top net winners (net > 0) and net losers (net < 0) by gained − lost."""
    t = changes.copy()
    t["Seller_Name"] = seller_info["Seller_Name"].reindex(t.index)
    t = t.reset_index()
    winners = t[t["Net"] > 0].sort_values(["Net", "Gained"], ascending=False).head(top)
    losers = t[t["Net"] < 0].sort_values(["Net", "Lost"], ascending=[True, False]).head(top)
    cols = ["Seller_ID", "Seller_Name", "Gained", "Lost", "Net"]
    return winners[cols].reset_index(drop=True), losers[cols].reset_index(drop=True)


def lost_to_whom(rows: pd.DataFrame, seller_id: object, top: int = 5) -> pd.DataFrame:
    """For a seller's competitive losing rows, count by the buy-box winner (top N)."""
    m = (rows["Seller_ID"] == seller_id) & rows["Is_Competitive"].astype(bool) & rows["Is_Loss"].astype(bool)
    vc = rows.loc[m, "BuyBox_Winner_Seller"].value_counts().head(top)
    return vc.rename_axis("BuyBox_Winner_Seller").reset_index(name="Count")


def seller_change_table(rows: pd.DataFrame, prev_rows: pd.DataFrame, changes: pd.DataFrame,
                        seller_info: pd.DataFrame) -> pd.DataFrame:
    """Per seller: previous / current competitive win rate, change (pp), gained, lost, net, new.

    Sorted with the biggest net losers first.
    """
    prev = prev_comp_by_seller(prev_rows)
    cur = comp_by_seller(rows)
    t = changes.join(prev.add_prefix("Prev_"), how="left")
    t["Prev_Rate"] = t["Prev_Comp_Wins"] / t["Prev_Comp_Variants"].where(t["Prev_Comp_Variants"] > 0)
    cur = cur.reindex(t.index)
    t["Cur_Rate"] = cur["Comp_Wins"] / cur["Comp_Variants"].where(cur["Comp_Variants"] > 0)
    t["Change_pp"] = (t["Cur_Rate"] - t["Prev_Rate"]) * 100
    t["Seller_Name"] = seller_info["Seller_Name"].reindex(t.index)
    cols = ["Seller_ID", "Seller_Name", "Prev_Rate", "Cur_Rate", "Change_pp", "Gained", "Lost", "Net", "New"]
    return t.reset_index()[cols].sort_values(["Net", "Lost"], ascending=[True, False]).reset_index(drop=True)


def lost_to_whom_all(rows: pd.DataFrame, seller_info: pd.DataFrame, top: int = 5) -> pd.DataFrame:
    """For every seller, the top buy-box winners against its competitive losing rows."""
    m = rows["Is_Competitive"].astype(bool) & rows["Is_Loss"].astype(bool)
    lost = rows.loc[m, ["Seller_ID", "BuyBox_Winner_Seller"]]
    if lost.empty:
        return pd.DataFrame(columns=["Seller_ID", "Seller_Name", "BuyBox_Winner_Seller", "Count", "Share"])
    counts = lost.groupby(["Seller_ID", "BuyBox_Winner_Seller"]).size().rename("Count").reset_index()
    totals = counts.groupby("Seller_ID")["Count"].transform("sum")
    counts["Share"] = counts["Count"] / totals
    counts["_total"] = totals
    counts = counts.sort_values(["_total", "Seller_ID", "Count"], ascending=[False, True, False])
    counts = counts.groupby("Seller_ID", sort=False).head(top)
    counts["Seller_Name"] = counts["Seller_ID"].map(seller_info["Seller_Name"])
    return counts[["Seller_ID", "Seller_Name", "BuyBox_Winner_Seller", "Count", "Share"]].reset_index(drop=True)
