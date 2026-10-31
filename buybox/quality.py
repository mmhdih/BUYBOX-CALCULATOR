"""Data-quality checks shown before any analysis (§5)."""

from __future__ import annotations

from dataclasses import dataclass

import pandas as pd

from buybox import labels as L
from buybox.core import market_summary

STATUS_OK = "ok"
STATUS_WARN = "warn"
STATUS_INFO = "info"


@dataclass
class QualityCheck:
    """One line of the data-quality checklist."""

    key: str
    label: str
    value: object
    status: str = STATUS_INFO
    rows: pd.DataFrame | None = None
    note: str = ""
    pct: bool = False

    @property
    def display(self) -> str:
        """Formatted value (thousands separators / percent with 1 decimal)."""
        v = self.value
        if self.pct and isinstance(v, (int, float)):
            return "-" if v != v else f"{v * 100:.1f}%"
        if isinstance(v, bool):
            return str(v)
        if isinstance(v, int):
            return f"{v:,}"
        return str(v)


def _row_view(rows: pd.DataFrame, mask: pd.Series, extra: tuple[str, ...] = ()) -> pd.DataFrame:
    cols = ["DKPC", "DKP", "Seller_ID", "Seller_Name", "Product_Name", "Weight", "Selling_Price", *extra]
    cols = [c for c in cols if c in rows.columns]
    return rows.loc[mask, cols].reset_index(drop=True)


def run_checks(
    rows: pd.DataFrame,
    duplicate_rows: pd.DataFrame,
    seller_info: pd.DataFrame,
    kam_merged: pd.DataFrame | None,
    current_columns: list[str],
    prev_columns: list[str] | None = None,
) -> list[QualityCheck]:
    """Run all data-quality checks and return them in display order."""
    s = market_summary(rows)
    checks: list[QualityCheck] = [
        QualityCheck("rows", "تعداد کل ردیف‌ها (DKPC)", s["rows"]),
        QualityCheck("dkp", "تعداد DKP یکتا", s["unique_dkp"]),
        QualityCheck("sellers", "تعداد فروشنده یکتا", s["unique_sellers"]),
        QualityCheck("groups", "تعداد گروه‌های DKP + وزن (ارزیابی‌شده)", s["groups"]),
        QualityCheck("single_pct", "درصد گروه‌های تک‌فروشنده", s["single_seller_group_pct"], pct=True),
    ]

    n_dup = len(duplicate_rows)
    checks.append(QualityCheck(
        "dup_dkpc", "ردیف‌های DKPC تکراری (اولین ردیف نگه داشته شد)", n_dup,
        STATUS_WARN if n_dup else STATUS_OK, duplicate_rows.reset_index(drop=True) if n_dup else None,
        "انتظار: صفر",
    ))

    bad_price = rows["Selling_Price"].isna() | (rows["Selling_Price"] <= 0)
    checks.append(QualityCheck(
        "bad_price", "قیمت فروش خالی یا ≤ ۰ (خارج از ارزیابی)", int(bad_price.sum()),
        STATUS_WARN if bad_price.any() else STATUS_OK, _row_view(rows, bad_price) if bad_price.any() else None,
    ))

    missing_id = rows["DKP"].isna() | rows["Seller_ID"].isna()
    if missing_id.any():
        checks.append(QualityCheck(
            "missing_id", "DKP یا Seller_ID خالی (خارج از ارزیابی)", int(missing_id.sum()), STATUS_WARN,
            _row_view(rows, missing_id),
        ))

    no_w = rows["Weight"].isna()
    by_seller = None
    if no_w.any():
        by_seller = (
            rows[no_w].groupby(["Seller_ID", "Seller_Name"]).size().rename("تعداد بدون وزن")
            .sort_values(ascending=False).reset_index()
        )
    checks.append(QualityCheck(
        "no_weight", "وزن خالی (خارج از ارزیابی) — به تفکیک فروشنده", int(no_w.sum()),
        STATUS_WARN if no_w.any() else STATUS_OK, by_seller,
    ))

    tie_rows = rows["Tie_Group"] & rows["Is_Win"]
    checks.append(QualityCheck(
        "ties", "گروه‌های دارای تساوی در کمترین قیمت (همه برنده‌اند)", s["tie_groups"],
        STATUS_INFO, _row_view(rows, tie_rows, ("Group_ID",)) if tie_rows.any() else None,
    ))

    dup_seller = rows["Same_Seller_Dup"]
    checks.append(QualityCheck(
        "same_seller_dup", "گروه‌هایی که یک فروشنده چند DKPC روی همان DKP + وزن دارد",
        s["same_seller_dup_groups"], STATUS_INFO,
        _row_view(rows, dup_seller, ("Group_ID", "Comment")).sort_values(["Group_ID"]) if dup_seller.any() else None,
    ))

    if kam_merged is not None:
        not_in = seller_info[~seller_info["In_KAM_List"].astype(bool)]
        counts = rows.groupby("Seller_ID").size()
        not_in_tbl = not_in.assign(Variants=counts.reindex(not_in.index).fillna(0).astype(int))
        not_in_tbl = not_in_tbl.reset_index()[["Seller_ID", "Seller_Name", "Variants"]].sort_values(
            "Variants", ascending=False)
        checks.append(QualityCheck(
            "not_in_kam", "فروشندگان موجود در اسورتمنت ولی نه در لیست KAM", len(not_in),
            STATUS_WARN if len(not_in) else STATUS_OK, not_in_tbl if len(not_in) else None,
        ))
        dups = kam_merged[kam_merged["Rows"] > 1]
        checks.append(QualityCheck(
            "kam_dup", "Seller Id تکراری در لیست KAM (ادغام شد)", len(dups),
            STATUS_WARN if len(dups) else STATUS_OK, dups.reset_index(drop=True) if len(dups) else None,
        ))
    else:
        checks.append(QualityCheck("kam_missing", "لیست KAM بارگذاری نشده است", "—", STATUS_WARN,
                                   note=f"KAM همه فروشندگان «{L.KAM_NOT_IN_LIST}» درج می‌شود."))

    if prev_columns is not None:
        added = [c for c in current_columns if c not in prev_columns]
        removed = [c for c in prev_columns if c not in current_columns]
        tbl = None
        if added or removed:
            tbl = pd.DataFrame(
                {"ستون": added + removed, "وضعیت": ["اضافه‌شده"] * len(added) + ["حذف‌شده"] * len(removed)}
            )
        checks.append(QualityCheck(
            "schema", "تفاوت ستون‌ها با فایل دوره قبل", f"+{len(added)} / −{len(removed)}",
            STATUS_INFO, tbl,
        ))

    if "Actual_BuyBox" in rows.columns:
        known = rows["Is_Evaluated"] & rows["Actual_BuyBox"].notna()
        if known.any():
            actual = rows.loc[known, "Actual_BuyBox"].astype(bool)
            pred = rows.loc[known, "Is_Win"].astype(bool)
            agree = float((actual == pred).mean())
            conf = pd.crosstab(pred.rename("قانون قیمت: برنده"), actual.rename("واقعی: برنده"))
            checks.append(QualityCheck(
                "actual_bb", "انطباق قانون قیمت با وضعیت واقعی بای باکس", agree,
                STATUS_OK if agree >= 0.99 else STATUS_WARN, conf.reset_index(),
                f"{int(known.sum()):,} ردیف با وضعیت واقعی معلوم", pct=True,
            ))
    return checks


def checks_to_frame(checks: list[QualityCheck]) -> pd.DataFrame:
    """Flat (label, value, status, note) frame for the Data_Quality sheet."""
    return pd.DataFrame(
        {
            "بررسی": [c.label for c in checks],
            "مقدار": [c.display for c in checks],
            "وضعیت": [{"ok": "OK", "warn": "هشدار", "info": "اطلاع"}[c.status] for c in checks],
            "توضیح": [c.note for c in checks],
        }
    )
