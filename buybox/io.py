"""Reading, validating and normalizing input files (assortment and KAM list)."""

from __future__ import annotations

import io
import re
import zipfile
from dataclasses import dataclass, field

import pandas as pd

from buybox import labels as L

# Canonical column -> list of accepted aliases (compared after normalization).
REQUIRED_COLUMNS: tuple[str, ...] = ("DKPC", "DKP", "Seller_ID", "Weight", "Selling_Price")
OPTIONAL_COLUMNS: tuple[str, ...] = (
    "Seller_Name",
    "Product_Name",
    "Brand_Name_En",
    "Leaf_Category_Name",
    "True_Wage",
    "Actual_BuyBox",
)
COLUMN_ALIASES: dict[str, tuple[str, ...]] = {
    "DKPC": ("dkpc", "variantid"),
    "DKP": ("dkp", "productid"),
    "Seller_ID": ("sellerid",),
    "Weight": ("weight",),
    "Selling_Price": ("sellingprice",),
    "Seller_Name": ("sellername",),
    "Product_Name": ("productname",),
    "Brand_Name_En": ("brandnameen", "brandname", "brand"),
    "Leaf_Category_Name": ("leafcategoryname", "leafcategory", "categoryname"),
    "True_Wage": ("truewage",),
    "Actual_BuyBox": ("actualbuybox", "buyboxstatus", "isbuybox", "actualbuyboxstatus"),
}

KAM_SHEET = "KA Selection"
KAM_COLUMNS: dict[str, tuple[str, ...]] = {
    "Seller_ID": ("sellerid",),
    "Seller_Name": ("sellername",),
    "Category": ("categorylevel1namefa", "category"),
    "KAM": ("kam",),
}


class InputFileError(Exception):
    """Raised with a user-facing Persian message when an input file cannot be used."""


@dataclass
class MappingResult:
    """Result of automatic column detection for an assortment file."""

    mapping: dict[str, str | None]
    missing_required: list[str] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return not self.missing_required


def normalize_name(name: object) -> str:
    """Normalize a column name: case-insensitive, spaces and underscores removed."""
    return re.sub(r"[\s_]+", "", str(name)).lower()


def _read_excel(data: bytes, sheet_name: str | int = 0) -> pd.DataFrame:
    buf = io.BytesIO(data)
    try:
        return pd.read_excel(buf, sheet_name=sheet_name, engine="calamine")
    except ImportError:
        buf.seek(0)
        return pd.read_excel(buf, sheet_name=sheet_name, engine="openpyxl")


def read_excel_bytes(data: bytes, sheet_name: str | int = 0) -> pd.DataFrame:
    """Read one sheet of an .xlsx file given as bytes, raising InputFileError on failure."""
    if not data:
        raise InputFileError(L.ERR_FILE_EMPTY)
    try:
        df = _read_excel(data, sheet_name)
    except PermissionError as exc:  # pragma: no cover - only for path based reads
        raise InputFileError(L.ERR_FILE_LOCKED) from exc
    except (zipfile.BadZipFile, ValueError, KeyError, OSError, Exception) as exc:  # noqa: BLE001
        raise InputFileError(L.ERR_FILE_MALFORMED.format(detail=type(exc).__name__)) from exc
    if df.empty:
        raise InputFileError(L.ERR_FILE_EMPTY)
    # Drop fully empty rows and unnamed empty columns that Excel sometimes leaves behind.
    df = df.dropna(how="all")
    df = df.loc[:, ~(df.columns.astype(str).str.startswith("Unnamed") & df.isna().all())]
    if df.empty:
        raise InputFileError(L.ERR_FILE_EMPTY)
    return df.reset_index(drop=True)


def list_sheets(data: bytes) -> list[str]:
    """Return the sheet names of an .xlsx file."""
    try:
        return list(pd.ExcelFile(io.BytesIO(data), engine="openpyxl").sheet_names)
    except Exception as exc:  # noqa: BLE001
        raise InputFileError(L.ERR_FILE_MALFORMED.format(detail=type(exc).__name__)) from exc


def detect_columns(columns: list[object]) -> MappingResult:
    """Map canonical column names to the file's actual column names."""
    norm = {normalize_name(c): c for c in reversed(list(columns))}  # first occurrence wins
    mapping: dict[str, str | None] = {}
    for canon, aliases in COLUMN_ALIASES.items():
        mapping[canon] = next((norm[a] for a in aliases if a in norm), None)
    missing = [c for c in REQUIRED_COLUMNS if mapping.get(c) is None]
    return MappingResult(mapping=mapping, missing_required=missing)


def normalize_id(series: pd.Series) -> pd.Series:
    """Normalize an ID column: integer IDs become Int64, anything else a stripped string."""
    s = series.copy()
    if s.dtype == object or pd.api.types.is_string_dtype(s):
        s = s.astype("string").str.strip().replace("", pd.NA)
    num = pd.to_numeric(s, errors="coerce")
    non_null = s.notna()
    if (num.notna() == non_null).all():
        valid = num.dropna()
        if (valid % 1 == 0).all():
            return num.astype("Int64")
    return s.astype("string")


def id_key(series: pd.Series) -> pd.Series:
    """String key for matching IDs across files (robust to int/str differences)."""
    return normalize_id(series).astype("string")


def build_work_frame(raw: pd.DataFrame, mapping: dict[str, str | None]) -> pd.DataFrame:
    """Build the typed canonical working frame (same index as ``raw``)."""
    work = pd.DataFrame(index=raw.index)
    for canon in REQUIRED_COLUMNS + OPTIONAL_COLUMNS:
        src = mapping.get(canon)
        if src is None or src not in raw.columns:
            continue
        col = raw[src]
        if canon in ("DKPC", "DKP", "Seller_ID"):
            work[canon] = normalize_id(col)
        elif canon in ("Weight", "Selling_Price", "True_Wage"):
            if col.dtype == object:
                col = col.astype("string").str.replace(",", "", regex=False).str.strip()
            work[canon] = pd.to_numeric(col, errors="coerce").astype("float64")
        else:
            work[canon] = col.astype("string").str.strip().replace("", pd.NA)
    return work


def parse_actual_buybox(series: pd.Series) -> pd.Series:
    """Parse an 'actual buy-box status' column to boolean (NA when unknown)."""
    truthy = {"1", "true", "yes", "y", "بله", "هست", L.COMMENT_WIN, "winner", "buybox"}
    falsy = {"0", "false", "no", "n", "خیر", "نیست", L.COMMENT_LOSS, "loser"}
    s = series.astype("string").str.strip().str.lower()
    out = pd.Series(pd.NA, index=series.index, dtype="boolean")
    out[s.isin(truthy)] = True
    out[s.isin(falsy)] = False
    return out


# ---------------------------------------------------------------------------
# KAM list
# ---------------------------------------------------------------------------
def read_kam_list(data: bytes) -> tuple[pd.DataFrame, list[str]]:
    """Read the KAM list ('KA Selection' sheet) and return (normalized frame, warnings)."""
    warnings: list[str] = []
    sheets = list_sheets(data)
    target = next((s for s in sheets if normalize_name(s) == normalize_name(KAM_SHEET)), None)
    if target is None:
        target = sheets[0]
        warnings.append(L.WARN_KAM_SHEET.format(sheet=target))
    df = read_excel_bytes(data, sheet_name=target)
    norm = {normalize_name(c): c for c in reversed(list(df.columns))}
    out = pd.DataFrame(index=df.index)
    for canon, aliases in KAM_COLUMNS.items():
        src = next((norm[a] for a in aliases if a in norm), None)
        if src is not None:
            out[canon] = df[src]
    missing = [c for c in ("Seller_ID", "KAM") if c not in out.columns]
    if missing:
        raise InputFileError(L.ERR_KAM_COLUMNS.format(missing=", ".join(missing)))
    for opt in ("Seller_Name", "Category"):
        if opt not in out.columns:
            out[opt] = pd.NA
    out["Seller_ID"] = normalize_id(out["Seller_ID"])
    out = out.dropna(subset=["Seller_ID"])
    return out[["Seller_ID", "Seller_Name", "Category", "KAM"]], warnings


def _join_unique(values: pd.Series) -> str:
    seen: list[str] = []
    for v in values.dropna().astype(str).str.strip():
        if v and v not in seen:
            seen.append(v)
    return L.JOIN_SEP.join(seen)


def merge_kam(kam: pd.DataFrame) -> pd.DataFrame:
    """Merge duplicate Seller Id rows into one, joining distinct names/KAMs with ' / '."""
    if kam.empty:
        return pd.DataFrame(columns=["Seller_ID", "Seller_Name", "Category", "KAM", "Rows", "KAM_Note"])
    grouped = kam.groupby("Seller_ID", sort=False)
    merged = grouped.agg(
        Seller_Name=("Seller_Name", _join_unique),
        Category=("Category", _join_unique),
        KAM=("KAM", _join_unique),
        Rows=("KAM", "size"),
    ).reset_index()
    merged["KAM_Note"] = merged["Rows"].gt(1).map({True: L.KAM_DUPLICATE_FLAG, False: ""})
    return merged


def split_kams(kam_value: str) -> list[str]:
    """Split a joined KAM string back into individual names."""
    return [k.strip() for k in str(kam_value).split(L.JOIN_SEP.strip()) if k.strip()]


def build_seller_info(work: pd.DataFrame, kam_merged: pd.DataFrame | None) -> pd.DataFrame:
    """One row per Seller_ID with display name, KAM and KAM-list flags.

    Name priority: assortment ``Seller_Name`` (most frequent) -> KAM list name -> the ID.
    Seller 1 is always labelled as Digikala retail.
    """
    sellers = pd.Index(work["Seller_ID"].dropna().unique(), name="Seller_ID")
    info = pd.DataFrame(index=sellers)
    name = pd.Series(pd.NA, index=sellers, dtype="string")
    if "Seller_Name" in work.columns:
        named = work.dropna(subset=["Seller_ID", "Seller_Name"])
        if not named.empty:
            counts = named.groupby(["Seller_ID", "Seller_Name"], observed=True).size()
            top = counts.sort_values(ascending=False).reset_index().drop_duplicates("Seller_ID")
            name = name.fillna(top.set_index("Seller_ID")["Seller_Name"].astype("string"))
    if kam_merged is not None and not kam_merged.empty:
        km = kam_merged.set_index("Seller_ID")
        km = km[~km.index.duplicated()]
        kam_names = km["Seller_Name"].replace("", pd.NA).astype("string")
        name = name.fillna(kam_names.reindex(sellers))
        info["KAM"] = km["KAM"].reindex(sellers).replace("", pd.NA).fillna(L.KAM_NOT_IN_LIST)
        info["In_KAM_List"] = sellers.isin(km.index)
        info["KAM_Note"] = km["KAM_Note"].reindex(sellers).fillna("")
    else:
        info["KAM"] = L.KAM_NOT_IN_LIST
        info["In_KAM_List"] = False
        info["KAM_Note"] = ""
    name = name.fillna(pd.Series(sellers.astype(str), index=sellers))
    if L.DIGIKALA_SELLER_ID in sellers:
        name.loc[L.DIGIKALA_SELLER_ID] = L.DIGIKALA_RETAIL
    info["Seller_Name"] = name.astype(str)
    return info[["Seller_Name", "KAM", "In_KAM_List", "KAM_Note"]]
