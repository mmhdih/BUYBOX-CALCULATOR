"""Run history: per-DKPC results saved as parquet + a small JSON of parameters and KPIs."""

from __future__ import annotations

import json
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import pandas as pd

from buybox.core import Params, comp_by_seller, market_summary
from buybox.export import safe_filename
from buybox.pipeline import history_frame


@dataclass(frozen=True)
class RunInfo:
    """A saved run on disk."""

    parquet: Path
    meta: dict

    @property
    def label(self) -> str:
        ts = self.meta.get("created_at", "")[:16].replace("T", " ")
        return f"{ts} — {self.meta.get('input_name', self.parquet.stem)}"


def _json_default(o: object) -> object:
    try:
        return float(o)  # numpy scalars
    except (TypeError, ValueError):
        return str(o)


def save_run(
    rows: pd.DataFrame,
    params: Params,
    input_name: str,
    runs_dir: Path,
    file_hash: str = "",
    now: datetime | None = None,
    original_columns: list[str] | None = None,
) -> RunInfo:
    """Persist a run as ``<YYYYMMDD_HHMM>_<input>.parquet`` plus a JSON with the same stem."""
    runs_dir.mkdir(parents=True, exist_ok=True)
    now = now or datetime.now()
    stem = f"{now.strftime('%Y%m%d_%H%M')}_{safe_filename(Path(input_name).stem, 80)}"
    parquet = runs_dir / f"{stem}.parquet"
    i = 2
    while parquet.exists():
        parquet = runs_dir / f"{stem}_{i}.parquet"
        i += 1
    frame = history_frame(rows)
    for c in frame.columns:
        if frame[c].dtype == object:
            frame[c] = frame[c].astype("string")
    frame.to_parquet(parquet, index=False)
    kpis = market_summary(rows)
    by_seller = comp_by_seller(rows)
    meta = {
        "created_at": now.isoformat(timespec="seconds"),
        "input_name": input_name,
        "file_hash": file_hash,
        "params": params.as_dict(),
        "columns": list(original_columns) if original_columns is not None else None,
        "kpis": kpis,
        "seller_comp": {str(k): [int(v.Comp_Variants), int(v.Comp_Wins)] for k, v in by_seller.iterrows()},
    }
    parquet.with_suffix(".json").write_text(
        json.dumps(meta, ensure_ascii=False, indent=1, default=_json_default), encoding="utf-8")
    return RunInfo(parquet=parquet, meta=meta)


def list_runs(runs_dir: Path) -> list[RunInfo]:
    """Saved runs, newest first (runs whose JSON is missing or broken are skipped)."""
    if not runs_dir.exists():
        return []
    runs: list[RunInfo] = []
    for pq in runs_dir.glob("*.parquet"):
        js = pq.with_suffix(".json")
        if not js.exists():
            continue
        try:
            meta = json.loads(js.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            continue
        runs.append(RunInfo(parquet=pq, meta=meta))
    runs.sort(key=lambda r: (r.meta.get("created_at", ""), r.parquet.name), reverse=True)
    return runs


def find_run(runs_dir: Path, file_hash: str) -> RunInfo | None:
    """A saved run of the same input file (by content hash), if any.

    Buy-box results do not depend on the parameters, so one run per input file is kept.
    """
    if not file_hash:
        return None
    for run in list_runs(runs_dir):
        if run.meta.get("file_hash") == file_hash:
            return run
    return None


def load_run(run: RunInfo) -> pd.DataFrame:
    """Load a saved run's per-DKPC rows (usable as a 'previous file')."""
    df = pd.read_parquet(run.parquet)
    for c in ("Is_Evaluated", "Is_Win", "Is_Loss", "Is_Competitive", "Near_Miss"):
        if c in df.columns:
            df[c] = df[c].fillna(False).astype(bool)
    return df


def delete_run(run: RunInfo) -> None:
    """Delete a saved run (parquet and JSON)."""
    for p in (run.parquet, run.parquet.with_suffix(".json")):
        if p.exists():
            p.unlink()


def trend_frame(runs: list[RunInfo], seller_ids: list[object] | None = None) -> pd.DataFrame:
    """Pooled competitive win rate per saved run (optionally pooled over given sellers only)."""
    recs = []
    wanted = {str(s) for s in seller_ids} if seller_ids else None
    for run in runs:
        created = run.meta.get("created_at")
        if wanted is None:
            k = run.meta.get("kpis", {})
            comp, wins = k.get("comp_rows", 0), k.get("comp_wins", 0)
        else:
            sc = run.meta.get("seller_comp", {})
            comp = sum(sc.get(s, [0, 0])[0] for s in wanted)
            wins = sum(sc.get(s, [0, 0])[1] for s in wanted)
        if not comp:
            continue
        recs.append({"created_at": pd.Timestamp(created), "run": run.label, "comp_rows": comp,
                     "comp_wins": wins, "comp_win_rate": wins / comp})
    if not recs:
        return pd.DataFrame(columns=["created_at", "run", "comp_rows", "comp_wins", "comp_win_rate"])
    return pd.DataFrame(recs).sort_values("created_at").reset_index(drop=True)


