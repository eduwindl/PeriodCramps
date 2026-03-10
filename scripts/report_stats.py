"""
statistics.py
=============
Computes all statistical summaries and KPIs from the cleaned DataFrames.
Results are dictionaries / DataFrames consumed by generate_report.py.
"""

from __future__ import annotations

import logging
from typing import Any
import pandas as pd
import numpy as np

from data_processing import VISITS_COLS, EQUIPMENT_COLS

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Helper
# ---------------------------------------------------------------------------

def _safe_mean(series: pd.Series) -> float:
    """Return mean of a numeric series, or 0 if empty / all-NaN."""
    valid = series.dropna()
    return float(valid.mean()) if not valid.empty else 0.0


# ---------------------------------------------------------------------------
# Visit statistics
# ---------------------------------------------------------------------------

def compute_visit_summary(visits: pd.DataFrame) -> dict[str, Any]:
    """
    High-level KPIs for the visits section.

    Returns dict with keys:
        total_visits, unique_centers, provinces_covered,
        ups_failures, ups_ok, avg_uptime, avg_bandwidth, avg_dhcp
    """
    ups_col = VISITS_COLS["ups"]
    uptime_col = VISITS_COLS["uptime"]
    bw_col = VISITS_COLS["bandwidth"]
    dhcp_col = VISITS_COLS["dhcp"]

    total_visits = len(visits)
    unique_centers: int = visits[VISITS_COLS["centro"]].nunique() if VISITS_COLS["centro"] in visits.columns else 0

    provinces = 0
    if VISITS_COLS["provincia"] in visits.columns:
        provinces = visits[VISITS_COLS["provincia"]].nunique()

    ups_failures = 0
    ups_ok = 0
    if ups_col in visits.columns:
        ups_failures = int((visits[ups_col].isin(["averiado", "falla", "malo", "dañado", "averiada"])).sum())
        ups_ok = total_visits - ups_failures

    avg_uptime = _safe_mean(visits[uptime_col]) if uptime_col in visits.columns else 0.0
    avg_bandwidth = _safe_mean(visits[bw_col]) if bw_col in visits.columns else 0.0
    avg_dhcp = _safe_mean(visits[dhcp_col]) if dhcp_col in visits.columns else 0.0

    summary = {
        "total_visits": total_visits,
        "unique_centers": unique_centers,
        "provinces_covered": provinces,
        "ups_failures": ups_failures,
        "ups_ok": ups_ok,
        "avg_uptime": round(avg_uptime, 2),
        "avg_bandwidth": round(avg_bandwidth, 2),
        "avg_dhcp": round(avg_dhcp, 2),
    }
    logger.info(f"Visit summary: {summary}")
    return summary


def get_ups_failed_centers(visits: pd.DataFrame) -> pd.DataFrame:
    """Return rows where UPS status indicates a failure."""
    ups_col = VISITS_COLS["ups"]
    if ups_col not in visits.columns:
        return pd.DataFrame()

    mask = visits[ups_col].isin(["averiado", "falla", "malo", "dañado", "averiada"])
    cols = [
        VISITS_COLS["centro"],
        VISITS_COLS["provincia"],
        VISITS_COLS["fecha"],
        ups_col,
        VISITS_COLS["obs"],
    ]
    cols_present = [c for c in cols if c in visits.columns]
    return visits[mask][cols_present].reset_index(drop=True)


def get_high_bandwidth_centers(
    visits: pd.DataFrame, threshold: float = 70.0, top_n: int = 10
) -> pd.DataFrame:
    """Return top N centers with highest bandwidth utilisation, filtered by threshold."""
    bw_col = VISITS_COLS["bandwidth"]
    if bw_col not in visits.columns:
        return pd.DataFrame()

    agg = (
        visits.groupby(VISITS_COLS["centro"])[bw_col]
        .mean()
        .reset_index()
        .rename(columns={bw_col: "Bandwidth_promedio"})
    )
    high = agg[agg["Bandwidth_promedio"] >= threshold].copy()
    high["Bandwidth_promedio"] = high["Bandwidth_promedio"].round(2)
    return high.sort_values("Bandwidth_promedio", ascending=False).head(top_n).reset_index(drop=True)


def get_dhcp_saturated_centers(
    visits: pd.DataFrame, threshold: float = 80.0
) -> pd.DataFrame:
    """Return centers where DHCP saturation >= threshold (per-visit, then max per center)."""
    dhcp_col = VISITS_COLS["dhcp"]
    if dhcp_col not in visits.columns:
        return pd.DataFrame()

    agg = (
        visits.groupby(VISITS_COLS["centro"])[dhcp_col]
        .max()
        .reset_index()
        .rename(columns={dhcp_col: "DHCP_max"})
    )
    saturated = agg[agg["DHCP_max"] >= threshold].copy()
    saturated["DHCP_max"] = saturated["DHCP_max"].round(2)
    return saturated.sort_values("DHCP_max", ascending=False).reset_index(drop=True)


def get_pending_ap_centers(visits: pd.DataFrame) -> pd.DataFrame:
    """Return centers with at least one pending AP to configure."""
    ap_col = VISITS_COLS["ap"]
    if ap_col not in visits.columns:
        return pd.DataFrame()

    agg = (
        visits.groupby(VISITS_COLS["centro"])[ap_col]
        .sum()
        .reset_index()
        .rename(columns={ap_col: "AP_pendientes_total"})
    )
    pending = agg[agg["AP_pendientes_total"] > 0].copy()
    return pending.sort_values("AP_pendientes_total", ascending=False).reset_index(drop=True)


def get_uptime_stats(visits: pd.DataFrame, low_threshold: float = 95.0) -> dict[str, Any]:
    """
    Return uptime statistics.

    Returns dict with keys:
        avg_uptime, max_uptime, min_uptime, low_uptime_centers (DataFrame)
    """
    uptime_col = VISITS_COLS["uptime"]
    if uptime_col not in visits.columns:
        return {"avg_uptime": 0, "max_uptime": 0, "min_uptime": 0, "low_uptime_centers": pd.DataFrame()}

    agg = (
        visits.groupby(VISITS_COLS["centro"])[uptime_col]
        .mean()
        .reset_index()
        .rename(columns={uptime_col: "Uptime_promedio"})
    )
    agg["Uptime_promedio"] = agg["Uptime_promedio"].round(2)

    low = agg[agg["Uptime_promedio"] < low_threshold].sort_values("Uptime_promedio").reset_index(drop=True)

    return {
        "avg_uptime": round(_safe_mean(visits[uptime_col]), 2),
        "max_uptime": round(float(visits[uptime_col].max()), 2) if not visits[uptime_col].dropna().empty else 0,
        "min_uptime": round(float(visits[uptime_col].min()), 2) if not visits[uptime_col].dropna().empty else 0,
        "low_uptime_centers": low,
    }


def get_hallazgos_summary(visits: pd.DataFrame) -> pd.DataFrame:
    """Return count of visits grouped by hallazgo."""
    h_col = VISITS_COLS.get("hallazgos", "Hallazgos")
    if h_col not in visits.columns:
        return pd.DataFrame()

    counts = (
        visits[h_col]
        .value_counts(dropna=True)
        .reset_index()
    )
    counts.columns = ["Hallazgos", "Cantidad"]
    
    # Add a Total row
    total = counts["Cantidad"].sum()
    if total > 0:
        counts.loc[len(counts)] = ["Total", total]

    return counts

# ---------------------------------------------------------------------------
# Equipment statistics
# ---------------------------------------------------------------------------

def compute_equipment_summary(equipment: pd.DataFrame) -> dict[str, Any]:
    """
    High-level KPIs for the equipment section.

    Returns dict with keys:
        total_replacements, unique_centers, equipment_types, technicians
    """
    total = len(equipment)
    centers = equipment[EQUIPMENT_COLS["centro"]].nunique() if EQUIPMENT_COLS["centro"] in equipment.columns else 0
    types = equipment[EQUIPMENT_COLS["equipo"]].nunique() if EQUIPMENT_COLS["equipo"] in equipment.columns else 0
    techs = equipment[EQUIPMENT_COLS["tecnico"]].nunique() if EQUIPMENT_COLS["tecnico"] in equipment.columns else 0

    summary = {
        "total_replacements": total,
        "unique_centers": centers,
        "equipment_types": types,
        "technicians": techs,
    }
    logger.info(f"Equipment summary: {summary}")
    return summary


def get_equipment_by_type(equipment: pd.DataFrame) -> pd.DataFrame:
    """Return count of replacements grouped by equipment type."""
    equipo_col = EQUIPMENT_COLS["equipo"]
    if equipo_col not in equipment.columns:
        return pd.DataFrame()

    counts = (
        equipment.groupby(equipo_col)
        .size()
        .reset_index(name="Cantidad")
        .sort_values("Cantidad", ascending=False)
        .reset_index(drop=True)
    )
    return counts


def get_equipment_detail(equipment: pd.DataFrame) -> pd.DataFrame:
    """Return full detail table of equipment changes."""
    cols = [
        EQUIPMENT_COLS["centro"],
        EQUIPMENT_COLS["fecha"],
        EQUIPMENT_COLS["equipo"],
        EQUIPMENT_COLS["serie_ant"],
        EQUIPMENT_COLS["serie_nueva"],
        EQUIPMENT_COLS["motivo"],
        EQUIPMENT_COLS["tecnico"],
    ]
    cols_present = [c for c in cols if c in equipment.columns]
    return equipment[cols_present].reset_index(drop=True)


# ---------------------------------------------------------------------------
# Master stats bundle
# ---------------------------------------------------------------------------

def build_all_stats(
    visits: pd.DataFrame,
    equipment: pd.DataFrame,
    config: dict,
) -> dict[str, Any]:
    """
    Compute all statistics needed for the report in one call.

    Parameters
    ----------
    visits : pd.DataFrame
        Month-filtered visits data.
    equipment : pd.DataFrame
        Month-filtered equipment data.
    config : dict
        Parsed config.yaml (thresholds section).

    Returns
    -------
    dict with all computed statistics.
    """
    thresholds = config.get("report", {}).get("thresholds", {})
    dhcp_thresh = thresholds.get("dhcp_saturation_pct", 80)
    bw_thresh = thresholds.get("bandwidth_high_pct", 70)
    uptime_thresh = thresholds.get("uptime_low_pct", 95)

    return {
        "visit_summary": compute_visit_summary(visits),
        "hallazgos_summary": get_hallazgos_summary(visits),
        "ups_failed": get_ups_failed_centers(visits),
        "high_bandwidth": get_high_bandwidth_centers(visits, threshold=bw_thresh),
        "dhcp_saturated": get_dhcp_saturated_centers(visits, threshold=dhcp_thresh),
        "pending_aps": get_pending_ap_centers(visits),
        "uptime": get_uptime_stats(visits, low_threshold=uptime_thresh),
        "visits_detail": visits,
        "equipment_summary": compute_equipment_summary(equipment),
        "equipment_by_type": get_equipment_by_type(equipment),
        "equipment_detail": get_equipment_detail(equipment),
    }
