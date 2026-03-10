"""
data_processing.py
==================
Responsible for loading and cleaning the Excel source files.
Outputs clean Pandas DataFrames consumed by statistics.py and generate_report.py.

Supports auto-detection of column names from different Excel formats.
"""

from __future__ import annotations

import logging
from pathlib import Path
from datetime import datetime
import pandas as pd

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Column name constants – internal names used throughout the system
# ---------------------------------------------------------------------------

# visitas_centros.xlsx - internal names
VISITS_COLS = {
    "centro": "Centro",
    "provincia": "Provincia",
    "fecha": "Fecha_visita",
    "ups": "UPS_estado",
    "bandwidth": "Bandwidth_utilizado",
    "dhcp": "DHCP_saturacion",
    "ap": "AP_pendientes",
    "uptime": "Uptime",
    "hallazgos": "Hallazgos",
    "obs": "Observaciones",
}

# cambios_equipos.xlsx - internal names
EQUIPMENT_COLS = {
    "centro": "Centro",
    "fecha": "Fecha",
    "equipo": "Equipo",
    "serie_ant": "Serie_anterior",
    "serie_nueva": "Serie_nueva",
    "motivo": "Motivo",
    "tecnico": "Tecnico",
}

# ---------------------------------------------------------------------------
# Column name auto-mapping: real Excel column -> internal column name
# Each internal name maps to a list of possible real column names (case-insensitive)
# ---------------------------------------------------------------------------

VISITS_COL_ALIASES = {
    "Centro": [
        "Centro", "Centros educativos", "Centro educativo",
        "Nombre de centro", "Escuela", "Nombre",
    ],
    "Provincia": [
        "Provincia", "Distrito", "Region", "Ubicacion",
    ],
    "Fecha_visita": [
        "Fecha_visita", "Fecha visita", "Fecha", "Date",
    ],
    "UPS_estado": [
        "UPS_estado", "UPS estado", "Estatus de UPS", "Estado UPS",
        "UPS", "Estado_UPS",
    ],
    "Bandwidth_utilizado": [
        "Bandwidth_utilizado", "Bandwidth utilizado", "Bandwidth",
        "Ancho de banda", "BW",
    ],
    "DHCP_saturacion": [
        "DHCP_saturacion", "DHCP saturacion", "DHCP", "Saturacion DHCP",
    ],
    "AP_pendientes": [
        "AP_pendientes", "AP pendientes", "Access Points",
        "AP", "APs pendientes",
    ],
    "Uptime": [
        "Uptime", "Disponibilidad", "Uptime %",
    ],
    "Hallazgos": [
        "Hallazgos", "Hallazgo", "Problema",
    ],
    "Observaciones": [
        "Observaciones", "Comentarios", "Notas", 
        "Descripcion", "Detalle",
    ],
}

EQUIPMENT_COL_ALIASES = {
    "Centro": [
        "Centro", "Nombre de centro", "Centro educativo",
        "Centros educativos", "Escuela",
    ],
    "Fecha": [
        "Fecha", "Fecha cambio", "Date",
    ],
    "Equipo": [
        "Equipo", "Dispositivo", "Tipo de equipo", "Tipo equipo",
        "Tipo", "Device",
    ],
    "Serie_anterior": [
        "Serie_anterior", "Serie anterior", "Serie anteri0r",
        "Serial anterior", "SN anterior",
    ],
    "Serie_nueva": [
        "Serie_nueva", "Serie nueva", "Serial nueva",
        "SN nueva", "Serial nuevo", "Serie nuevo",
    ],
    "Motivo": [
        "Motivo", "Razon", "Razon de cierre",
        "Causa", "Descripcion",
    ],
    "Tecnico": [
        "Tecnico", "Tecnico responsable", "Responsable",
        "Ingeniero", "Nombre tecnico",
    ],
}


# ---------------------------------------------------------------------------
# Generic helpers
# ---------------------------------------------------------------------------

def _resolve_path(filepath: str | Path) -> Path:
    """Return an absolute Path, searching from the project root when relative."""
    p = Path(filepath)
    if not p.is_absolute():
        project_root = Path(__file__).parent.parent
        p = project_root / filepath
    return p


def _normalize_col_name(name: str) -> str:
    """Normalize a column name for comparison: lowercase, strip, remove accents/special chars."""
    import unicodedata
    name = str(name).strip().lower()
    # Remove accents
    nfkd = unicodedata.normalize('NFKD', name)
    name = ''.join(c for c in nfkd if not unicodedata.combining(c))
    return name


def _auto_map_columns(df: pd.DataFrame, aliases: dict[str, list[str]]) -> pd.DataFrame:
    """
    Auto-detect and rename columns based on alias mappings.
    For each internal column name, try to find a matching column in the DataFrame.
    """
    df = df.copy()
    current_cols = {_normalize_col_name(c): c for c in df.columns}
    rename_map = {}

    for internal_name, possible_names in aliases.items():
        # Skip if the internal name already exists
        if internal_name in df.columns:
            continue

        # Try each alias
        found = False
        for alias in possible_names:
            normalized_alias = _normalize_col_name(alias)
            if normalized_alias in current_cols:
                real_col = current_cols[normalized_alias]
                if real_col != internal_name:
                    rename_map[real_col] = internal_name
                    logger.info(f"Column mapping: '{real_col}' -> '{internal_name}'")
                found = True
                break

        if not found:
            logger.debug(f"No match found for internal column '{internal_name}'")

    if rename_map:
        df = df.rename(columns=rename_map)
        logger.info(f"Renamed {len(rename_map)} columns: {rename_map}")
    else:
        logger.info("No column renaming needed - all expected columns found.")

    return df


def load_excel(filepath: str | Path, sheet_name: int | str = 0) -> pd.DataFrame:
    """Load an Excel file into a DataFrame."""
    path = _resolve_path(filepath)
    if not path.exists():
        raise FileNotFoundError(f"Excel file not found: {path}")

    logger.info(f"Loading Excel file: {path}")
    df = pd.read_excel(path, sheet_name=sheet_name, engine="openpyxl")
    logger.info(f"  -> {len(df)} rows loaded from '{path.name}'")
    logger.info(f"  -> Columns found: {list(df.columns)}")
    return df


# ---------------------------------------------------------------------------
# Date parsing
# ---------------------------------------------------------------------------

def _parse_dates_column(series: pd.Series) -> pd.Series:
    """
    Parse a date column trying multiple strategies.
    Handles Excel-native datetime objects, American format, and ISO format.
    """
    # If already datetime, return as-is
    if pd.api.types.is_datetime64_any_dtype(series):
        logger.info(f"Date column already datetime64, {series.notna().sum()} valid dates.")
        return series

    # Try American format first: MM/DD/YYYY
    parsed = pd.to_datetime(series, format="%m/%d/%Y", errors="coerce")
    valid_count = parsed.notna().sum()
    total = len(series)
    logger.info(f"Date parse attempt (MM/DD/YYYY): {valid_count}/{total} valid.")

    if valid_count > 0 and valid_count >= total * 0.5:
        return parsed

    # Try pandas auto-detection
    parsed_auto = pd.to_datetime(series, errors="coerce", dayfirst=False)
    valid_auto = parsed_auto.notna().sum()
    logger.info(f"Date parse attempt (auto): {valid_auto}/{total} valid.")

    if valid_auto > valid_count:
        return parsed_auto

    return parsed if valid_count >= valid_auto else parsed_auto


# ---------------------------------------------------------------------------
# Visits data
# ---------------------------------------------------------------------------

MESES_STR = {
    1: ["enero", "january", "jan"],
    2: ["febrero", "february", "feb"],
    3: ["marzo", "march", "mar"],
    4: ["abril", "april", "apr"],
    5: ["mayo", "may"],
    6: ["junio", "june", "jun"],
    7: ["julio", "july", "jul"],
    8: ["agosto", "august", "aug"],
    9: ["septiembre", "september", "sep"],
    10: ["octubre", "october", "oct"],
    11: ["noviembre", "november", "nov"],
    12: ["diciembre", "december", "dec"],
}


def clean_visits(df: pd.DataFrame) -> pd.DataFrame:
    """Clean and normalise the visits DataFrame, auto-mapping columns."""
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    # Auto-map column names from real Excel to internal names
    df = _auto_map_columns(df, VISITS_COL_ALIASES)

    fecha_col = VISITS_COLS["fecha"]
    if fecha_col in df.columns:
        df[fecha_col] = _parse_dates_column(df[fecha_col])
        logger.info(f"Visits date column parsed. dtype={df[fecha_col].dtype}, "
                    f"valid dates: {df[fecha_col].notna().sum()}/{len(df)}")
    else:
        logger.warning(f"Column '{fecha_col}' not found in visits data. "
                       f"Available columns: {list(df.columns)}")

    numeric_cols = [
        VISITS_COLS["bandwidth"], VISITS_COLS["dhcp"],
        VISITS_COLS["uptime"], VISITS_COLS["ap"],
    ]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    ups_col = VISITS_COLS["ups"]
    if ups_col in df.columns:
        df[ups_col] = df[ups_col].astype(str).str.strip().str.lower()

    obs_col = VISITS_COLS["obs"]
    if obs_col in df.columns:
        df[obs_col] = df[obs_col].fillna("—")

    return df


def _filter_by_date(df: pd.DataFrame, year: int, month: int, date_col: str) -> pd.DataFrame:
    """
    Filter DataFrame by year and month using datetime column.
    This is the primary and most reliable filtering method.
    """
    if df.empty:
        logger.warning(f"Empty DataFrame, nothing to filter for {year}-{month:02d}.")
        return df

    logger.info(f"Filtering {len(df)} rows for period: {year}-{month:02d}")

    # 1. Try the specified date column first
    if date_col in df.columns and pd.api.types.is_datetime64_any_dtype(df[date_col]):
        valid_dates = df[date_col].notna()
        mask = valid_dates & (df[date_col].dt.year == year) & (df[date_col].dt.month == month)
        matched = mask.sum()
        logger.info(f"Date filter on '{date_col}': {matched}/{len(df)} rows match {year}-{month:02d}")
        if matched > 0:
            return df[mask].copy()
        else:
            valid_df = df[valid_dates]
            if not valid_df.empty:
                min_date = valid_df[date_col].min()
                max_date = valid_df[date_col].max()
                logger.warning(f"No match for {year}-{month:02d}. "
                               f"Data range: {min_date} to {max_date}")

    # 2. Try ANY datetime column as fallback
    for col in df.columns:
        if col == date_col:
            continue
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            valid_dates = df[col].notna()
            mask = valid_dates & (df[col].dt.year == year) & (df[col].dt.month == month)
            matched = mask.sum()
            if matched > 0:
                logger.info(f"Fallback date filter on '{col}': {matched} rows match")
                return df[mask].copy()

    # 3. No datetime column found or no match — return empty
    logger.warning(f"NO DATA found for {year}-{month:02d}. Returning empty DataFrame.")
    return df.iloc[0:0].copy()


def filter_visits_by_month(df: pd.DataFrame, year: int, month: int) -> pd.DataFrame:
    filtered = _filter_by_date(df, year, month, VISITS_COLS["fecha"])
    logger.info(f"Visits filtered to {year}-{month:02d}: {len(filtered)} rows")
    return filtered


# ---------------------------------------------------------------------------
# Equipment data
# ---------------------------------------------------------------------------

def clean_equipment(df: pd.DataFrame) -> pd.DataFrame:
    """Clean and normalise the equipment-change DataFrame, auto-mapping columns."""
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    # Auto-map column names from real Excel to internal names
    df = _auto_map_columns(df, EQUIPMENT_COL_ALIASES)

    fecha_col = EQUIPMENT_COLS["fecha"]
    if fecha_col in df.columns:
        df[fecha_col] = _parse_dates_column(df[fecha_col])
        logger.info(f"Equipment date column parsed. dtype={df[fecha_col].dtype}, "
                    f"valid dates: {df[fecha_col].notna().sum()}/{len(df)}")
    else:
        logger.warning(f"Column '{fecha_col}' not found in equipment data. "
                       f"Available columns: {list(df.columns)}")

    for col in [EQUIPMENT_COLS["serie_ant"], EQUIPMENT_COLS["serie_nueva"]]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    motivo_col = EQUIPMENT_COLS["motivo"]
    if motivo_col in df.columns:
        df[motivo_col] = df[motivo_col].fillna("No especificado")

    return df


def filter_equipment_by_month(df: pd.DataFrame, year: int, month: int) -> pd.DataFrame:
    filtered = _filter_by_date(df, year, month, EQUIPMENT_COLS["fecha"])
    logger.info(f"Equipment filtered to {year}-{month:02d}: {len(filtered)} rows")
    return filtered


# ---------------------------------------------------------------------------
# Convenience loader
# ---------------------------------------------------------------------------

def load_and_prepare(
    visits_path: str | Path,
    equipment_path: str | Path,
    year: int,
    month: int,
    config: dict | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Load both Excel files, clean them, and return month-filtered DataFrames.
    If `config` is provided and SQL is enabled, DHCP/Uptime columns in the
    month-filtered visits DataFrame will be enriched from the database.

    Returns
    -------
    visits_all, visits_month, equipment_all, equipment_month
    """
    visits_raw = load_excel(visits_path)
    visits_all = clean_visits(visits_raw)
    visits_month = filter_visits_by_month(visits_all, year, month)

    equip_raw = load_excel(equipment_path)
    equip_all = clean_equipment(equip_raw)
    equip_month = filter_equipment_by_month(equip_all, year, month)

    # --------------------------------------------------------
    # SQL enrichment (DHCP & Uptime) – non-breaking
    # --------------------------------------------------------
    if config is not None:
        try:
            import sql_connector as sc   # sibling module in scripts/
            logger.info("SQL enrichment: attempting to load DHCP / Uptime from database …")
            visits_month = sc.enrich_visits_with_sql(visits_month, config, year, month)
        except ImportError:
            logger.warning("sql_connector module not found; skipping SQL enrichment.")
        except Exception as exc:
            logger.warning(f"SQL enrichment error (non-fatal): {exc}")

    return visits_all, visits_month, equip_all, equip_month
