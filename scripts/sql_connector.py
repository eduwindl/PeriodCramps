"""
sql_connector.py
================
Handles SQL Server connectivity for the automated reporting system.
Loads DHCP saturation and Uptime data directly from the database,
with automatic fallback to Excel if the connection fails or is disabled.

Usage (standalone test):
    python scripts/sql_connector.py
"""

from __future__ import annotations

import logging
from pathlib import Path
from datetime import datetime, date
from typing import Optional
import pandas as pd

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _get_db_config(config: dict) -> dict:
    """Extract and validate the database section from config."""
    return config.get("database", {})


def _build_connection_string(db_cfg: dict) -> str:
    """Build the pyodbc connection string from config."""
    driver   = db_cfg.get("driver", "ODBC Driver 18 for SQL Server")
    server   = db_cfg.get("server", "")
    database = db_cfg.get("database", "")
    timeout  = db_cfg.get("timeout", 30)

    trusted = db_cfg.get("trusted_connection", False)
    if trusted:
        auth_part = "Trusted_Connection=yes;"
    else:
        user     = db_cfg.get("username", "")
        password = db_cfg.get("password", "")
        auth_part = f"UID={user};PWD={password};"

    # TrustServerCertificate=yes is required by ODBC Driver 18 for non-SSL servers
    db_part = f"DATABASE={database};" if database else ""
    conn_str = (
        f"DRIVER={{{driver}}};"
        f"SERVER={server};"
        f"{db_part}"
        f"Connect Timeout={timeout};"
        f"TrustServerCertificate=yes;"
        f"{auth_part}"
    )
    return conn_str


def get_connection(config: dict):
    """
    Return a pyodbc Connection object using settings from config.yaml.

    Returns None if pyodbc is not installed or connection fails.
    The caller is responsible for closing the connection.
    """
    try:
        import pyodbc  # type: ignore
    except ImportError:
        logger.warning(
            "pyodbc not installed. Run: pip install pyodbc  "
            "→ Falling back to Excel data."
        )
        return None

    db_cfg = _get_db_config(config)
    if not db_cfg.get("enabled", False):
        logger.info("SQL connection disabled in config (database.enabled = false).")
        return None

    conn_str = _build_connection_string(db_cfg)
    try:
        conn = pyodbc.connect(conn_str, autocommit=True)
        logger.info(
            f"SQL connection OK → {db_cfg.get('server')} / {db_cfg.get('database')}"
        )
        return conn
    except Exception as exc:
        logger.warning(f"SQL connection failed: {exc}  → Falling back to Excel data.")
        return None


# ---------------------------------------------------------------------------
# Introspect table columns
# ---------------------------------------------------------------------------

def _get_actual_columns(conn, table: str) -> list[str]:
    """Return list of column names for the given table."""
    try:
        cursor = conn.cursor()
        cursor.execute(f"SELECT TOP 0 * FROM {table}")
        cols = [desc[0] for desc in cursor.description]
        cursor.close()
        return cols
    except Exception as exc:
        logger.warning(f"Could not read columns from {table}: {exc}")
        return []


def _find_column(candidates: list[str], actual_cols: list[str]) -> Optional[str]:
    """
    Case-insensitive search: return the first actual column name that
    matches any of the candidate names.
    """
    lower_actual = {c.lower(): c for c in actual_cols}
    for c in candidates:
        if c.lower() in lower_actual:
            return lower_actual[c.lower()]
    return None


# ---------------------------------------------------------------------------
# DHCP loader
# ---------------------------------------------------------------------------

# Candidate column names for auto-detection (most → least likely)
DHCP_CENTER_CANDIDATES    = ["CenterName",  "Centro", "NombreCentro", "center_name",  "nombre_centro", "school", "site"]
DHCP_DATE_CANDIDATES      = ["RecordDate",  "Fecha",  "FechaRegistro", "record_date",  "fecha_registro", "date",  "fecha_visita"]
DHCP_PCT_CANDIDATES       = ["Cantidad", "SaturationPct", "DHCP_saturacion", "DHCPPct", "saturation_pct", "dhcp_pct", "dhcp", "saturacion", "pct"]
DHCP_PROV_CANDIDATES      = ["Province",    "Provincia", "province", "region", "distrito"]


def load_dhcp_from_sql(
    config: dict,
    year: int,
    month: int,
    conn=None,
) -> Optional[pd.DataFrame]:
    """
    Query dbo.DHCPRecord (or configured table) for the given year/month.
    Returns a DataFrame with internal column names matching VISITS_COLS,
    or None if the query fails / SQL is disabled.

    The returned DataFrame has columns:
        Centro, DHCP_saturacion, Fecha_visita  (+ optional Provincia)
    """
    db_cfg = _get_db_config(config)
    table  = db_cfg.get("tables", {}).get("dhcp", "dbo.DHCPRecord")

    if not table:
        logger.info("No DHCP table configured → skipping SQL for DHCP.")
        return None

    # Allow caller to pass existing connection (useful for tests)
    _conn = conn
    _owned = False
    if _conn is None:
        _conn = get_connection(config)
        _owned = True

    if _conn is None:
        return None   # fallback to Excel

    try:
        # Discover actual columns
        actual_cols = _get_actual_columns(_conn, table)
        logger.info(f"Columns found in {table}: {actual_cols}")

        # Map to internal names using override from config, then auto-detect
        col_map_cfg = db_cfg.get("column_map", {}).get("dhcp", {})

        def resolve(candidates, cfg_key):
            """Return the actual SQL column to use."""
            # 1. Explicit config override
            if cfg_key in col_map_cfg and col_map_cfg[cfg_key] in actual_cols:
                return col_map_cfg[cfg_key]
            # 2. Auto-detect
            if actual_cols:
                return _find_column(candidates, actual_cols)
            # 3. Try config override even if not in actual_cols (maybe wrong schema)
            return col_map_cfg.get(cfg_key)

        col_centro = resolve(DHCP_CENTER_CANDIDATES, "centro")
        col_fecha  = resolve(DHCP_DATE_CANDIDATES,   "fecha")
        col_dhcp   = resolve(DHCP_PCT_CANDIDATES,    "dhcp_pct")
        col_prov   = resolve(DHCP_PROV_CANDIDATES,   "provincia")

        if not col_dhcp:
            logger.warning(f"Could not locate DHCP saturation column in {table}.")
            return None

        # Build SELECT columns list
        select_parts = []
        if col_centro: select_parts.append(f"[{col_centro}] AS Centro")
        if col_fecha:  select_parts.append(f"[{col_fecha}] AS Fecha_visita")
        if col_dhcp:   select_parts.append(f"[{col_dhcp}] AS DHCP_saturacion")
        if col_prov:   select_parts.append(f"[{col_prov}] AS Provincia")

        select_clause = ", ".join(select_parts) if select_parts else "*"

        # Filter by year/month if date column exists
        if col_fecha:
            query = f"""
                SELECT {select_clause}
                FROM   {table}
                WHERE  YEAR([{col_fecha}])  = {year}
                  AND  MONTH([{col_fecha}]) = {month}
                ORDER BY [{col_fecha}]
            """
        else:
            query = f"SELECT {select_clause} FROM {table}"

        logger.info(f"Executing DHCP query → {table} ({year}-{month:02d})")
        df = pd.read_sql(query, _conn)
        logger.info(f"DHCP SQL → {len(df)} rows returned.")

        # Coerce numeric
        if "DHCP_saturacion" in df.columns:
            df["DHCP_saturacion"] = pd.to_numeric(df["DHCP_saturacion"], errors="coerce")

        return df

    except Exception as exc:
        logger.warning(f"DHCP SQL query failed: {exc}  → Falling back to Excel.")
        return None
    finally:
        if _owned and _conn is not None:
            try:
                _conn.close()
            except Exception:
                pass


# ---------------------------------------------------------------------------
# UPTIME loader
# ---------------------------------------------------------------------------

UPTIME_CENTER_CANDIDATES  = ["CenterName", "Centro", "NombreCentro", "center_name", "nombre_centro", "school", "site"]
UPTIME_DATE_CANDIDATES    = ["RecordDate", "Fecha",  "FechaRegistro", "record_date",  "fecha_registro", "date", "fecha_visita"]
UPTIME_PCT_CANDIDATES     = ["UptimePct",  "Uptime", "uptime_pct",   "disponibilidad", "uptime_%", "pct_uptime"]
UPTIME_PROV_CANDIDATES    = ["Province",   "Provincia", "province",  "region", "distrito"]


def load_uptime_from_sql(
    config: dict,
    year: int,
    month: int,
    conn=None,
) -> Optional[pd.DataFrame]:
    """
    Query the Uptime table for the given year/month.
    Returns a DataFrame with columns matching VISITS_COLS (Uptime column),
    or None if table is not configured / query fails.
    """
    db_cfg = _get_db_config(config)
    table  = db_cfg.get("tables", {}).get("uptime", "")

    if not table:
        logger.info("No Uptime table configured → skipping SQL for Uptime.")
        return None

    # Specifically for PingResults table
    if "PingResults" in table:
        _conn = conn
        _owned = False
        if _conn is None:
            _conn = get_connection(config)
            _owned = True

        if _conn is None:
            return None
        
        try:
            query = f"""
                SELECT Long_Name AS Centro,
                       CAST(SUM(CAST(PingStatus AS INT)) * 100.0 / NULLIF(COUNT(*), 0) AS FLOAT) AS Uptime
                FROM {table}
                GROUP BY Long_Name
            """
            logger.info(f"Executing Uptime query (PingResults aggregated) → {table} ({year}-{month:02d})")
            df = pd.read_sql(query, _conn)
            logger.info(f"Uptime SQL → {len(df)} rows returned.")
            
            if "Uptime" in df.columns:
                df["Uptime"] = pd.to_numeric(df["Uptime"], errors="coerce").round(2)
                
            return df
        except Exception as exc:
            logger.warning(f"Uptime SQL query failed: {exc} → Falling back to Excel.")
            return None
        finally:
            if _owned and _conn is not None:
                try:
                    _conn.close()
                except Exception:
                    pass

    _conn = conn
    _owned = False
    if _conn is None:
        _conn = get_connection(config)
        _owned = True

    if _conn is None:
        return None

    try:
        actual_cols = _get_actual_columns(_conn, table)
        logger.info(f"Columns found in {table}: {actual_cols}")

        col_map_cfg = db_cfg.get("column_map", {}).get("uptime", {})

        def resolve(candidates, cfg_key):
            if cfg_key in col_map_cfg and col_map_cfg[cfg_key] in actual_cols:
                return col_map_cfg[cfg_key]
            if actual_cols:
                return _find_column(candidates, actual_cols)
            return col_map_cfg.get(cfg_key)

        col_centro = resolve(UPTIME_CENTER_CANDIDATES, "centro")
        col_fecha  = resolve(UPTIME_DATE_CANDIDATES,   "fecha")
        col_uptime = resolve(UPTIME_PCT_CANDIDATES,    "uptime_pct")
        col_prov   = resolve(UPTIME_PROV_CANDIDATES,   "provincia")

        if not col_uptime:
            logger.warning(f"Could not locate Uptime column in {table}.")
            return None

        select_parts = []
        if col_centro: select_parts.append(f"[{col_centro}] AS Centro")
        if col_fecha:  select_parts.append(f"[{col_fecha}] AS Fecha_visita")
        if col_uptime: select_parts.append(f"[{col_uptime}] AS Uptime")
        if col_prov:   select_parts.append(f"[{col_prov}] AS Provincia")

        select_clause = ", ".join(select_parts) if select_parts else "*"

        if col_fecha:
            query = f"""
                SELECT {select_clause}
                FROM   {table}
                WHERE  YEAR([{col_fecha}])  = {year}
                  AND  MONTH([{col_fecha}]) = {month}
                ORDER BY [{col_fecha}]
            """
        else:
            query = f"SELECT {select_clause} FROM {table}"

        logger.info(f"Executing Uptime query → {table} ({year}-{month:02d})")
        df = pd.read_sql(query, _conn)
        logger.info(f"Uptime SQL → {len(df)} rows returned.")

        if "Uptime" in df.columns:
            df["Uptime"] = pd.to_numeric(df["Uptime"], errors="coerce")

        return df

    except Exception as exc:
        logger.warning(f"Uptime SQL query failed: {exc}  → Falling back to Excel.")
        return None
    finally:
        if _owned and _conn is not None:
            try:
                _conn.close()
            except Exception:
                pass


# ---------------------------------------------------------------------------
# Merge SQL data into visits DataFrame
# ---------------------------------------------------------------------------

def enrich_visits_with_sql(
    visits_df: pd.DataFrame,
    config: dict,
    year: int,
    month: int,
) -> pd.DataFrame:
    """
    Attempt to load DHCP and/or Uptime data from SQL.

    Strategy:
    - If SQL returns data, REPLACE the corresponding columns in visits_df
      (or append rows if visits_df is empty).
    - If SQL fails/returns None, the original visits_df is left unchanged.

    This means Excel is always the base; SQL just overrides certain columns.
    """
    db_cfg = _get_db_config(config)
    if not db_cfg.get("enabled", False):
        return visits_df

    conn = get_connection(config)
    if conn is None:
        return visits_df

    try:
        def get_base_name(name):
            if not isinstance(name, str): return ""
            return name.split(" -")[0].strip().lower()

        visits_df["base_name"] = visits_df.get("Centro", "").apply(get_base_name)

        # --- DHCP from SQL ---
        dhcp_df = load_dhcp_from_sql(config, year, month, conn=conn)
        if dhcp_df is not None and not dhcp_df.empty and "DHCP_saturacion" in dhcp_df.columns:
            if visits_df.empty:
                visits_df = dhcp_df.copy()
                logger.info("Visits DataFrame populated from SQL DHCP data.")
            else:
                if "Centro" in dhcp_df.columns:
                    dhcp_df["base_name"] = dhcp_df["Centro"].apply(get_base_name)
                    dhcp_lookup = dhcp_df.groupby("base_name")["DHCP_saturacion"].max()
                    
                    visits_df["DHCP_saturacion"] = visits_df["base_name"].map(dhcp_lookup).fillna(visits_df.get("DHCP_saturacion", pd.Series(dtype=float)))
                    logger.info("Merged SQL DHCP saturation into visits DataFrame.")

        # --- Uptime from SQL ---
        uptime_df = load_uptime_from_sql(config, year, month, conn=conn)
        if uptime_df is not None and not uptime_df.empty and "Uptime" in uptime_df.columns:
            if visits_df.empty:
                visits_df = uptime_df.copy()
                logger.info("Visits DataFrame populated from SQL Uptime data.")
            else:
                if "Centro" in uptime_df.columns:
                    uptime_df["base_name"] = uptime_df["Centro"].apply(get_base_name)
                    uptime_lookup = uptime_df.groupby("base_name")["Uptime"].mean()
                    
                    visits_df["Uptime"] = visits_df["base_name"].map(uptime_lookup).fillna(visits_df.get("Uptime", pd.Series(dtype=float)))
                    logger.info("Merged SQL Uptime into visits DataFrame.")

        if "base_name" in visits_df.columns:
            visits_df = visits_df.drop(columns=["base_name"])

    except Exception as exc:
        logger.warning(f"enrich_visits_with_sql failed: {exc}")
    finally:
        try:
            conn.close()
        except Exception:
            pass

    return visits_df


# ---------------------------------------------------------------------------
# Standalone test / diagnostic
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import os, sys, yaml, pprint
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

    project_root = Path(__file__).parent.parent
    cfg_path = project_root / "config" / "config.yaml"

    with open(cfg_path, "r", encoding="utf-8") as f:
        config = yaml.safe_load(f)

    print("\n=== SQL Connector Diagnostic ===\n")
    db_cfg = _get_db_config(config)
    print(f"Enabled    : {db_cfg.get('enabled')}")
    print(f"Server     : {db_cfg.get('server')}")
    print(f"Database   : {db_cfg.get('database')}")
    print(f"Driver     : {db_cfg.get('driver')}")
    print(f"Auth       : {'Windows' if db_cfg.get('trusted_connection') else 'SQL'}")
    print(f"DHCP table : {db_cfg.get('tables', {}).get('dhcp')}")
    print(f"Uptime tbl : {db_cfg.get('tables', {}).get('uptime', '(none)')}")
    print()

    conn = get_connection(config)
    if conn:
        print("SUCCESS: Connection successful!\n")

        # Quick test: fetch first 5 rows of DHCP table
        dhcp_table = db_cfg.get("tables", {}).get("dhcp", "dbo.DHCPRecord")
        if dhcp_table:
            try:
                cursor = conn.cursor()
                cursor.execute(f"SELECT TOP 5 * FROM {dhcp_table}")
                rows = cursor.fetchall()
                cols = [d[0] for d in cursor.description]
                print(f"Sample from {dhcp_table}:")
                print("  Columns:", cols)
                for r in rows:
                    row_dict = dict(zip(cols, r))
                    print("  ", row_dict)
                cursor.close()
            except Exception as exc:
                print(f"  Query error: {exc}")

        # Test month load
        now = datetime.now()
        print(f"\\nLoading DHCP for {now.year}-{now.month:02d} ...")
        df = load_dhcp_from_sql(config, now.year, now.month, conn=conn)
        if df is not None:
            print(f"  -> {len(df)} rows. Columns: {list(df.columns)}")
            print(df.head())
        else:
            print("  -> None returned (check column mapping in config.yaml).")

        conn.close()
    else:
        print("✗ Could not connect. Check config.yaml and ODBC driver installation.")
        print("  Install ODBC driver: https://aka.ms/downloadmsodbcsql")
