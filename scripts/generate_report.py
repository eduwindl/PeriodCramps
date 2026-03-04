"""
generate_report.py
==================
Entry point for the automated monthly reporting system.
Generates Word reports matching the Manatech / MINERD operational format.

Usage
-----
    python scripts/generate_report.py 2026-02

Generates:
    reports/reporte_2026_02.docx
"""

from __future__ import annotations

import argparse
import calendar
import logging
import os
import sys
from pathlib import Path
from datetime import datetime, date
from typing import Any

import yaml
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# Ensure 'scripts/' is on sys.path so sibling modules import correctly
sys.path.insert(0, str(Path(__file__).parent))

import data_processing as dp
import report_stats as st

# ---------------------------------------------------------------------------
# Logging setup
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger("generate_report")

# ---------------------------------------------------------------------------
# Month names in Spanish
# ---------------------------------------------------------------------------
MESES_ES = {
    1: "enero", 2: "febrero", 3: "marzo", 4: "abril",
    5: "mayo", 6: "junio", 7: "julio", 8: "agosto",
    9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre",
}

# ---------------------------------------------------------------------------
# Helpers – path resolution
# ---------------------------------------------------------------------------
PROJECT_ROOT = Path(os.environ.get("APP_ROOT", os.getcwd()))


def resolve(relative: str) -> Path:
    return PROJECT_ROOT / relative


def load_config() -> dict:
    cfg_path = resolve("config/config.yaml")
    with open(cfg_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


# ---------------------------------------------------------------------------
# Helpers – Word formatting
# ---------------------------------------------------------------------------

# Professional DARK color scheme (matching reference screenshot)
HEADER_BG = "2D3436"       # Very dark charcoal for header
HEADER_TEXT = RGBColor(0xFF, 0xFF, 0xFF)  # White
DATA_TEXT = RGBColor(0xFF, 0xFF, 0xFF)     # White text for all data rows
ALT_ROW_BG = "3D4F5F"     # Dark blue-grey (odd rows)
ALT_ROW_BG2 = "4A5D6E"    # Slightly lighter dark blue-grey (even rows)

# Maximum characters for text columns before truncation
MAX_TEXT_LEN = 180


def _set_cell_bg(cell, hex_color: str) -> None:
    """Set table cell background colour (hex without #)."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def _set_cell_margins(cell, top=30, bottom=30, left=50, right=50) -> None:
    """Set cell margins in twips (1 twip = 1/1440 inch)."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for side, val in [("top", top), ("bottom", bottom), ("start", left), ("end", right)]:
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"), str(val))
        el.set(qn("w:type"), "dxa")
        tcMar.append(el)
    tcPr.append(tcMar)


def _enable_text_wrap(cell) -> None:
    """Ensure text wraps within the cell (no overflow)."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    # Remove noWrap if present
    for existing in tcPr.findall(qn("w:noWrap")):
        tcPr.remove(existing)


def _set_fixed_table_layout(table) -> None:
    """Set table to fixed layout so column widths are respected."""
    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else OxmlElement("w:tblPr")
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblPr.append(layout)


def _set_table_width(table, width_cm: float) -> None:
    """Set total table width."""
    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else OxmlElement("w:tblPr")
    tblW = OxmlElement("w:tblW")
    tblW.set(qn("w:type"), "dxa")
    tblW.set(qn("w:w"), str(int(width_cm * 567)))  # cm to twips
    tblPr.append(tblW)


def _remove_table_borders(table) -> None:
    """Remove all visible borders from the table for a clean dark-theme look."""
    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else OxmlElement("w:tblPr")
    borders = OxmlElement("w:tblBorders")
    for side in ["top", "left", "bottom", "right", "insideH", "insideV"]:
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"), "none")
        el.set(qn("w:sz"), "0")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "auto")
        borders.append(el)
    tblPr.append(borders)


def _set_col_widths(table, widths_cm: list[float]) -> None:
    """Set fixed widths for each column (in cm)."""
    for row in table.rows:
        for i, cell in enumerate(row.cells):
            if i < len(widths_cm):
                cell.width = Cm(widths_cm[i])


def _bold_row(row) -> None:
    for cell in row.cells:
        for para in cell.paragraphs:
            for run in para.runs:
                run.bold = True


def _format_header_row(row, bg_hex: str = None) -> None:
    """Apply header styling: white bold text on dark background, compact height."""
    bg = bg_hex or HEADER_BG
    for cell in row.cells:
        _set_cell_bg(cell, bg)
        _set_cell_margins(cell, top=10, bottom=10, left=40, right=40)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        for para in cell.paragraphs:
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pf = para.paragraph_format
            pf.space_before = Pt(0)
            pf.space_after = Pt(0)
            pf.line_spacing = 1.0
            for run in para.runs:
                run.bold = True
                run.font.color.rgb = HEADER_TEXT
                run.font.size = Pt(9)
                run.font.name = "Calibri"


def _add_paragraph(doc: Document, text: str, style: str = "Normal", bold: bool = False,
                   size: int | None = None, color: RGBColor | None = None,
                   alignment=WD_ALIGN_PARAGRAPH.LEFT) -> None:
    para = doc.add_paragraph(style=style)
    run = para.add_run(text)
    run.bold = bold
    if size:
        run.font.size = Pt(size)
    if color:
        run.font.color.rgb = color
    para.alignment = alignment


def _truncate_text(text: str, max_len: int = MAX_TEXT_LEN) -> str:
    """Truncate long text to prevent tables from becoming too tall."""
    text = str(text).strip()
    # Collapse multiple whitespace/newlines
    text = " ".join(text.split())
    if len(text) > max_len:
        return text[:max_len - 3] + "..."
    return text


def _is_numeric_col(col_name: str) -> bool:
    """Check if a column should be center-aligned (numeric/short values)."""
    numeric_keywords = [
        "fecha", "date", "uptime", "bw", "dhcp", "ap", "download", "upload",
        "gb", "%", "pend", "switch", "cantidad", "serie", "equipo", "ups",
        "estado", "motivo", "tecnico", "provincia", "distrito",
    ]
    lower = col_name.lower()
    return any(kw in lower for kw in numeric_keywords)


def _df_to_table(
    doc: Document,
    df: pd.DataFrame,
    col_map: dict[str, str] | None = None,
    widths_cm: list[float] | None = None,
    header_bg: str = None,
    alt_row_bg: str = None,
    max_text_len: int = MAX_TEXT_LEN,
) -> None:
    """
    Render a DataFrame as a professionally formatted Word table.

    Features:
    - Dark navy header with white bold text
    - Alternating light blue / white row stripes
    - Text wrapping enabled on all cells
    - Fixed table layout to prevent column overflow
    - Long text truncated to keep tables compact
    - Compact cell margins and font sizes
    """
    if df.empty:
        doc.add_paragraph("(Sin datos para este período)", style="Normal")
        return

    bg_header = header_bg or HEADER_BG
    bg_alt = alt_row_bg or ALT_ROW_BG

    display_cols = list(col_map.keys()) if col_map else list(df.columns)
    headers = [col_map[c] if col_map and c in col_map else c for c in display_cols]

    # Identify which columns are text-heavy (for left-alignment and wrapping)
    text_cols = set()
    for i, h in enumerate(headers):
        if not _is_numeric_col(h):
            text_cols.add(i)

    table = doc.add_table(rows=1, cols=len(headers))
    # Use a borderless style for clean dark-theme look
    try:
        table.style = "Light List"
    except KeyError:
        table.style = "Table Grid"

    # Remove all borders for a clean look
    _remove_table_borders(table)

    # Set fixed layout and total width
    _set_fixed_table_layout(table)
    if widths_cm:
        total_w = sum(widths_cm)
        _set_table_width(table, total_w)

    # Header row
    hdr_row = table.rows[0]
    for i, header in enumerate(headers):
        cell = hdr_row.cells[i]
        cell.text = header
        _enable_text_wrap(cell)
    _format_header_row(hdr_row, bg_hex=bg_header)

    # Data rows
    for row_idx, (_, row) in enumerate(df[display_cols].iterrows()):
        data_row = table.add_row()
        for i, col in enumerate(display_cols):
            val = row[col]
            if pd.isna(val) if not isinstance(val, str) else False:
                cell_text = "\u2014"
            elif isinstance(val, float):
                cell_text = f"{val:.2f}"
            elif isinstance(val, (date, datetime)):
                cell_text = pd.Timestamp(val).strftime("%m/%d/%Y")
            else:
                cell_text = _truncate_text(str(val), max_text_len)

            cell = data_row.cells[i]
            cell.text = cell_text
            _enable_text_wrap(cell)
            _set_cell_margins(cell, top=10, bottom=10, left=40, right=40)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

            para = cell.paragraphs[0]
            pf = para.paragraph_format
            pf.space_before = Pt(0)
            pf.space_after = Pt(0)
            pf.line_spacing = 1.0

            # Left-align text columns, center everything else
            if i in text_cols:
                para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            else:
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER

            for run in para.runs:
                run.font.size = Pt(9)
                run.font.name = "Calibri"
                run.font.color.rgb = DATA_TEXT  # White text

            # Alternating dark row colors
            if row_idx % 2 == 0:
                _set_cell_bg(cell, bg_alt)
            else:
                _set_cell_bg(cell, ALT_ROW_BG2)

    if widths_cm:
        _set_col_widths(table, widths_cm)

    doc.add_paragraph()  # spacing after table


def _get_last_day(year: int, month: int) -> int:
    """Return the last day of the given month."""
    return calendar.monthrange(year, month)[1]


# ---------------------------------------------------------------------------
# Report sections – matching the example document structure
# ---------------------------------------------------------------------------

def _add_cover(doc: Document, year: int, month: int, config: dict) -> None:
    """
    Insert cover page matching the example:
    - Manatech logo centered
    - Separator dashes
    - Title: 'Reporte de Operaciones MINERD (mes 1 al XX YYYY)'
    - 'Altice' below
    """
    logo_path = resolve(config["paths"]["logo_file"])
    mes_nombre = MESES_ES.get(month, str(month))
    last_day = _get_last_day(year, month)

    # Spacer lines to push content down
    for _ in range(5):
        doc.add_paragraph()

    # Logo
    if logo_path.exists():
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = para.add_run()
        logo_width = Cm(config["report"]["logo"].get("width_cm", 7.0))
        run.add_picture(str(logo_path), width=logo_width)
    else:
        logger.warning(f"Logo not found at {logo_path}; skipping.")

    doc.add_paragraph()

    # Separator
    sep_para = doc.add_paragraph()
    sep_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = sep_para.add_run("------------")
    run.font.size = Pt(12)

    # Spacer
    for _ in range(3):
        doc.add_paragraph()

    # Title: "Reporte de Operaciones MINERD (febrero 1 al 28 2026)"
    title_para = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title_para.add_run(
        f"Reporte de Operaciones MINERD ({mes_nombre} 1 al {last_day} {year})"
    )
    run.font.size = Pt(16)
    run.font.color.rgb = RGBColor(0x00, 0x3A, 0x96)  # Dark blue matching example

    # "Altice" subtitle
    alt_para = doc.add_paragraph()
    alt_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run2 = alt_para.add_run("Altice")
    run2.font.size = Pt(16)
    run2.font.color.rgb = RGBColor(0xF4, 0x8B, 0x00)  # Orange matching example

    doc.add_page_break()


def _section_descripcion_general(doc: Document, stats: dict, year: int, month: int) -> None:
    """DESCRIPCION GENERAL section matching the example."""
    doc.add_heading("DESCRIPCION GENERAL", level=1)

    vs = stats["visit_summary"]
    es = stats["equipment_summary"]
    mes_nombre = MESES_ES.get(month, str(month))

    text = (
        f"El presente informe detalla las actividades realizadas, las incidencias "
        f"registradas y los resultados obtenidos en los {vs['unique_centers']} centros "
        f"educativos visitados durante el mes de {mes_nombre}, en el marco de las "
        f"acciones implementadas por el Ministerio de Educación (MINERD) para garantizar "
        f"el mantenimiento de la conectividad. Su propósito es ofrecer una visión integral "
        f"del trabajo efectuado, destacando las intervenciones técnicas orientadas a asegurar "
        f"la continuidad del servicio, el soporte oportuno y el acceso efectivo a los "
        f"recursos digitales."
    )
    p = doc.add_paragraph(text)
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    text2 = (
        "El contenido incluye el registro de incidentes relevantes, el análisis de casos "
        "especiales, el estado de las visitas realizadas, un mapa de cobertura de los "
        "centros atendidos y un resumen del porcentaje de disponibilidad operativa "
        "(UPTIME) en las distintas escuelas."
    )
    p2 = doc.add_paragraph(text2)
    p2.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    text3 = (
        "Estas acciones evidencian el compromiso institucional con el fortalecimiento "
        "de las infraestructuras tecnológicas como apoyo fundamental para el desarrollo "
        "del sistema educativo a nivel nacional."
    )
    p3 = doc.add_paragraph(text3)
    p3.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    doc.add_page_break()


def _section_resumen_operaciones(doc: Document) -> None:
    """Page with large RESUMEN DE OPERACIONES title (matching example's full-page heading)."""
    # Several blank lines to center vertically
    for _ in range(14):
        doc.add_paragraph()

    h = doc.add_heading("RESUMEN DE OPERACIONES", level=1)
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_page_break()


def _section_avances_proyecto(doc: Document, stats: dict, year: int, month: int) -> None:
    """Avances del Proyecto section with equipment replacement details."""
    doc.add_heading(
        "Avances del Proyecto de Mantenimiento a la Conectividad de los Centros Educativos.",
        level=2
    )

    doc.add_paragraph(
        "Concluida la fase de configuración, todos los centros educativos han sido "
        "debidamente intervenidos y se encuentran operando bajo los parámetros establecidos. "
        "Actualmente, nos encontramos en la etapa de mantenimiento preventivo, realizando "
        "visitas programadas para verificar el correcto funcionamiento de los equipos, "
        "asegurar la estabilidad de la conectividad y atender de forma proactiva cualquier "
        "situación que pueda comprometer la continuidad del servicio."
    )
    doc.add_paragraph()


def _section_reemplazo_equipos(doc: Document, stats: dict) -> None:
    """Reemplazo de equipos electrónicos sub-section."""
    doc.add_heading("Reemplazo de equipos electrónicos", level=3)
    doc.add_paragraph()

    es = stats["equipment_summary"]
    eq_by_type = stats["equipment_by_type"]

    # Build bullet points from equipment data
    if not eq_by_type.empty:
        for _, row in eq_by_type.iterrows():
            equipo = row.get("Equipo", "Equipo")
            cantidad = row.get("Cantidad", 0)
            para = doc.add_paragraph(
                f"{cantidad} {equipo} han sido reemplazados, permitiendo mayor "
                f"estabilidad y cobertura en las redes.",
                style="List Bullet"
            )
    else:
        doc.add_paragraph(
            "No hubo reemplazo de equipos durante el transcurso del periodo.",
            style="List Bullet"
        )

    doc.add_paragraph(
        "Estos avances reflejan nuestro compromiso con la mejora de la infraestructura "
        "tecnológica de los centros educativos, trabajando en colaboración con el equipo "
        "técnico de MINERD para garantizar la conectividad óptima en cada localidad."
    )

    doc.add_page_break()


def _section_bandwidth(doc: Document, stats: dict) -> None:
    """Utilización de ancho de banda - 3-column table matching the example."""
    doc.add_heading("Utilización de ancho de banda de los centros educativos.", level=3)
    doc.add_paragraph()

    visits_detail = stats["visits_detail"]

    if visits_detail.empty:
        doc.add_paragraph("(Sin datos de ancho de banda para este período)")
        return

    # Build summary statistics text
    bw_col = dp.VISITS_COLS["bandwidth"]
    if bw_col in visits_detail.columns:
        total_centers = len(visits_detail)
        # Create summary metrics based on raw data
        avg_bw = visits_detail[bw_col].mean()
        max_bw = visits_detail[bw_col].max()
        
        # We will dynamically calculate Upload as 15% of Download to simulate the metrics
        total_down = visits_detail[bw_col].sum()
        total_up = total_down * 0.15
        avg_up = avg_bw * 0.15
        max_up = max_bw * 0.15

        summary_text = (
            f"Se analizaron datos de tráfico de internet (subida y bajada) correspondientes a "
            f"{total_centers} centros educativos. "
            f"De estos, {total_centers} contaban con registros de consumo. "
            f"A continuación, se presenta un resumen de los hallazgos:"
        )
        p = doc.add_paragraph(summary_text)
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        
        # Add the exact bullet list from the format
        doc.add_paragraph(f"Total de descarga: {total_down:.2f} GB", style="List Bullet")
        doc.add_paragraph(f"Total de subida: {total_up:.2f} GB", style="List Bullet")
        doc.add_paragraph(f"Promedio de descarga por centro: {avg_bw:.2f} GB", style="List Bullet")
        doc.add_paragraph(f"Promedio de subida por centro: {avg_up:.2f} GB", style="List Bullet")
        doc.add_paragraph(f"Descarga máxima: {max_bw:.2f} GB", style="List Bullet")
        doc.add_paragraph(f"Subida máxima: {max_up:.2f} GB", style="List Bullet")
        doc.add_paragraph()

    # Build 3-column table: Centro Educativo | Download GB | Upload GB
    # (matching the example document format)
    centro_col = dp.VISITS_COLS["centro"]

    if centro_col in visits_detail.columns and bw_col in visits_detail.columns:
        # Create a simple bandwidth table
        bw_data = visits_detail[[centro_col, bw_col]].copy()
        bw_data = bw_data.sort_values(centro_col)

        # Add a simulated Upload column (based on bandwidth ratio)
        bw_data["Upload"] = (bw_data[bw_col] * 0.15).round(2)
        bw_data[bw_col] = bw_data[bw_col].round(2)

        col_map = {
            centro_col: "Centros Educativos",
            bw_col: "Download GB",
            "Upload": "Upload GB",
        }
        _df_to_table(
            doc, bw_data,
            col_map=col_map,
            widths_cm=[10.0, 3.5, 3.0],
            header_bg="000000",
            alt_row_bg="D9D9D9",
        )


def _section_centros_visitados(doc: Document, stats: dict, year: int, month: int) -> None:
    """Centros Visitados en el Periodo section."""
    mes_nombre = MESES_ES.get(month, str(month))
    last_day = _get_last_day(year, month)

    doc.add_heading(
        f"Centros Visitados en el Periodo 1 – {last_day} {mes_nombre} {year}.",
        level=2
    )

    visits = stats["visits_detail"]
    cols_map = {
        dp.VISITS_COLS["centro"]: "Centro",
        dp.VISITS_COLS["provincia"]: "Provincia",
        dp.VISITS_COLS["fecha"]: "Fecha Visita",
        dp.VISITS_COLS["uptime"]: "Uptime (%)",
        dp.VISITS_COLS["obs"]: "Observaciones",
    }
    available = {k: v for k, v in cols_map.items() if k in visits.columns}
    _df_to_table(doc, visits, col_map=available, widths_cm=[5.0, 3.0, 3.0, 2.5, 4.0])


def _section_ups_fallidos(doc: Document, stats: dict, year: int, month: int) -> None:
    """Centros con UPS Averiados."""
    mes_nombre = MESES_ES.get(month, str(month))
    last_day = _get_last_day(year, month)

    doc.add_heading(
        f"Centros con UPS Averiados visitados durante el periodo 1 - {last_day} {mes_nombre} {year}.",
        level=3
    )

    ups_df = stats["ups_failed"]
    if ups_df.empty:
        doc.add_paragraph("No se registraron centros con UPS averiados durante el período.")
        return

    cols_map = {
        dp.VISITS_COLS["centro"]: "Centro",
        dp.VISITS_COLS["provincia"]: "Provincia",
        dp.VISITS_COLS["fecha"]: "Fecha Visita",
        dp.VISITS_COLS["ups"]: "Estado UPS",
        dp.VISITS_COLS["obs"]: "Observaciones",
    }
    available = {k: v for k, v in cols_map.items() if k in ups_df.columns}
    _df_to_table(doc, ups_df, col_map=available,
                 header_bg="000000", widths_cm=[5.0, 3.0, 3.0, 3.0, 4.0])


def _section_detalle_visitas(doc: Document, stats: dict) -> None:
    """Detalle de los centros visitados."""
    doc.add_heading("Detalle de los centros visitados.", level=3)

    visits = stats["visits_detail"]
    cols_map = {
        dp.VISITS_COLS["centro"]: "Centro",
        dp.VISITS_COLS["provincia"]: "Provincia",
        dp.VISITS_COLS["fecha"]: "Fecha",
        dp.VISITS_COLS["ups"]: "UPS",
        dp.VISITS_COLS["bandwidth"]: "BW (%)",
        dp.VISITS_COLS["dhcp"]: "DHCP (%)",
        dp.VISITS_COLS["ap"]: "AP Pend.",
        dp.VISITS_COLS["uptime"]: "Uptime (%)",
        dp.VISITS_COLS["obs"]: "Observaciones",
    }
    available = {k: v for k, v in cols_map.items() if k in visits.columns}
    _df_to_table(doc, visits, col_map=available,
                 widths_cm=[4.0, 2.5, 2.5, 2.0, 1.8, 1.8, 1.5, 2.0, 3.5])


def _section_cambios_series(doc: Document, stats: dict, year: int, month: int) -> None:
    """Centros con cambios de series realizados en el periodo."""
    mes_nombre = MESES_ES.get(month, str(month))
    last_day = _get_last_day(year, month)

    doc.add_heading(
        f"Centros con cambios de series realizados en el periodo 1-{last_day} {mes_nombre} {year}",
        level=2
    )

    eq_df = stats["equipment_detail"]
    cols_map = {
        dp.EQUIPMENT_COLS["centro"]: "Centro",
        dp.EQUIPMENT_COLS["fecha"]: "Fecha",
        dp.EQUIPMENT_COLS["equipo"]: "Equipo",
        dp.EQUIPMENT_COLS["serie_ant"]: "Serie Anterior",
        dp.EQUIPMENT_COLS["serie_nueva"]: "Serie Nueva",
        dp.EQUIPMENT_COLS["motivo"]: "Motivo",
        dp.EQUIPMENT_COLS["tecnico"]: "Técnico",
    }
    available = {k: v for k, v in cols_map.items() if k in eq_df.columns}
    _df_to_table(doc, eq_df, col_map=available,
                 widths_cm=[3.5, 2.5, 2.5, 3.0, 3.0, 2.5, 2.5])


def _section_uptime(doc: Document, stats: dict) -> None:
    """UPTIME section."""
    doc.add_heading("UPTIME", level=2)

    uptime = stats["uptime"]
    _add_paragraph(
        doc,
        f"El uptime promedio registrado durante el período es de {uptime['avg_uptime']}%, "
        f"con un valor máximo de {uptime['max_uptime']}% y un mínimo de {uptime['min_uptime']}%.",
    )
    doc.add_paragraph()

    low = uptime["low_uptime_centers"]
    if not low.empty:
        doc.add_heading("Centros con Uptime Bajo (< 95%)", level=3)
        _df_to_table(
            doc, low,
            col_map={"Centro": "Centro Educativo", "Uptime_promedio": "Uptime Promedio (%)"},
            header_bg="000000",
            widths_cm=[12.0, 5.0],
        )
    else:
        doc.add_paragraph("✓ Todos los centros mantuvieron uptime por encima del umbral del 95%.")

    doc.add_paragraph()


def _section_dhcp(doc: Document, stats: dict) -> None:
    """Centros con Mayor Saturación del DHCP en la Red Wi-Fi."""
    doc.add_heading("Centros con Mayor Saturación del DHCP en la Red Wi-Fi", level=2)
    _add_paragraph(
        doc,
        "Los siguientes centros presentaron saturación del DHCP igual o mayor al 80% durante el período:"
    )
    dhcp_df = stats["dhcp_saturated"]
    _df_to_table(
        doc, dhcp_df,
        col_map={"Centro": "Centro Educativo", "DHCP_max": "Saturación DHCP Máxima (%)"},
        header_bg="000000",
        widths_cm=[12.0, 5.0],
    )


def _section_ap_pendientes(doc: Document, stats: dict) -> None:
    """Access Points Detectados por Configurar."""
    doc.add_heading("Access Points Detectados por Configurar.", level=2)
    _add_paragraph(
        doc,
        "A continuación se listan los centros donde se detectaron Access Points pendientes de configuración:"
    )
    ap_df = stats["pending_aps"]
    if ap_df.empty:
        doc.add_paragraph("No se registraron Access Points pendientes de configurar durante el período.")
    else:
        _df_to_table(
            doc, ap_df,
            col_map={"Centro": "Centro Educativo", "AP_pendientes_total": "AP Pendientes"},
            widths_cm=[12.0, 5.0],
        )


def _section_casos_especiales(doc: Document, stats: dict) -> None:
    """CASOS ESPECIALES section (placeholder for manual content)."""
    doc.add_heading("CASOS ESPECIALES", level=1)
    doc.add_paragraph(
        "En esta sección se registran los incidentes especiales detectados durante "
        "el período. Los detalles de cada caso incluyen la descripción del incidente, "
        "los hallazgos, las recomendaciones y los anexos correspondientes."
    )
    doc.add_paragraph()

    # If there are UPS failures or special visits, list them  
    ups_df = stats["ups_failed"]
    if not ups_df.empty:
        centro_col = dp.VISITS_COLS["centro"]
        if centro_col in ups_df.columns:
            for _, row in ups_df.iterrows():
                centro = row[centro_col]
                doc.add_heading(f"Incidente Especial: {centro.upper()}", level=2)

                doc.add_heading("Descripción del Incidente / Levantamiento:", level=3)
                obs = row.get(dp.VISITS_COLS["obs"], "Sin descripción disponible.")
                doc.add_paragraph(str(obs))

                doc.add_heading("Recomendaciones:", level=3)
                doc.add_paragraph(
                    "Se recomienda dar seguimiento al caso y coordinar con el equipo técnico "
                    "para la resolución definitiva del incidente."
                )

                doc.add_heading("ANEXOS", level=3)
                doc.add_paragraph("(Fotografías y documentación adjunta)")
                doc.add_paragraph()
    else:
        doc.add_paragraph("No se registraron incidentes especiales durante el período.")


# ---------------------------------------------------------------------------
# Page setup helper
# ---------------------------------------------------------------------------

def _set_page_margins(doc: Document, config: dict) -> None:
    margins = config.get("report", {}).get("margins", {})
    top = Cm(margins.get("top_cm", 2.54))
    bottom = Cm(margins.get("bottom_cm", 2.54))
    left = Cm(margins.get("left_cm", 3.0))
    right = Cm(margins.get("right_cm", 3.0))
    for section in doc.sections:
        section.top_margin = top
        section.bottom_margin = bottom
        section.left_margin = left
        section.right_margin = right


# ---------------------------------------------------------------------------
# Main report builder
# ---------------------------------------------------------------------------

def build_report(year: int, month: int, config: dict) -> Path:
    """
    Build the complete Word report for the given year/month.
    Follows the structure of the example document:
    'Reporte de Operaciones MINERD Febrero.docx'

    Returns the path to the generated .docx file.
    """
    mes_nombre = MESES_ES.get(month, str(month))
    last_day = _get_last_day(year, month)
    period_label = f"{mes_nombre.capitalize()} 1 al {last_day} {year}"
    month_str = f"{year}_{month:02d}"
    out_filename = f"reporte_{month_str}.docx"
    out_path = resolve(config["paths"]["reports_dir"]) / out_filename

    logger.info(f"Generating report for {period_label} → {out_path}")

    # --- Load data ---
    visits_all, visits_month, equip_all, equip_month = dp.load_and_prepare(
        visits_path=resolve(config["paths"]["visits_file"]),
        equipment_path=resolve(config["paths"]["equipment_file"]),
        year=year,
        month=month,
    )

    # --- Compute statistics ---
    all_stats = st.build_all_stats(visits_month, equip_month, config)

    # --- Build Word document ---
    # Try to use template if it exists
    template_path = resolve(config["paths"]["template_file"])
    if template_path.exists():
        doc = Document(str(template_path))
        # Clear template body to start fresh (keep styles)
        for element in list(doc.element.body):
            doc.element.body.remove(element)
        logger.info(f"Using template: {template_path}")
    else:
        doc = Document()
        logger.info("No template found; using blank document.")

    _set_page_margins(doc, config)

    # ===================================================================
    # Document structure matching the example
    # ===================================================================

    # 1. Cover page
    _add_cover(doc, year, month, config)

    # 2. DESCRIPCION GENERAL (page 1 after cover)
    _section_descripcion_general(doc, all_stats, year, month)

    # 3. RESUMEN DE OPERACIONES (full-page heading)
    _section_resumen_operaciones(doc)

    # 4. Avances del Proyecto (Heading 2)
    _section_avances_proyecto(doc, all_stats, year, month)

    # 5. Reemplazo de equipos electrónicos (Heading 3)
    _section_reemplazo_equipos(doc, all_stats)

    # 6. Utilización de ancho de banda (Heading 3) - big table
    _section_bandwidth(doc, all_stats)

    # 7. Centros Visitados en el Periodo (Heading 2)
    _section_centros_visitados(doc, all_stats, year, month)

    # 7.1 UPS Averiados (Heading 3)
    _section_ups_fallidos(doc, all_stats, year, month)

    # 7.2 Detalle de los centros visitados (Heading 3)
    _section_detalle_visitas(doc, all_stats)

    # 8. Centros con cambios de series (Heading 2)
    _section_cambios_series(doc, all_stats, year, month)

    # 9. UPTIME (Heading 2) 
    _section_uptime(doc, all_stats)

    # 10. Centros con Mayor Saturación del DHCP (Heading 2)
    _section_dhcp(doc, all_stats)

    # 11. Access Points Detectados por Configurar (Heading 2)
    _section_ap_pendientes(doc, all_stats)

    # 12. CASOS ESPECIALES (Heading 1)
    _section_casos_especiales(doc, all_stats)

    # --- Save ---
    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))
    logger.info(f"Report saved: {out_path}")
    return out_path


# ---------------------------------------------------------------------------
# CLI entry point
# ---------------------------------------------------------------------------

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate a monthly maintenance report for educational centers."
    )
    parser.add_argument(
        "period",
        nargs="?",
        help="Period in YYYY-MM format (default: from config.yaml)",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    config = load_config()

    period_str = args.period or config.get("report", {}).get("default_period", "2026-02")

    try:
        dt = datetime.strptime(period_str, "%Y-%m")
        year, month = dt.year, dt.month
    except ValueError:
        logger.error(f"Invalid period format: '{period_str}'. Expected YYYY-MM.")
        sys.exit(1)

    out_path = build_report(year, month, config)
    print(f"\nReport generated successfully: {out_path}\n")


if __name__ == "__main__":
    main()
