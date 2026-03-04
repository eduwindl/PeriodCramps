"""
create_sample_data.py
=====================
One-time script to generate realistic sample Excel files for testing.

Run from the project root:
    python scripts/create_sample_data.py
"""

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))

import pandas as pd
import numpy as np
from datetime import date, timedelta
import random

ROOT = Path(__file__).parent.parent
RANDOM_SEED = 42
rng = random.Random(RANDOM_SEED)
np.random.seed(RANDOM_SEED)

CENTROS = [
    "Escuela Básica Juan Pablo Duarte",
    "Liceo Secundario Salomé Ureña",
    "Centro Educativo Los Pinos",
    "Escuela Primaria La Esperanza",
    "Instituto Politécnico Noreste",
    "Colegio Santo Tomás",
    "Escuela Rural El Limón",
    "Liceo Nocturno Pedro Mir",
    "Centro UASD Extensión Este",
    "Escuela Básica Enrique Henríquez",
    "Instituto de Formación Técnica Norte",
    "Centro Educativo Divina Providencia",
    "Escuela Primaria Villa Juana",
    "Liceo Bilingüe Las Américas",
    "Centro Educativo Fe y Alegría",
]

PROVINCIAS = [
    "Distrito Nacional", "Santo Domingo", "Santiago", "La Vega",
    "San Pedro de Macorís", "La Romana", "San Cristóbal", "Duarte",
    "Espaillat", "Puerto Plata",
]

TECNICOS = ["Carlos Rodríguez", "Ana Martínez", "Luis Pérez", "María García", "José Santana"]

EQUIPOS = ["Router", "Switch", "Access Point", "UPS", "PC", "Servidor", "Patch Panel"]

MOTIVOS = [
    "Fallo técnico", "Daño por voltaje", "Fin de vida útil",
    "Actualización tecnológica", "Robo o extravío", "Daño por humedad",
]

UPS_ESTADOS = ["bueno", "bueno", "bueno", "bueno", "averiado", "averiada", "bueno"]


def random_date_in_month(year: int, month: int) -> date:
    if month == 12:
        end_month = date(year + 1, 1, 1) - timedelta(days=1)
    else:
        end_month = date(year, month + 1, 1) - timedelta(days=1)
    start = date(year, month, 1)
    delta = (end_month - start).days
    return start + timedelta(days=rng.randint(0, delta))


def create_visits(year: int = 2026, month: int = 2, n: int = 25) -> pd.DataFrame:
    rows = []
    for _ in range(n):
        centro = rng.choice(CENTROS)
        provincia = rng.choice(PROVINCIAS)
        fecha = random_date_in_month(year, month)
        ups = rng.choice(UPS_ESTADOS)
        bw = round(rng.uniform(20.0, 95.0), 2)
        dhcp = round(rng.uniform(30.0, 99.0), 2)
        ap = rng.randint(0, 5)
        uptime = round(rng.uniform(80.0, 99.99), 2)
        obs = rng.choice([
            "Sin novedad", "Red estable", "Se reconfiguraron VLANs",
            "Se actualizó firmware", "Revisión preventiva completada",
            "Se reemplazaron cables dañados", "Monitoreo de red instalado",
        ])
        rows.append({
            "Centro": centro,
            "Provincia": provincia,
            "Fecha_visita": fecha,
            "UPS_estado": ups,
            "Bandwidth_utilizado": bw,
            "DHCP_saturacion": dhcp,
            "AP_pendientes": ap,
            "Uptime": uptime,
            "Observaciones": obs,
        })
    return pd.DataFrame(rows)


def create_equipment(year: int = 2026, month: int = 2, n: int = 18) -> pd.DataFrame:
    rows = []
    for _ in range(n):
        centro = rng.choice(CENTROS)
        fecha = random_date_in_month(year, month)
        equipo = rng.choice(EQUIPOS)
        serie_ant = f"SN{rng.randint(10000, 99999)}"
        serie_nueva = f"SN{rng.randint(10000, 99999)}"
        motivo = rng.choice(MOTIVOS)
        tecnico = rng.choice(TECNICOS)
        rows.append({
            "Centro": centro,
            "Fecha": fecha,
            "Equipo": equipo,
            "Serie_anterior": serie_ant,
            "Serie_nueva": serie_nueva,
            "Motivo": motivo,
            "Tecnico": tecnico,
        })
    return pd.DataFrame(rows)


def save_excel(df: pd.DataFrame, path: Path, sheet_name: str = "Datos") -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(str(path), engine="openpyxl", date_format="YYYY-MM-DD") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        ws = writer.sheets[sheet_name]
        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col) + 4
            ws.column_dimensions[col[0].column_letter].width = min(max_len, 40)
    print(f"  ✓ Saved: {path} ({len(df)} rows)")


if __name__ == "__main__":
    print("Creating sample Excel data files...\n")
    visits_df = create_visits(2026, 2, 25)
    save_excel(visits_df, ROOT / "data" / "visitas_centros.xlsx", sheet_name="Visitas")

    equip_df = create_equipment(2026, 2, 18)
    save_excel(equip_df, ROOT / "data" / "cambios_equipos.xlsx", sheet_name="Cambios")

    print("\nDone! You can now run:")
    print("  python scripts/generate_report.py 2026-02")
