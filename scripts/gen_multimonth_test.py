"""Generate multi-month test data with American date format MM/DD/YYYY."""
import pandas as pd
import random
from pathlib import Path

random.seed(42)
data_dir = Path(__file__).parent.parent / "data"
data_dir.mkdir(exist_ok=True)

centros = ['Escuela A', 'Escuela B', 'Liceo Central', 'Escuela Norte', 'Instituto Sur',
           'Colegio Este', 'Centro Oeste', 'Escuela Rural', 'Liceo Bilingue', 'Politecnico']
provincias = ['Santo Domingo', 'Santiago', 'Duarte', 'La Vega', 'San Cristobal']

rows_v = []
for month in [1, 2, 3]:
    for i in range(8):
        day = random.randint(1, 28)
        rows_v.append({
            'Fecha_visita': '%02d/%02d/2026' % (month, day),
            'Centro': random.choice(centros),
            'Provincia': random.choice(provincias),
            'UPS_estado': random.choice(['ok', 'ok', 'ok', 'averiado']),
            'Bandwidth_utilizado': round(random.uniform(20, 95), 2),
            'DHCP_saturacion': round(random.uniform(10, 99), 2),
            'AP_pendientes': random.randint(0, 3),
            'Uptime': round(random.uniform(85, 99.9), 2),
            'Observaciones': random.choice(['Sin novedad', 'Revision preventiva', 'Cable reemplazado']),
        })

rows_e = []
for month in [1, 2, 3]:
    for i in range(3):
        day = random.randint(1, 28)
        rows_e.append({
            'Fecha': '%02d/%02d/2026' % (month, day),
            'Centro': random.choice(centros),
            'Equipo': random.choice(['Router', 'Switch', 'Access Point', 'UPS']),
            'Serie_anterior': 'SN%d' % random.randint(10000, 99999),
            'Serie_nueva': 'SN%d' % random.randint(10000, 99999),
            'Motivo': random.choice(['Fallo tecnico', 'Fin de vida util', 'Dano por voltaje']),
            'Tecnico': random.choice(['Carlos', 'Ana', 'Luis']),
        })

df_v = pd.DataFrame(rows_v)
df_e = pd.DataFrame(rows_e)
df_v.to_excel(str(data_dir / 'visitas_centros.xlsx'), index=False)
df_e.to_excel(str(data_dir / 'cambios_equipos.xlsx'), index=False)

print("Created %d visits" % len(df_v))
print("Created %d equipment changes" % len(df_e))
print("Visits by month:")
for idx, row in df_v.iterrows():
    print("  %s -> %s" % (row['Fecha_visita'], row['Centro']))
