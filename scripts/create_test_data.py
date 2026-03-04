import pandas as pd
import os
from pathlib import Path

# Destinos
data_dir = Path(r"C:\Users\Eduwi\.gemini\antigravity\scratch\reporting-system\data")
dist_data_dir = Path(r"C:\Users\Eduwi\.gemini\antigravity\scratch\reporting-system\dist-standalone\ManatechReports-win32-x64\resources\app\data")

for d in [data_dir, dist_data_dir]:
    d.mkdir(parents=True, exist_ok=True)

# Datos para FEBRERO y MARZO (formato americano MM/DD/YYYY)
data_visitas = {
    'Fecha_visita': ['02/01/2026', '02/15/2026', '03/01/2026', '03/05/2026'],
    'Centro': ['Escuela A', 'Escuela B', 'Liceo Central', 'Escuela Norte'],
    'Provincia': ['Santo Domingo', 'Santiago', 'Duarte', 'La Vega'],
    'UPS_estado': ['ok', 'averiado', 'ok', 'ok'],
    'Bandwidth_utilizado': [45, 88, 30, 92],
    'DHCP_saturacion': [10, 85, 20, 15],
    'AP_pendientes': [0, 2, 0, 1],
    'Uptime': [99.9, 92.1, 100, 98.5],
    'Observaciones': ['Todo normal', 'UPS requiere cambio', 'Nuevo liceo', 'AP pendiente configurar']
}

data_equipos = {
    'Fecha': ['02/05/2026', '03/02/2026'],
    'Centro': ['Escuela A', 'Liceo Central'],
    'Equipo': ['Access Point', 'Switch 24p'],
    'Serie_anterior': ['SN12345_old', 'SN98765_old'],
    'Serie_nueva': ['SN12345_new', 'SN98765_new'],
    'Motivo': ['Falla puerto', 'Mantenimiento'],
    'Tecnico': ['Eduwi', 'Eduwi']
}

df_v = pd.DataFrame(data_visitas)
df_e = pd.DataFrame(data_equipos)

# Guardar en ambos sitios
df_v.to_excel(data_dir / "visitas_centros.xlsx", index=False)
df_e.to_excel(data_dir / "cambios_equipos.xlsx", index=False)

df_v.to_excel(dist_data_dir / "visitas_centros.xlsx", index=False)
df_e.to_excel(dist_data_dir / "cambios_equipos.xlsx", index=False)

print("Datos de Febrero y Marzo generados correctamente.")
