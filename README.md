# 📡 Sistema Automatizado de Reportes de Mantenimiento

Generador automático de reportes mensuales en Word para el Proyecto de Mantenimiento a la Conectividad de los Centros Educativos.

## Inicio Rápido

```bash
# 1. Instalar dependencias
pip install -r requirements.txt

# 2. Crear datos de ejemplo
python scripts/create_sample_data.py

# 3. Generar reporte de febrero 2026
python scripts/generate_report.py 2026-02
```

El reporte se genera en `reports/reporte_2026_02.docx`.

## Documentación Completa

Ver [`docs/README.md`](docs/README.md) para instrucciones detalladas.  
Ver [`docs/data_format.md`](docs/data_format.md) para el formato de los archivos Excel.
