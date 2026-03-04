# 📡 Sistema Automatizado de Reportes de Mantenimiento

> Generador automático de reportes mensuales en Word para el Proyecto de Mantenimiento a la Conectividad de los Centros Educativos.

---

## 📋 Descripción General

Este sistema lee datos operacionales desde archivos Excel y genera automáticamente un documento Word estructurado y profesional que incluye:

- Descripción general del período
- Resumen de operaciones (KPIs)
- Detalle de centros visitados
- Análisis de UPS averiados
- Estadísticas de ancho de banda
- Cambios de equipos electrónicos
- Análisis de uptime
- Saturación DHCP
- Access Points pendientes de configurar

---

## 🗂️ Estructura del Proyecto

```
reporting-system/
├── assets/
│   └── logo_empresa.png          # Logo de la empresa (reemplazar con el real)
├── config/
│   └── config.yaml               # Configuración global del sistema
├── data/
│   ├── visitas_centros.xlsx       # Base de datos de visitas
│   └── cambios_equipos.xlsx      # Base de datos de cambios de equipos
├── docs/
│   ├── README.md                 # Este archivo
│   └── data_format.md            # Formato requerido de los Excel
├── reports/                      # Aquí se generan los reportes .docx
├── scripts/
│   ├── data_processing.py        # Carga y limpieza de datos
│   ├── statistics.py             # Cálculo de métricas y KPIs
│   ├── generate_report.py        # Generador del reporte Word (punto de entrada)
│   └── create_sample_data.py     # Generador de datos de ejemplo para pruebas
├── templates/
│   └── reporte_template.docx     # Plantilla Word (opcional, se crea si no existe)
├── .gitignore
└── requirements.txt
```

---

## ⚙️ Instalación de Dependencias

### Requisitos previos

- Python 3.9 o superior
- pip

### Instalar dependencias

```bash
pip install -r requirements.txt
```

---

## 📊 Formato de los Archivos Excel

Ver [`docs/data_format.md`](data_format.md) para la especificación completa de columnas requeridas.

### Resumen rápido

**`data/visitas_centros.xlsx`** — Una fila por visita:
| Centro | Provincia | Fecha_visita | UPS_estado | Bandwidth_utilizado | DHCP_saturacion | AP_pendientes | Uptime | Observaciones |

**`data/cambios_equipos.xlsx`** — Una fila por cambio:
| Centro | Fecha | Equipo | Serie_anterior | Serie_nueva | Motivo | Tecnico |

---

## 🚀 Cómo Generar el Reporte

### 1. Preparar los datos (primera vez o de prueba)

```bash
python scripts/create_sample_data.py
```

Esto crea archivos Excel de ejemplo en `data/` con datos ficticios para febrero 2026.

### 2. Generar el reporte mensual

```bash
python scripts/generate_report.py YYYY-MM
```

**Ejemplos:**

```bash
# Reporte de febrero 2026
python scripts/generate_report.py 2026-02

# Reporte de enero 2026
python scripts/generate_report.py 2026-01

# Reporte del mes configurado por defecto en config.yaml
python scripts/generate_report.py
```

### 3. Resultado

El reporte se genera en:

```
reports/reporte_YYYY_MM.docx
```

---

## ⚙️ Configuración

El archivo `config/config.yaml` controla todos los parámetros del sistema:

| Parámetro | Descripción | Valor por defecto |
|---|---|---|
| `report.thresholds.dhcp_saturation_pct` | Umbral de alerta de DHCP (%) | `80` |
| `report.thresholds.bandwidth_high_pct` | Umbral de alto consumo de BW (%) | `70` |
| `report.thresholds.uptime_low_pct` | Umbral de uptime bajo (%) | `95` |
| `report.logo.width_cm` | Ancho del logo en el reporte | `5.0` |
| `report.default_period` | Período por defecto si no se pasa argumento | `2026-02` |
| `company.name` | Nombre de la empresa | `TechNet Soluciones` |

---

## 🎨 Personalización de la Plantilla

1. Crea o edita `templates/reporte_template.docx` con los estilos que prefieras en Word.
2. El generador detecta automáticamente el template y lo usa como base.
3. Si no existe el template, genera el documento con estilos predeterminados.

---

## 🖼️ Reemplazar el Logo

1. Coloca el logo de tu empresa en `assets/logo_empresa.png`
2. El logo se insertará automáticamente en la portada del reporte.
3. Ajusta el tamaño en `config.yaml` → `report.logo.width_cm`

---

## 🔄 Flujo del Sistema

```
Excel Files
    ↓
data_processing.py  →  Carga y limpieza
    ↓
statistics.py       →  KPIs y análisis
    ↓
generate_report.py  →  Construcción del Word
    ↓
reports/reporte_YYYY_MM.docx
```

---

## 🐛 Solución de Problemas

| Problema | Solución |
|---|---|
| `FileNotFoundError: Excel file not found` | Verifica que los `.xlsx` estén en `data/` |
| `ModuleNotFoundError` | Ejecuta `pip install -r requirements.txt` |
| Logo no aparece | Verifica que `assets/logo_empresa.png` exista |
| Columnas no reconocidas | Revisa `docs/data_format.md` para el nombre exacto de columnas |

---

## 📁 Control de Versiones

El proyecto está listo para Git. Para inicializar:

```bash
git init
git add .
git commit -m "Initial commit: automated reporting system"
```

El archivo `.gitignore` ya excluye:
- Los reportes generados (`reports/`)
- Los datos Excel (`data/`)
- Archivos de entorno virtual y caché Python

---

## 👥 Autores

Sistema desarrollado por el equipo de **TechNet Soluciones** – Dirección de Mantenimiento de Conectividad.
