# 📊 Formato de Datos – Especificación de Archivos Excel

Este documento describe la estructura exacta requerida para los archivos Excel de entrada del sistema de reportes.

---

## 1. `visitas_centros.xlsx`

### Descripción

Contiene el registro de cada visita técnica realizada a los centros educativos.  
**Sheet name:** `Visitas` (o primera hoja del archivo)  
**Una fila = una visita**

### Columnas Requeridas

| # | Nombre de Columna | Tipo | Obligatorio | Descripción |
|---|---|---|---|---|
| 1 | `Centro` | Texto | ✅ Sí | Nombre completo del centro educativo |
| 2 | `Provincia` | Texto | ✅ Sí | Provincia donde se ubica el centro |
| 3 | `Fecha_visita` | Fecha (AAAA-MM-DD) | ✅ Sí | Fecha en que se realizó la visita |
| 4 | `UPS_estado` | Texto | ✅ Sí | Estado del UPS: `bueno`, `averiado`, `averiada`, `falla`, `malo`, `dañado` |
| 5 | `Bandwidth_utilizado` | Numérico (0–100) | ✅ Sí | Porcentaje de utilización del ancho de banda |
| 6 | `DHCP_saturacion` | Numérico (0–100) | ✅ Sí | Porcentaje de saturación del servidor DHCP |
| 7 | `AP_pendientes` | Entero (≥0) | ✅ Sí | Número de Access Points pendientes de configurar |
| 8 | `Uptime` | Numérico (0–100) | ✅ Sí | Porcentaje de tiempo en que la red estuvo disponible |
| 9 | `Observaciones` | Texto | ❌ Opcional | Notas del técnico sobre la visita |

### Valores válidos para `UPS_estado`

El sistema detecta automáticamente UPS averiados cuando el valor es:
- `averiado`
- `averiada`
- `falla`
- `malo`
- `dañado`

Cualquier otro valor (como `bueno`, `óptimo`, `funcionando`) se interpreta como UPS en buen estado.

### Ejemplo

| Centro | Provincia | Fecha_visita | UPS_estado | Bandwidth_utilizado | DHCP_saturacion | AP_pendientes | Uptime | Observaciones |
|---|---|---|---|---|---|---|---|---|
| Escuela Juan Pablo Duarte | Distrito Nacional | 2026-02-05 | bueno | 65.3 | 72.1 | 0 | 99.5 | Sin novedad |
| Liceo Salomé Ureña | Santiago | 2026-02-12 | averiado | 88.7 | 91.4 | 3 | 87.2 | Se reportó UPS sin batería |
| Centro Los Pinos | La Vega | 2026-02-18 | bueno | 45.0 | 55.0 | 1 | 98.1 | Se actualizó firmware |

---

## 2. `cambios_equipos.xlsx`

### Descripción

Contiene el registro de cada cambio o reemplazo de equipo electrónico realizado en los centros.  
**Sheet name:** `Cambios` (o primera hoja del archivo)  
**Una fila = un cambio de equipo**

### Columnas Requeridas

| # | Nombre de Columna | Tipo | Obligatorio | Descripción |
|---|---|---|---|---|
| 1 | `Centro` | Texto | ✅ Sí | Nombre del centro donde se realizó el cambio |
| 2 | `Fecha` | Fecha (AAAA-MM-DD) | ✅ Sí | Fecha del cambio de equipo |
| 3 | `Equipo` | Texto | ✅ Sí | Tipo de equipo (Router, Switch, AP, UPS, PC, etc.) |
| 4 | `Serie_anterior` | Texto | ✅ Sí | Número de serie del equipo retirado |
| 5 | `Serie_nueva` | Texto | ✅ Sí | Número de serie del equipo instalado |
| 6 | `Motivo` | Texto | ✅ Sí | Razón del cambio (fallo técnico, fin de vida útil, etc.) |
| 7 | `Tecnico` | Texto | ✅ Sí | Nombre del técnico responsable |

### Ejemplo

| Centro | Fecha | Equipo | Serie_anterior | Serie_nueva | Motivo | Tecnico |
|---|---|---|---|---|---|---|
| Escuela Juan Pablo Duarte | 2026-02-07 | Router | SN45231 | SN67890 | Fallo técnico | Carlos Rodríguez |
| Liceo Salomé Ureña | 2026-02-14 | UPS | SN11234 | SN99871 | Daño por voltaje | Ana Martínez |
| Centro Los Pinos | 2026-02-20 | Switch | SN33410 | SN55123 | Fin de vida útil | Luis Pérez |

---

## ⚠️ Notas Importantes

1. **Nombres de columnas exactos**: Los nombres deben coincidir exactamente (incluyendo mayúsculas y guiones bajos).
2. **Formato de fechas**: Usar formato `AAAA-MM-DD` o dejar que Excel maneje el formato de fecha nativo.
3. **Valores numéricos**: No incluir el símbolo `%` en las celdas numéricas; solo el número (ej: `65.3` no `65.3%`).
4. **Celdas vacías**: Las columnas obligatorias no deben tener celdas vacías. Para `Observaciones` (opcional), las celdas vacías se reemplazarán automáticamente con `—`.
5. **Datos del período**: El sistema filtra automáticamente por mes y año, por lo que el archivo puede contener datos de múltiples meses.

---

## 🔍 Validación de Datos

El sistema aplica automáticamente las siguientes correcciones al cargar los datos:

| Tipo de dato | Corrección aplicada |
|---|---|
| Fechas en texto | Conversión automática a datetime |
| Valores numéricos inválidos | Se convierten a `NaN` y se excluyen de promedios |
| Espacios en nombres de columnas | Se eliminan automáticamente |
| Estado de UPS | Normalizado a minúsculas para comparación |
| Observaciones vacías | Reemplazadas con `—` |
