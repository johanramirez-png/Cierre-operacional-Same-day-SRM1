# Dashboards Same Day MLC

Dashboards de análisis de ciclos Same Day — MercadoLibre Chile.

Actualización automática desde Google Sheets via Apps Script.

## Dashboards disponibles

| Dashboard | Descripción |
|-----------|-------------|
| [Dashboard Histórico SRM1](Dashboard_Historico_SameDay_MLC.html) | Histórico SRM1: PM1/PM2/PM3/PM4 |
| [Dashboard Histórico SRM2](Dashboard_Historico_SameDay_MLC_SRM2.html) | Histórico SRM2: PM1/PM2/PM3 |
| [Análisis Ciclos SRM1](Analisis_Ciclos_SameDay_MLC_SRM1.html) | Timelines y comparación SRM1 |
| [Análisis Ciclos SRM2](Analisis_Ciclos_SameDay_MLC_SRM2.html) | Timelines y comparación SRM2 |

## Flujo de actualización

```
Google Sheets (carga de datos)
        ↓
Apps Script (onChange trigger)
        ↓ ↓
   Drive   GitHub API
              ↓
       GitHub Pages (live)
```

Los dashboards se actualizan en línea en segundos después de cargar datos en la hoja.
