# ESTADO – Etapa de cálculo de indicadores viables

**Punto de control:** 9-nov-2026, fin de la tanda 1. Avance detallado en `avance_1.md`.

## Encargo
El profesor Francisco pidió dos cosas:
- organizar los datos de los indicadores calculables y construir las variables obtenibles de información secundaria;
- calcular los indicadores "Viable" y "Viable con ajuste" en sus años calculables.

Para los "No viables por ahora", solo se construyen las variables.

Reglas:
- Las decisiones metodológicas pendientes no las toma el equipo. Se modelan como escenarios en `config/supuestos.yaml`, con un valor por defecto PROVISIONAL, y se calculan todos.
- Nada "Validado": resultados "Provisional" y variables "Pendiente".
- Sin datos personales.

## Cómo está organizado
| Ruta | Contenido |
|---|---|
| `outputs/indicadores/plan/plan_calculo.md` y `supuestos_borrador.yaml` | Plan del Analizador: fórmulas y semáforos de la Matriz, fuentes, esquemas, asignación |
| `config/supuestos.yaml` | 36 parámetros con opciones y valor por defecto PROVISIONAL (zona 1 de PP NE/N/D; DSI con 7 metas del ICPD-6 por defecto) |
| `scripts/indicadores/comun.py` | Escenarios, semáforos exactos de la Matriz, reglas de la Guía, escritura validada (`--pruebas`) |
| `scripts/indicadores/<sis>_<ind>.py` | Un script por indicador: `cs_pcs.py`, `pdm_metas.py`, `pdm_ca.py`, `pc_i5.py`, `pp_epp.py` |
| `scripts/indicadores/descargas.py` | Descargas (DANE, MinSalud, Policía) a `data_raw_IIC/06_descargas_indicadores/` con manifiesto |
| `scripts/indicadores/ensamblar.py` | Consolida `parciales/<SIS>/` en las tablas finales (CSV `;` UTF-8 con BOM + `base_indicadores.xlsx` + diccionario) |
| `scripts/indicadores/llenar_matriz.py` | Copia de la matriz con las hojas Registro y PP_Proyectos diligenciadas (Pendiente). Con `--verificar --zonas` compara contra la hoja Resultados usando una copia temporal |
| `outputs/indicadores/` | `valores_variables.csv` (912), `resultados_indicadores.csv` (388), `pp_proyectos.csv` (129), `soporte_conteos.csv` (12), `diccionario_datos.csv`, `base_indicadores.xlsx`, `matriz_diligenciada.xlsx` |

Ejecución, desde la raíz y sin `-I`:
1. `python3 scripts/indicadores/<script>.py` para cada indicador.
2. `python3 scripts/indicadores/ensamblar.py`.
3. `python3 scripts/indicadores/llenar_matriz.py --verificar --zonas`.

Requiere `libreoffice-calc` (instalado con apt; reinstalarlo si el Codespace se recrea).

## Tandas
| Tanda | Contenido | Estado |
|---|---|---|
| 1 | Módulo común, CS.PCS, metas del PDM, PDM.CA, PC.I5, PP.EPP 2025-2026 | **Terminada**: verificación contra la matriz con 21 resultados iguales y 14 diferencias esperadas (PP 2026 semestral y zona 1) |
| 2 | PDM.CM, PC.I2, PC.I3, CS.DSI, variables de los no viables | Pendiente |
| 3 | CS.VRC y CS.RPO (clasificación con una fila por evento u observación + muestra del 20 % para la usuaria), PP.EPP 2024 | Pendiente |
| 4 | CS.ICRCS, revisión crítica del Analizador, `resumen_calculos.md`, entregables finales | Pendiente |

## Pendientes y alertas
- **Forensis 2024:** la usuaria lo subirá a `data_raw_IIC/06_descargas_indicadores/`. Mientras tanto se usa 7,25 (p. 440) en estado Pendiente.
- **OPP de PP 2026:** la cifra correcta es $5.702.528.148 (Hacienda y ROCA). La cifra $11.555.041.247 citada en la versión 1.1 del Excel de disponibilidad era errónea; falta la decisión de la usuaria sobre corregir lo ya entregado.
- **PP_Proyectos:** el proyecto 0553 (costo del proceso, sin zona) se omite en la hoja, porque el valor por defecto lo excluye.

## Cómo retomar
Lee este archivo, `avance_1.md`, `plan/plan_calculo.md` (asignación por tanda y sistema) y `scripts/indicadores/README.md`. Lanza la tanda 2 con un Ejecutor por sistema. Cada uno escribe solo sus scripts y `parciales/<SIS>/` y no modifica `comun.py`. Después ensambla, verifica contra la matriz y guarda el punto de control.
