# ESTADO – Etapa de cálculo de indicadores viables

**Punto de control:** 9-nov-2026, fin de la tanda 3. Avances detallados en `avance_1.md`, `avance_2.md` y `avance_3.md`.

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
| `scripts/indicadores/<sis>_<ind>.py` | Un script por indicador: `cs_pcs.py`, `cs_dsi.py`, `cs_vrc.py`, `cs_rpo.py` (con `cs_segunda_lectura.py`), `pdm_metas.py`, `pdm_ca.py`, `pdm_cm.py`, `pdm_no_viables.py`, `pc_i5.py`, `pc_i2.py`, `pc_i3.py`, `pc_no_viables.py`, `pp_epp.py` (importa `pp_epp_2024.py`), `pp_no_viables.py` (orden de ejecución en el README) |
| `scripts/indicadores/descargas.py` | Descargas (DANE, MinSalud, Policía) a `data_raw_IIC/06_descargas_indicadores/` con manifiesto |
| `scripts/indicadores/ensamblar.py` | Consolida `parciales/<SIS>/` en las tablas finales (CSV `;` UTF-8 con BOM + `base_indicadores.xlsx` + diccionario) |
| `scripts/indicadores/llenar_matriz.py` | Copia de la matriz con las hojas Registro y PP_Proyectos diligenciadas (Pendiente). Con `--verificar --zonas` compara contra la hoja Resultados usando una copia temporal |
| `outputs/indicadores/` | `valores_variables.csv`, `resultados_indicadores.csv` (939), `pp_proyectos.csv` (181), `soporte_conteos.csv` (582), `clasificacion_vrc.csv` (17), `clasificacion_rpo.csv` (56), `clasificacion_pp_2024.csv` (51), `clasificacion_pp_entregas.csv` (180), `segunda_lectura/` (muestras del 20 % para la usuaria), `diccionario_datos.csv`, `base_indicadores.xlsx`, `matriz_diligenciada.xlsx` |

Ejecución, desde la raíz y sin `-I`:
1. `python3 scripts/indicadores/<script>.py` para cada indicador, en el orden del README (en PDM el orden importa).
2. `python3 scripts/indicadores/ensamblar.py`.
3. `python3 scripts/indicadores/llenar_matriz.py --verificar --zonas`.

Requiere `libreoffice-calc` (instalado con apt; reinstalarlo si el Codespace se recrea).

## Tandas
| Tanda | Contenido | Estado |
|---|---|---|
| 1 | Módulo común, CS.PCS, metas del PDM, PDM.CA, PC.I5, PP.EPP 2025-2026 | **Terminada**: verificación contra la matriz con 21 resultados iguales y 14 diferencias esperadas (PP 2026 semestral y zona 1) |
| 2 | PDM.CM, PC.I2, PC.I3, CS.DSI, variables de los no viables | **Terminada**: 42 resultados comparados con la matriz, 28 iguales y 14 diferencias esperadas (las mismas de la tanda 1) |
| 3 | CS.VRC y CS.RPO (clasificación con una fila por evento u observación + muestra del 20 % para la usuaria), PP.EPP 2024 | **Terminada**: RPO 2024 = 94,1; VRC = N/D por defecto (presencia desconocida en la audiencia del Alcalde); EPP 2024 = 98,69 aproximado. 60 resultados comparados con la matriz, 46 iguales y 14 diferencias esperadas |
| 4 | CS.ICRCS, revisión crítica del Analizador, `resumen_calculos.md`, entregables finales | Pendiente |

## Pendientes y alertas
- **Forensis 2024:** la usuaria lo subirá a `data_raw_IIC/06_descargas_indicadores/`. Mientras tanto se usa 7,25 (p. 440) en estado Pendiente.
- **OPP de PP 2026:** la cifra correcta es $5.702.528.148 (Hacienda y ROCA). La cifra $11.555.041.247 citada en la versión 1.1 del Excel de disponibilidad era errónea; falta la decisión de la usuaria sobre corregir lo ya entregado.
- **CS.DSI = 0 (Rojo) en 2024 y 2025** por la ventana documental por defecto (31-mar del año siguiente). Con todo lo publicado hasta hoy, 2024 da 28,6. La lista de metas la definen los profesores.
- **Por confirmar:** PP 2024 tiene 51 proyectos con entrega exigible (no 50; registro dudoso en la zona 8); la cifra de 55 entregas cumplidas en 2025 no se reproduce (rango de 45 a 50); en PC.I2 se asignaron las asambleas de PP a las zonas 2 a 13; en PC.I3 se excluyó el diplomado en Gestión Pública; ILE.A usa el rango 0-6 por defecto y el informe dice "≤5".
- **Segunda lectura pendiente:** la usuaria debe llenar `segunda_lectura/muestra_vrc.csv` (4) y `muestra_rpo.csv` (12).
- **PP.EPP 2024:** se cruzan 50 de 51 proyectos ($18,6 mil millones), no 47 de 50 ($17,33) como en la nota previa; la descripción de `pp_2024_metodo` en `supuestos.yaml` quedó desactualizada. Hay dos cruces débiles por confirmar con Planeación.
- **PP_Proyectos:** el proyecto 0553 (costo del proceso, sin zona) se omite en la hoja, porque el valor por defecto lo excluye.

## Cómo retomar
Lee este archivo, `avance_1.md`, `plan/plan_calculo.md` (asignación por tanda y sistema) y `scripts/indicadores/README.md`. Lanza la tanda 4: CS.ICRCS (N/D por defecto porque VRC es N/D), revisión crítica del Analizador y `resumen_calculos.md`. Cada uno escribe solo sus scripts y `parciales/<SIS>/` y no modifica `comun.py`. Después ensambla, verifica contra la matriz y guarda el punto de control.
