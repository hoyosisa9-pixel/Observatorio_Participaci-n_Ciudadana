# Grafo de conocimiento del OCE (memoria del proyecto)

Memoria tipo grafo: **nodos** (dimensiones, indicadores, variables, fuentes, decisiones, archivos) y **aristas**
(`A --relación--> B`). Actualízala cada vez que cambie el estado de un indicador. Decisiones: `decisiones.md`.
Memoria histórica previa: `.memory/knowledge_graph.md`.

Estados de indicador: `Se mantiene` · `Se ajusta` · `Se reemplaza` · `Eliminado` · `Pendiente` · `Nuevo`.
Revisión indicador por indicador: `outputs/revision_indicadores/plan_revision.md` (columna "Paso").

## 1. Nodos: dimensiones
- DIM.PP Presupuesto participativo · DIM.PC Participación ciudadana · DIM.CS Control social · DIM.PDM Cumplimiento del PDM 2024–2027

## 2. Nodos: indicadores (estado al 15-nov-2026)
| Nodo | Dimensión | Estado | Decisión | Serie hoy |
|---|---|---|---|---|
| PP.FID Fidelidad a lo votado | PP | Se mantiene (definido) | D12 | 2024, 2025 |
| PP.EPP Ejecución presupuestal del PP | PP | Se mantiene; revisar D4 | D4, D12 | 2024, 2025, 2026-S1 |
| PP.resp (compromiso, ≥90 %, equidad zonas, terminados) | PP | Respaldo | — | 2024–2025 |
| PC.I1 Incorporación de aportes | PC | Pendiente: diferenciar de PP.FID (D13) | D13 | proxy 2024–2025 |
| PC.I2 Cobertura territorial | PC | Se mantiene | D13 | 2024–2025 |
| PC.I3 Oferta formativa | PC | Por revisar (2024 N/D) | — | 2025 |
| PC.I4 Publicación de resultados | PC | Pendiente (sin inventario) | C3 | — |
| PC.VEED Veedurías inscritas en la Personería | PDM (reemplaza a CP) | Movido al PDM | D26 | 2012–2026 (listado incompleto) |
| PC.I5, PC.I6 | PC | Eliminado | D1 | — |
| PC.ISPC Índice | PC | Recalcular con la nueva selección | — | — |
| CS.PCS, CS.DSI, CS.VRC, CS.RPO, CS.ICRCS | CS | Se mantienen (definido); PCS es autorreporte (C6) | D12 | 2024–2025 |
| PDM.SALUD Cobertura de salud | PDM | Se reemplaza (antes: aseguramiento subsidiado) | D3, D5 | 2020–2025 |
| PDM.CM Mortalidad vial | PDM | Se mantiene; faltan 2025+ | D6 | 2024 |
| PDM.CA Espacio público verde | PDM | Se ajusta (LB 4,66; referente 9) | D7 | 2022, medición 5,75 |
| PDM.ILE Logro educativo | PDM | Se ajusta (cobertura neta + primera infancia) | D2, D8 | PI 2019–2025 |
| PDM.SP Servicios públicos domiciliarios | PDM | Nuevo | D9 | por construir |
| PDM.SEXTO | PDM | Pendiente (deporte/seguridad/vivienda/pobreza) | D10 | — |
| PDM.CP Procesos participativos | PDM | Se reemplaza por veedurías inscritas (PC.VEED se mueve aquí) | D26 | veedurías 2012–2026 |
| PDM.CD Servicios deportivos | PDM | Se reemplaza (candidato a SEXTO) | D10, D11 | — |
| PDM.ICPD Índice agregado | PDM | Recalcular con la nueva selección | — | — |
| PDM.POLPUB Logros de políticas públicas | PDM | Pendiente | — | — |

### 2.1 Hallazgos verificados (Analizador + QA, 15-nov-2026)
- **Apéndice 9.1.2 del PDM** (`116989_apendice91planindicativo27052024.xlsx`): fuente de LB y metas de los indicadores del PDM. Verificado: homicidio 3,7→2,7; hurto 632,5→631,5; lesiones personales 113,7→112,7; mortalidad vial 10,2→8 (Movilidad); espacio verde 3,88 estable; cobertura neta 100,44 «Monitoreo»; primera infancia 0–6 16,3→19; acueducto/alcantarillado/energía 99,73/99,56/99,87→99,9 (ECV); aseo sin indicador; subsidiado 99,5.
- **Veedurías**: datos.gov.co `jju5-br3w` (45 filas; inscritas 2019–2026: 2, 4, 11, 4, 6, 1, 5, 1). Faltan filas (consecutivos 118–293; `n` solo impares): la serie no es confiable tal cual.
- **Cobertura neta**: MinEducación `nudc-7mev`, 2011–2024 (código 5266/05266).
- **Policía**: tránsito `ha6j-pa2r` (43, 20, 30 muertes en 2023–2025); homicidio `m8fd-ahd9` (10, 7, 3, 9, 9, 12, 8 en 2019–2025); hurto a personas `4rxi-8m8d`.
- **MDM** `nkjx-rsq7`: solo 2016–2020 (descartada). Forensis y TerriData no se pueden consultar automáticamente.

```
PDM.SALUD --sin_meta_mismo_universo--> Ap.9.1.2 (solo subsidiado)
PDM.CA    --tres_cifras_en_conflicto--> 3,88 (PDM 2023) | 4,66 (AMVA 2022) | 5,75 (POT)
PDM.CM    --no_mezclar--> Forensis vs. Policía ; meta de Movilidad
PDM.SEXTO --candidatos--> homicidio (1), hurto (2), lesiones personales (por evaluar)
PDM.ILE   --limitación--> numerador con gestantes; fuente cambia en 2024
```

### 2.2 PDM: series construidas (Ejecutor + QA, 15-nov-2026)
- Script: `scripts/revision_indicadores/pdm_series.py`. Salidas en `outputs/revision_indicadores/PDM/` (`series_pdm.csv`, `series_candidatos_sexto.csv`, `reporte_ejecutor.md`, `qa_series_pdm.md`). Descargas en `data_raw_IIC/07_revision_nov2026/`.
- **TerriData**: usar el archivo COMPLETO `TerriData05266.txt.zip`, no el resumido «f» (que usa `calcular.py` y que no trae el REC anual ni la salud de 2018–2019).
- Salud 2018–2025: 82,71 → 83,06. ILE 2018–2024: 83,39 → 95,15. Servicios públicos (promedio de 4) 2019–2024: 91,97; 77,22; 89,74; 89,56; 89,76; 89,31. Homicidio 2019–2025: 4,25 … 3,02. Mortalidad vial: 9,35 … 11,34 (2023: 16,79 sin explicación).
- Sexta medida: candidatos homicidio, hurto y lesiones; NO reemplaza a CP (D19).
- Entrega al profe: `outputs/entrega_profe/` (1_PDM.md listo, dudas_para_el_profe.md; PC, PP y CS pendientes, CS de último).

## 3. Nodos: fuentes
FUE.TERRIDATA (DNP) · FUE.DANE (proyecciones de población) · FUE.MINSALUD (cifras de afiliación) · FUE.FORENSIS (INMLCF) ·
FUE.POLICIA · FUE.POT (diagnóstico de revisión, tabla 39) · FUE.AMVA · FUE.PDM (PDM 2024–2027 y Apéndice 9.1) ·
FUE.IGD (Informe de Gestión por Dependencias) · FUE.EFPC (listados de PP) · FUE.HACIENDA (ejecución) · FUE.DATOSGOV (datos.gov.co) ·
FUE.PERSONERIA · FUE.SUI (Superservicios) · FUE.ICFES · FUE.MDM (Medición de Desempeño Municipal)

## 4. Aristas
```
DIM.PDM --contiene--> PDM.SALUD, PDM.CM, PDM.CA, PDM.ILE, PDM.SP, PDM.SEXTO --agrega--> PDM.ICPD
DIM.PC  --contiene--> PC.I1, PC.I2, PC.I3, PC.I4 --agrega--> PC.ISPC
PDM.CP  --reemplazado_por--> PC.VEED (veedurías inscritas, D26)
DIM.CS  --contiene--> CS.PCS, CS.DSI, CS.VRC, CS.RPO --agrega--> CS.ICRCS
DIM.PP  --contiene--> PP.FID, PP.EPP
PDM.SALUD --reemplaza--> aseguramiento subsidiado ; --usa--> FUE.MINSALUD, FUE.DANE, FUE.TERRIDATA
PDM.CM    --usa--> FUE.FORENSIS, FUE.POLICIA, FUE.DANE
PDM.CA    --usa--> FUE.POT, FUE.AMVA, FUE.PDM ; --línea_base--> 4,66 m²/hab (2022) ; --referente--> 9 m²/hab
PDM.ILE   --usa--> FUE.TERRIDATA (cobertura neta, primera infancia), FUE.DANE ; --reemplaza_componente--> Saber 11
PDM.SP    --usa--> FUE.TERRIDATA, FUE.SUI
PC.VEED   --usa--> FUE.PERSONERIA, FUE.DATOSGOV
PC.I1     --posible_doble_conteo--> PP.FID
PDM.CD    --candidato_a--> PDM.SEXTO
CS.DSI    --depende_de--> metas seleccionadas del PDM (7)
```

## 5. Nodos: archivos clave
- Estado anterior: `outputs/sistema_simplificado/` (CSV, Excel con fórmulas, DOCX/PDF) · `propuesta_aprobada.md`
- Cálculos base: `outputs/indicadores/` · scripts en `scripts/indicadores/` y `scripts/sistema_simplificado/`
- Reunión 12-nov: `outputs/guia/Informe tecnico reunion OCE 12 noviembre 2026.docx`
- Revisión en curso: `outputs/revision_indicadores/`
- Datos crudos: **no están en este Codespace** (solo inventarios); ZIP en Google Drive.

## 6. Estado al cierre del 15-nov-2026 (las 4 dimensiones revisadas)
- Entregas: `outputs/entrega_profe/` → `indicadores_cambios_y_datos_OCE.docx` (para el profe, cuadros sencillos) y `detalle_revision_indicadores_OCE.docx` (para Isabella). Se generan con `scripts/entregables/generar_entrega_profe.py` a partir de los .md de la misma carpeta (las dudas se leen de `dudas_para_el_profe.md`).
- QA de construibilidad: `outputs/revision_indicadores/qa_construibilidad.md`. Ningún índice se puede calcular sin resolver las dudas. IREPP: 2024–2025 apenas se decidan las dudas 9 y 10. ICPD-6: 2022 y 2024. ISPC: 2025. ICRCS: ninguno sin respaldos.
- Cambios vigentes: PDM.CP → veedurías; PDM.CD → servicios públicos; PC.I1 → IPCM (votación PP); PC.I4 → lista fija de 3 procesos; PP.CEP → ejecución ≥ 90 % (propuesta); PP.CMA → fidelidad; CS.PCS → oportunidad de respuesta PQRSDF (propuesta).
- Descargas de hoy con manifiesto: `data_raw_IIC/07_revision_nov2026/` (efpc_listados_pp, pp_votaciones, informes_gestion, pqrs, hacienda_ejecucion, terridata, policía, MEN). No se versionan en git.
