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
| CS.PQRSDF (reemplaza a CS.PCS), CS.DSI (10 metas), CS.VRC, CS.RPO, CS.ICRCS | CS | Construido (20-nov): ICRCS 2024 = 60,59, 2025 = 38,18 | D45, D75–D82 | ver §7 |
| PDM.SALUD Cobertura de salud | PDM | Se reemplaza (antes: aseguramiento subsidiado) | D3, D5 | 2020–2025 |
| PDM.CM Mortalidad vial | PDM | Se mantiene; faltan 2025+ | D6 | 2024 |
| PDM.CA Espacio público verde | PDM | Se ajusta (urbano; LB 2019 = 4,54; referente 9) | D7, D30, D31 | cortes 2019, 2022, 2024 |
| PDM.ILE Logro educativo | PDM | Se ajusta (cobertura neta + primera infancia) | D2, D8 | PI 2019–2025 |
| PDM.SP Servicios públicos domiciliarios | PDM | Nuevo | D9 | por construir |
| PDM.HOM Homicidio (7.º componente) | PDM | Nuevo, construido | D34, D58 | 2003–2025 |
| PDM.CP Procesos participativos | PDM | Se reemplaza por veedurías inscritas (PC.VEED se mueve aquí) | D26 | veedurías 2012–2026 |
| PDM.CD Servicios deportivos | PDM | Reemplazado por servicios públicos | D17 | — |
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
PDM.HOM  --es--> 7.º componente del ICPD-7 (D34)
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
DIM.PDM --contiene--> PDM.SALUD, PDM.ILE, PDM.SP, PDM.CM, PDM.CA, PDM.VEED, PDM.HOM --agrega--> PDM.ICPD7 (promedio simple)
DIM.PC  --contiene--> PC.I1, PC.I2, PC.I3, PC.I4 --agrega--> PC.ISPC
PDM.CP  --reemplazado_por--> PC.VEED (veedurías inscritas, D26)
DIM.CS  --contiene--> CS.PQRSDF, CS.DSI, CS.VRC, CS.RPO --agrega--> CS.ICRCS
DIM.PP  --contiene--> PP.FID, PP.EPP
PDM.SALUD --reemplaza--> aseguramiento subsidiado ; --usa--> FUE.MINSALUD, FUE.DANE, FUE.TERRIDATA
PDM.CM    --usa--> FUE.FORENSIS, FUE.POLICIA, FUE.DANE
PDM.CA    --usa--> FUE.POT, FUE.AMVA ; --línea_base--> 4,54 m²/hab (2019, AMVA-UNAL; D31) ; --referente--> 9 m²/hab ; --solicitar_a--> Planeación (serie anual, método AMVA)
PDM.ILE   --usa--> FUE.TERRIDATA (cobertura neta, primera infancia), FUE.DANE ; --reemplaza_componente--> Saber 11
PDM.SP    --usa--> FUE.TERRIDATA, FUE.SUI
PC.VEED   --usa--> FUE.PERSONERIA, FUE.DATOSGOV
PC.I1     --se_reemplaza_por--> IPCM (votación PP; D28, D32) ; --línea_base--> 3,68 % (2017, población VP; D69) ; --referente--> 7 % (tope de la escala del IPCM) ; --solicitar_a--> Planeación (2018–2024)
PDM.CD    --reemplazado_por--> PDM.SP
CS.DSI    --depende_de--> 10 metas del PDM ligadas al ICPD-7 (D75)
CS.VRC    --se_reporta_como--> rango + brecha de información (D33, decidido; la brecha debe bajar año a año) ; --necesita--> listas de asistencia (Alcaldía), veedurías inscritas (Personería), observación directa
```

## 5. Nodos: archivos clave
- Estado anterior: `outputs/sistema_simplificado/` (CSV, Excel con fórmulas, DOCX/PDF) · `propuesta_aprobada.md`
- Cálculos base: `outputs/indicadores/` · scripts en `scripts/indicadores/` y `scripts/sistema_simplificado/`
- Reunión 12-nov: `outputs/guia/Informe tecnico reunion OCE 12 noviembre 2026.docx`
- Revisión en curso: `outputs/revision_indicadores/`
- Datos crudos: **no están en este Codespace** (solo inventarios); ZIP en Google Drive.

## 6. Estado al cierre del 15-nov-2026 (HISTÓRICO: superado por la §7)
- Entregas: `outputs/entrega_profe/` → `indicadores_cambios_y_datos_OCE.docx` (para el profe, cuadros sencillos) y `detalle_revision_indicadores_OCE.docx` (para Isabella). Se generan con `scripts/entregables/generar_entrega_profe.py` a partir de los .md de la misma carpeta (las dudas se leen de `dudas_para_el_profe.md`).
- QA de construibilidad: `outputs/revision_indicadores/qa_construibilidad.md`. Ningún índice se puede calcular sin resolver las dudas. IREPP: 2024–2025 apenas se decidan las dudas 9 y 10. ICPD-6: 2019, 2022 y 2024. ISPC: 2025. ICRCS: ninguno sin respaldos.
- Cambios vigentes: PDM.CP → veedurías; PDM.CD → servicios públicos; PC.I1 → IPCM (votación PP); PC.I4 → lista fija de 3 procesos; PP.CEP → ejecución ≥ 90 % (propuesta); PP.CMA → fidelidad; CS.PCS → oportunidad de respuesta PQRSDF (propuesta).
- Descargas de hoy con manifiesto: `data_raw_IIC/07_revision_nov2026/` (efpc_listados_pp, pp_votaciones, informes_gestion, pqrs, hacienda_ejecucion, terridata, policía, MEN). No se versionan en git.
- PDM.CA --línea_base--> AMVA-UNAL 4,54 (2019; D31) ; --corte--> AMVA-Humboldt 4,66 (2022) ; --medición--> POT 5,75 (2024) ; --excluye--> autorreporte 5,00 (2020/2023) y 3,88 PDM ; primer dato 2013 (ICAU 4,16). Fuentes: Reporte ODS 11.7 AMVA/ONU-Hábitat 2022, ICAU 2013 Minambiente, PDM 2016-2019, PDM 2020-2023, evaluación del POT.

## 7. Estado al 18-nov-2026: dudas respondidas, fase de CONSTRUCCIÓN
- Respuestas del profe: D34–D45 en `decisiones.md`. Plan de construcción: `outputs/revision_indicadores/plan_construccion.md`.
- Regla transversal: cada indicador e índice reporta su **brecha de información** (años o eventos sin dato).
- PDM.ICPD pasa a **7 componentes** (+ homicidio). PC.I2 con 12 zonas. Decidido el 18-nov: LB veedurías = promedio 2019–2023 (5,4); referente vial 7,33 (D47); ICPD-7 = promedio simple (D48); I3 según D49.
- **PP CONSTRUIDO (18-nov; actualizado 19-nov con D88/D89)**: IREPP 2020 = 57,98; 2021 = 90,71 (lista oficial, 45 proyectos sin INDER); 2022 = 85,51; 2023 = 89,15; 2024 = 93,56; 2025 = 95,83. Brecha 0/6. Se cruza el valor votado con la apropiación inicial de Hacienda, sumando por sufijo de 4 dígitos (cada sufijo tiene zona fija). Salidas en `outputs/construccion/PP/`. Siguiente: PDM (ICPD-7).
PP.EPP, PP.EJ90, PP.FID, PP.IREPP --serie--> 2020–2025 ; --universo_2021--> lista oficial 29399 (D88) ; --rubros--> 2.3 + 2.4 (D89)
- **PDM CONSTRUIDO (18-nov)**: ICPD-7 2019 = 62,41; 2022 = 68,79; 2024 = 66,85 (Naranja); brecha 4/7. Salidas en `outputs/construccion/PDM/`. Población única = DANE VP (TerriData «Descripción general»). Siguiente: Participación ciudadana (ISPC).
PDM.ICPD7 --componentes--> SALUD, ILE, SP, CM, CA, VEED, HOM (1/7 cada uno) ; --años--> 2019, 2022, 2024
- **PC CONSTRUIDO (18-nov)**: ISPC 2025 = 61,62 (Naranja; 69,95 si se publica la rendición de PP 2025). I1 (votos de PP): 2017 = 8.221, 2020 = 9.956, 2021 = 16.828, 2022 = 12.541, 2023 = 9.156, 2025 = 10.662; 2018 sin votación. Brecha del ISPC: 5/6. Salidas en `outputs/construccion/PC/`. Siguiente: Control social (ICRCS).
PC.I1 --convención--> año de la votación (vigencia t+1) ; --LB--> 3,68 (VP) ; PC.I2 --denominador--> 12 ; PC.ISPC --años--> 2025
- **CS CONSTRUIDO PARCIAL (19-nov)**: el ICRCS es N/D en todos los años porque falta PQRSDF (portal caído). Rangos posibles: 2024 = 40,20–65,20 y 2025 = 17,50–42,50. DSI = 10 metas (D75), 2024 0–40 y 2025 0–10. VRC 2024 = 66,67 y 2025 = 20. RPO 2021 = 84,62–100, 2023 = 0, 2024 = 94,12, 2025 = 50. CS.PCS fue reemplazado por CS.PQRSDF. Salidas en `outputs/construccion/CS/`.
CS.PQRSDF --fuente--> informes trimestrales BPMS (SAC como respaldo) ; --pendiente--> cs_descargas.py cuando responda el portal
- **Cierre 19-nov**: QA final aprobado con observaciones. Entregas: los Word en `outputs/entrega_profe/` y la base para Power BI en `outputs/entrega_powerbi/base_indicadores_OCE.xlsx`. Pendiente: PQRSDF (servidor de archivos de la Alcaldía con error 500).
- **CS COMPLETADO (20-nov)**: PQRSDF 24/24 trimestres extraídos del ZIP de Drive (D85). ICRCS 2024 = 60,59 Naranja; 2025 = 38,18 Rojo. Supera la línea «CS CONSTRUIDO PARCIAL».
FUE.DRIVE_ZIP --respaldo_de--> FUE.PORTAL_ENVIGADO (archivos idénticos al inventario; IDs fuera del repositorio)

## 8. Estado al 19-nov-2026 (noche): brechas del Drive cerradas
- Aplicadas D88–D91: PP con brecha 0/6; PC I1 2/8 e I2 1/6; el ISPC sigue solo en 2025.
- Entregables vigentes: el Excel de 5 hojas (`outputs/entrega_powerbi/base_indicadores_OCE.xlsx`) y en `outputs/entrega_profe/` el Word resumido, el Word detallado del profe, la ficha metodológica y el Word de Isabella.
PC.I1.ZONAS --zonas--> 2–13 (la zona 1, industrial, no tiene plan zonal) ; PC.I1 --denominador--> población total (Decreto 242/2017, p. 31)
