# Plan de revisión indicador por indicador (OCE, Producto 2)

**Estado:** PROVISIONAL · elaborado por el Analizador el 15-nov-2026 · nada validado.
**Insumos:** `.claude/memoria/grafo_conocimiento.md`, `.claude/memoria/decisiones.md` (C1–C8, D1–D14), `outputs/sistema_simplificado/trabajo/propuesta_aprobada.md`, `outputs/sistema_simplificado/indicadores_simplificados.csv`, `outputs/sistema_simplificado/variables_por_dimension.csv`, informe de la reunión del 12-nov-2026 (`outputs/guia/Informe tecnico reunion OCE 12 noviembre 2026.docx`) y consultas web del 15-nov-2026.

**Convención de verificación**
- **[V]** = verificado hoy: consulté la API, descargué el documento o leí la fila.
- **[V-rep]** = verificado en trabajo anterior del repositorio (URL y cifra en los CSV), no reconsultado hoy.
- **[V-reu]** = visto en pantalla en la reunión del 12-nov, según el informe. No lo reconsulté en la web.
- **[S]** = supuesto o pendiente de verificar.

**Revisión de QA (15-nov-2026):** aprobado con observaciones; ya incorporadas en este archivo. Reporte completo: `outputs/revision_indicadores/qa_plan_revision.md`.

**Hallazgo transversal [V]:** el Apéndice 9.1.2 del PDM 2024–2027 (indicadores de resultado, con línea base, año, meta 2027 y fuente) contiene metas para casi todos los indicadores del PDM que se revisan aquí. Es la referencia para aplicar el criterio C8 (metas del PDM) y el C7 (mismo universo). URL: https://alcaldia-municipal-de-envigado-en-antioquia.micolombiadigital.gov.co/sites/alcaldia-municipal-de-envigado-en-antioquia/content/files/002340/116989_apendice91planindicativo27052024.xlsx (hoja «Apéndice9.1.2.»). En la tabla, «Ap. 9.1.2» remite a esa hoja.

---

## 1. Tabla maestra

Veredictos: Se mantiene / Se ajusta / Se reemplaza / Eliminado / Pendiente / Nuevo. Prioridad: A (alta: el docente lo pidió o la entrega lo exige), M (media), B (baja).

### 1.1 Cumplimiento del PDM (PDM)

| Paso | Código | Indicador | Estado anterior | Decisión de las reuniones | Veredicto propuesto | ¿Serie posible? | Fuente(s) pública(s) verificada(s) | Qué falta | Prior. |
|---|---|---|---|---|---|---|---|---|---|
| 2 | PDM.SEXTO | Sexto indicador (seguridad, deporte, vivienda o pobreza) | No existía | D10: pendiente; debe tener vínculo con el PDM y 4–5 años de serie | **Nuevo**. Recomendado: **tasa de homicidio** (alternativa: tasa de hurto a personas). Ver §2.3 | Sí: 2019–2025 [V] (conteos anuales) | Homicidio: https://www.datos.gov.co/resource/m8fd-ahd9.json (MinDefensa/Policía, 2003–sep-2026) [V]. Hurto a personas: https://www.datos.gov.co/resource/4rxi-8m8d.json [V]. Meta y línea base: Ap. 9.1.2, línea 4 [V] | Que la usuaria o el docente elijan. Calcular la tasa con la población DANE única. Redactar la ficha | A |
| 3 | PDM.CA | Espacio público verde | 100 (2024; 5,75 vs. 3,88), Verde | D7: se mantiene. LB AMVA 4,66 (2022), referente 9; verificar 2024–2025; no mezclar verde, efectivo y total | **Se ajusta**: logro = 100·(valor − 4,66)/(9 − 4,66), tope 0–100, solo con espacio **verde urbano** | Débil: 2022 (4,66) y una medición de 5,75 [V-rep]. 2023 en el PDM: 3,88 [V]. 2024–2025 sin confirmar [S] | Ap. 9.1.2: «Espacio público verde» LB **3,88 m²/hab (2023)**, meta 2027 **3,88**, tendencia «Estable» [V]. POT, tabla 39: https://alcaldia-municipal-de-envigado-en-antioquia.micolombiadigital.gov.co/sites/alcaldia-municipal-de-envigado-en-antioquia/content/files/002925/146211_8-dimension-funcional-_-espacio-publico.pdf [V-rep] | Buscar la cifra 2024 y 2025 en los informes de gestión 2024 y 2025 (URLs en §5) y en el informe de seguimiento del PDM. Explicar por qué difieren 3,88 (PDM), 4,66 (AMVA) y 5,75 (POT) | A |
| 4 | PDM.CM | Mortalidad vial | 100 (2024; 7,25 Forensis vs. meta 8); Policía 8,05 → 99,32 | D6: se mantiene; faltan los años posteriores a 2024 | **Se mantiene** (con ajuste del universo de la meta) | Sí con la Policía: 2019–2025 [V]. Forensis: 2024 [V-rep]; 2025 no verificado [S] | Policía, homicidio culposo en accidente de tránsito: https://www.datos.gov.co/resource/ha6j-pa2r.json, Envigado: 22, 12, 21, 29, 43, 20, **30** muertes (2019–2025), datos hasta 31-jul-2026 [V]. Forensis 2024: https://www.medicinalegal.gov.co/documents/20143/1269218/Forensis_2024.pdf [V-rep]. El sitio del INMLCF devolvió 403 a la consulta automática | Calcular 2025 con la población DANE única. Hoy el respaldo de la Policía 2024 (8,05) usa una población de 248.304 y la salud usa 260.430: unificar. La meta (8) y la LB 10,2 (2023) del Ap. 9.1.2 tienen como fuente a la **Secretaría de Movilidad**, no a Forensis (C7). Revisar el salto de 2023 de la Policía (43) | A |
| 5 | PDM.ILE | Índice de logro educativo: cobertura neta + atención integral a primera infancia | Primera infancia 2019–2025 (17,53 en 2024 y 17,12 en 2025, población 0–5); Saber 11 como respaldo | D2, D8: sustituir Saber 11 por la cobertura neta total; conservar primera infancia y el nombre; fórmula por definir; conciliar 0–5 con 0–6 | **Se ajusta** | Cobertura neta: **2011–2024** [V] (usar 2018+, ver nota). Primera infancia: 2019–2025 [V-rep] | MinEducación (datos.gov.co): https://www.datos.gov.co/resource/nudc-7mev.json. Envigado: cobertura neta 96,96 (2018), 97,06, 96,85, 98,04, **100,44 (2022)**, 99,68, 98,45 (2024) [V]. Ap. 9.1.2: «Tasa de cobertura escolar neta» LB 100,44 (2022), meta 100,44, tendencia «Monitoreo» (coincide exactamente con el MEN) [V]. «Tasa de cobertura en atención integral de niños y niñas de **0 a 6 años**» LB 16,3 (2023), meta 19, Secretaría de Educación [V]. TerriData (atendidos): https://terridata.dnp.gov.co/ [V-rep] | Fórmula (pregunta P1). Rango de edad (P2): el PDM dice 0–6, pero 0–5 reproduce mejor la LB (16,56 frente a 16,3; con 0–6 da 14,56). Cobertura neta 2025 aún no publicada. La serie salta entre 2017 (64,24) y 2018 (96,96), probablemente por el cambio de proyección del DANE [S] | A |
| 6 | PDM.SP | Servicios públicos domiciliarios (acueducto, alcantarillado, energía, aseo; cobertura total) | No existía | D9: nuevo; cobertura total, no solo urbana; sin internet; precisar el índice, las fichas y la ponderación | **Nuevo** | Probable: TerriData separa REC/SUI (anual) y censo, y urbano/rural/total [V-reu]. La serie de Envigado no se reconsultó hoy [S] | Ap. 9.1.2, línea 1, nodo 4: acueducto LB **99,73**, alcantarillado **99,56**, energía **99,87** (2023), meta **99,9**, fuente: **Encuesta de Calidad de Vida (DAP)**; **aseo no tiene indicador de resultado** [V]. Superservicios, informe nacional de coberturas AAA 2024: https://superservicios.gov.co/sites/default/files/2025-12/Informe-Nacional-de-coberturas-AAA-vigencia-2024.pdf (existe, 7,6 MB; no leí la tabla municipal) [V]; vigencia 2023: https://superservicios.gov.co/sites/default/files/inline-files/Informe-Nacional-de-Coberturas-AAA-vigencia-2023.pdf [S, solo en el buscador]. UPME, ICEE municipal (energía): https://docs.upme.gov.co/SIMEC/Energia%20Electrica/Documents/Informes_cobertura/Boletin_ICEE_2024.pdf [S, solo en el buscador] | Descargar de TerriData las 4 coberturas totales, años y fuente (REC). C7: la meta del PDM viene de la ECV (encuesta de hogares) y TerriData del REC (suscriptores). Decidir la ponderación (P3) y qué hacer con el aseo sin meta. Coberturas ≈ 99–100 %: el indicador casi no varía | A |
| 7 | PDM.SALUD | Cobertura de salud (total) | Serie 2020–2025 (83,76 en 2024; 83,06 en 2025), una sola proyección DANE | D3, D5: se mantiene la cobertura total; la meta debe ser del mismo universo | **Se mantiene**, sin meta del PDM (solo serie y variación) | Sí: 2020–2025 [V-rep] | TerriData (MinSalud) y DANE [V-rep]. Ap. 9.1.2: el único indicador del PDM es «Cobertura del aseguramiento en salud del **Régimen Subsidiado**», LB 99,5 (2022), meta 99,5 [V]: otro universo, no sirve como meta (C7) | Decidir la referencia de lectura (P4): sin meta, con la variación anual o con un referente externo. Documentar que la meta del PDM no aplica | M |
| 8 | PDM.ICPD | Índice agregado del PDM | ICPD-6 con promedio parcial, referencial | Recalcular con la nueva selección | **Se ajusta** | Depende de los componentes. Con espacio verde casi sin serie y servicios públicos sin serie confirmada, y la regla de no calcular con entradas faltantes, el índice saldría **solo para 1 o 2 años** [QA] | Los 6 componentes | Nueva composición: SALUD, CM, CA, ILE, SP, SEXTO. Redefinir IBS/ITG o usar un promedio simple (P5). Regla de la Guía: no calcular con entradas faltantes | M |
| 12 | PDM.CP | Procesos participativos | N/D; respaldo índice situacional 3,23 (2022) | D11: sale del PDM porque la participación tiene su propia dimensión | **Eliminado** | — | Ap. 9.1.2: «Implementación de los procesos de participación ciudadana» LB 75 (2023), meta 100, fuente DAP [V]: es un autorreporte (C6) | Dejar la nota de por qué sale | B |
| 12 | PDM.CD | Servicios deportivos | N/D; respaldo de disciplinas 38/40 | D10, D11: sin tasa oficial del Índer; candidato al sexto indicador | **Se reemplaza** (por PDM.SEXTO) | No hay serie pública [V-rep] | Ap. 9.1.2: cobertura deportiva LB 5,3 (2023), meta 6; actividad física LB 8 (promedio 2020–2023), meta 9; fuente: bases de datos del Índer [V] | Nada, si el sexto indicador es de seguridad | B |
| 13 | PDM.POLPUB | Logros de políticas públicas | N/D; sin inventario oficial | Ninguna reunión lo trató | **Pendiente** | No (solo 2025, construible) [V-rep] | Acuerdos del Concejo [S] | Decidir si sigue (P6). Se propuso como «nueva dimensión», pero C2 dice que las dimensiones no cambian | B |

### 1.2 Participación ciudadana (PC)

| Paso | Código | Indicador | Estado anterior | Decisión de las reuniones | Veredicto propuesto | ¿Serie posible? | Fuente(s) pública(s) verificada(s) | Qué falta | Prior. |
|---|---|---|---|---|---|---|---|---|---|
| 1 | PC.VEED | Veedurías ciudadanas inscritas en la Personería | No existía | D13: número anual de veedurías inscritas en la Personería, con línea base | **Nuevo** | Sí: 2012–2026 por fecha de inscripción [V]. Inscritas por año en el listado: 2019: 2 · 2020: 4 · 2021: 11 · 2022: 4 · 2023: 6 · 2024: 1 · 2025: 5 · 2026: 1 | Personería de Envigado (datos.gov.co, «Veedurías Ciudadanas Personería Envigado»): https://www.datos.gov.co/resource/jju5-br3w.json, 45 filas, actualizado el 07-oct-2026 [V] | **Alerta:** el número consecutivo va de 118 a 293 (176 números) y solo hay 45 filas: el conteo por año probablemente subestima las inscripciones [S]. QA encontró que **no** es solo un listado de vigentes (hay al menos 3 vencidas: n.º 278, 218 y 283) y que la columna `n` solo trae impares del 1 al 89, señal de filas perdidas al publicar. Verificada la **fuente**, no la serie. Con 1 inscripción en 2024, una LB en ese año es inestable. Definir si se mide el flujo (inscritas en el año) o el stock (vigentes al 31-dic) (P7). El conjunto trae nombres y correos: usar solo conteos | A |
| 9 | PC.I1 | Incorporación de aportes | N/D; respaldo 50/51 (2024) y 61/61 (2025) con los datos de PP | D13: puede correlacionarse con la fidelidad de PP; diferenciarlos | **Pendiente**: D13 pide **diferenciarlo** de PP.FID, no reemplazarlo. Opciones: redefinirlo para que no use los datos de PP, dejarlo como variable o reemplazarlo por PC.VEED | No: no hay universo oficial de procesos [V-rep] | — | Confirmar con el docente (P8). El respaldo duplica PP.FID (doble conteo) | M |
| 9 | PC.I2 | Cobertura territorial | 100 (2024), 92,31 (2025) | D13: se mantiene | **Se mantiene** | Solo 2024–2025 [V-rep]. Antes de 2024: no se buscó [S] | Informes de gestión 2024 y 2025 [V-rep] | Intentar 2022–2023 con los informes de gestión anteriores (C5) | M |
| 9 | PC.I3 | Oferta formativa | 83,33 (2025); 2024 N/D | Sin decisión | **Se ajusta**: usar como principal el número de acciones culminadas (7 en 2024, 21 en 2025), porque la tasa no tiene denominador para 2024 | Conteo: 2024–2025 [V-rep] | Informes de gestión 2024 y 2025 [V-rep] | Confirmarlo con el docente. Buscar 2022–2023 | M |
| 9 | PC.I4 | Publicación de resultados | N/D; universo exploratorio 3 y 5 | C3: los indicadores sin información se cambian | **Pendiente** → probable **Eliminado** o solo variable | No [V-rep] | Portal de la Alcaldía [V-rep] | Decidirlo (P8) | B |
| 9 | PC.ISPC | Índice de participación | Opción redistribuida 88,32 (2025) | Recalcular con la nueva selección | **Se ajusta** | Depende: con I2, I3 (conteo) y VEED, 2024–2025 | Componentes | Pesos nuevos (P9). VEED es un conteo, no un porcentaje: hay que normalizarlo (p. ej., contra la línea base o un máximo) | M |
| — | PC.I5, PC.I6 | Sesiones y participación virtual | — | D1 | **Eliminado** (hecho) | — | — | — | — |

### 1.3 Control social (CS): definido en D12

| Paso | Código | Indicador | Estado anterior | Decisión de las reuniones | Veredicto propuesto | ¿Serie posible? | Fuente(s) | Qué falta | Prior. |
|---|---|---|---|---|---|---|---|---|---|
| 11 | CS.PCS | Promoción del control social | 100 (2024 y 2025), Verde | D12: definido | **Se mantiene, con alerta C6**: es un autorreporte del Plan Indicativo que da 100 %. Mostrar el número de estrategias (1 por año) y no usar el semáforo Verde como logro | 2024–2025 [V-rep] | Plan Indicativo de Productos (DAP) [V-rep] | Pregunta P10: ¿se acepta el autorreporte o se usa solo el conteo? | M |
| 11 | CS.DSI | Disponibilidad de información de seguimiento | 0 (2024 y 2025), Rojo; respaldo 28,57 | D12 | **Se mantiene** | 2024–2025 [V-rep] | Apéndice 9.1 e informes [V-rep] | Hay que actualizar la lista de 7 metas cuando cambie la selección del PDM (CS.DSI depende de esas metas) | M |
| 11 | CS.VRC | Veedurías en la rendición de cuentas | N/D; respaldo 66,67 y 20 | D12 | **Se mantiene** | 2024–2025 [V-rep] | Actas de rendición [V-rep] | Segunda lectura pendiente (`outputs/indicadores/segunda_lectura/muestra_vrc.csv`) | B |
| 11 | CS.RPO | Respuesta a observaciones | 94,12 (2024); 2025 N/D (respaldo 50) | D12 | **Se mantiene** | 2024–2025 [V-rep] | Cuadros de preguntas y respuestas [V-rep] | Segunda lectura (`muestra_rpo.csv`) | B |
| 11 | CS.ICRCS | Índice de control social | N/D; respaldos 65,20 y 42,50 | D12 | **Se mantiene** | 2024–2025 | Los 4 componentes | Nada nuevo | B |

### 1.4 Presupuesto participativo (PP): definido en D12; D4 por aclarar

| Paso | Código | Indicador | Estado anterior | Decisión de las reuniones | Veredicto propuesto | ¿Serie posible? | Fuente(s) | Qué falta | Prior. |
|---|---|---|---|---|---|---|---|---|---|
| 10 | PP.FID | Fidelidad a lo votado | 76,47 (2024; 39/51), 95,08 (2025) | D12: definido | **Se mantiene** | 2024–2025; 2026 cuando se tenga el valor votado [V-rep] | Listado EFPC: https://efpc.envigado.gov.co/wp-content/uploads/2025/07/Proyectos_PP_2024.pdf (responde 200 hoy) [V]. Hacienda y plan de acción [V-rep] | No encontré listados de 2022–2023 por URL directa (404) [V]; buscarlos en el sitio de la EFPC para cumplir C5 | M |
| 10 | PP.EPP | Ejecución presupuestal del PP | 98,69 (2024 aprox.), 98,83 (2025), 23,26 (2026-S1) | D4: «revisar lo de la ejecución de presupuesto»; D12: definido | **Se mantiene; D4 queda como pregunta** (P11) | 2024–2025 (+2026-S1) [V-rep]. 2023: CUIPO tiene periodos 2023–2026 (`outputs/limpieza_bpin_cuipo.md`), pero falta la lista de proyectos de PP de ese año [S] | Ejecución de Hacienda 2024: https://alcaldia-municipal-de-envigado-en-antioquia.micolombiadigital.gov.co/sites/alcaldia-municipal-de-envigado-en-antioquia/content/files/002660/132983_ejecucion-presupuestal-egresos-al-31122024.pdf [V-rep] | Temas posibles de D4: (a) excluir el rubro de déficit fiscal de 2024 (98,69 → 98,64); (b) proyecto 0553, costo del proceso (98,83 → 98,86); (c) la cifra OPP 2026 errada en el Excel de disponibilidad v1.1 ($11.555.041.247; la correcta es $5.702.528.148); (d) obligaciones frente a pagos; (e) extender la serie | A (por D4) |
| 10 | PP.resp | Respaldos: compromiso, ≥90 %, equidad entre zonas, terminados | Ver §5 | — | **Se mantienen como respaldo**. «Proyectos terminados» es un autorreporte (C6): no presentarlo como logro | 2024–2025 | Hacienda y plan de acción [V-rep] | Nada | B |

---

## 2. Verificación rápida de las fuentes nuevas o pendientes (15-nov-2026)

### 2.1 Veedurías (PC.VEED) [V]
- El conjunto **existe** y lo publica la Personería de Envigado en datos.gov.co: id `jju5-br3w`, «Veedurías Ciudadanas Personería Envigado», actualizado el 07-oct-2026. Tiene 45 filas con estas columnas: N., año de registro inicial, número de veeduría, fecha de inscripción, objeto, duración/fecha final, integrantes, correo, municipio y número de veedores.
- Inscripciones por año, según la fecha de inscripción: 2012: 1 · 2014: 2 · 2015: 2 · 2016: 2 · 2017: 1 · 2018: 3 · 2019: 2 · 2020: 4 · 2021: 11 · 2022: 4 · 2023: 6 · 2024: 1 · 2025: 5 · 2026: 1. Hay dos filas con año de registro 2026 y fecha de 2025.
- Limitación: el número consecutivo va de 118 a 293 y solo hay 45 registros. Lo más probable es que el listado muestre solo las veedurías vigentes [S]. Para medir el flujo real se puede usar el rango de consecutivos por año (por ejemplo, 2021: del 262 al 275), pero eso sigue siendo una aproximación [S]. Confirmarlo con el esquema de publicación de la Personería: datos.gov.co, id `rmdh-csmg` [V, existe].
- El conjunto trae datos personales (nombres y correos): solo deben usarse conteos agregados.

### 2.2 ILE: cobertura neta y primera infancia [V]
- MinEducación, «Estadísticas en educación preescolar, básica y media por municipio», id `nudc-7mev`: Envigado tiene `cobertura_neta` de 2011 a 2024. También trae la cobertura neta en transición: 64,78 (2018) … 63,47 (2024).
- El Apéndice 9.1.2 usa este mismo dato (LB 100,44 en 2022) con tendencia «Monitoreo» y meta igual a la LB. No hay una meta de mejora, así que el «avance hacia la meta» no aplica a este componente.
- Primera infancia: el PDM nombra el indicador como de **0 a 6 años** (LB 16,3 en 2023, meta 19). La hoja simplificada usa 0–5 (16,56 en 2023), que es la más cercana a la LB. Hay una contradicción interna en el repositorio: `outputs/indicadores/ESTADO.md` dice que ILE.A usa 0–6 por defecto y el CSV simplificado usa 0–5.

### 2.3 Sexto indicador del PDM: candidatos ordenados
Los cuatro candidatos tienen un indicador de resultado en el Apéndice 9.1.2 [V]. Los conteos son de datos.gov.co [V]. Las tasas indicativas dividen por la población DANE de la hoja de salud y **no son resultados**: las debe recalcular el Ejecutor.

| Orden | Candidato | Vínculo con el PDM (Ap. 9.1.2) | Serie pública verificada | Ventajas | Riesgos |
|---|---|---|---|---|---|
| 1 | **Tasa de homicidio por 100.000 hab.** | Línea 4, programa 1, R1: LB **3,7 (2023)**, meta **2,7**, decreciente, fuente Policía Nacional, ODS 16.1.1 | `m8fd-ahd9` (MinDefensa/Policía), Envigado: 10, 7, 3, 9, 9, 12, 8 homicidios (2019–2025); datos hasta 30-sep-2026. Tasa indicativa 2023: 9/256.042 = 3,52 (el PDM dice 3,7: otra población) | Fuente oficial, anual, poco subregistro y relevancia ciudadana alta | Con conteos pequeños la tasa salta mucho (±1 homicidio ≈ ±0,4 puntos). Se podría usar un promedio móvil de 3 años [S] |
| 2 | **Tasa de hurto a personas** | Línea 4, programa 1, R3: LB **632,5 (2023)**, meta **631,5**, fuente Policía Nacional | `4rxi-8m8d` (MinDefensa/Policía), Envigado: 1.601, 1.080, 1.560, 1.554, 1.643, 1.329, 1.327 (2019–2025). Tasa indicativa 2023: 641,7 | Conteos grandes y serie estable | Depende de que la gente denuncie. La meta casi no exige (baja 1 punto) |
| 3 | Deporte y actividad física (reemplazo directo de CD) | Línea 1, nodo 2: cobertura deportiva LB 5,3 (2023), meta 6; actividad física LB 8 (prom. 2020–2023), meta 9; población inscrita en programas deportivos LB 12,2 (2021, ECV), meta 13 | No hay una serie pública anual: depende de las bases del Índer o de la ECV [V-rep] | Continuidad con CD | Va contra C4, porque habría que pedir los datos |

Faltó evaluar (señalado por QA): **tasa de lesiones personales** (violencia interpersonal, mencionada en la reunión). Está en el Ap. 9.1.2 con LB 113,7 y meta 112,7. Revisar su serie en la Policía antes de decidir P12.

Descartados por ahora:
- **Déficit habitacional**: el Ap. 9.1.2 trae el déficit cuantitativo (LB 5,07 en 2023, meta 4,07) y el cualitativo (11,23), pero su fuente es la ECV, que no es anual [S]. El conjunto de datos.gov.co «Déficit habitacional, según municipio 2018» (`tt8i-w5qr`) es censal y de un solo año.
- **Pobreza**: el Ap. 9.1.2 trae «Personas con ingreso per cápita por debajo de la línea de pobreza» (LB 13,69 en 2023, fuente ECV). La prensa informa ECV de 2021 y 2024 (https://www.elcolombiano.com/antioquia/informe-calidad-de-vida-en-envigado-JI26595279) [S, solo en el buscador]: la serie no es anual.
- **Medición de Desempeño Municipal** (`nkjx-rsq7`): en datos.gov.co, para Envigado, solo trae de 2016 a 2020 [V].

### 2.4 Espacio público verde 2024–2025
- El Ap. 9.1.2 da una LB de 3,88 m²/hab (2023) y una meta 2027 de 3,88 (estable). Además trae un indicador distinto: «Espacio público efectivo», LB 3,33 (2023), meta 4 [V]. Tenemos tres cifras de verde que no coinciden: 3,88 (PDM, 2023), 4,66 (AMVA, 2022, tabla 39 del POT) y 5,75 (medición del POT).
- No encontré en la web una cifra de 2024 ni de 2025 para Envigado [V: búsqueda sin resultado]. Hay que revisar a mano los informes de gestión 2024 y 2025 y la evaluación del Plan Indicativo, que tiene la columna de resultado (URLs en §5).

### 2.5 Mortalidad vial 2025
- **Policía [V]**: 30 muertes en 2025, del conjunto `ha6j-pa2r` (actualizado el 01-sep-2026, con datos hasta 31-jul-2026). Tasa indicativa: 11,34 con 264.637 habitantes. **Si se confirma, 2025 queda por encima de la meta de 8.**
- Forensis 2025: no pude verificarlo, porque el sitio del INMLCF responde 403 a la consulta automática. La usuaria debe revisarlo a mano.
- Anomalía: la Policía da 43 muertes en 2023, una tasa de unos 16,8, mientras que la LB del PDM para 2023 es 10,2 (fuente: Secretaría de Movilidad). Hay que revisar si el conjunto cuenta hechos ocurridos en Envigado o víctimas residentes en Envigado [S].

### 2.6 Servicios públicos
- Las metas del PDM vienen de la ECV (DAP) y no incluyen el aseo [V].
- TerriData separa series censales y de registros (REC), y urbano/rural/total [V-reu]. No pude leerla con la herramienta web porque la aplicación se carga con JavaScript. Para descargarla: TerriData → Descargas o Fichas → Envigado (05266) → dimensión Vivienda y servicios públicos.
- Como contraste anual, sirven los informes nacionales de coberturas AAA de Superservicios (2023 y 2024) y el ICEE de la UPME para la energía.

---

## 3. Orden de trabajo (pasos)

El orden va de lo que el docente pidió de forma explícita (informe, §6: sexto indicador, variable de PC, espacio verde, ajustes de variables) a lo que se cierra más fácil.

1. **PC.VEED**: la fuente existe, pero la serie no es confiable tal como está publicada (faltan filas). Contrastar el listado con la Personería (página web, informes de gestión o rendición de cuentas) o usar el rango de consecutivos por año como aproximación. Construir la ficha y resolver P7 (flujo o stock). *Pedido; no es tan fácil como parecía.*
2. **PDM.SEXTO**: presentar a la usuaria y al docente la tabla 2.3. El Ejecutor descarga los conteos de homicidio y hurto 2019–2025 y calcula las tasas con la población DANE única. *Pedido.*
3. **PDM.CA**: buscar 2024–2025 en los informes de gestión y en la evaluación del Plan Indicativo. Redactar la ficha con el concepto «verde urbano», la LB 4,66 y el referente 9, y explicar la diferencia con 3,88 y 5,75. *Pedido.*
4. **PDM.CM**: elegir **una sola fuente** para toda la serie (no mezclar Forensis 2024 con Policía 2025; miden distinto según `variables_por_dimension.csv`, PDM.CM.O). Opción natural: Policía 2019–2025, con Forensis como respaldo. Unificar la población. La meta 8 y la LB 10,2 son de Movilidad: no comparar la tasa de la Policía (11,34 en 2025) con esa meta sin anotarlo (C7). Revisar el salto de 2023 (43 muertes).
5. **PDM.ILE**: descargar la cobertura neta 2018–2024 (en `nudc-7mev` el código de Envigado cambia de 5266 a 05266: filtrar por ambos) y proponer la fórmula (P1) y el rango de edad (P2). Documentar dos limitaciones ya conocidas de primera infancia: el numerador incluye **mujeres gestantes** y el denominador solo niños; la fuente cambia en 2024 de TerriData al informe de gestión (`scripts/sistema_simplificado/calcular.py:409-412`). Con la misma proyección DANE de la hoja, 0–6 da **14,02** en 2023 (no 14,56).
6. **PDM.SP**: descargar de TerriData las 4 coberturas totales con su fuente y sus años, y armar las fichas. Proponer la ponderación (P3).
7. **PDM.SALUD**: documentar en la ficha que la meta del PDM es del régimen subsidiado y no aplica. Elegir la referencia de lectura (P4). *Fácil.*
8. **PDM.ICPD**: recalcular con los 6 componentes cuando cierren los pasos 2 a 7 (P5).
9. **PC.I1, I2, I3, I4 e ISPC**: cerrar los veredictos y los pesos (P8, P9). Buscar I2 e I3 de 2022–2023.
10. **PP (D4)**: llevar al docente la lista de posibles temas de D4 (P11). Buscar los listados de PP de 2022–2023.
11. **CS**: solo la alerta C6 de PCS (P10) y las segundas lecturas pendientes. Se actualiza la lista de metas de DSI con la selección final del PDM.
12. **CP y CD**: dejar documentado por qué salen.
13. **POLPUB**: decidir si sigue o se archiva (P6).
14. **Base de datos**: organizarla solo después de cerrar la selección (compromiso del 12-nov).

---

## 4. Preguntas para la usuaria o el docente

- **P1. Fórmula del ILE.** ¿Promedio simple de los dos componentes? ¿Cómo se trata la cobertura neta mayor de 100 (100,44 en 2022)? ¿Tope en 100? ¿Avance hacia la meta solo para primera infancia, ya que la cobertura neta es de «Monitoreo»?
- **P2. Edad de primera infancia.** El PDM dice 0–6, pero 0–5 reproduce mejor la LB 16,3. ¿Cuál se usa? ¿Se cambia la LB?
- **P3. Servicios públicos.** ¿Promedio simple o ponderado de las 4 coberturas? ¿El aseo entra sin meta del PDM? ¿Qué fuente se usa: el REC de TerriData (anual) o la ECV, que es la de la meta (C7)?
- **P4. Salud.** Sin meta del mismo universo en el PDM, ¿se lee solo la serie y su variación, o se fija un referente (p. ej., 95 %)?
- **P5. ICPD.** ¿Se mantienen los subíndices IBS/ITG con los nuevos componentes o se usa un promedio simple de 6?
- **P6. Logros de políticas públicas.** ¿Sigue? Choca con C2 (no hay dimensiones nuevas).
- **P7. Veedurías.** ¿Inscritas en el año (flujo) o vigentes al 31-dic (stock)? ¿Se acepta que el listado de datos.gov.co parece mostrar solo las vigentes? ¿Encaja mejor en PC o en CS (CS.VRC ya usa veedurías)?
- **P8. PC.I1 y PC.I4.** ¿Se eliminan, se reemplazan por PC.VEED o quedan solo como variables?
- **P9. Pesos del ISPC** con I2, I3 y VEED, y forma de normalizar el conteo de VEED.
- **P10. PCS.** ¿Se acepta el autorreporte que da 100 % (C6) o se usa solo el conteo de estrategias?
- **P11. D4 (ejecución del PP).** ¿A qué se refería la nota: el rubro de déficit fiscal, el proyecto 0553, obligaciones frente a pagos, la corrección de la OPP 2026 o la serie?
- **P12. Sexto indicador.** ¿Homicidio o hurto a personas? ¿Tasa anual o promedio móvil de 3 años?
- **P13. Espacio verde.** ¿Se acepta como referencia de evaluación 4,66 (AMVA 2022) → 9 en lugar de la meta del PDM (3,88, estable)?

---

## 5. Datos ya trabajados (resumen)

| Qué | Valor | Fuente | Archivo |
|---|---|---|---|
| PP: fidelidad | 76,47 (2024, 39/51); 95,08 (2025, 58/61); monetaria 103,55 y 100,74 | Listado EFPC + ejecución de Hacienda/plan de acción | `outputs/sistema_simplificado/indicadores_simplificados.csv` |
| PP: EPP | 98,69 (2024 aprox., 50 de 51 proyectos); 98,83 (2025); 23,26 (2026-S1) | Hacienda (ejecución de egresos 2024; plan de acción IV trim. 2025; II trim. 2026) | ídem; `outputs/disponibilidad/epp_pp_2025.md`; scripts `scripts/indicadores/pp_epp.py` y `pp_epp_2024.py` |
| PP: respaldos | Compromiso 99,56 (2025); ≥90 %: 48/50 y 57/61; brecha entre zonas 17,42 y 2,51; terminados 49/61 (autorreporte) | Hacienda y plan de acción | `indicadores_simplificados.csv` |
| PC | I2 100 y 92,31; I3 83,33 (2025); acciones culminadas 7 y 21; respaldo de I1 50/51 y 61/61; ISPC redistribuido 88,32 (2025) | Informes de gestión 2024 y 2025 | ídem; `scripts/indicadores/pc_i2.py` y `pc_i3.py` |
| CS | PCS 100/100; DSI 0/0 (respaldo 28,57); VRC N/D (respaldos 66,67 y 20); RPO 94,12 (2024); ICRCS respaldos 65,20 y 42,50 | Plan Indicativo, actas y cuadros de rendición | ídem; `outputs/indicadores/clasificacion_vrc.csv` y `clasificacion_rpo.csv` |
| Salud | 82,29 (2020) … 83,76 (2024), 83,06 (2025); MinSalud 87,65 (dic-2025) y 82,59 (ago-2026) | TerriData (MinSalud) + DANE 2018–2042 | ídem; `scripts/sistema_simplificado/calcular.py` |
| Mortalidad vial | 7,25 (2024, Forensis p. 440); Policía 8,05 (población de 248.304) | Forensis 2024; Policía | ídem; `scripts/indicadores/pdm_cm.py` |
| Espacio verde | 5,75 (POT) vs. meta 3,88; EPV total del POT 69,41 | Diagnóstico del POT | ídem; `scripts/indicadores/pdm_ca.py` |
| Primera infancia | 13,36 (2019) … 17,53 (2024), 17,12 (2025), población 0–5 | TerriData/informes de gestión + DANE por edad simple | ídem; `outputs/sistema_simplificado/trabajo/poblacion_envigado_edad_pped.csv` |
| Saber 11 (antiguo respaldo) | Oficiales urbanas 267,16 (2019), 263,6 (2020), 269,69 (2022); sin 2021 ni 2023–2025 | ICFES en datos.gov.co (`kgxf-xxbe`, `rnvb-vnyh`) | ídem; `scripts/sistema_simplificado/descargas_nuevas.py` |
| CD / CP | Disciplinas 38/40 (2025); índice situacional 3,23 (2022) | Informe de gestión (Índer); diagnóstico del PDM | ídem |

URLs de los documentos de la Alcaldía usados [V-rep]:
- Informe de gestión 2024: https://alcaldia-municipal-de-envigado-en-antioquia.micolombiadigital.gov.co/sites/alcaldia-municipal-de-envigado-en-antioquia/content/files/002640/131997_24informegestionpordependenciasalcaldiamunicipalenvigado2024.pdf
- Informe de gestión 2025: https://alcaldia-municipal-de-envigado-en-antioquia.micolombiadigital.gov.co/sites/alcaldia-municipal-de-envigado-en-antioquia/content/files/003058/152890_2025informegestiondependencias.pdf
- PDM 2024–2027: https://alcaldia-municipal-de-envigado-en-antioquia.micolombiadigital.gov.co/sites/alcaldia-municipal-de-envigado-en-antioquia/content/files/002562/128075_plandesarrollo20242027.pdf
- Evaluación del Plan Indicativo de Productos (II trim. 2026): https://alcaldia-municipal-de-envigado-en-antioquia.micolombiadigital.gov.co/sites/alcaldia-municipal-de-envigado-en-antioquia/content/files/003611/180505_evaluaciondefinitivaplanindicativoproductos2026segundotrimestre20260730.xlsx

Nota: los cálculos base (commit d723c4c, `outputs/indicadores/`) no cambian. Lo nuevo se marca como provisional.

---

## 6. Propuesta de actualización del grafo (no aplicada)

Propuesta para que el coordinador la registre en `.claude/memoria/grafo_conocimiento.md`:

**Nodos de fuente nuevos**
- FUE.AP912 (Apéndice 9.1.2 del PDM: indicadores de resultado con LB y meta)
- FUE.ECV (Encuesta de Calidad de Vida de Envigado, DAP; no anual)
- FUE.MEN_DGOV (`nudc-7mev`)
- FUE.POL_HOM (`m8fd-ahd9`), FUE.POL_HURTO (`4rxi-8m8d`), FUE.POL_TRANS (`ha6j-pa2r`)
- FUE.PERS_VEED (`jju5-br3w`)
- FUE.UPME_ICEE, FUE.SSPD_AAA

**Cambios de estado**
- PDM.SEXTO: candidatos ordenados → homicidio, hurto a personas, deporte
- PDM.CM: serie de la Policía 2019–2025; 2025 = 30 muertes
- PDM.ILE: cobertura neta 2011–2024 verificada
- PC.VEED: fuente verificada; serie 2012–2026 (alerta: solo vigentes)
- PC.I1 → Se reemplaza (propuesto); PC.I4 → Pendiente/probable Eliminado
- PC.I3 → Se ajusta (conteo)
- PDM.POLPUB → Pendiente (conflicto con C2)

**Aristas nuevas**
```
PDM.SALUD --meta_no_aplica--> AP912 «Cobertura aseguramiento régimen subsidiado» (LB 99,5 2022)   [C7]
PDM.CM    --meta--> AP912 (LB 10,2 2023, meta 8; fuente Sec. Movilidad) ; --usa--> FUE.POL_TRANS
PDM.CA    --meta_pdm--> AP912 (LB 3,88 2023, meta 3,88 estable) ; --distinto_de--> «Espacio público efectivo» (LB 3,33, meta 4)
PDM.ILE   --componente--> cobertura neta (FUE.MEN_DGOV; AP912 LB 100,44 2022, Monitoreo)
PDM.ILE   --componente--> primera infancia (AP912 «0 a 6 años» LB 16,3 2023, meta 19) ; --conflicto_edad--> hoja 0–5
PDM.SP    --meta--> AP912 acueducto 99,73 / alcantarillado 99,56 / energía 99,87 (2023, ECV) → 99,9 ; aseo --sin_meta_pdm-->
PDM.SP    --usa--> FUE.TERRIDATA (REC), FUE.SSPD_AAA, FUE.UPME_ICEE ; --conflicto_universo--> FUE.ECV
PDM.SEXTO --candidato_1--> homicidio (AP912 LB 3,7 2023, meta 2,7; FUE.POL_HOM)
PDM.SEXTO --candidato_2--> hurto a personas (AP912 LB 632,5 2023, meta 631,5; FUE.POL_HURTO)
PDM.SEXTO --candidato_3--> deporte/actividad física (AP912; sin serie pública)
PC.VEED   --usa--> FUE.PERS_VEED ; --posible_solape--> CS.VRC
PP.EPP    --pregunta_abierta--> D4 (déficit fiscal 2024, proyecto 0553, OPP 2026)
CS.PCS    --alerta--> C6 (autorreporte que da 100 %)
```

**Decisiones a registrar en `decisiones.md` cuando se respondan:** P1–P13 de la §4.
