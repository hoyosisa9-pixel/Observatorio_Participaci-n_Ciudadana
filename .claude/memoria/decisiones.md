# Registro de decisiones (reuniones con el docente Francisco Correa)

Fuente de cada entrada entre corchetes. Las notas manuscritas de la usuaria se transcribieron el 15-nov-2026.
Todo resultado sigue siendo **Provisional**.

## Criterios generales (aplican a todos los indicadores)
- C1. Simplificar la forma de mostrar la información. [Notas R1]
- C2. Los componentes (dimensiones) son los mismos: PP, PC, CS, PDM. [Notas R1]
- C3. Los indicadores sin información se cambian por otros similares que representen la dimensión (no tienen que ser idénticos). [Notas R1]
- C4. Trabajar con indicadores cuyos datos se puedan encontrar sin pedirlos a Planeación. [Notas R1, nota #2]
- C5. Todo debe poder verse como **evolución en el tiempo** (serie anual). [Notas R1]
- C6. No usar datos de autorreporte donde la Alcaldía se da 100 %. [Notas R1]
- C7. Comparar solo con metas y líneas base del **mismo universo** (población, concepto, periodo). [Informe 12-nov, §3 y §7]
- C8. Plantear los indicadores del PDM en función de las metas del POT/PDM cuando existan. [Notas R2]

## Reunión 1 (≈10-nov-2026; ver también outputs/sistema_simplificado/trabajo/propuesta_aprobada.md)
- D1. Eliminar PC.I5 (cumplimiento de sesiones) y PC.I6 (participación virtual). **Hecho.** [Notas R1, punto 1]
- D2. ILE: eliminar el índice de logro educativo o transformarlo: cobertura de niños de 0–6 años. [Notas R1, punto 2]
- D3. Aseguramiento al régimen subsidiado → **cobertura en salud del municipio**. **Hecho (serie construida).** [Notas R1]
- D4. Pendiente: revisar lo que se dijo de la dimensión PP, algo sobre **ejecución de presupuesto**. [Notas R1, final]

## Reunión 2 (12-nov-2026; informe: outputs/guia/Informe tecnico reunion OCE 12 noviembre 2026.docx)
- D5. Salud: se mantiene la cobertura total; la meta tiene que ser del mismo universo. Revisar la tabla de afiliación subsidiada 2022 del PDM. [04:55–08:48]
- D6. Mortalidad vial: se mantiene (Forensis/Policía); faltan los años posteriores a 2024. [04:55–08:48]
- D7. Espacio público verde: se mantiene. Línea base AMVA **4,66 m²/hab (2022)**, medición 5,75, referente **9 m²/hab** (tabla 39 del diagnóstico del POT). Verificar los cortes 2024–2025 en el PDM y el POT. No confundir verde, efectivo y total. [08:48–23:00; Notas R2]
- D8. ILE: reemplazar Saber 11 por **cobertura neta total en educación** (TerriData) + **atención integral a primera infancia**; se conserva el nombre. Fórmula por definir. Conciliar 0–5 (hoja) con 0–6 (reunión/documento). [29:35–33:00; Notas R2]
- D9. **Servicios públicos domiciliarios** (nuevo): acueducto, alcantarillado, energía y aseo; cobertura **total** (no solo urbana). Internet queda fuera. Precisar el índice, las fichas y la posible ponderación. [33:00–44:20]
- D10. **Sexto indicador del PDM**: pendiente. Candidatos: deporte/recreación/juventudes (reemplazo de CD), seguridad y convivencia (hurto, homicidio, violencia interpersonal), pobreza, vivienda/déficit habitacional. Debe tener vínculo con el PDM y una serie de 4–5 años. [44:30–58:59]
- D11. CP (procesos participativos) sale del PDM porque la participación ya tiene su propia dimensión; CD (servicios deportivos) sin tasa oficial del Índer. [23:20–25:45]
- D12. Se dan por definidos el indicador de PP y el componente de control social. [59:30–61:24]
- D13. Participación ciudadana: falta precisar una variable. Notas: **número anual de veedurías ciudadanas inscritas en la Personería de Envigado** (datos.gov.co), con línea base; I1 (incorporación de aportes) puede correlacionarse con PP: diferenciarlos; mantener la **cobertura territorial de la participación**. [Notas, hoja 3; informe §6]
- D14. La encuesta de percepción fue revisada y remitida; su aplicación sigue sin confirmar. [59:30–61:24]

## Respuestas de la usuaria (15-nov-2026, revisión de la dimensión PDM)
- D15. **ILE** = promedio simple de cobertura neta y cobertura de primera infancia, según la fórmula de la matriz (`outputs/disponibilidad/trabajo/matriz.csv`, PDM.ILE): LA = 100·A/MA y LN = 100·N/MN, cada una con tope 100, e ILE = (LA + LN)/2. Fuente principal: **TerriData**.
- D16. Primera infancia con **0–5 años**, porque reproduce mejor la LB del PDM (16,3). Hay que anotar que el PDM dice 0–6.
- D17. **Servicios públicos reemplaza a CD (cobertura deportiva)**: acueducto, alcantarillado, energía y aseo, desde TerriData (accesibilidad). Ponderación: la recomienda Claude (simple, como el resto de la matriz).
- D18. **Salud**: la cobertura total reemplaza el aseguramiento subsidiado; serie de TerriData con su variación anual; sin meta del PDM.
- D19. **Sexta medida**: la usuaria pide recomendación entre homicidio, hurto a personas y lesiones personales. **NO reemplaza a CP** (la usuaria lo corrigió el 15-nov; fue una inferencia errónea de Claude). El estado de CP queda por confirmar con la usuaria.
- D26. **PDM.CP (procesos participativos) → número anual de veedurías ciudadanas inscritas en la Personería** (datos.gov.co, jju5-br3w). La usuaria lo confirma el 15-nov: la nota «cambiar por número de veedurías ciudadanas» era para CP (en la matriz, la categoría de CP es «PARTICIPACIÓN CIUDADANA»). Ya no es un candidato de la dimensión PC. Alertas: el listado parece incompleto (consecutivos 118–293 con 45 filas) y CS.VRC también usa veedurías (es otra medida, pero hay que anotarlo).
- D27. **ILE**: la «cobertura en educación» de TerriData Comparaciones = cobertura neta total (ya usada: 2015–2024) + primera infancia 0–5. Confirmado por la usuaria.
- Regla de la usuaria: donde no hay datos, buscar una variable parecida que sí tenga datos.
- D28. **PC.I1 → opción A** (usuaria, 15-nov): la fidelidad a lo votado se queda SOLO en PP (la foto era del índice de PP: fidelidad en lugar de CEP/CMA). I1 se reemplaza por otro indicador, en principio la participación en la votación de PP (verificar que tenga serie y valga la pena). Así se cumple el «diferenciar» del profe.
- Hallazgo PC.I1 (15-nov): existe un indicador oficial de la Alcaldía, el **IPCM** (Índice de Participación Ciudadana Municipal = votos de PP / población). 2017: 8.221 votos, IPCM 3,62 % (informe DAP, 11-oct-2017, p. 7 y 12). 2019: sin votación. 2025: 10.662 votantes (informe de gestión 2025). 2018 y 2020–2024: no publicados. El sitio envigado.gov.co exige `curl -k` (cadena de certificado incompleta); la usuaria lo autorizó el 15-nov solo para ese sitio.
- D11 (corrección): que CP «sale del PDM» fue una lectura de Claude del informe, no una decisión confirmada.
- Fuente de datos que indica la usuaria: **TerriData, sección «Comparaciones»**. Tiene series completas (p. ej., cobertura de acueducto REC 2008–2024 en el archivo masivo; la web muestra 2025; y homicidios).
- Prioridad de la usuaria: el entregable es la **disponibilidad de datos** para construir los índices y lo que cambió. No hace falta calcular índices.
- Instrucciones de la «reunión fructífera» (chat anterior): por componente, qué indicadores se pueden crear y qué variables se podrían solicitar en el futuro (Alcaldía/Planeación); cada indicador con uno de más como reemplazo o proxy.
- D20. **Espacio verde**: avance de 4,66 hacia 9 m² por habitante.
- D21. **Índice del PDM**: no se decide todavía. Al profe se le entregan las variables disponibles para construir los índices.
- D22. **Logros de políticas públicas**: sigue, como dimensión para el observatorio futuro, con variables por solicitar.
- D23. **Control social**: revisar bien. El autorreporte al 100 % no sirve (PCS y otros).
- D24. **PP**: el profe propuso (a) **ejecución según los proyectos que sí se hicieron** y (b) **fidelidad con los proyectos votados**. Probablemente es lo que la nota D4 llamaba «algo sobre ejecución». Medidas ya planteadas antes: fidelidad, tasa de compromiso, proyectos con ejecución ≥ 90 %, equidad entre zonas, avance físico (autorreporte), ritmo de ejecución y rezago.
- D25. **Entregable**: un Word corto para el profe, por dimensión e indicador (qué se mantiene, qué se modificó, por qué se reemplazó y con qué variables contamos). Nada de informes largos.
- Pendiente: la usuaria traerá los reemplazos que había acordado en el repositorio anterior (el Codespace que se le acabó).

## Compromisos
- Isabella: priorizar ajustes; buscar 1–2 alternativas para el sexto indicador del PDM y la variable de PC; verificar espacio verde 2024–2025; luego organizar la base de datos.
- Francisco: actualizar el documento del Producto 2 (matrices por dimensión y definiciones) cuando reciba los insumos.
- D29. **PC.I4 se conserva** (usuaria, 15-nov): sin inventario oficial, el observatorio fija una **lista de procesos participativos recurrentes** y verifica cada año en el portal si se publicaron sus resultados (no es autorreporte). El Índice de Participación del DNP (ICM) se descarta como reemplazo porque mide participación electoral; puede quedar solo como contexto.
- PC cerrada (15-nov): I2 respaldo = zonas con proyectos de PP votados 2020–2025 (12/13 cada año; zona 1 nunca aparece; listados EFPC). I3 conteo: 2023 = 17 (informe de Control Interno vigencia 2023, proyecto 0541; dato del DAP sin evidencia en G+), 2024 = 7, 2025 = 20. I4: la rendición de PP 2023 sí existe (112893). Registro de cambios para el profe: outputs/entrega_profe/registro_cambios.md (actualizarlo en cada dimensión).
- Zona 1 de Envigado = zona industrial (dicho por el profe a la usuaria): explica que no tenga proyectos de PP. Duda abierta: denominador de I2, 13 o 12 zonas.
- CS (15-nov): propuesta PCS → oportunidad de respuesta a PQRSDF (informes trimestrales 2020–2026, BPMS; 4T-2023 = 3.544/4.149 = 85 %). DSI: actualizar la lista de metas. VRC extensible a 2021–2023 y RPO a 2021 y 2023. Formato pedido: cuadros sencillos para el profe + documento detallado para la usuaria.
- Espacio verde, serie histórica (Analizador, 15-nov; outputs/revision_indicadores/PDM/espacio_verde_serie.md): primer dato propio de Envigado = 2013 (4,16, ICAU de Minambiente). Cortes con la metodología del AMVA: 4,54 (2019, AMVA-UNAL), 4,66 (2022, AMVA-Humboldt) y 5,75 (2024, POT tabla 39: 1.382.762 m² / 240.655 hab. urbanos). Excluir: 5,00 (2020 y 2023, autorreporte) y 3,88 (PDM, método sin declarar). Ningún documento concilia 3,88 con 4,66. No hay serie anual. Pendiente: QA de esta búsqueda. Corregir los escenarios de outputs/indicadores/valores_variables.csv que usan 69,41 (verde total) frente a una meta urbana.
- D30. **Espacio verde = URBANO** (usuaria, 15-nov): cortes 2019 (4,54), 2022 (4,66) y 2024 (5,75); el total 2024 (69,41) solo como contexto. QA (qa_espacio_verde.md): cifras confirmadas, pero los tres cortes NO son comparables (productores, poblaciones y delimitación distintos; el 4,54 no se reproduce con 92,06 ha / 232.903 hab.). Presentarlos como cortes, no como tendencia. Índice del PDM: años con todos los componentes = 2019, 2022 y 2024.
- D31. **Espacio verde: línea base 2019 = 4,54 (AMVA-UNAL)** (usuaria, 15-nov). Reemplaza la LB 4,66 (2022) de D7/D20. Avance = (valor − 4,54)/(9 − 4,54): 2022 = 2,7 %; 2024 = 27,1 %. No se espera una serie continua. Prioridad: que la fuente de cada corte quede clara (documento, tabla, página, URL) para la trazabilidad. Recomendación para el informe: pedir a Planeación la serie completa desde 2019 con la metodología del AMVA; que Planeación la mida cada año en el seguimiento del POT (expediente municipal) y que la Secretaría de Medio Ambiente la reporte en el PDM. Cuadro de fuentes en `outputs/entrega_profe/1_PDM.md` y en los dos Word.
- D32. **PC.I1 (IPCM): opción A** (usuaria, 15-nov). Línea base **2017 = 3,62 %** (oficial, informe DAP 11-oct-2017, p. 7 y 12); referente **7 %** = tope de la escala oficial de tramos del IPCM (p. 6), porque el PDM 2024-2027 no tiene meta (solo el producto «ejercicios de PP realizados», apéndice 9.1.1, fila 539). Avance = (IPCM − 3,62)/(7 − 3,62). 2025: 10.662 votantes (informe de gestión 2025, p. 92 del PDF, **verificado**) / 264.637 (DANE) = 4,03 % → 12,1 %. Ojo: 2017 cuenta votos y 2025 votantes. Solicitar a Planeación 2018–2024 por zona, la población de referencia y la unidad.
- D33 (**decidido por la usuaria**, 15-nov: «lo dejamos así», con la nota de método y que la información mejore año a año; se le informa al profe en la duda 14). **CS.VRC con brecha de información** (pedido de la usuaria, 15-nov): mínimo = documentada/eventos (entra al índice), máximo = (documentada + desconocida)/eventos, brecha = desconocida/eventos. Universo principal: Alcalde + sectoriales con acta o memoria. 2024: 3 eventos → 66,7–100 %, brecha 33,3 %; 2025: 5 → 20–80 %, brecha 60 %. Sensibilidad (todos los eventos): 2024 50–100 (50); 2025 28,6–85,7 (57,1). Cambia la regla N/D de la matriz. Script `scripts/revision_indicadores/vrc_brecha.py` → `outputs/revision_indicadores/CS/vrc_brecha.csv`. Para seguir midiéndolo: lista fija de eventos, listas de asistencia con casilla de veeduría (Alcaldía), listado de veedurías (Personería), registro propio y observación directa de la audiencia del Alcalde.

## Respuestas del profe a las dudas (la usuaria las trae el 18-nov-2026, conversadas con el profe)
- D34 (duda 1). **ICPD con 7 componentes**: 6 con veedurías (en lugar de procesos participativos) + un **séptimo de seguridad = homicidio**. Meta PDM 2,7 (LB 3,7 de 2023), fuente Policía Nacional = mismo universo que el dato (datos.gov.co m8fd-ahd9). Abierto: cómo agregar 7 (los subíndices de la matriz eran 3 + 3).
- D35 (duda 2). **Salud: referente = cobertura universal (100 %)**.
- D36 (duda 3). Veedurías: la usuaria dice que la respuesta está en lo conversado. Lo único registrado es la nota de la reunión 2 (D13): «número anual de veedurías … **con línea base**». Referente = línea base; el año no se fijó (propuesta: 2023, año de LB del PDM). Por confirmar el año.
- D37 (duda 4). Espacio verde = D31 (LB 2019 = 4,54; referente 9; cortes 2019, 2022, 2024; pedir la serie a Planeación).
- D38 (duda 4b). Lo conversado (informe del 12-nov y C7): **no comparar con metas de otro universo**. Servicios públicos (meta de la ECV) y mortalidad vial (meta de Movilidad) no usan esas metas. Propuesta coherente con D35: servicios públicos → referente 100 % (cobertura universal). Mortalidad vial: referente por confirmar.
- D39 (duda 5). Servicios públicos = **promedio simple** de las 4 coberturas.
- D40 (dudas 6 y 7). **Sí**: I1 = IPCM (D32) e I4 = lista fija de 3 procesos (D29).
- D41 (duda 7b). **I2 con 12 zonas** (sin la zona 1, industrial): dejar 13 sabiendo que en una no hay actividad no se justifica y sesga el indicador.
- D42 (duda 8). **ISPC mantiene los pesos 30/25/20/25.**
- D43 (duda 9). Lo conversado (15-nov): CEP → **proyectos con ejecución ≥ 90 %** (Hacienda, verificable) como principal; «terminados» (autorreporte) solo como respaldo con advertencia.
- D44 (duda 10). **IREPP mantiene 20/40/40.** (duda 11) PP se calcula con 2024–2025 y se reporta la **brecha de información** (2020–2023 sin ejecución por proyecto).
- D45 (duda 12). **Sí**: PCS → oportunidad de respuesta a las PQRSDF. (duda 13) **Sí**: DSI con la lista de metas actualizada, **con brecha de información**. Regla de la usuaria: **siempre** reportar la brecha de información. (duda 14) VRC = D33 (rango + brecha).
- D46 (usuaria, 18-nov). **Veedurías: LB = promedio 2019–2023 = 5,4** (2, 4, 11, 4, 6). Puntaje = min(100, 100·valor/5,4).
- D47 (usuaria, 18-nov). **Mortalidad vial: referente del mismo universo** = la reducción del PDM (8/10,2 = −21,6 %) aplicada a la LB Policía 2019 (9,35) → **7,33**. CM = min(100, 100·7,33/O). La meta 8 de Movilidad no se usa.
- D48 (usuaria, 18-nov). **ICPD-7 = promedio simple de los 7 componentes** (cada uno 1/7), para conservar la regla de igual peso por componente.
- D49 (usuaria, 18-nov). **I3**: tasa solo en 2025 (denominador reconstruido) + respaldo conteo 17, 7, 20 (criterio unificado: solo ciudadanía). **Corrección**: 2024 = 7 formaciones para la ciudadanía (plan indicativo, consecutivo 521: 95 certificados); la nota «incluye servidores» era errónea (funcionarios formados 2024 = 0). Solicitar la programación anual a la Subdirección de Participación Ciudadana.
- D50 (usuaria, 18-nov). **Servicios públicos: opción A.** El indicador usa las 4 coberturas desde 2019, como se acordó con el profe. Acueducto, alcantarillado y aseo se muestran aparte con su serie de 2008 a 2024, solo como contexto. Regla general: cada indicador lleva todos los años que tenga; el índice solo los años en que coinciden todos los componentes.
- D51 (usuaria, 18-nov). **PP: excluir el rubro de déficit fiscal 2.3.7.06** (pagos de vigencias anteriores) de los proyectos de PP. 2024: $627.204.058 en 0346, 0347, 0272, 0348, 0746, 0714 y 0349. La cifra con el rubro incluido va en nota. Según QA: FID 46/51 = 90,20. EPP 2024: 98,64 con 50 proyectos; queda en 99,13 al aplicar D52, porque los cruces dudosos salen de EPP.
- D52 (usuaria, 18-nov). **PP: los cruces dudosos cuentan como «no verificados»**: en 2024, 0709, 0758, 0757 y 0733 (la apropiación inicial ya difiere del valor votado). Se tratan igual que EFPC2024-Z9-4: cuentan en el denominador de EJ90 y FID, no suman al numerador y quedan fuera de EPP. Regla general: la llave de cruce es valor votado = apropiación inicial de Hacienda; el nombre solo confirma.
- D53 (usuaria, 18-nov). **PP se extiende a 2020–2023 ya**, antes de pasar al PDM. Cruce por apropiación inicial (QA: 2020 = 53/68 montos iguales; prueba 2023 = 29/31).
- Hallazgo QA (18-nov): en `pp_proyectos.csv` de 2025, los valores votados de 0673 y 0786 (zona 12) estaban **intercambiados** por el emparejamiento «en orden» de `scripts/indicadores/pp_epp.py:188-191`. Corregido: FID 2025 = 59/61 = 96,72.
- D54 (usuaria, 18-nov). **PP 2024: se acepta el cruce EFPC2024-Z9-4 = sufijo 0755.** Los 8 rubros suman $500.398.021, igual al peso que el valor votado, tanto en apropiación inicial como en definitiva y obligación. QA lo confirmó en eje2024q4, pp. 80–82. IREPP 2024 ≈ 93,56.
- D55 (usuaria, 18-nov). **PP 2022: se resuelven el duplicado y las zonas.** Z6-2 y Z6-5 son el mismo proyecto (0347, $275.041.582) y se cuentan una vez, así que el universo queda en 39. 0346 se asigna a Z3 y 0348 a Z9, porque cada sufijo conserva su zona en todos los años. IREPP 2022 ≈ 85,51.
- D56 (usuaria, 18-nov). **PP 2021 = N/D con brecha.** Regla: un año se calcula solo si se verifica al menos el **50 %** de sus proyectos. 2021 tiene 39/89. QA buscó los 50 que faltan en las 11 columnas de Hacienda y no están: el problema es el dato, no el cálculo. Hay que pedir a Planeación el plan de acción 2021 con códigos de proyecto, el decreto de liquidación o las fichas BPIN.
- PP 2020 (QA, 18-nov): 20 de 69 proyectos tuvieron un contracrédito total ($3.627.742.979; la causa no consta) y cuentan como no ejecutados y no fieles. Se muestran en una fila de respaldo. Ojo: el EPP de 2020 (82,42) no los refleja; obligación / votado = 46,34 %.

## PDM: fichas finales (Analizador 18-nov) y decisiones de la usuaria
- D57 (usuaria). **ILE 2015–2017 se muestra con alerta de comparabilidad.** En 2019 cambia la composición del numerador (antes solo ICBF; desde 2019 ICBF + preescolar MEN) y 2016 es atípico. Estos años no afectan el ICPD.
- D58 (usuaria). **Homicidio: se mantiene la meta 2,7.** Va con nota: la LB 3,7 del PDM se calculó con la población post-COVID; con la VP, la LB 2023 da 3,52.
- D59 (usuaria). **Espacio verde 2024: se usa el valor publicado, 5,75.** Va con nota de sensibilidad: con la población VP daría 5,54 y el ICPD 2024 = 66,18.
- D60 (usuaria). **Semáforo del PDM = el mismo de PP**, como propuesta: Verde ≥ 90, Amarillo 75 a < 90, Naranja 50 a < 75, Rojo < 50.
- D61 (técnica, Analizador, verificado). **Población única = DANE VP (8-ago-2025)**, es decir, TerriData «Descripción general» / «Población total» (retroproyección 2005–2017 + PPED 2018–2042; 2018 = 228.836). El 238.221 sale de la dimensión «Censo 2005 y proyecciones DANE»: no usar. Filtrar siempre por dimensión. No cambia ningún valor calculado desde 2018. El referente vial 7,3301 se mantiene.
- D62 (técnica, Analizador). **Salud**: numerador = contributivo + subsidiado + especiales en 2012–2025. «Afiliados al SGSSS» de TerriData excluye los especiales en 2012–2017. 2017 se marca como atípico.
- D63 (técnica, Analizador). **Homicidio incluye el feminicidio.** Conteo propio de la Policía (m8fd-ahd9) / población VP; coincide con TerriData en 23 de 23 años. **Mortalidad vial**: conteo propio de ha6j-pa2r para 2010–2025 y tasa de TerriData para 2003–2009.
- D64 (técnica, Analizador). **Veedurías**: se cuentan por año de `fecha_inscripcion`. Puntaje con las veedurías listadas. Brecha = consecutivos faltantes del año / (listadas + faltantes); en 2022–2024 = 0.
- ICPD-7 provisional (Analizador): 2019 = 62,41; 2022 = 68,79; 2024 = 66,85 (Naranja). Brecha 2019–2025: 4 de 7 años.
- D65 (usuaria, 18-nov; QA del PDM aprobado con observaciones). **SP 2019 se mantiene (91,97) con alerta.** Entre 2016 y 2020 el REC rural es 0 y el urbano 100 en las tres coberturas, así que esa serie no es comparable con la de 2021 en adelante. La alerta de comparabilidad también va en el ICPD 2019. Hay que solicitar a Planeación el REC rural de 2016 a 2020.
- D66 (usuaria). **Brecha de veedurías, regla simple** (fiel a D64): brecha del año = consecutivos que faltan dentro del año / (listadas + esos faltantes), en un solo valor. Aparte se anota «consecutivos sin año atribuible: N». La brecha del periodo 2019–2025 es 19/52 = 36,5 %; reemplaza a «0 de 14 años».
- D67 (usuaria). **Veedurías 2012–2018 llevan puntaje**, como 2019–2025, con nota de listado incompleto (11 de 120 consecutivos).
- D68 (usuaria). **La LB de veedurías se mantiene en 5,4** (D46). Se anota el límite: si se contaran los faltantes seguros de 2019–2021, la LB sería 7,8.
- Pendientes de trazabilidad del PDM (QA): archivar las fuentes de espacio verde (VMR del AMVA, nota del Humboldt y PDF 146211 del POT) y la PPED por edad con su manifiesto; corregir la nota de SP; 66,18 → 66,17; anotar las veedurías 291–292 (año de registro 2026); anotar el universo de ILE (gestantes frente a solo niños en 2024–2025); «causa no documentada» en la cobertura neta de 2015; que el script falle si falta pypdf.
