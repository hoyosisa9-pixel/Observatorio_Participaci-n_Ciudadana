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
