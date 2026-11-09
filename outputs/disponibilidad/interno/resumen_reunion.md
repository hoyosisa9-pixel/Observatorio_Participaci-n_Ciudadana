# Disponibilidad y brecha de información del sistema de indicadores
**Para:** profesores Francisco Correa y Eber · **Observatorio Ciudadano de Envigado (OCE)** · **Corte:** 14-oct-2026

Alcance: 4 sistemas, 39 variables y 23 indicadores e índices, para 2024, 2025 y lo publicado de 2026 (sin cierre anual). No se calculó el IIC ni ningún índice. Todas las cifras son exploratorias, con estado **Pendiente**; ninguna está validada. Detalle en `disponibilidad_brecha_indicadores.xlsx`.

## 1. Semáforo por sistema

| Sistema | Variables (Disp. / Constr. / Parcial / Proxy / N.D.) | Indicadores | Semáforo |
|---|---|---|---|
| CS - Control social | 2 / 3 / 3 / 0 / 0 (8) | 1 viable, 3 con ajuste, ICRCS con ajuste | Amarillo-verde: ICRCS calculable solo en 2024, tras fijar definiciones |
| PC - Participación ciudadana | 1 / 0 / 11 / 0 / 0 (12) | 3 con ajuste (I2, I3, I5), 4 no viables (I1, I4, I6, ISPC) | Rojo: casi todo es Parcial por falta de soportes y de un inventario de procesos |
| PDM - Cumplimiento del PDM | 7 / 1 / 3 / 2 / 1 (14) | 2 con ajuste (CM, CA), 5 no viables (CS, CP, CD, ILE, ICPD-6) | Rojo: las metas existen, pero casi no hay valores observados con el método del plan |
| PP - Presupuesto participativo | 0 / 3 / 2 / 0 / 0 (5) | EPP con ajuste, 3 no viables (CEP, CMA, IREPP) | Amarillo: EPP reconstruible; CEP y CMA sin soporte |
| **Total** | **10 / 7 / 19 / 2 / 1 (39)** | **1 viable, 10 con ajuste, 12 no viables** | |

De las 10 variables Disponibles, 7 son metas del PDM; solo 3 son valores observados (CS.PCS.EE, CS.PCS.EP, PC.I2.V2). Por la regla de no redistribuir pesos (Guía 15), ningún índice se calcula con todas sus entradas, salvo ICRCS en 2024 una vez definidos los universos.

## 2. Qué es viable ya
- **CS.PCS** (control social promovido): programación y ejecución publicadas en el plan indicativo; resultado 100 en 2024 y 2025, con la salvedad de que es autorreporte del DAP ("Sí, pero con observaciones").
- **PC.I2** (cobertura territorial): 13 zonas publicadas en el PDM (p. 558); la numeración por zona sale de reportes agregados.
- **PP.EPP** 2025: reconstruible con el plan de acción (ver sección 4).
- **PDM.CM** 2024: Forensis da 7,25 por 100.000, con meta 8.
- **PC.I3 y PC.I5**, acotados (ver ajustes).

## 3. Indicadores a ajustar y cómo (priorizado)
1. **Definir universos antes de medir (Guía 17):** lista de unas 10 metas para CS.DSI; qué cuenta como evento de rendición de cuentas (CS.VRC); universo de observaciones y rúbrica de "respuesta de fondo" (CS.RPO). Desbloquea ICRCS 2024.
2. **Regla de identificación del PP:** listado de la EFPC cruzado con el plan de acción, excluyendo el proyecto 0553 (costo del proceso); CUIPO solo como contraste.
3. **Inventario anual de procesos participativos cerrados** (con acto de decisión y enlace de publicación), solicitado a Participación; sin él PC.I1, PC.I4 y el ISPC quedan N/D.
4. **Fuente oficial por resultado del PDM:** INMLCF para mortalidad vial; Secretaría de Salud con método BDUA para aseguramiento (no usar TerriData, que da ~12 % sobre población total); Medio Ambiente para espacio público verde; ficha técnica del DAP para procesos participativos; Índer para cobertura deportiva; ICFES para el logro por plantel.
5. **Metas del PDM:** evaluar 2024-2025 como logro al corte frente a la meta 2027, o pedir metas anuales al DAP.
6. **PC.I3:** medir en personas (producto 521 del plan indicativo) o exigir a la EFPC la programación fechada de acciones.
7. **PP.CEP:** o se exigen actas de recibo e informes finales, o se admite un CEP provisional con autorreporte y regla de agregación fija (rango hoy 74-90 %), rotulado.
8. **Zonas:** fijar 13 como denominador y declarar que la zona 1 no tiene comité ni proyectos PP.
9. **PC.I5 y PC.I6:** seleccionar las instancias (mínimo CMPC y CTP) y pedir relación anual de actividades con modalidad.
10. **Correcciones formales de la Matriz:** códigos repetidos, páginas del PDM, estados no definidos en la Guía, nombre de PP.CMA.NA.

## 4. Reconstrucción de la EPP 2025 como complemento a la revisión del profesor Francisco
La conclusión del profesor, con las cuatro fuentes que revisó, es que EPP es N/D; eso sigue siendo cierto con esas fuentes. Esta tarea aporta **una quinta fuente oficial** (plan de acción IV trimestre 2025 y ROCA del DAP) que permite construir la variable sin inferir nada: se suman dos columnas oficiales (apropiación definitiva y obligación) sobre 61 proyectos con "ZONA n" que cruzan 1:1 con el listado de la EFPC (56 con valor idéntico; todo de libre destinación, sin cofinanciación que excluir).

- **Resultado PRELIMINAR, POR VALIDAR (lo debe confirmar Planeación):** apropiación $20.928.690.218; obligación $20.684.427.720; EPP 2025 = 98,83 % sin el proyecto 0553 (98,86 % con él). El cálculo es reproducible con `scripts/epp_pp_2025.py` (detalle por zona y por proyecto en `epp_pp_2025_zonas.csv` y `epp_pp_2025_proyectos.csv`): 61 proyectos en ambas fuentes, misma cantidad por zona, 56 con valor idéntico al peso. Por zona, la EPP va de 97,32 % (zona 9) a 99,83 % (zona 3).
- **Límites:** 2026 solo provisional (obligación al 30-jun-2026); 2024 aproximado (47 de 50 proyectos), por lo que es Parcial. CUIPO difiere en 3 %, sin explicar, y se usa solo como contraste.
- **CEP y CMA siguen N/D, como concluyó el profesor**, y por tanto el IREPP no es calculable. El autorreporte de entrega de 2025 oscila entre 45 y 55 de 61 proyectos según la regla de agregación.
- **Formulación sugerida:** "EPP 2025 reconstruible en 98,83 % (preliminar, por validar); CEP y CMA sin soporte; IREPP no calculable".

## 5. Preguntas abiertas para los profesores
1. PP: ¿se acepta el EPP reconstruido como "Construible, Pendiente" y la exclusión del proyecto 0553? ¿Quién en Planeación confirma los códigos PP 2024-2026?
2. PP.CEP: ¿se admite un CEP provisional con autorreporte (y con qué regla de agregación) o se exige acta de recibo/informe final?
3. Control social: ¿qué es un "evento municipal de rendición de cuentas" (1 a 8 por año según el criterio)? ¿La mención de veedurías en un acta acredita presencia? ¿Cuáles son las metas de DSI? ¿Qué es "respuesta de fondo"?
4. PC.I1 y PC.I4: ¿se solicita el inventario oficial o se fija un universo cerrado ex ante (PP, PDM/POT, rendición), aceptando el riesgo de subconteo?
5. ¿Se evalúan 2024-2026 contra la meta 2027 o se piden metas anuales al DAP?
6. ¿Se acepta Forensis para mortalidad vial y el POT (5,75, no reconciliable con la base 3,88) para espacio público verde? ¿Se descarta TerriData para aseguramiento?
7. PDM.CP y PDM.CD: ¿se redefinen (no hay método ni cifra) o se solicitan al DAP y al Índer?
8. PC.I3: ¿personas formadas (dato oficial) o acciones formativas (requiere programación de la EFPC)?
9. ¿Cómo tratar datos con corte distinto al 31-dic y 2026 sin cierre? ¿Se presentan cifras provisionales de 2026?
10. Inconsistencias a consultar con la Alcaldía: personas capacitadas 2025 (65 vs 170), fecha de la audiencia del Alcalde 2024 (20 vs 26-nov) y de la rendición de PP 2024 (12-ene vs 12-feb-2025).

## 6. Documentos por descargar manualmente
Detalle en `documentos_por_descargar.md`. Los de mayor efecto:
- Clasificación de planteles ICFES (`https://www.icfes.gov.co/analisis-de-datos/`; la nota de 10 años devolvió 404): cambiaría PDM.ILE.S de Parcial a Construible.
- Cifras de aseguramiento MinSalud, tablero histórico (`https://minsalud.gov.co/Proteccion-Social/Paginas/cifras-aseguramiento-salud.aspx`) y método BDUA de Salud: cambiaría PDM.CS.O.
- Forensis 2025 / observatorio ANSV (`https://ansv.gov.co`): PDM.CM.O 2025.
- Rendición de cuentas de la Alcaldía (`https://www.envigado.gov.co/rendicion-de-cuentas`): CS.VRC.ET/EV, CS.RPO.OT/OR, PC.I6.V1.
- Rendición de Cuentas 2025 p. 61 (`.../002947/147324_rc2025.pdf`) y Acta 096 de 2026 del Concejo (`https://www.concejoenvigado.gov.co/wp-content/uploads/2026/09/Acta-096-de-2026.pdf`): PP.EPP.APP 2026, PP.CEP.PE, PP.CEP.PEC, PP.CMA.NA.
- Acuerdo 007 de 2025 (`https://www.concejoenvigado.gov.co/wp-content/uploads/2025/02/Acuerdo-007-de-2025.pdf`): PC.I2.V2.
- Sección Observatorio Ciudadano de la EFPC (`https://efpc.envigado.gov.co/observatorio-ciudadano/`, 404): PC.I1 y PC.I4.
