# Dimensión: Cumplimiento del Plan de Desarrollo Municipal (PDM)

**Estado:** cerrada (provisional). Pendiente solo lo que está en `dudas_para_el_profe.md`.
Series: `outputs/revision_indicadores/PDM/series_pdm.csv`.

## Cuadro de cambios

| Indicador original (matriz) | Qué se hizo | Motivo | Origen | Datos disponibles hoy |
|---|---|---|---|---|
| Logro de aseguramiento al régimen subsidiado | **Reemplazado** por la cobertura de salud total | La información del subsidiado no era fácil de conseguir; la cobertura total cubre a toda la población | Profe (reunión 1 propuso discapacidad, inclusión o cobertura de salud); Isabella eligió cobertura; confirmado el 12-nov | 2018–2025, TerriData (MinSalud) |
| Logro de reducción de mortalidad vial | **Se mantiene**, con una sola fuente (Policía) | Forensis 2025 no se puede consultar y mezclar fuentes cambia la medida | Se mantiene: profe (12-nov). Fuente única: Claude | 2019–2025, Policía / TerriData |
| Logro de espacio público verde | **Ajustado**: avance de 4,66 hacia 9 m²/hab | Con la meta del PDM (3,88) daba 100 sin exigencia | Profe (12-nov) | Solo 2022 (4,66, AMVA) y 2024 (5,75, POT) |
| Logro de implementación de procesos participativos | **Reemplazado** por el número anual de veedurías inscritas en la Personería | No hay cifra ni método oficial publicado | Profe (notas: «cambiar por número de veedurías»); Isabella lo ubicó aquí | 2012–2026 (sin registros en 2013), datos.gov.co (Personería); listado al parecer incompleto |
| Logro de cobertura deportiva | **Reemplazado** por servicios públicos (acueducto, alcantarillado, aseo y energía) | No hay tasa oficial del Índer | Servicios públicos: profe (12-nov). Que reemplaza a deporte: Isabella | Acueducto, alcantarillado y aseo 2008–2024 (la página web de TerriData muestra también 2025); energía 2019–2024 |
| Índice de logro educativo (primera infancia + Saber 11) | **Ajustado**: Saber 11 → cobertura neta; primera infancia con 0–5 años | La clasificación de planteles no es pública; hay que ver la evolución en años | Transformar el ILE: profe (reunión 1). Cobertura neta: Isabella (aceptada el 12-nov). 0–5 años: Isabella, con el cálculo de Claude | 2018–2024, TerriData / MinEducación |
| — (sexta medida de la reunión) | **Pregunta para el profe**: con las veedurías en el lugar de procesos participativos, el índice ya tiene sus 6 componentes. Una medida de seguridad sería un **séptimo** componente | En la reunión se acordaron 5 líneas y quedó abierta una sexta | Candidatos: profe (12-nov) | Homicidio, hurto y lesiones 2019–2025, Policía |
| — (dimensión futura) | **Logros de políticas públicas**, para el observatorio futuro | Seguimiento de las políticas municipales | Profe (reunión 1) | Sin datos; variables por solicitar |

## Variables con las que contamos

| Componente | Variables | Años | Fuente | ¿Tiene meta en el PDM? |
|---|---|---|---|---|
| Cobertura de salud | Afiliados al SGSSS; población DANE | 2018–2025 | TerriData, DANE | No (la meta es del subsidiado) |
| Logro educativo | Cobertura neta; niños atendidos en atención integral; población de 0–5 años | 2018–2024 | TerriData, MinEducación, DANE | Sí (19 % y 100,44 %) |
| Servicios públicos | Coberturas totales de acueducto, alcantarillado, aseo y energía | 2019–2024 (las 4 juntas) | TerriData | Sí, 99,9 %, pero de otro universo (Encuesta de Calidad de Vida, no el registro); el aseo no tiene |
| Mortalidad vial | Muertes en tránsito; población DANE | 2019–2025 | Policía / TerriData | Sí, 8 por 100.000, pero de otro universo (Secretaría de Movilidad, no Policía) |
| Espacio verde | m² de espacio público verde urbano por habitante | 2022, 2024 | AMVA, POT | Referente 9 m² (propuesto por el profe) |
| Veedurías | Veedurías inscritas por año | 2012–2026 | Personería (datos.gov.co) | No |

**Viabilidad del índice del PDM (revisado por QA):** todos los componentes tienen variables con datos públicos, pero hoy el índice **no se puede calcular** hasta resolver las dudas 1 a 5: los referentes de salud y veedurías, y metas del mismo universo para servicios públicos y mortalidad vial. Resueltas, solo saldría para los años con espacio verde: **2022** (línea base 4,66, avance 0) **y 2024** (5,75; falta confirmar que la medición del POT es de 2024). Para tener serie hace falta espacio verde año por año.

Subíndices con los reemplazos: **bienestar social** = salud + logro educativo + servicios públicos; **territorio y gobernanza** = mortalidad vial + espacio verde + veedurías.

## Notas
- **Primera infancia:** el PDM dice 0–6 años, pero 0–5 reproduce mejor su línea base (16,56 frente a 16,3; con 0–6 da 14,02).
- **Servicios públicos:** el registro tiene saltos (caída a ~68 % en 2020; zona rural en 0 de 2016 a 2020) y la cobertura rural es baja (acueducto 17,9 % en 2024), probablemente porque los acueductos veredales no están registrados. La meta del PDM viene de la Encuesta de Calidad de Vida y no es comparable con el registro.
- **Veedurías:** los consecutivos van del 118 al 293, pero solo hay 45 registros publicados (en 2024 aparece 1). Control social también usa veedurías (en la rendición de cuentas), pero mide otra cosa.

## Variables por solicitar (informe futuro)
- Espacio público verde por habitante, año por año (AMVA o Planeación).
- Clasificación de planteles en Saber 11 (Secretaría de Educación).
- Listado histórico completo de veedurías inscritas (Personería).
- Logros de políticas públicas: políticas vigentes, informes de seguimiento y metas cumplidas.
