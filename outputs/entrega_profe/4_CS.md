**HISTÓRICO (revisión del 15-nov). Los resultados vigentes están en los Word y en outputs/construccion/; las decisiones, en registro_cambios.md.**

# Dimensión: Control social (CS)

**Estado:** en revisión. Lo marcado **«Propuesta (Claude), por confirmar»** no está decidido.
Índice de la matriz: ICRCS = (DSI + PCS + VRC + RPO) / 4.
Criterio de la usuaria: el autorreporte que da 100 % no sirve.

## Cuadro de cambios

| Indicador original | Situación | Propuesta | Origen | Datos |
|---|---|---|---|---|
| **PCS** Promoción del control social (estrategias ejecutadas / programadas) | Es un **autorreporte** del Plan Indicativo: 1 de 1 estrategia = 100 % en 2024 y 2025. No muestra nada | **Propuesta (Claude), por confirmar:** reemplazar por la **oportunidad de respuesta a las PQRSDF** (solicitudes atendidas en términos de ley / solicitudes recibidas). Es la vía con la que la ciudadanía pide cuentas, y el dato sale del sistema de registro (BPMS), no de una autoevaluación | Quitar el autorreporte: Isabella. Reemplazo: Claude | Informes trimestrales de PQRSDF (Secretaría General). Años completos: **2020–2025** (2026 solo tiene dos trimestres). Verificado: 4.º trimestre de 2023: 3.544 / 4.149 = 85 %. Serie anual: sumar los trimestres |
| **DSI** Disponibilidad de información de seguimiento (metas con valor publicado / metas seleccionadas) | Verificable. Da 0 % en 2024 y 2025 (respaldo con todo lo publicado hasta hoy: 28,57 % en 2024) | **Se mantiene**; hay que **actualizar la lista de metas**. Ojo: DSI mide metas **del PDM**, así que la lista debe usar indicadores que el PDM sí tiene con meta (por ejemplo, aseguramiento subsidiado, no la cobertura total) | Profe (12-nov: «control social definido») | 2024–2025 (el PDM vigente empieza en 2024) |
| **VRC** Veedurías en la rendición de cuentas (eventos con veeduría / eventos) | Con la regla de la matriz queda N/D: en varios eventos no se sabe si hubo veeduría (incluida la audiencia del Alcalde de 2024 y 2025) | **Se mantiene**, reportado como **rango con la brecha de información**; la brecha debe bajar año a año. Ver cuadro abajo | Se mantiene: profe. Rango y brecha: Isabella (15-nov), con la propuesta de Claude | 2024: 66,7 %–100 % (brecha 33,3 %). 2025: 20 %–80 % (brecha 60 %). Se podría extender a 2021–2023 con las actas publicadas |
| **RPO** Respuesta a observaciones ciudadanas (con respuesta de fondo / observaciones) | 2024: 16/17 (94,12 %). 2025: el cuadro publicado está incompleto (respaldo: 1/2) | **Se mantiene** | Profe | 2024 y 2025 (frágil). **Se puede extender a 2021 y 2023**: esos cuadros de preguntas y respuestas están publicados |
| **ICRCS** Índice | Hoy **no se puede calcular**: VRC no tiene dato en 2024 ni en 2025 y RPO no lo tiene en 2025. Con los respaldos: 65,20 (2024) y 42,50 (2025), pero el respaldo de VRC va contra la regla de la matriz | Se recalcula con el reemplazo de PCS | — | 2024–2025, solo con respaldos. Extender VRC y RPO hacia atrás no da más años al índice, porque DSI empieza en 2024 |

## Variables con las que contamos

| Componente | Variables | Años con datos | Fuente |
|---|---|---|---|
| Oportunidad de respuesta a PQRSDF (en lugar de PCS) | Solicitudes atendidas en términos; solicitudes recibidas | 2020–2025 completos (trimestral) | Informes de PQRSDF, Secretaría General (sistemas BPMS y SAC) |
| DSI | Metas del PDM seleccionadas; metas con valor observado publicado | 2024–2025 | Plan Indicativo, informes de gestión, fuentes oficiales |
| VRC | Eventos de rendición de cuentas; eventos con veeduría documentada | 2024–2025 (ampliable a 2021–2023) | Actas y cuadros de rendición de cuentas |
| RPO | Observaciones ciudadanas; observaciones con respuesta de fondo | 2024–2025 (ampliable a 2021 y 2023) | Cuadros de preguntas y respuestas de la rendición |

## VRC con la brecha de información (decisión de Isabella, 15-nov)

| Año | Eventos | Con veeduría documentada | Sin veeduría | No se sabe | VRC mínimo | VRC máximo | Brecha |
|---|---|---|---|---|---|---|---|
| 2024 | 3 | 2 | 0 | 1 (audiencia del Alcalde) | **66,7 %** | 100 % | **33,3 %** |
| 2025 | 5 | 1 | 1 | 3 (audiencia del Alcalde; Discapacidad S1 y S2) | **20 %** | 80 % | **60 %** |

- **Cómo se construye:** mínimo = documentada / eventos (lo que no se sabe cuenta como no documentado, coherente con «presencia documentada»); máximo = (documentada + no se sabe) / eventos; brecha = no se sabe / eventos. Al índice entraría el **mínimo**, siempre con la brecha al lado.
- **Eventos contados:** audiencia del Alcalde + rendiciones sectoriales con acta o memoria (Salud, Discapacidad, Víctimas). Con todos los eventos publicados: 2024: 8 → 50 %–100 % (brecha 50 %); 2025: 7 → 28,6 %–85,7 % (brecha 57,1 %).
- **Trazabilidad:** clasificación por evento con documento, página, enlace y cita en `outputs/indicadores/clasificacion_vrc.csv`; cálculo en `scripts/revision_indicadores/vrc_brecha.py` → `outputs/revision_indicadores/CS/vrc_brecha.csv`.
- **Cómo se hizo:** se revisaron los documentos publicados de cada evento de 2024 y 2025 (actas, memorias, presentaciones, cuadros de preguntas y justificaciones). «Con veeduría» = el documento lo dice expresamente; «sin veeduría» = el documento integral del evento la describe sin veedurías; «no se sabe» = el documento no permite saberlo. Una invitación no cuenta como asistencia.
- **Se deja así y se mejora año a año:** la brecha se reporta cada año y debe bajar a medida que la Alcaldía publique las listas de asistencia y el observatorio lleve su registro. Cuando llegue a 0, mínimo y máximo coinciden y VRC vuelve a ser un solo número, como pide la matriz.
- **Cambia una regla de la matriz** (presencia desconocida = N/D), por eso se le informa al profe en la duda 14.

### Qué se necesita para que el observatorio lo siga midiendo
| Qué | Para qué | Lo produce | Lo recoge |
|---|---|---|---|
| Lista fija de eventos del año (Alcalde, Salud ×2, Discapacidad ×2, Víctimas, PP, grupos étnicos), confirmada con el cronograma de rendición de cuentas que publica la Alcaldía | Denominador estable (como en I4) | Alcaldía | Observatorio, en enero |
| Lista de asistencia de cada evento con la casilla «¿pertenece a una veeduría? ¿cuál?» | Que la presencia quede documentada | Dependencia que organiza el evento (por confirmar quién coordina la rendición en Envigado) | Observatorio, después de cada evento |
| Listado de veedurías inscritas | Verificar que el veedor pertenece a una veeduría inscrita | Personería | Observatorio, una vez al año |
| Registro propio: una fila por evento (fecha, dependencia, documento, enlace, veeduría sí/no/sin soporte) | Trazabilidad | — | Observatorio; cierre al 31-dic, búsqueda hasta el 31-mar (regla 17) |
| Observación directa de la audiencia del Alcalde | Es el evento que hoy queda sin dato | — | Observatorio |

## Notas
- **PQRSDF:** el propio informe advierte que algunas solicitudes se responden a tiempo pero se cierran tarde en el sistema, lo que baja el porcentaje. Además, el total incluye solicitudes pendientes que aún están dentro del plazo legal (224 en el 4.º trimestre de 2023) y el informe usa dos sistemas (BPMS y SAC). Hay que fijar la regla antes de calcular.
- **Veedurías:** el número de veedurías inscritas se usa en el PDM (en lugar de procesos participativos). VRC mide otra cosa: si hay veedurías en los eventos de rendición de cuentas.

## Variables por solicitar (informe futuro)
- Programación y ejecución de las estrategias de promoción del control social, con soportes (para el PCS original).
- Listas de asistencia de las rendiciones de cuentas 2024–2025, con la presencia de veedurías (Alcaldía), para cerrar la brecha de VRC.
- Cuadro completo de preguntas y respuestas de la rendición de cuentas de 2025 (para RPO).
