# Guía de estudio: el sistema de indicadores del Observatorio Ciudadano de Envigado

**Para:** Isabella (uso interno, no se entrega a los profesores)
**Preparada:** martes 10 de noviembre de 2026, para la reunión del miércoles 11 con el profesor Francisco Correa
**Fuentes:** solo archivos del repositorio. Las cifras se tomaron de los entregables enviados: `resultados_indicadores.csv`, `base_indicadores_OCE.xlsx`, `informe_avance_1` y el Excel de disponibilidad versión 1.2. Esta guía no recalcula ni cambia ningún resultado.

> **Cómo usar esta guía.** La Parte 1 cuenta lo que se hizo, en orden. La Parte 2 reúne los problemas. La Parte 3 es la propuesta de variables que vas a discutir. La Parte 4 explica la base de datos del tablero. La Parte 5 te prepara para la reunión: si solo tienes diez minutos, lee la hoja de reunión (5.1) y las respuestas a preguntas (5.2). Al final hay un glosario con todas las siglas.

---

# Parte 1. Qué se ha hecho

## 1.1 Punto de partida

### Qué es el OCE y qué es el OPCEN

- El **Observatorio Ciudadano de Envigado (OCE)** es, según la propuesta de 2025 (`Documentos guias/02032026_Propuesta  Observatorio Ciudadano.docx`), una instancia **ciudadana y autónoma** que vigila, evalúa y pide cuentas sobre la gestión pública municipal. La propuesta lleva los nombres de la Alcaldía de Envigado, el Departamento Administrativo de Planeación (DAP), la Subdirección de Participación Ciudadana y la Institución Universitaria de Envigado (IUE).
- El **Observatorio de Participación Ciudadana de Envigado (OPCEN)** es la instancia **institucional** (de la Alcaldía). La propuesta dice que el OPCEN "diagnostica y describe" los procesos participativos, mientras el OCE "vigila, evalúa críticamente y exige rendición de cuentas".
- El marco normativo clave es el **Acuerdo Municipal 033 de 2022** (Política Pública de Participación Incidente).
- Uno de los objetivos de la propuesta es "diseñar un sistema de indicadores y monitoreo" y ofrecer a la ciudadanía "tableros de control interactivos" para explorar indicadores, comparar periodos y filtrar por zona.

### Qué pidieron los profesores

- Los profesores **Francisco Correa y Eber Gutiérrez** definieron un sistema preliminar de indicadores para la Alcaldía de Envigado (con entrega prevista el 15 de octubre de 2026) y lo plasmaron en la **Matriz de información de indicadores** (archivo Excel en `Documentos guias/`, que no se modifica).
- **Primer encargo (octubre):** decir, variable por variable, si el dato existe, dónde está, en qué formato y qué falta. Se llamó "disponibilidad y brecha". No incluía calcular nada.
- **Segundo encargo (noviembre), del profesor Francisco:** organizar los datos de los indicadores que sí se pueden calcular, construir las variables que se obtienen de información pública y calcular los indicadores clasificados como "Viable" y "Viable con ajuste". Para los "No viables por ahora", solo construir las variables.
- **Para qué sirve:** el sistema alimenta un informe a Envigado y un **tablero** (una página interactiva con los indicadores). Otra persona hará el tablero y necesita, al final de esta semana, la lista de variables y la base de datos.

### Antecedente: el IIC

Antes de estos encargos, el repositorio trabajó en el **Índice de Incidencia Ciudadana (IIC)**: se hizo un inventario de archivos fuente y se probaron las "llaves de cruce" entre bases (por ejemplo, el código BPIN de un proyecto). Ese trabajo sirvió como insumo, pero **ninguno de los dos encargos calcula el IIC**.

## 1.2 La matriz de los profesores

La Matriz es un libro de Excel con cinco hojas:

| Hoja | Qué tiene |
|---|---|
| Guía | 20 reglas de uso (cómo registrar datos, cuándo un valor se puede usar, cómo leer los semáforos) |
| Matriz | El catálogo: 19 indicadores y 4 índices, con fórmula, fuentes, semáforo e interpretación |
| Registro | Donde se escribe el valor de cada variable por año, periodo y territorio: 39 variables × 4 años (2024 a 2027) = 156 filas |
| PP_Proyectos | Una fila por proyecto de presupuesto participativo, para las personas atendidas y la meta de atención |
| Resultados | Fórmulas que calculan los puntajes y los semáforos |

**Nota sobre la "guía metodológica".** En `Documentos guias/` solo hay dos archivos: la propuesta del OCE y la Matriz. La guía metodológica es la hoja **Guía** de la Matriz. Los cuatro anexos en Word que la Guía cita como fuente metodológica (uno por sistema) y el documento de diagnóstico del OPCEN **no están en el repositorio**.

### Conceptos básicos (para leer todo lo demás)

- **Variable:** un dato que se busca en un documento. Ejemplo: "número de estrategias de promoción del control social ejecutadas".
- **Indicador:** una cuenta hecha con dos variables, casi siempre 100 × numerador / denominador. Da un **puntaje** de 0 a 100.
- **Índice:** un promedio ponderado de varios indicadores. Ejemplo: el ICRCS promedia cuatro indicadores de control social.
- **Semáforo:** el color que se asigna al puntaje según unos umbrales (cortes) fijados en la Matriz. Se clasifica **sin redondear**: 99,99 en un indicador del PDM no es Verde.
- **Peso:** cuánto cuenta cada indicador dentro de un índice (por ejemplo, 25 %).

### Reglas generales de la Guía que hay que conocer

- **N/D (no disponible):** información faltante, no validada o fuera del corte. Sin color.
- **NE (no evaluable):** base cero; no hay universo que evaluar. Ejemplo: la zona 1 no tiene proyectos de presupuesto participativo. Sin color.
- **Revisar:** claves duplicadas, numerador mayor que el universo u otra inconsistencia. Sin color.
- **Un cero observado y validado con base positiva sí da 0 puntos** (y su color).
- **Tope de 100:** solo los indicadores que lo dicen expresamente (PCS y los del PDM) se limitan a 100. En los demás, si el numerador supera al denominador, el resultado es "Revisar".
- **Índices sin redistribuir pesos (regla 15):** si falta un componente, el índice queda N/D. No se reparte el peso entre los que sí hay.
- **"Validado" (regla 3):** la Matriz solo usa un valor si tiene estado "Validado", fuente, fecha de corte y valor numérico. Validar es decisión del responsable de la matriz (regla 19). **En este trabajo nada está validado:** las variables están en "Pendiente" y los resultados en "Provisional".
- **Metas del PDM (reglas 8 y 9):** las metas son de 2027; comparar 2024 con la meta 2027 es "logro al corte", no cumplimiento anual.
- **Control social (regla 17):** los universos se definen antes de medir; una convocatoria no prueba que las veedurías asistieron; un acuse de recibo no es respuesta de fondo; la búsqueda documental llega hasta el 31 de marzo del año siguiente.
- **Alertas de PP (regla 20):** si CEP o CMA es menor de 50, se muestra una alerta aunque el índice sea alto. No se compensa una entrega débil con buena ejecución financiera.

### Los cuatro sistemas, indicador por indicador

**Sistema 1. Control social (CS).** Mide si existen condiciones para que la ciudadanía vigile la gestión y si la Alcaldía responde.

| Código | Qué mide, en palabras simples | Fórmula y variables | Semáforo (umbrales exactos de la Matriz) |
|---|---|---|---|
| CS.DSI – Disponibilidad de información de seguimiento | De una lista fija de metas del plan, cuántas tienen una cifra pública utilizable del año | DSI = 100 × MD / MS. MD = metas de la lista con valor observado publicado; MS = metas de la lista | Verde ≥ 90; Amarillo 70 a < 90; Rojo < 70; Gris N/D |
| CS.PCS – Cumplimiento de la promoción del control social | Si se hicieron las estrategias de promoción del control social que estaban programadas | PCS = 100 × EE / EP, con tope de 100. EE = ejecutadas; EP = programadas | Verde ≥ 75; Amarillo 50 a < 75; Rojo < 50 |
| CS.VRC – Presencia de veedurías en la rendición de cuentas | En cuántos eventos de rendición de cuentas consta que hubo veedurías | VRC = 100 × EV / ET. EV = eventos con presencia documentada; ET = eventos del año. Si hay casos con presencia desconocida, VRC = N/D | Verde ≥ 75; Amarillo 50 a < 75; Rojo < 50 |
| CS.RPO – Respuesta pública a observaciones ciudadanas | De las observaciones que hizo la ciudadanía en la rendición de cuentas, cuántas recibieron respuesta de fondo publicada | RPO = 100 × OR / OT. OR = con respuesta de fondo; OT = total de observaciones sobre el PDM | Verde ≥ 90; Amarillo 70 a < 90; Rojo < 70 |
| CS.ICRCS – Índice de Condiciones y Respuesta para el Control Social | Resumen de los cuatro anteriores | (DSI + PCS + VRC + RPO) / 4, cada uno al 25 % | Verde ≥ 75; Amarillo 50 a < 75; Rojo < 50 |

**Sistema 2. Participación ciudadana (PC).** Mide si la participación llega al territorio, si incide en decisiones y si la oferta formativa se cumple.

| Código | Qué mide | Fórmula y variables | Semáforo |
|---|---|---|---|
| PC.I1 – Incorporación de aportes ciudadanos | De los procesos participativos cerrados, en cuántos se incorporó al menos un aporte ciudadano en una decisión formal | 100 × V1 / V2 | Verde ≥ 70; Amarillo 40 a < 70; Rojo < 40 |
| PC.I2 – Cobertura territorial de la participación | En cuántas de las zonas oficiales hubo al menos una actividad participativa | 100 × V1 / V2 (V2 = zonas oficiales) | Verde ≥ 80; Amarillo 50 a < 80; Rojo < 50 |
| PC.I3 – Cumplimiento de la oferta formativa | De las acciones formativas programadas para la ciudadanía, cuántas se realizaron | 100 × V1 / V2 | Verde ≥ 75; Amarillo 50 a < 75; Rojo < 50 |
| PC.I4 – Publicación de resultados participativos | De los procesos cerrados, cuántos publicaron sus resultados | 100 × V1 / V2 | Verde ≥ 90; Amarillo 70 a < 90; Rojo < 70 |
| PC.I5 – Cumplimiento de sesiones de participación (complementario) | De las sesiones previstas de unas instancias seleccionadas, cuántas se hicieron | 100 × V1 / V2 | Verde ≥ 75; Amarillo 50 a < 75; Rojo < 50 |
| PC.I6 – Disponibilidad de participación virtual (complementario) | De las actividades realizadas, cuántas tuvieron un canal virtual para intervenir | 100 × V1 / V2 | Verde ≥ 60; Amarillo 30 a < 60; Rojo < 30 |
| PC.ISPC – Índice Simplificado de Participación Ciudadana de Envigado | Resumen de I1 a I4 (I5 e I6 no entran) | 0,30 × I1 + 0,25 × I2 + 0,20 × I3 + 0,25 × I4 | Verde ≥ 75; Amarillo 50 a < 75; Rojo < 50 |

**Sistema 3. Cumplimiento del Plan de Desarrollo Municipal (PDM).** Compara seis resultados del plan con su meta de 2027. En todos: puntaje = 100 × observado / meta, con tope de 100 (en mortalidad vial es al revés, porque menos es mejor: 100 × meta / observado).

| Código | Qué mide | Meta de formulación (2027) | Semáforo (igual para todo el PDM) |
|---|---|---|---|
| PDM.CS – Logro de la cobertura de aseguramiento subsidiado | Cobertura del régimen subsidiado de salud | 99,5 % | Verde solo en 100 sin redondear; Amarillo 75 a < 100; Rojo 0 a < 75 |
| PDM.CM – Logro de la reducción de mortalidad vial | Muertes por siniestros viales por cada 100.000 habitantes | 8 | Igual |
| PDM.CA – Logro del mantenimiento de espacio público verde | Metros cuadrados de espacio público verde por habitante | 3,88 | Igual |
| PDM.CP – Logro de implementación de procesos participativos | Porcentaje oficial de implementación de procesos de participación | 100 % | Igual |
| PDM.CD – Logro de cobertura de servicios deportivos | Cobertura de servicios deportivos | 6 % | Igual |
| PDM.ILE – Índice de Logro Educativo | Promedio de dos logros: atención integral de 0 a 6 años (LA = 100 × A / MA) y porcentaje de instituciones oficiales que mejoran su categoría en Saber 11 (LS = 100 × S / MS) | MA = 19 %; MS = 14 % | Igual |
| PDM.ICPD-6 – Índice de cumplimiento de resultados seleccionados del PDM | Promedio de dos subíndices: bienestar social, IBS = (CS + ILE + CD) / 3, y territorio y gobernanza, ITG = (CM + CA + CP) / 3; ICPD-6 = (IBS + ITG) / 2 | — | Igual |

**Sistema 4. Presupuesto participativo (PP).** Mide si los recursos que la comunidad priorizó se ejecutan, se entregan y atienden a la población.

| Código | Qué mide | Fórmula y variables | Semáforo (cuatro colores) |
|---|---|---|---|
| PP.EPP – Ejecución presupuestal del PP | Qué parte del dinero disponible del PP se convirtió en obligaciones | EPP = 100 × OPP / APP. OPP = obligaciones; APP = apropiación definitiva | Verde 90 a 100; Amarillo 75 a < 90; Naranja 50 a < 75; Rojo 0 a < 50; N/D o NE sin puntaje |
| PP.CEP – Cumplimiento de entrega de proyectos | Qué parte de los proyectos completó las entregas exigibles, verificadas con acta o informe | CEP = 100 × PEC / PE | Igual |
| PP.CMA – Cumplimiento de metas de atención | Promedio, por proyecto, de personas atendidas frente a la meta | Por proyecto: 100 × mín(PA / MP; 1); CMA = SC / NA | Igual |
| PP.IREPP – Índice de Resultados de Ejecución del PP | Resumen con 80 % del peso en entrega y atención | 0,20 × EPP + 0,40 × CEP + 0,40 × CMA | Igual |

Para el municipio, el EPP es **la suma de obligaciones sobre la suma de apropiaciones** de todos los proyectos, no un promedio de zonas (regla 13).

## 1.3 Etapa 1: disponibilidad y brecha (octubre de 2026)

### Qué se buscó y dónde

- Para cada una de las **39 variables** de la hoja Registro: si el dato existe, en qué documento, en qué página, para qué años y si viene por zona.
- Para cada uno de los **23 indicadores e índices**: si se puede calcular hoy.
- **Dónde:** primero en el archivo del observatorio (los 2.787 archivos fuente, ver 1.5) y después en la web, en fuentes primarias: Ministerio de Salud (MinSalud), Medicina Legal (Forensis), DANE, TerriData, ICFES, ADRES, datos de la Policía, sitio de la Escuela de Formación (EFPC) y portal de la Alcaldía. Regla: ninguna variable podía quedar "No disponible" o "Proxy" sin búsqueda en fuentes primarias.
- **Periodo:** 2024, 2025 y lo publicado de 2026. **Territorio:** municipio y, cuando había, las 13 zonas.

### La escala de disponibilidad, con un ejemplo real de cada una

| Categoría | Definición (hoja Léeme del Excel) | Ejemplo real |
|---|---|---|
| **Disponible** | Publicado con la unidad y la definición del indicador para al menos un año entre 2024 y 2026 | CS.PCS.EE: el Plan Indicativo reporta 1 estrategia de promoción del control social ejecutada en 2024 y 1 en 2025 (consecutivo 544) |
| **Construible** | No está publicado como cifra, pero se puede contar o calcular con documentos que ya se tienen | PP.EPP.APP: no existe una cifra oficial de "apropiación del PP"; se reconstruyó cruzando los 61 proyectos del plan de acción 2025 con el listado de la EFPC ($20.928.690.218) |
| **Parcial** | Existe, pero faltan años, la desagregación o el soporte exigido, o la definición difiere | PC.I2.V1: hay reportes de asambleas por zona, pero no las actas o listas de asistencia que pide la Matriz |
| **Proxy** | Solo hay un dato parecido, con otra definición, unidad o universo | PDM.CD.O: la Matriz pide cobertura deportiva en porcentaje; solo se halló "38 de 40 disciplinas implementadas" en 2025 |
| **No disponible** | No está en el archivo ni en fuentes públicas; hay que pedirlo | PDM.CP.O: no hay cifra ni método oficial del porcentaje de implementación de procesos participativos |

### Resultados (verificados en el Excel de disponibilidad, versión 1.2)

**Variables (39):** Disponible 10 · Construible 7 · Parcial 19 · Proxy 2 · No disponible 1.

| Sistema | Disponible | Construible | Parcial | Proxy | No disponible | Total |
|---|---|---|---|---|---|---|
| CS | 2 | 3 | 3 | 0 | 0 | 8 |
| PC | 1 | 0 | 11 | 0 | 0 | 12 |
| PDM | 7 | 1 | 3 | 2 | 1 | 14 |
| PP | 0 | 3 | 2 | 0 | 0 | 5 |
| **Total** | **10** | **7** | **19** | **2** | **1** | **39** |

Qué significan esos números:
- **De los 10 "Disponibles", 7 son metas del PDM** (99,5; 8; 3,88; 100; 6; 19 y 14). Solo **3 son datos observados**: CS.PCS.EE, CS.PCS.EP y PC.I2.V2 (13 zonas). Es decir, casi nada de lo que se mide está publicado tal cual.
- **Los 7 Construibles** (CS.DSI.MD, CS.DSI.MS, CS.VRC.ET, PDM.ILE.A, PP.EPP.OPP, PP.EPP.APP, PP.CEP.PE) se pueden armar con documentos que ya se tienen.
- **Los 19 Parciales** son la mitad del sistema: el dato existe, pero falta soporte, años o definición. Participación ciudadana tiene 11 de sus 12 variables en Parcial.
- **Los 2 Proxy** son PDM.CS.O y PDM.CD.O. **El único No disponible** es PDM.CP.O.

**Indicadores e índices (23):** 1 Viable · 10 Viables con ajuste · 12 No viables por ahora.
- **Viable (1):** CS.PCS.
- **Viables con ajuste (10):** CS.DSI, CS.VRC, CS.RPO, CS.ICRCS, PC.I2, PC.I3, PC.I5, PDM.CM, PDM.CA y PP.EPP. "Con ajuste" significa que se pueden calcular si se fija una definición (por ejemplo, el universo de eventos) o se acepta una fuente con salvedades.
- **No viables por ahora (12):** PC.I1, PC.I4, PC.I6, PC.ISPC, PDM.CS, PDM.CP, PDM.CD, PDM.ILE, PDM.ICPD-6, PP.CEP, PP.CMA y PP.IREPP. Los tres índices (ISPC, ICPD-6, IREPP) quedan bloqueados por la regla 15: les falta al menos un componente.

### Qué se entregó en la etapa 1

- `outputs/disponibilidad/disponibilidad_brecha_indicadores.xlsx`: hojas Léeme, Indicadores (23), Variables (39), Fuentes (66), Solicitudes (32 datos para pedir a la Alcaldía, por dependencia) y Enlaces (218 hipervínculos).
- Informe formal en Word (`informe_disponibilidad.docx`) y en LaTeX/PDF con su paquete para Overleaf.
- Un hallazgo adicional que se ofreció como complemento a la revisión previa del profesor Francisco (que había dejado EPP en N/D): **EPP 2025 = 98,83, marcado "preliminar, por validar"**.
- Documentos internos (no se entregan): resumen de la reunión y lista de documentos por descargar, en `outputs/disponibilidad/interno/`.

## 1.4 Etapa 2: cálculo provisional (noviembre de 2026)

### Qué significa "provisional"

- **Ninguna cifra está validada.** Las variables están en estado "Pendiente" y los resultados en "Provisional".
- Con la regla estricta de la Matriz (solo calcula con datos "Validado"), **todo saldría N/D**. Por eso se calculó con la regla "preliminar" (calcular con datos "Pendiente") y se rotuló cada resultado como provisional.
- Además, muchos resultados dependen de **decisiones metodológicas que no tomó el equipo** sino que corresponden a los profesores.

### Qué es un escenario y qué es un supuesto por defecto

- Cuando la Matriz deja una pregunta abierta (por ejemplo, "¿qué cuenta como evento de rendición de cuentas?"), en vez de decidir por los profesores se calcularon **todas las respuestas razonables**. Cada combinación de respuestas es un **escenario**.
- Una de esas respuestas se marcó como **valor por defecto**: el supuesto de trabajo que se usa mientras los profesores deciden. Todos los valores por defecto están rotulados "PROVISIONAL".
- Las preguntas y sus opciones están en el archivo **`config/supuestos.yaml`** (un archivo de texto de configuración): **36 parámetros**, cada uno con sus opciones y su valor por defecto. La hoja Supuestos de la base tiene esos 36, más 3 sensibilidades adicionales y 2 casos dudosos de la segunda lectura (41 filas).

Tres ejemplos de supuestos por defecto:
1. **Ventana de publicación (control social):** solo cuenta un documento publicado a más tardar el 31 de marzo del año siguiente (regla 17 de la Guía). Alternativa: aceptar todo lo publicado hasta la fecha de consulta.
2. **Proyecto 0553 (costo del proceso de PP):** se excluye del cálculo del EPP. Si se incluye, EPP 2025 pasa de 98,83 a 98,86.
3. **Zona 1:** en PP se marca NE (no tiene proyectos). Alternativa: N/D.

Otros: 2026 se trata como primer semestre provisional (corte 30 de junio de 2026), las metas del PDM se comparan con la meta 2027, el DSI usa 7 metas.

### Resultados por indicador (nivel municipal, escenario por defecto)

| Indicador | Año | Numerador / denominador | Puntaje | Semáforo |
|---|---|---|---|---|
| CS.PCS – Promoción del control social | 2024 | 1 / 1 estrategias | 100 | Verde |
| CS.PCS | 2025 | 1 / 1 | 100 | Verde |
| CS.DSI – Disponibilidad de información de seguimiento | 2024 | 0 / 7 metas | 0 | Rojo |
| CS.DSI | 2025 | 0 / 7 | 0 | Rojo |
| CS.VRC – Veedurías en la rendición de cuentas | 2024 | 3 eventos; presencia desconocida en 1 | N/D | Gris |
| CS.VRC | 2025 | 5 eventos; presencia desconocida | N/D | Gris |
| CS.RPO – Respuesta a observaciones ciudadanas | 2024 | 16 / 17 observaciones | 94,12 | Verde |
| CS.RPO | 2025 | Cuadro publicado incompleto | N/D | Gris |
| CS.ICRCS – Índice de control social | 2024 | Falta VRC | N/D | Gris |
| CS.ICRCS | 2025 | Faltan VRC y RPO | N/D | Gris |
| PC.I2 – Cobertura territorial | 2024 | 13 / 13 zonas | 100 | Verde |
| PC.I2 | 2025 | 12 / 13 | 92,31 | Verde |
| PC.I3 – Oferta formativa | 2024 | No hay programación publicada | N/D | Gris |
| PC.I3 | 2025 | 20 / 24 acciones formativas | 83,33 | Verde |
| PC.I5 – Sesiones de instancias | 2025 | 8 / 8 reuniones (CMSSS y COPACO) | 100 | Verde |
| PDM.CA – Espacio público verde | 2024 | 5,75 observado / meta 3,88 m² por habitante | 100 | Verde |
| PDM.CM – Mortalidad vial | 2024 | 7,25 observado / meta 8 por 100.000 | 100 | Verde |
| PP.EPP – Ejecución del PP | 2024 (aproximado) | $18.357.265.731 / $18.601.236.386 | 98,69 | Verde |
| PP.EPP | 2025 | $20.684.427.720 / $20.928.690.218 | 98,83 | Verde |
| PP.EPP | 2026, primer semestre | $5.702.528.148 / $24.519.738.737 | 23,26 | Sin semáforo |

Cómo se lee un ejemplo: **EPP 2025 = 100 × 20.684.427.720 / 20.928.690.218 = 98,83**. Como 98,83 está entre 90 y 100, el semáforo de PP es Verde.

**EPP por zona (escenario por defecto):** en 2024 y 2025 todas las zonas con puntaje están en Verde, salvo la **zona 10 en 2024: 82,58, Amarillo**, que se explica sobre todo por un cruce aproximado del proyecto 0709. La zona 1 es NE en los tres años.

| Territorio | 2024 (aproximado) | Semáforo 2024 | 2025 | Semáforo 2025 | 2026, primer semestre | Semáforo 2026 |
|---|---|---|---|---|---|---|
| Municipio | 98,69 | Verde | 98,83 | Verde | 23,26 | Sin semáforo |
| Zona 1 | NE | Gris | NE | Gris | NE | Gris |
| Zona 2 | 97,32 | Verde | 97,38 | Verde | 13,61 | Sin semáforo |
| Zona 3 | 99,12 | Verde | 99,83 | Verde | 30,14 | Sin semáforo |
| Zona 4 | 97,48 | Verde | 98,70 | Verde | 17,08 | Sin semáforo |
| Zona 5 | 99,07 | Verde | 99,14 | Verde | 22,54 | Sin semáforo |
| Zona 6 | 99,85 | Verde | 98,70 | Verde | 23,54 | Sin semáforo |
| Zona 7 | 99,99 | Verde | 99,77 | Verde | 30,96 | Sin semáforo |
| Zona 8 | 98,68 | Verde | 99,50 | Verde | 26,11 | Sin semáforo |
| Zona 9 | 99,99 | Verde | 97,32 | Verde | 28,76 | Sin semáforo |
| Zona 10 | 82,58 | Amarillo | 98,40 | Verde | 9,54 | Sin semáforo |
| Zona 11 | 98,33 | Verde | 98,11 | Verde | 13,97 | Sin semáforo |
| Zona 12 | 98,87 | Verde | 99,12 | Verde | 13,29 | Sin semáforo |
| Zona 13 | 97,97 | Verde | 99,82 | Verde | 14,19 | Sin semáforo |

**Escenarios que cambian el resultado (los más importantes):**

| Indicador | Año | Escenario alternativo | Puntaje | Semáforo |
|---|---|---|---|---|
| CS.DSI | 2024 | Aceptar publicaciones hasta la fecha de consulta | 28,57 | Rojo |
| CS.VRC | 2024 | Evento sin mención de veedurías cuenta como ausencia | 66,67 | Amarillo |
| CS.VRC | 2025 | Igual | 20 | Rojo |
| CS.RPO | 2024 | Todas las observaciones, con felicitaciones | 72,73 | Amarillo |
| CS.RPO | 2024 | Todos los cuadros de preguntas publicados | 89,47 | Amarillo |
| CS.RPO | 2024 | Criterio laxo de respuesta de fondo | 100 | Verde |
| CS.RPO | 2025 | Calcular con lo publicado | 50 | Rojo |
| CS.ICRCS | 2024 | VRC con ausencia (0 + 100 + 66,67 + 94,12) / 4 | 65,20 | Amarillo |
| CS.ICRCS | 2024 | Control: consecutivos 545 y 546 (28,57 + 100 + 50 + 94,12) / 4 | 68,17 | Amarillo |
| PC.I2 | 2024 | Solo asambleas de PP | 92,31 | Verde |
| PC.I3 | 2025 | Medir en personas (186 / 15) | Revisar | Sin semáforo |
| PDM.CM | 2024 | Datos de la Policía y población DANE posterior a la pandemia | 99,32 | Amarillo |
| PP.EPP | 2024 | Solo cruces exactos (35 de 51 proyectos) | 99,28 | Verde |
| PP.EPP | 2024 | Sin el rubro de déficit fiscal (calculado aparte, fuera de las opciones configuradas) | 98,64 | Verde |
| PP.EPP | 2025 | Incluir el proyecto 0553 | 98,86 | Verde |

Para el ICRCS se calculó también un **rango de sensibilidad** (el menor y el mayor puntaje posible de cada componente): 55,68 a 73,81 en 2024 (Amarillo en ambos extremos) y 42,50 a 57,14 en 2025 (Rojo a Amarillo). La revisión crítica estimó que una combinación más extrema llega a 43,18 (Rojo) en 2024; esa combinación no está calculada en los archivos. Es un rango, no un valor oficial.

### Variables de los indicadores no viables (construidas, sin puntaje)

| Indicador | Lo que se construyó | Por qué no se calcula |
|---|---|---|
| PC.I1, I4, I6 | Universo exploratorio de procesos: 3 (2024) y 5 (2025). Participación virtual: al menos 1 actividad por año | No hay inventario oficial de procesos ni registro de actividades |
| PDM.CS | Afiliados al régimen subsidiado: 29.794 (dic-2025) y 31.997 (ago-2026) | El método del plan (99,5 %) no se puede reproducir con fuentes públicas |
| PDM.ILE | ILE.A: 15,58 (2024) y 15,40 (2025) con 0 a 6 años; 18,43 y 18,23 con 0 a 5 años. Al menos 1 institución mejora por año | Falta la clasificación ICFES por plantel (componente LS) |
| PDM.CP y PDM.CD | Índice situacional 3,23 (2022); 38 de 40 disciplinas | No son la misma medida que pide la Matriz |
| PP.CEP | Proyectos con entrega exigible: 51 (2024), 61 (2025), 66 (2026). Entrega cumplida 2025: 49 por defecto (45 a 50 según la regla), como autorreporte | No hay actas de recibo ni informes finales |
| PP.CMA | Tabla de clasificación por proyecto (180 filas) | No hay meta de personas por proyecto |
| PP.IREPP | Nada adicional | Depende de CEP y CMA |

### Qué se entregó en la etapa 2 (10 de noviembre)

- `outputs/indicadores/entregables/base_indicadores_OCE.xlsx`: la base de resultados en 13 hojas (se explica en la Parte 4).
- `informe_avance_1` en PDF, Word, LaTeX y paquete para Overleaf (24 páginas).
- `nota_tecnica_participacion.docx` (4 páginas), para el profesor Eber Gutiérrez, sobre el sistema de participación.
- Excel de disponibilidad versión 1.2 (corrige una cifra, ver Parte 2).

## 1.5 Cómo se hizo

### Las bases recolectadas: 2.787 archivos

- **Qué son exactamente:** 2.787 **archivos fuente** (no son 2.787 registros de datos ni 2.787 indicadores). Son documentos oficiales descargados: 2.512 PDF, 118 XLS, 56 XLSX, 44 DOC, 26 CSV, 16 DOCX y algunas presentaciones. Pesan en total unos 7,8 GB.
- Están listados en un **inventario** (`data_raw_IIC/inventario_detallado_IIC.csv`): una fila por archivo, con 11 datos de cada uno (carpeta, nombre, ruta, enlace de origen, página de donde se bajó, título, año, emisor, fecha de publicación y tamaño).
- Se organizan en cinco carpetas: 01 Propuestas y acuerdos (389), 02 Ejecución y presupuesto (1.067), 03 Actas y comités técnicos (754), 04 Datos abiertos y SECOP (6) y 05 Fuentes externas y POT (571).
- Los emisores principales son el DAP (722 archivos), el Concejo Municipal (449), la Oficina de Evaluación y Control (389), la Subdirección de Participación Ciudadana (347), la Secretaría de Hacienda (216), la Contraloría Municipal (162) y la EFPC (78).
- Llegaron en cinco archivos ZIP, que hoy están respaldados en Google Drive (carpeta Bases_Datos_IIC_Envigado). Se verificó que los 2.787 archivos están completos y con el tamaño exacto. **El repositorio no registra quién hizo la descarga original**; el inventario se incorporó el 25 de septiembre de 2026.
- En la etapa 2 se descargaron además, aparte del inventario, archivos del DANE (población), MinSalud (afiliación) y la Policía (homicidios en accidentes de tránsito), en `data_raw_IIC/06_descargas_indicadores/`.

### Búsqueda en fuentes oficiales y el dominio espejo

- El portal `www.envigado.gov.co` **falla por un problema de certificado de seguridad** (TLS, el protocolo que cifra la conexión). Por regla, no se intentó saltarse esa protección.
- Los mismos documentos de la Alcaldía están publicados con la misma ruta en un **dominio espejo**: `alcaldia-municipal-de-envigado-en-antioquia.micolombiadigital.gov.co`. En la versión 1.1 todos los enlaces a documentos se cambiaron a ese espejo: 36 de 36 responden con certificado válido. Las 15 páginas de sección del portal que no tienen espejo quedaron con la nota "puede requerir abrirse desde otro navegador".

### Cálculos en scripts reproducibles

- Un **script** es un programa corto (aquí, en el lenguaje Python) que lee los documentos fuente, cuenta o suma y escribe las tablas de resultados. Se escribió un script por indicador (por ejemplo, `pp_epp.py` para el EPP).
- **Reproducible** significa que cualquiera puede volver a correrlo y obtener exactamente las mismas cifras; si cambia un supuesto, se corre de nuevo y todo se actualiza solo.
- Todo se guarda en un **repositorio privado en GitHub** (un servicio que guarda el historial de un proyecto). Cada **commit** es una "foto" guardada del proyecto, con fecha y descripción; sirve para saber qué cambió y cuándo, y para volver atrás si algo se rompe.

### Agentes Analizador y Ejecutor

El trabajo lo hizo **Claude Code**, un asistente de programación, organizado en dos papeles:
- **Agente Analizador:** diseña el plan (fórmulas, fuentes, supuestos) y hace las revisiones críticas. No escribe el código definitivo.
- **Agente Ejecutor:** escribe y corre los scripts, busca los documentos y arma los entregables siguiendo el plan.

Separar los papeles sirve para que quien revisa no sea quien calculó.

### Verificaciones y por qué dan confianza

| Verificación | Qué se hizo | Resultado | Por qué da confianza |
|---|---|---|---|
| **Comparación con la hoja Resultados de la Matriz** | Se llenó una copia temporal de la Matriz con los datos marcados como "Validado" (solo para que sus fórmulas calcularan), se recalculó con LibreOffice y se comparó con los scripts. La copia se borró | **62 resultados comparados: 48 iguales y 14 diferencias esperadas** | Las fórmulas de los profesores dan lo mismo que los scripts. Las 14 diferencias son todas del PP del primer semestre de 2026 (municipio y sus 13 zonas): la Matriz solo calcula PP en periodo anual y devuelve N/D; en la zona 1, además, el equipo marca NE y la Matriz N/D |
| **Recálculo a mano** (revisión crítica del Analizador, 9 de noviembre) | El Analizador leyó él mismo los documentos fuente, sin usar los scripts, y rehízo 23 cuentas (15 valores de indicador, 5 observaciones del RPO y 3 proyectos de PP 2024) | **23 de 23 coinciden** | Los datos se leyeron bien de los documentos. Única salvedad: la cifra de mortalidad vial (7,25) no se pudo verificar contra Forensis porque el PDF no está en el repositorio |
| **Fórmulas y semáforos** (misma revisión) | Recalculó por programa 975 filas | 951 reproducen exactamente la regla de la Matriz; las otras 24 son una sensibilidad declarada de PC.I3 | Topes, "Revisar", NE y umbrales se aplican como dice la Matriz |
| **Segunda lectura** | Tú clasificaste, sin mirar la clasificación del primer lector, una muestra aleatoria del 20 %: 4 eventos del VRC y 12 observaciones del RPO. Se hizo **sobre las citas y el contexto registrados en la muestra, no releyendo cada documento completo** | **16 de 16 en acuerdo** (VRC 4/4 y RPO 12/12). Ningún resultado cambia. Dos casos con duda pasaron a ser decisiones para los profesores | Dos lectores llegan a la misma clasificación. Límite: como no se releyeron los documentos completos, mide el acuerdo sobre la regla, no la extracción de las citas |
| **Control de calidad de la etapa 1** | Fase 3: el Analizador reabrió 5 hallazgos de alto impacto (EPP 2025, entrega de PP, PCS, PP 2024 y datos personales) e hizo 20 correcciones (6 cambiaron la categoría de una variable). Versión 1.1: revisión de enlaces | EPP 2025 reproducido al peso; 5 errores de presentación corregidos en la etapa 2; 0 datos personales | Las clasificaciones se revisaron con una regla de coherencia escrita y los enlaces funcionan |

La revisión crítica encontró **5 errores, todos de presentación**, que se corrigieron sin cambiar ningún puntaje por defecto: el semáforo de PP 2026 (que salía Rojo y pasó a "sin semáforo"), el nombre del CMSSS, el nombre y tipo en filas del EPP, los rótulos del rango del ICRCS y textos desactualizados de supuestos.

## 1.6 Línea de tiempo de entregables y versiones (según el historial de commits)

| Fecha | Commit | Qué pasó |
|---|---|---|
| 24-sep-2026 | b4ef62c | Se crea el repositorio |
| 25-sep-2026 | 8ff6c4f, cf3496d | Inventario de los archivos fuente y matriz de relaciones del IIC; primer informe en LaTeX para Overleaf |
| 14-oct-2026 | 9aca547 | Perfil de llaves de cruce del IIC y limpieza del código BPIN |
| 14-oct-2026 | 5fed56a, 89f620c | Etapa 1, fases 1 y 2: búsqueda de las 39 variables (PP al final) |
| 14-oct-2026 | ce8a5df | Fases 3 y 4: revisión crítica y **Excel de disponibilidad versión 1.0** |
| 14-oct-2026 | 0efe801 | EPP 2025 reproducible y resumen de la reunión en Word |
| 14-oct-2026 | 0546308, 2ab79df, ca50937 | **Informe de disponibilidad en LaTeX/Overleaf** y Excel pulido para los profesores |
| 14-oct-2026 | f09792c | **Informe de disponibilidad en Word** |
| 9-nov-2026 | 7a29057 | **Disponibilidad versión 1.1**: correcciones del control de calidad (enlaces al espejo), sin cambiar cifras |
| 9-nov-2026 | 9f5483e | Etapa 2: plan, supuestos e infraestructura común |
| 9-nov-2026 | 1c0e78a, e3df303, 9cfdb46, d723c4c | Tandas 1 a 4 de cálculo; la tanda 4 incluye el ICRCS y la revisión crítica |
| 10-nov-2026 | a34719d | **Disponibilidad versión 1.2**: corrige las obligaciones de PP 2026 (la portada dice "Versión 1.2 – 9 de noviembre de 2026") |
| 10-nov-2026 | 3b0086a | **Base de resultados, informe de avance 1** (PDF, Word, LaTeX, Overleaf) y **nota técnica de participación** |
| 10-nov-2026 | 52590ac | Segunda lectura terminada e incorporada al informe de avance y a la base |

---

# Parte 2. Problemas y limitaciones

Para cada problema: qué pasó, por qué importa y cómo se manejó o qué falta.

## 2.1 Problemas de información (lo que no existe o no se puede ver)

1. **No hay inventario oficial de procesos participativos.**
   - *Qué pasó:* nadie publica la lista de procesos participativos "cerrados" ni el vínculo entre cada aporte ciudadano y una decisión formal.
   - *Por qué importa:* es el numerador y el denominador de PC.I1 e I4, que pesan 55 % del ISPC. Sin ellos, el índice de participación no se calcula.
   - *Cómo se manejó:* se armó un universo exploratorio (3 procesos en 2024, 5 en 2025) que corre riesgo de subconteo. Falta la solicitud a la Subdirección de Participación.

2. **Universo de eventos de rendición de cuentas.**
   - *Qué pasó:* no hay un balance que diga cuántos eventos de rendición hubo. Según la definición, salen de 1 a 8 por año.
   - *Por qué importa:* es el denominador del VRC.
   - *Cómo se manejó:* por defecto, audiencia del Alcalde más rendiciones sectoriales con acta (3 eventos en 2024 y 5 en 2025). La definición final es de los profesores.

3. **Presencia de veedurías desconocida.**
   - *Qué pasó:* de la audiencia del Alcalde solo se publicó el cuadro de preguntas, sin acta ni lista de asistentes.
   - *Por qué importa:* la Matriz dice que con presencia desconocida el VRC es N/D, y sin VRC el ICRCS también es N/D.
   - *Cómo se manejó:* se dejó N/D. Se calcularon escenarios (66,67 en 2024 si el silencio se toma como ausencia), pero todos exigen apartarse de la regla. Además, en una de las actas la presencia se acredita con un párrafo de plantilla que se repite en otras actas.

4. **Ventana de publicación del DSI.**
   - *Qué pasó:* con la regla de la Guía (documentos publicados hasta el 31 de marzo del año siguiente), ninguna de las 7 metas tenía cifra a tiempo. El POT se publicó el 24 de octubre de 2025; Forensis no tiene fecha de publicación verificable.
   - *Por qué importa:* DSI da 0 (Rojo) en 2024 y 2025 y arrastra al ICRCS.
   - *Cómo se manejó:* con todo lo publicado hasta hoy, 2024 da 28,57 (sigue Rojo). Si Forensis 2024 hubiera salido antes del límite, sería 14,3. Las evaluaciones del PDM de la Alcaldía solo publican indicadores de producto, no de resultado.

5. **Metas solo a 2027, sin metas anuales.**
   - *Qué pasó:* las 7 metas del PDM coinciden en el PDM, el documento técnico de soporte (DTS) del 22-dic-2025 y el Apéndice 9.1.2, pero todas son de 2027.
   - *Por qué importa:* comparar 2024 con la meta de 2027 es "logro al corte", no cumplimiento del año. Si se exigieran metas anuales, todo el PDM quedaría N/D.
   - *Cómo se manejó:* por defecto se compara con la meta 2027, como permite la regla 9 de la Guía.

6. **Fuentes con otra definición.**
   - *TerriData* publica "cobertura del régimen subsidiado" como afiliados sobre población total (cerca de 12 %), no con el método del plan (99,5 %). La Matriz prohíbe sustituir.
   - *POT:* da 5,75 m² de espacio público verde urbano por habitante en 2024; la línea base del plan es 3,88 y el Área Metropolitana daba 4,66 en 2022. No se reconcilian. Se usa el POT rotulado "serie distinta".
   - *Forensis frente a Policía:* Medicina Legal cuenta muertes por transporte (18 en 2024; tasa 7,25); la Policía cuenta "homicidio en accidente de tránsito" (20 en 2024). Con Policía y población DANE posterior a la pandemia, el CM 2024 pasa de 100 (Verde) a 99,32 (Amarillo).

7. **Forensis 2024 pendiente.** La cifra 7,25 es citada (p. 440): el PDF no está en el repositorio porque el sitio de Medicina Legal responde "403" (rechaza descargas automáticas). Debes descargarlo a mano y subirlo a `data_raw_IIC/06_descargas_indicadores/inmlcf/`. Falta además confirmar si cuenta por lugar del hecho o por residencia.

8. **Actas en SharePoint sin acceso.** Las evidencias de las sesiones de COPACO y CMSSS de 2024 son carpetas de SharePoint (el sistema interno de archivos de la Alcaldía) privadas. Por eso PC.I5 solo se calcula para 2025.

9. **Portales que no abren.** Además del certificado de `www.envigado.gov.co`, fallaron `mi.envigado.gov.co`, el sitio de la Agencia Nacional de Seguridad Vial (ANSV), el tablero histórico de MinSalud, la descarga del ICFES por plantel y la sección Observatorio del sitio de la EFPC (error 404, "página no encontrada").

## 2.2 Problemas del presupuesto participativo

10. **No hay marca oficial de qué proyectos son de PP.** Ninguna fuente (plan de acción, Hacienda, CUIPO) dice "este proyecto es de PP". La lista se **reconstruyó** cruzando el listado de la EFPC con los proyectos que dicen "ZONA n" en el plan de acción. En 2025 los 61 proyectos cruzan uno a uno y 56 coinciden al peso. Falta que Planeación confirme la lista.

11. **2024 es aproximado y tiene cruces débiles.** No hay plan de acción de 2024 por proyecto, así que se cruzó el listado con los rubros de Hacienda: 35 coincidencias exactas, 15 aproximadas y 1 sin correspondencia (zona 9, $500.398.021, que queda fuera). Hay tres cruces débiles por confirmar con Planeación:
    - **13-2 ↔ 0714:** la apropiación de Hacienda supera al listado en 60,8 % ($685.162.170 frente a $425.998.943).
    - **12-3 ↔ 0733:** coincidencia de nombre baja (0,5) y 17,2 % menos que el listado.
    - **10-2 ↔ 0709:** 32,5 % más que el listado. Es la causa principal del **Amarillo de la zona 10 (82,58)**.

12. **Rubro 2.3.7.06 de déficit fiscal.** Siete proyectos de 2024 incluyen **$627.204.058** del rubro "Disminución de pasivos – Financiación de déficit fiscal" (pagos de vigencias anteriores, ejecutados al 100 %). Probablemente no son PP de la vigencia. **Sin ese rubro, EPP 2024 = 98,64** (en lugar de 98,69). No se excluyó por defecto; los proyectos quedan marcados "posible no PP". Ese escenario se calculó aparte y no está entre las opciones del archivo de supuestos.

13. **51 frente a 50 proyectos en 2024.** El listado de la EFPC 2024 tiene 51 registros. La etapa 1 contaba 50 porque omitía uno de la zona 8 cuyo valor viene sin el signo "$" en el PDF. Ese valor coincide al peso con Hacienda (proyecto 0726), así que se incluyó como registro dudoso.

14. **Las 55 entregas cumplidas no se reproducen.** En la etapa 1 se citó que con la regla "última fila por indicador" 55 de 61 proyectos tenían la entrega cumplida en 2025. En la etapa 2, con cerca de 40 definiciones probadas, el resultado va de 45 a 50 (49 por defecto). El Excel de disponibilidad v1.2 todavía dice "entre 45 y 55".

15. **Cifra de obligaciones de PP 2026 errónea, corregida en la v1.2.** La v1.1 decía $11.555.041.247; la cifra correcta es **$5.702.528.148** (Hacienda al 30 de junio de 2026 y la hoja ROCA del plan de acción coinciden al peso). Además, los encabezados de la hoja del plan de acción dicen "01-04-2026", pero las cifras son del 30 de junio.

16. **El Excel de disponibilidad conserva cifras de 2024 superadas.** La v1.2 solo corrigió la obligación de 2026. Para 2024 sigue diciendo "47 de 50 proyectos", OPP $17.199.119.162 y PE = 50, cifras que la etapa 2 reemplazó por 50 de 51 proyectos, OPP $18.357.265.731 y PE = 51. Si te preguntan, la cifra vigente es la de la etapa 2.

17. **PP 2026 no es anual.** La Matriz solo prevé el PP en periodo anual. El 23,26 compara medio año de ejecución con una apropiación anual, así que se muestra sin semáforo.

## 2.3 Problemas de calidad y presentación

18. **Datos personales en los cuadros de preguntas.** Los documentos originales tienen nombres, teléfonos y correos de ciudadanos. Las tablas del repositorio no los reproducen: las citas dicen "[ciudadano/a]" o "[funcionario/a]". Un archivo cuyo nombre contenía el nombre de una persona se citó como "[archivo con nombre de persona omitido]".

19. **El nombre de la EFPC no está unificado.** En el repositorio aparece de tres formas:
    - "Escuela de Formación para la Participación Ciudadana": así la registra el inventario (como emisor de 78 archivos, adscrita a la Subdirección de Participación Ciudadana del DAP), el Excel de disponibilidad y la base de resultados.
    - "Escuela de Formación Ciudadana": en el informe de disponibilidad (LaTeX y Word).
    - "Escuela de Formación Política y Ciudadana": en los avances 2 y 3 y en la nota técnica de participación.
    - Recomendación: usar la del inventario y confirmar con el profesor.

20. **Nombres de indicadores distintos a los de la Matriz.** En los archivos de resultados el ICRCS se llama "Índice de control y rendición social" (la Matriz dice "Índice de Condiciones y Respuesta para el Control Social"), y en el resumen de cálculos y la hoja Léeme de la base el DSI aparece como "Divulgación de la situación" (la Matriz dice "Disponibilidad de información de seguimiento"). Para el tablero conviene usar los nombres de la Matriz.

21. **CMSSS mal nombrado, ya corregido.** En entregas previas se llamó "Comité Municipal de Participación Social en Salud". El nombre correcto, según el acta de la primera rendición de Salud 2024, es **Consejo Municipal de Seguridad Social en Salud**.

22. **65 frente a 170 personas capacitadas.** El plan indicativo dice 65 en 2025 y el informe de gestión dice 170. La nota técnica lo explica: 65 es lo ejecutado en 2025 y 170 coincide con el acumulado 2024-2025 (105 + 65). Es una lectura aritmética que el informe no declara y que debe confirmar la Subdirección.

23. **Asignaciones y reconstrucciones inferidas.**
    - PC.I2 2025: la fuente dice "12 asambleas… cubriendo la totalidad de las zonas" sin listarlas; se asignaron a las zonas 2 a 13. Si "totalidad" son 13 zonas, sería 13/13.
    - PC.I3 2025: el denominador se reconstruyó como culminadas más no ejecutadas, y se excluyó el diplomado en Gestión Pública por ser para servidores, algo que la fuente no acredita.

24. **Reparos del DAP en PCS.** La fila del plan indicativo trae observaciones del DAP ("no concuerda la información de ejecución anual…" y evidencias "Sí, pero con observaciones"). El resultado coincide con lo reportado, pero hay que atenderlo antes de validar.

25. **Fechas contradictorias entre documentos.** Audiencia del Alcalde 2024: 20 o 26 de noviembre. Rendición de PP de la vigencia 2024: 12 de enero o 12 de febrero de 2025. No cambian los resultados por defecto.

26. **Inconsistencias de la Matriz misma.** Las páginas del PDM que cita la Matriz ("copia del Concejo") no corresponden a los documentos; el programa de mortalidad vial probablemente es el 3.2.3 y no el 3.2.2; "CS" es a la vez el sistema de control social y el indicador de salud del PDM; "MS" es a la vez "metas seleccionadas" y "meta Saber 11". La propuesta del OCE habla de 12 zonas, pero el PDM (p. 558) establece 13 zonas con 12 comités zonales (la zona 1 no tiene comité).

27. **Huecos de trazabilidad.** En PP, la columna "fuente" trae una ruta de archivo en vez del nombre de la publicación; la fecha de publicación está vacía en muchas variables; algunos conteos (PCS, las 13 zonas) no tienen fila de soporte. No cambian cálculos, pero hay que arreglarlos para la base final.

---

# Parte 3. Qué variables quedan: propuesta para la reunión

> **Importante:** esta clasificación es una **propuesta razonada** para discutir. La decisión es del profesor.

## 3.1 Cómo se clasificó

- **A. Lista para la base de datos:** ya tiene dato y fuente; no depende de una decisión que cambie el resultado.
- **B. Lista si los profesores deciden algo:** el dato existe, pero falta una decisión concreta. Se indica la opción recomendada.
- **C. Depende de una solicitud a la Alcaldía:** el dato no se puede obtener con lo público; se indica la dependencia y qué pedir (según la hoja Solicitudes del Excel de disponibilidad).
- **D. Propuesta de ajustar, reemplazar o eliminar:** aun con una solicitud, la variable no se sostiene en esta versión.

**Regla de desempate:** cuando una variable depende de una solicitud y además de una decisión, se clasificó por lo que la bloquea primero. Por ejemplo, PDM.CP.O queda en C porque primero hay que preguntar al DAP si existe la ficha; si no existe, pasaría a D.

Las 41 filas son las **39 variables de la hoja Registro** más las **2 variables de la hoja PP_Proyectos** (personas atendidas, PA, y meta de atención, MP). La tabla también está en `outputs/guia/trabajo/propuesta_variables.csv`.

## 3.2 Tabla de variables

| ID | Nombre | Indicador | Cat. | Dato actual | Fuente | Decisión, solicitud o ajuste | Recomendación | ¿Tablero? |
|---|---|---|---|---|---|---|---|---|
| CS.DSI.MD | Número de metas seleccionadas con valor observado publicado para la vigencia | CS.DSI | **B** | 0 (2024) y 0 (2025); 2 de 7 si se acepta todo lo publicado hasta hoy | Soporte meta por meta (POT, Forensis, MinSalud, evaluaciones del PDM) | Decidir: lista de metas (7 del ICPD-6 o 10), ventana de publicación (31 de marzo del año siguiente o hasta la fecha de consulta) y qué fuente cuenta como valor publicado | Aprobar 7 metas, ventana del 31 de marzo (es la regla 17 de la Guía) y fuente oficial con la misma definición. Pedir además al DAP los valores observados de resultado 2024-2025 | Sí |
| CS.DSI.MS | Número total de metas incluidas previamente en la selección de seguimiento | CS.DSI | **B** | 7 metas (2024 y 2025) | Apéndice 9.1.2 del Plan Indicativo del PDM | Decidir la lista de metas que forman el denominador | Aprobar las 7 metas del ICPD-6 y fijarlas para todo el periodo; la Matriz sugiere unas 10, pero las 3 adicionales no tienen evidencia | Sí |
| CS.PCS.EE | Número de estrategias de promoción del control social ejecutadas y reportadas en el periodo | CS.PCS | **A** | 1 (2024), 1 (2025); 0 al 30-jun-2026 | Plan Indicativo de Productos, evaluación definitiva 2T-2026 (DAP), consecutivo 544 | Ninguna para el cálculo. Anotar que es autorreporte con observaciones del DAP | Dejar como está; pedir los soportes a la Subdirección de Participación cuando se busque validar | Sí |
| CS.PCS.EP | Número de estrategias de promoción del control social programadas oficialmente para el mismo periodo | CS.PCS | **A** | 1 (2024), 1 (2025), 1 (2026) | Plan Indicativo de Productos, evaluación definitiva 2T-2026 (DAP), consecutivo 544 | Ninguna | Dejar como está | Sí |
| CS.VRC.EV | Número de eventos municipales de rendición de cuentas realizados con presencia documentada de veedurías | CS.VRC | **C** | Sin valor: presencia de veedurías desconocida (audiencia del Alcalde) | Actas de rendición de Salud; cuadro de preguntas del Alcalde | Solicitud (hoja Solicitudes): actas, memorias y listas de asistencia (solo conteos) a la Subdirección de Participación y a la Secretaría de Salud. También hay una decisión: cómo tratar el evento sin mención de veedurías | Mantener N/D (regla de la Matriz) y enviar la solicitud. No contar el evento sin mención como ausencia | No |
| CS.VRC.ET | Número total de eventos municipales de rendición de cuentas realizados en el periodo | CS.VRC | **B** | 3 eventos (2024), 5 (2025) | Actas, memorias y cuadros de rendición; Plan Indicativo de Productos, evaluación definitiva 2T-2026 (DAP), consecutivos 545 y 546 | Decidir qué es un evento municipal de rendición de cuentas (solo Alcalde; Alcalde y sectoriales con acta; todo documento) | Aprobar el valor por defecto: audiencia del Alcalde y rendiciones sectoriales con acta o memoria publicada | No |
| CS.RPO.OR | Número de observaciones ciudadanas del universo evaluado con respuesta de fondo publicada al corte | CS.RPO | **B** | 16 (2024); 2025 sin dato (cuadro incompleto) | Preguntas y recomendaciones, audiencia pública de rendición de cuentas del Alcalde 2024 | Decidir qué ejercicios cuentan, si se excluyen felicitaciones y qué es respuesta de fondo. Para 2025, solicitud de la matriz completa de preguntas (DAP y Comunicaciones) | Aprobar el valor por defecto (Alcalde, sin felicitaciones, criterio estricto); 2025 en N/D hasta recibir la matriz completa | Sí |
| CS.RPO.OT | Número total de observaciones ciudadanas sobre el PDM registradas públicamente en ejercicios cerrados de rendición de cuentas del periodo | CS.RPO | **B** | 17 (2024); 2025 sin dato | Igual que CS.RPO.OR | Las mismas decisiones de CS.RPO.OR | Igual que CS.RPO.OR | Sí |
| PC.I1.V1 | Número de procesos participativos cerrados que incorporaron al menos un aporte ciudadano en una decisión formal | PC.I1 | **C** | Sin dato (no se hizo el vínculo aporte-decisión) | Informe de gestión por dependencias 2025 y otros (narrativo) | Solicitud a la Subdirección de Participación: inventario de procesos participativos cerrados con el acto donde se incorporó el aporte | Dejar fuera del tablero por ahora y enviar la solicitud | No |
| PC.I1.V2 | Número total de procesos participativos cerrados en los que se recibieron aportes ciudadanos | PC.I1 | **C** | 3 (2024) y 5 (2025), universo exploratorio no oficial | Informe de gestión 2025 y documentos de PP, PDM, POT y rendición | Misma solicitud de inventario. Alternativa: que los profesores acepten el universo exploratorio | Dejar fuera hasta tener el inventario; el universo exploratorio corre riesgo de subconteo | No |
| PC.I2.V1 | Número de zonas con al menos una actividad participativa realizada para su comunidad | PC.I2 | **B** | 13 zonas (2024), 12 (2025) | Informe de gestión 2025 p. 92; anteproyecto del PDM (2024); asambleas del POT (2025) | Decidir qué actividades cubren una zona (solo PP o también PDM y POT) y qué evidencia basta (reporte o acta por zona) | Aprobar el valor por defecto (asambleas de PP, PDM y POT; basta el reporte oficial) y pedir actas por zona para validar | Sí |
| PC.I2.V2 | Número total de zonas oficiales del municipio | PC.I2 | **A** | 13 zonas | PDM p. 558 | Confirmar 13 zonas (y no 12 con comité) | Aprobar 13; pedir el texto del Acuerdo 007 de 2025 | Sí |
| PC.I3.V1 | Número de acciones formativas programadas para la ciudadanía que se realizaron | PC.I3 | **B** | 7 (2024), 20 (2025) | Informe de gestión por dependencias 2025 pp. 87-88 | Decidir la unidad (acciones formativas o personas), si se excluye la formación para servidores y si cuentan los semilleros | Aprobar acciones formativas, excluir lo exclusivo para servidores y no contar semilleros (como dice la Matriz) | Sí |
| PC.I3.V2 | Número total de acciones formativas programadas para la ciudadanía con finalización prevista en el periodo | PC.I3 | **B** | Sin dato (2024); 24 (2025, reconstruido) | Informe de gestión 2025 pp. 87-88 (culminadas más no ejecutadas) | Mismas decisiones de PC.I3.V1. Para 2024, solicitud a la EFPC del plan de formación con fechas | Aprobar; 2024 queda N/D hasta recibir la programación | Sí |
| PC.I4.V1 | Número de procesos participativos cerrados con resultados publicados y consultables | PC.I4 | **C** | Sin dato (no auditado documento por documento) | Documentos de resultados de los procesos (exploratorio) | Misma solicitud de inventario de PC.I1, con enlace público por proceso | Dejar fuera por ahora | No |
| PC.I4.V2 | Número total de procesos participativos cerrados | PC.I4 | **C** | 3 (2024) y 5 (2025), universo exploratorio | Igual que PC.I1.V2 | Misma solicitud de inventario | Dejar fuera por ahora | No |
| PC.I5.V1 | Número de sesiones previstas de las instancias seleccionadas que se realizaron | PC.I5 | **B** | 8 sesiones realizadas (2025); 2024 sin dato | Seguimiento al Plan de Acción de la Política de Participación Social en Salud (PPSS) 2025 | Decidir qué instancias se miden (solo COPACO y CMSSS, o también CMPC, CTP y comités zonales) | Aprobar el indicador acotado a COPACO y CMSSS para 2025, rotulado así, y pedir calendarios y actas del CMPC y el CTP | Sí |
| PC.I5.V2 | Número total de sesiones previstas de esas mismas instancias para el periodo | PC.I5 | **B** | 8 sesiones previstas (2025) | PPSS 2025 | La misma decisión de PC.I5.V1 | Igual que PC.I5.V1 | Sí |
| PC.I6.V1 | Número de actividades participativas realizadas que habilitaron un canal virtual para intervenir | PC.I6 | **D** | Al menos 1 actividad por año (2024 y 2025), cota inferior | Rendición del Alcalde 2024; votación de PP en plataforma web (2025) | Ajuste: no existe un registro de actividades realizadas; el indicador es complementario y no entra a ningún índice | Sacarlo del tablero en esta versión y pedir a la Subdirección y a la EFPC un registro de actividades con modalidad desde 2026 | No |
| PC.I6.V2 | Número total de actividades participativas realizadas | PC.I6 | **D** | Sin dato (universo sin definir) | Cifras parciales (asambleas, Semana de la Participación) | Igual que PC.I6.V1 | Igual que PC.I6.V1 | No |
| PDM.CS.O | Cobertura porcentual observada del aseguramiento en el régimen subsidiado | PDM.CS | **C** | Sin dato con el método del plan (proxy: 29.794 afiliados al régimen subsidiado, dic-2025) | MinSalud, cifras de aseguramiento (solo proxy) | Solicitud a la Secretaría de Salud, Dirección de Aseguramiento: cobertura del régimen subsidiado con el método del plan (fórmula, numerador, denominador, fuente) | Dejar fuera y no sustituir con TerriData (cerca de 12 %, otra definición) | No |
| PDM.CS.M | Cobertura meta del aseguramiento en el régimen subsidiado | PDM.CS | **A** | 99,5 (meta 2027) | PDM 2024-2027 (Acuerdo 017 de 2024), p. 602 | Ninguna para la variable (el indicador sale por falta del observado) | Conservar en la base | No |
| PDM.CM.O | Tasa anual observada de mortalidad vial por 100.000 habitantes | PDM.CM | **A** | 7,25 muertes por 100.000 hab. (2024), cifra citada | INMLCF, Forensis 2024, p. 440 (PDF por verificar) | Confirmar Medicina Legal como fuente oficial. Pendiente: subir el PDF de Forensis 2024 y verificar la cifra | Aprobar Medicina Legal; verificar el PDF antes del viernes | Sí |
| PDM.CM.M | Tasa meta de mortalidad vial por 100.000 habitantes | PDM.CM | **A** | 8 (meta 2027) | PDM 2024-2027 (Acuerdo 017 de 2024), p. 706 | Decidir el horizonte: meta 2027 (no hay metas anuales) | Aprobar meta 2027, leída como "logro al corte" (regla 9 de la Guía) | Sí |
| PDM.CA.O | Espacio público verde observado, en metros cuadrados por habitante | PDM.CA | **B** | 5,75 m² por habitante (2024) | Documento diagnóstico de la revisión del POT, espacio público (2025), tabla 39 | Decidir si se acepta la cifra del POT, que es una serie distinta de la línea base (3,88) | Aceptarla rotulada como "serie distinta" y pedir a Medio Ambiente la ficha del indicador | Sí |
| PDM.CA.M | Meta de espacio público verde, en metros cuadrados por habitante | PDM.CA | **A** | 3,88 (meta 2027) | PDM 2024-2027 (Acuerdo 017 de 2024), p. 717 | Horizonte de la meta (igual que PDM.CM.M) | Aprobar meta 2027 | Sí |
| PDM.CP.O | Porcentaje observado de implementación de procesos de participación ciudadana, según el método oficial | PDM.CP | **C** | Sin dato ni método oficial | Ninguna | Solicitud a la Subdirección de Participación (DAP): ficha técnica y valor del indicador "Implementación de los procesos de participación ciudadana" | Dejar fuera. Si el DAP confirma que no existe la ficha, pasar a D y redefinirlo | No |
| PDM.CP.M | Porcentaje meta de implementación de esos procesos | PDM.CP | **A** | 100 (meta 2027) | PDM 2024-2027 (Acuerdo 017 de 2024), p. 742 | Ninguna para la variable | Conservar en la base | No |
| PDM.CD.O | Cobertura porcentual observada de los servicios deportivos incluidos en el resultado | PDM.CD | **C** | Sin dato (proxy: 38 de 40 disciplinas en 2025) | Informe de gestión 2025 (solo proxy) | Solicitud al Índer: cobertura de servicios deportivos con el método de la línea base (5,3 %) | Dejar fuera; la Matriz prohíbe convertir disciplinas en cobertura | No |
| PDM.CD.M | Cobertura porcentual meta de esos servicios deportivos | PDM.CD | **A** | 6 (meta 2027) | PDM 2024-2027 (Acuerdo 017 de 2024), p. 609 | Ninguna para la variable | Conservar en la base | No |
| PDM.ILE.A | Cobertura observada de atención integral de niños y niñas de 0 a 6 años | PDM.ILE | **B** | 15,58 (2024), 15,40 (2025) con niños de 0 a 6 años | Balance del PDM 2024 e informe de gestión 2025 (Educación) y población DANE | Decidir la edad del denominador (0 a 6 años, como la Matriz, o 0 a 5, como el informe). También solicitud a Educación del dato oficial | Aprobar 0 a 6 años; el índice igual queda fuera mientras falte PDM.ILE.S | No |
| PDM.ILE.MA | Meta de cobertura de atención integral inicial | PDM.ILE | **A** | 19 (meta 2027) | PDM 2024-2027 (Acuerdo 017 de 2024), p. 650 | Ninguna para la variable | Conservar en la base | No |
| PDM.ILE.S | Porcentaje observado de instituciones oficiales que mejoran su categoría en Saber 11 | PDM.ILE | **C** | Sin dato (al menos 1 institución mejora por año, cota inferior) | Sin clasificación ICFES por plantel | Solicitud a la Secretaría de Educación: listado de instituciones oficiales con su categoría ICFES 2023, 2024 y 2025 | Dejar fuera y enviar la solicitud | No |
| PDM.ILE.MS | Meta de instituciones oficiales que mejoran su categoría en Saber 11 | PDM.ILE | **A** | 14 (meta 2027) | PDM 2024-2027 (Acuerdo 017 de 2024), p. 652 | Ninguna para la variable | Conservar en la base | No |
| PP.EPP.OPP | Valor de las obligaciones presupuestales registradas con cargo a recursos de presupuesto participativo | PP.EPP | **A** | 2025: $20.684.427.720; 2024 aprox.: $18.357.265.731; 2026 S1: $5.702.528.148 | Plan de acción IV trimestre 2025 (DAP) cruzado con el listado de la EFPC; Hacienda (2024 y 2026) | 2025 lista. Para 2024: aceptar cruces aproximados y decidir sobre el rubro de déficit fiscal. Para todos: proyecto 0553. Solicitud a Planeación: lista oficial de proyectos de PP | Aprobar 2025; 2024 rotulado "aproximado"; 2026 sin semáforo; excluir el 0553 | Sí |
| PP.EPP.APP | Valor de la apropiación definitiva de recursos de presupuesto participativo | PP.EPP | **A** | 2025: $20.928.690.218; 2024 aprox.: $18.601.236.386; 2026: $24.519.738.737 | Igual que PP.EPP.OPP | Las mismas de PP.EPP.OPP | Igual que PP.EPP.OPP | Sí |
| PP.CEP.PE | Número total de proyectos financiados con entrega exigible en la vigencia | PP.CEP | **B** | 51 (2024), 61 (2025), 66 (2026) proyectos | Listados de proyectos de PP de la EFPC 2024 y 2025; plan de acción 2026 | Decidir qué es "entrega exigible en la vigencia". Solicitud del listado oficial 2026 | Aprobar "todo proyecto financiado en la vigencia"; el indicador igual queda fuera mientras falte PP.CEP.PEC | No |
| PP.CEP.PEC | Número de proyectos con todas las entregas exigibles completas y verificadas al corte | PP.CEP | **C** | 49 de 61 (2025) como autorreporte (45 a 50 según la regla); 2024 y 2026 sin dato | Plan de acción IV trimestre 2025, ejecución física neta (autorreporte) | Solicitud a las unidades ejecutoras: actas de recibo, informes finales o de supervisión por proyecto | Dejar fuera y no reemplazar por el autorreporte (la Guía, regla 20, pide no compensar entrega débil) | No |
| PP.CMA.NA | Número de proyectos con meta de atención aplicable en el universo completo y verificado | PP.CMA | **C** | Sin dato | Ninguna | Solicitud a las unidades ejecutoras: fichas con meta anual de personas por proyecto. Además, decidir qué tipos de proyecto aplican | Dejar fuera y enviar la solicitud | No |
| PP.CMA.PA | Personas efectivamente atendidas por proyecto (hoja PP_Proyectos) | PP.CMA | **C** | Sin dato (sin datos personales) | Ninguna | Solicitud: informe de personas únicas atendidas por proyecto | Dejar fuera | No |
| PP.CMA.MP | Meta anual de personas por atender del proyecto (hoja PP_Proyectos) | PP.CMA | **C** | Sin dato | Ninguna | Solicitud: ficha técnica con meta anual de personas | Dejar fuera | No |

**Resumen:** A = 13 · B = 13 · C = 13 · D = 2 (total 41).

## 3.3 Clasificación por indicador e índice

| Grupo | Indicadores | Qué haría falta |
|---|---|---|
| **Entran al tablero ya** (3) | CS.PCS (2024 y 2025), PDM.CM (2024), PP.EPP (2025; 2024 aproximado; 2026 sin semáforo) | Para PDM.CM, verificar el PDF de Forensis 2024. Para PP.EPP, rótulos de "aproximado" y "sin semáforo" |
| **Entran con una decisión** (6) | CS.DSI, CS.RPO, PC.I2, PC.I3, PC.I5, PDM.CA | Aprobar los valores por defecto (o elegir otra opción) en la lista de metas del DSI, el universo y el criterio del RPO, las actividades del I2, la unidad del I3, las instancias del I5 y la serie del POT |
| **Se pueden mostrar como N/D explicado** (2) | CS.VRC, CS.ICRCS | Mantener N/D hasta recibir actas o listas de asistencia. Mostrarlos "en construcción" es opcional |
| **Salen por ahora** (12) | PC.I1, PC.I4, PC.I6, PC.ISPC, PDM.CS, PDM.CP, PDM.CD, PDM.ILE, PDM.ICPD-6, PP.CEP, PP.CMA, PP.IREPP | Solicitudes a la Alcaldía (inventario de procesos, cobertura de salud con el método del plan, ficha de participación, cobertura deportiva, clasificación ICFES, actas de recibo, fichas de meta de personas). PC.I6 se propone sacar de esta versión |

Total: 3 + 6 + 2 + 12 = 23.

## 3.4 Conteo final

**Si se aprueban las recomendaciones, el tablero tendría 9 indicadores con resultado y 18 variables.**

Cómo se cuenta:
- **Los 9 indicadores** son los 3 que entran ya más los 6 que entran con una decisión: CS.PCS, CS.DSI, CS.RPO, PC.I2, PC.I3, PC.I5, PDM.CA, PDM.CM y PP.EPP. Ningún índice tiene resultado: a los cuatro les falta al menos un componente.
- **Las 18 variables** son las dos de cada uno de esos 9 indicadores (numerador y denominador): 8 de la categoría A (CS.PCS.EE, CS.PCS.EP, PC.I2.V2, PDM.CM.O, PDM.CM.M, PDM.CA.M, PP.EPP.OPP y PP.EPP.APP) y 10 de la categoría B (CS.DSI.MD, CS.DSI.MS, CS.RPO.OR, CS.RPO.OT, PC.I2.V1, PC.I3.V1, PC.I3.V2, PC.I5.V1, PC.I5.V2 y PDM.CA.O).
- **No se cuentan:** 5 variables A que son metas de indicadores que salen (PDM.CS.M, PDM.CP.M, PDM.CD.M, PDM.ILE.MA y PDM.ILE.MS), 3 variables B cuyos indicadores siguen bloqueados (CS.VRC.ET, PDM.ILE.A y PP.CEP.PE), las 13 C y las 2 D. Igual quedan en la base para no perder el trabajo.
- **Variante:** si el profesor quiere mostrar VRC e ICRCS como N/D explicado, serían 11 indicadores visibles (9 con resultado) y 20 variables (se suman CS.VRC.EV y CS.VRC.ET).
- Las **variables auxiliares** (por ejemplo, 18 muertes viales o 29.794 afiliados) están en la base como contexto y no se cuentan entre las 18.

---

# Parte 4. La base de datos para el tablero

## 4.1 Qué hay hoy

Existen dos versiones del mismo contenido:
- **El libro de Excel para los profesores:** `outputs/indicadores/entregables/base_indicadores_OCE.xlsx`, con encabezados en español y 13 hojas.
- **Los archivos CSV** en `outputs/indicadores/`. Un **CSV** es una tabla en texto plano que cualquier programa puede leer; es el formato más cómodo para un tablero.

### Hojas del Excel (y su CSV equivalente)

| Hoja | Filas | Qué contiene | CSV equivalente |
|---|---|---|---|
| Léeme | 16 | Avisos, contenido, llaves y convenciones | — |
| Resultados | 947 | Un resultado por indicador × año × periodo × territorio × escenario: entradas, puntaje, semáforo, estado y motivo si es N/D | `resultados_indicadores.csv` (1.003 filas = 947 + 56) |
| No calculados | 56 | Indicadores no viables: por qué no se calculan y qué variables se construyeron | Incluidos en `resultados_indicadores.csv` (columna `no_calculado_por_etapa`) |
| Variables | 3.578 | El valor de cada variable con fuente, enlace, página, fechas, estado y escenario | `valores_variables.csv` |
| PP por zona | 42 | EPP por zona y año, con APP, OPP, porcentaje, semáforo y tres alternativas | Se deriva de `resultados_indicadores.csv` |
| PP por proyecto | 181 | Cada proyecto de PP con apropiación, compromisos, obligaciones y porcentaje; tipo de cruce de 2024 y marca "posible no PP" | `pp_proyectos.csv` + `clasificacion_pp_2024.csv` |
| Soporte de conteos | 611 | Una fila por cada cosa contada (zona, evento, observación, proyecto, meta), con documento y página | `soporte_conteos.csv` |
| Clasificación VRC | 17 | Cada evento de rendición y cómo se clasificó | `clasificacion_vrc.csv` |
| Clasificación RPO | 56 | Cada observación ciudadana y su respuesta (anonimizadas) | `clasificacion_rpo.csv` |
| Entregas PP | 180 | Proyectos con entrega exigible y las tres reglas de "entrega cumplida" | `clasificacion_pp_entregas.csv` |
| Supuestos | 41 | Los 36 parámetros, 3 sensibilidades y 2 casos dudosos, con su pregunta para los profesores | `config/supuestos.yaml` |
| Efecto de alternativas | 200 | Cuánto cambia cada resultado municipal con cada alternativa | — |
| Diccionario de datos | 248 | Qué significa cada columna de cada hoja y cuáles son llaves | `diccionario_datos.csv` |

También existe `matriz_diligenciada.xlsx`: una copia de la Matriz con las hojas Registro y PP_Proyectos llenas (todo en "Pendiente", por eso su hoja Resultados muestra N/D).

## 4.2 Cómo se conectan las tablas

Una **llave** es una columna que permite unir dos tablas: si dos filas tienen el mismo valor en la llave, hablan de lo mismo.

| Llave | Qué es | Dónde aparece |
|---|---|---|
| Código del indicador | Por ejemplo `PP.EPP` | Resultados, No calculados, Variables, Soporte, Efecto de alternativas |
| Código (id) de la variable | Por ejemplo `PP.EPP.OPP` | Variables, Soporte de conteos; en Resultados va en "Variables usadas" (varios separados por ";") |
| Año | 2024, 2025, 2026 | Casi todas |
| Periodo | Anual, Semestre 1, Semestre 2 | Casi todas |
| Territorio | Municipio, Zona 1 … Zona 13 | Resultados, Variables, Soporte, PP por zona. En PP por proyecto la zona es un número (2 a 13) y hay que convertirla a "Zona n" |
| Clave de escenario | "defecto" o la lista de cambios parámetro=opción | Resultados, Variables, Soporte, PP por zona, Efecto de alternativas. El nombre del parámetro es la llave con la hoja Supuestos |
| ID de soporte | Por ejemplo `CS-RPO-2024-001` | Variables ("Identificadores de soporte", varios separados por coma), Soporte de conteos y las hojas de Clasificación VRC y RPO |
| Código del proyecto | 4 dígitos, por ejemplo `0709` | PP por proyecto y Entregas PP (2025 y 2026). En 2024, Entregas PP usa un código del listado (por ejemplo `EFPC2024-Z2-1`) y se une a través de la tabla de cruce de 2024 |

### Esquema simple de tablas y relaciones

```
                         SUPUESTOS (clave del parámetro)
                                   |
                                   | el parámetro va dentro de la "clave de escenario"
                                   v
EFECTO DE ALTERNATIVAS --(indicador, año, periodo, escenario)--> RESULTADOS <--(indicador, año, periodo)-- NO CALCULADOS
                                                                    |
                                                                    | "Variables usadas" + año + periodo + territorio + escenario
                                                                    v
                                                                VARIABLES
                                                                    |
                                                                    | "Identificadores de soporte"
                                                                    v
                                                           SOPORTE DE CONTEOS
                                                             |            |
                                         (ID de soporte = ID del evento)  (ID de soporte = ID de la observación)
                                                             v            v
                                                   CLASIFICACIÓN VRC   CLASIFICACIÓN RPO

RESULTADOS (PP.EPP) --(año, periodo, territorio)--> PP POR ZONA <--(año, zona -> "Zona n")-- PP POR PROYECTO
PP POR PROYECTO --(año, código del proyecto; en 2024 vía cruce del listado)--> ENTREGAS PP
DICCIONARIO DE DATOS: describe todas las columnas (no se une con nada)
```

## 4.3 Formato

- **CSV:** separados por **punto y coma (;)**, codificación **UTF-8 con BOM** (una marca invisible al inicio que hace que Excel muestre bien las tildes). Abrir en Excel con "Datos > Desde texto/CSV".
- **Ojo con los decimales:** en los CSV el separador decimal es el **punto** (98.83288), no la coma. Las fechas van como aaaa-mm-dd.
- **Excel:** el mismo contenido con encabezados en español.
- **Vocabulario del semáforo:** el CSV usa Verde, Amarillo, Rojo y Gris; el Excel distingue además "Sin semáforo" (PP 2026) y "Revisar" (PC.I3 en personas). Para el tablero conviene usar el del Excel.
- Ningún dato está "Validado" ni contiene datos personales.

## 4.4 Cómo se actualiza cuando cambia un supuesto

En palabras simples:
1. Se abre `config/supuestos.yaml` y, en el parámetro que el profesor decidió (por ejemplo, `vrc_evento_sin_mencion`), se cambia el valor por defecto por la opción elegida. La opción tiene que estar entre las que el archivo ya ofrece; si no está (como el rubro de déficit fiscal), primero hay que agregarla.
2. Se vuelven a correr los scripts de los indicadores afectados y luego el que junta todo (`ensamblar.py`).
3. Se vuelve a comparar con la hoja Resultados de la Matriz (`llenar_matriz.py --verificar --zonas`) y se regenera el Excel de la base.

Esto lo hace Claude Code; toma minutos. Los escenarios alternativos se regeneran solos, y el "por defecto" pasa a ser lo que el profesor decidió.

## 4.5 Qué falta para que quede "final" esta semana

1. **Aplicar las decisiones de la reunión** en `config/supuestos.yaml` y recalcular (y, si se decide, agregar la opción de excluir el rubro de déficit fiscal).
2. **Verificar Forensis 2024** cuando subas el PDF.
3. **Unificar nombres** de indicadores con la Matriz y el nombre de la EFPC.
4. **Unificar el vocabulario del semáforo** entre CSV y Excel.
5. **Unificar la columna de disponibilidad**: en 10 de las 39 variables, la categoría de la hoja Variables (por año y escenario) no coincide con la categoría final de la etapa 1.
6. **Corregir la trazabilidad:** nombre de la publicación en lugar de ruta en PP, fechas de publicación y filas de soporte faltantes.
7. **Agregar una hoja "Catálogo"** con código, nombre de la Matriz, sistema, tipo, fórmula y umbrales de cada indicador. Hoy la base no la tiene y el tablero la necesita para pintar los colores.
8. **Preparar una versión liviana para el tablero:** solo el escenario por defecto y los 9 indicadores aprobados, más la hoja de efecto de alternativas. Falta que el profesor diga si el tablero muestra escenarios.
9. **Arreglar las uniones:** zona como número frente a "Zona n" y el código de proyecto de 2024 en Entregas PP.

---

# Parte 5. Preparación de la reunión

## 5.1 Hoja de reunión (una página)

**Objetivo:** salir con la lista aprobada de indicadores y variables del tablero y las decisiones que permiten cerrar la base el viernes 13 de noviembre.

| # | Pregunta exacta | Opciones | Efecto en los resultados | Recomendación (propuesta; decide el profesor) |
|---|---|---|---|---|
| 1 | ¿Qué indicadores entran a la primera versión del tablero? | 9 con resultado (+ VRC e ICRCS como N/D) / solo los 3 sin decisión / todos los 23 | 9 indicadores y 18 variables (11 y 20 con los N/D) | 9 con resultado; VRC e ICRCS como N/D explicado |
| 2 | ¿Se aprueban en bloque los valores por defecto (36 parámetros) como versión 1 provisional, salvo lo que se discuta hoy? | Sí / revisar uno a uno | Si sí, los resultados quedan como en el informe de avance | Sí, rotulados "provisional" |
| 3 | VRC: si el documento de un evento no menciona veedurías, ¿la presencia es desconocida o ausencia? | Desconocida (regla de la Matriz) / ausencia si el documento es integral / ausencia siempre | VRC 2024: N/D o 66,67 Amarillo; ICRCS 2024: N/D o 65,20 Amarillo | Desconocida (N/D) y pedir actas y listas de asistencia |
| 4 | DSI: ¿qué metas forman la lista y hasta cuándo se acepta una publicación? | 7 metas del ICPD-6 / 10 metas; 31 de marzo / hasta hoy | 2024: 0 (Rojo) o 28,57 (Rojo) | 7 metas y 31 de marzo (regla 17) |
| 5 | PP 2024: ¿se excluye el rubro de déficit fiscal ($627.204.058) y se aceptan los cruces aproximados? | Incluir / excluir; completa / solo exactos | 98,69 / 98,64 / 99,28 (Verde en todos); la zona 10 sigue Amarilla | Excluir el rubro, aceptar los cruces rotulados "aproximado" y confirmar 0714, 0733 y 0709 con Planeación |
| 6 | ¿Las metas del PDM se leen contra la meta 2027 y PP 2026 se muestra sin semáforo? | Meta 2027 / metas anuales; 2026 sin semáforo / no mostrar | Con metas anuales, CA y CM pasan a N/D | Meta 2027 como "logro al corte"; 2026 sin semáforo |
| 7 | ¿Quién envía las solicitudes a la Alcaldía y cuáles van primero? | Profesores / Isabella con carta del profesor | Desbloquearían VRC, ICRCS, I1, I4, PDM.CS, CP, CD, ILE, CEP y CMA | Primero: inventario de procesos, actas de rendición, cobertura de salud y lista oficial de proyectos de PP |

**Necesito que me confirmes para cerrar la base esta semana:**
- La lista de indicadores y variables (puntos 1 y 2).
- Si el tablero muestra solo el escenario por defecto o también las alternativas.
- Los nombres oficiales: los de la Matriz para los indicadores y el de la EFPC.
- El formato de entrega a la persona del tablero (Excel, CSV o ambos) y su contacto.

## 5.2 Preguntas que podría hacer el profesor y cómo responderlas

**¿De dónde sale el 98,83?**
Es el EPP de 2025. Se tomaron los 61 proyectos del plan de acción del IV trimestre de 2025 cuyo nombre dice "ZONA n" y se compararon uno a uno con el listado de la EFPC: coinciden en número por zona y 56 coinciden al peso. Se sumaron sus obligaciones y su apropiación: 100 × 20.684.427.720 / 20.928.690.218 = 98,83. No incluye el proyecto 0553 (costo del proceso); con él daría 98,86. Lo reprodujo la revisión crítica al peso. Falta que Planeación confirme la lista de proyectos.

**¿Por qué VRC queda en N/D?**
Porque la Matriz dice que si hay eventos con presencia desconocida de veedurías, el VRC es N/D. De la audiencia del Alcalde solo se publicó el cuadro de preguntas, sin acta ni asistentes, así que no se sabe si hubo veedurías. Si se decidiera contar ese silencio como ausencia, 2024 daría 66,67 (Amarillo) y 2025, 20 (Rojo), pero eso es apartarse de la regla. Lo correcto es pedir las actas o listas de asistencia.

**¿Qué tan confiables son los resultados?**
Son provisionales, no validados. A favor: las fórmulas coinciden con las de la Matriz (48 de 62 resultados iguales y las 14 diferencias explicadas), el Analizador recalculó a mano 23 valores leyendo los documentos y todos coinciden, y la segunda lectura dio 16 de 16 en acuerdo. En contra: muchos datos son autorreporte de la Alcaldía, la mortalidad vial es una cifra citada sin verificar, PP 2024 es aproximado y varios resultados dependen de decisiones que están pendientes. La confiabilidad aritmética es alta; la de las fuentes, la da la Alcaldía cuando valide.

**¿Qué falta para el tablero?**
Que se aprueben la lista de indicadores y los supuestos; aplicar esas decisiones y recalcular; verificar Forensis; unificar nombres y vocabulario; agregar una hoja de catálogo con umbrales; y una versión liviana con el escenario por defecto. Con las decisiones de hoy, se puede tener el viernes.

**¿Por qué DSI da 0?**
Porque de las 7 metas seleccionadas, ninguna tenía una cifra oficial publicada, con la misma definición, antes del 31 de marzo del año siguiente, que es el límite de la regla 17 de la Guía. La Alcaldía solo publica la evaluación de indicadores de producto, no de resultado; el POT salió en octubre de 2025 y Forensis no tiene fecha verificable. Con todo lo publicado hasta hoy, 2024 sube a 28,57, que sigue siendo Rojo. El 0 es en sí un hallazgo: la información de resultados del plan no se publica a tiempo.

**¿Por qué 2026 no tiene semáforo?**
Porque solo hay datos al 30 de junio de 2026 y la Matriz solo prevé el PP en periodo anual. El 23,26 compara seis meses de ejecución con la apropiación de todo el año; pintarlo de Rojo sugeriría un mal desempeño que no se puede afirmar. Se muestra como dato informativo.

**¿Qué es un escenario?**
Es una forma de calcular el indicador con una respuesta distinta a una pregunta que la Matriz deja abierta. Ejemplo: para el RPO, si se cuentan o no las felicitaciones (94,12 o 72,73). En vez de decidir por los profesores, se calcularon todas las opciones y se marcó una como valor por defecto provisional. La hoja "Efecto de alternativas" muestra cuánto cambia cada resultado.

**¿Por qué no se calcularon I1, I4 e I6?**
Porque no existe un inventario oficial de procesos participativos cerrados ni un registro que vincule cada aporte con una decisión (I1 e I4), ni un registro de actividades con canal virtual (I6). Se armó un universo exploratorio (3 procesos en 2024, 5 en 2025), pero corre riesgo de subconteo. Sin I1 e I4, el ISPC queda bloqueado (pesan 55 %). Hay que pedirle el inventario a la Subdirección de Participación.

## 5.3 Plan de trabajo hasta el viernes 13 de noviembre

| Día | Isabella | Claude Code | Depende de los profesores |
|---|---|---|---|
| **Miércoles 11** (reunión) | Lleva la hoja de reunión; anota cada decisión con la opción elegida; pide el contacto de la persona del tablero | Nada antes de la reunión | Las 7 decisiones de la hoja |
| **Miércoles 11** (tarde) | Descarga a mano el PDF de Forensis 2024 y lo sube al repositorio; envía a Claude Code la lista de decisiones | Aplica las decisiones en el archivo de supuestos (agrega la opción de déficit fiscal si se aprobó), recalcula, verifica Forensis y compara con la Matriz | — |
| **Jueves 12** | Revisa los resultados nuevos frente a esta guía; redacta el borrador de las solicitudes a partir de la hoja Solicitudes | Unifica nombres, vocabulario del semáforo y disponibilidad; agrega la hoja Catálogo; arregla trazabilidad y llaves; prepara la versión liviana para el tablero y el diccionario | Firmar y enviar las solicitudes a la Alcaldía |
| **Viernes 13** | Entrega la lista de variables y la base a la persona del tablero y al profesor | Verificación final (matriz, recálculo, datos personales, "Validado"), guarda la versión en GitHub y actualiza la memoria | Visto bueno final |

Las respuestas de la Alcaldía no llegarán esta semana: lo que está en C queda fuera de la primera versión del tablero y entrará cuando llegue la información.

---

# Glosario de siglas

| Sigla | Significado | Nota |
|---|---|---|
| ADRES | Administradora de los Recursos del Sistema General de Seguridad Social en Salud | Significado general; el repositorio solo cita "ADRES – BDUA" |
| AMVA | Área Metropolitana del Valle de Aburrá | — |
| ANSV | Agencia Nacional de Seguridad Vial (Observatorio Nacional de Seguridad Vial) | — |
| APP | Apropiación definitiva de recursos de presupuesto participativo | Denominador del EPP |
| BDUA | Base de Datos Única de Afiliados (salud) | Significado general; no está escrito en el repositorio |
| BOM | Marca de orden de bytes: marca invisible al inicio de un archivo que indica UTF-8 | — |
| BPIN | Banco de Programas y Proyectos de Inversión Nacional (código de un proyecto de inversión) | Significado general; no está escrito en el repositorio |
| CA | Logro del mantenimiento de espacio público verde (PDM.CA) | — |
| CD | Logro de cobertura de servicios deportivos (PDM.CD) | — |
| CEP | Cumplimiento de entrega de proyectos (PP.CEP) | — |
| CM | Logro de la reducción de mortalidad vial (PDM.CM) | — |
| CMA | Cumplimiento de metas de atención (PP.CMA) | — |
| CMPC | Consejo Municipal de Participación Ciudadana | — |
| CMSSS | Consejo Municipal de Seguridad Social en Salud | Antes mal nombrado; corregido |
| COPACO | Comité de Participación Comunitaria en Salud | — |
| CP | Logro de implementación de procesos participativos (PDM.CP) | — |
| CS | Dos significados: el sistema de Control social, y el indicador de salud del PDM (PDM.CS, Logro de la cobertura de aseguramiento subsidiado) | Ambigüedad de la Matriz |
| CSV | Valores separados por comas: tabla en texto plano (aquí con punto y coma) | — |
| CTP | Consejo Territorial de Planeación | — |
| CUIPO | Categoría Única de Información del Presupuesto Ordinario (reporte presupuestal a la Contraloría General) | Significado general; el repositorio solo dice "Contraloría General de la República – CUIPO" |
| DANE | Departamento Administrativo Nacional de Estadística | Significado general; no está escrito en el repositorio |
| DAP | Departamento Administrativo de Planeación (Alcaldía de Envigado) | — |
| DNP | Departamento Nacional de Planeación (publica TerriData) | Significado general |
| DSI | Disponibilidad de información de seguimiento (CS.DSI) | En algunos archivos aparece como "Divulgación de la situación"; el nombre correcto es el de la Matriz |
| DTS | Documento técnico de soporte del PDM | Así se usa en el repositorio |
| EFPC | Escuela de Formación para la Participación Ciudadana | Nombre según el inventario; en otros archivos aparece como "Escuela de Formación Ciudadana" o "Escuela de Formación Política y Ciudadana" |
| EPP | Ejecución presupuestal del presupuesto participativo (PP.EPP) | — |
| EPV | Espacio público verde | — |
| Forensis | Publicación anual de datos de Medicina Legal | No es sigla |
| I1 a I6 | Indicadores de participación ciudadana: I1 Incorporación de aportes, I2 Cobertura territorial, I3 Oferta formativa, I4 Publicación de resultados, I5 Sesiones de participación, I6 Participación virtual | — |
| IBS / ITG | Subíndice de bienestar social / Subíndice de territorio y gobernanza (partes del ICPD-6) | — |
| ICFES | Instituto Colombiano para la Evaluación de la Educación (pruebas Saber 11) | Significado general; no está escrito en el repositorio |
| ICLD | Ingresos corrientes de libre destinación | — |
| ICPD-6 | Índice de cumplimiento de resultados seleccionados del PDM (6 dimensiones) | — |
| ICRCS | Índice de Condiciones y Respuesta para el Control Social | En los archivos de resultados aparece como "Índice de control y rendición social" |
| IIC | Índice de Incidencia Ciudadana | Trabajo previo; no se calcula en estos encargos |
| ILE | Índice de Logro Educativo (PDM.ILE); sus partes son LA (logro de acceso inicial) y LS (logro en Saber 11) | — |
| INMLCF | Instituto Nacional de Medicina Legal y Ciencias Forenses | — |
| IREPP | Índice de Resultados de Ejecución del Presupuesto Participativo | — |
| ISPC | Índice Simplificado de Participación Ciudadana de Envigado | — |
| IUE | Institución Universitaria de Envigado | — |
| JAC | Juntas de Acción Comunal | — |
| MD / MS | Metas con valor publicado / metas seleccionadas (variables del DSI). MS también es la meta de Saber 11 en el ILE | Ambigüedad de la Matriz |
| MGA | Metodología General Ajustada (fichas de proyectos de inversión) | Significado general; no está escrito en el repositorio |
| MinSalud | Ministerio de Salud y Protección Social | — |
| N/D | No disponible: información faltante, no validada o fuera del corte | Sin color |
| NA, PA, MP, SC | Variables del CMA: proyectos aplicables (NA), personas atendidas (PA), meta de personas (MP), suma de puntajes (SC) | — |
| NE | No evaluable: base cero, sin universo | Sin color |
| OCE | Observatorio Ciudadano de Envigado | — |
| ODS | Objetivos de Desarrollo Sostenible | Aparecen en la columna de categoría de la Matriz |
| OPCEN | Observatorio de Participación Ciudadana de Envigado (institucional) | — |
| OPP | Obligaciones presupuestales con cargo a recursos de presupuesto participativo | Numerador del EPP |
| PC | Sistema de Participación ciudadana | — |
| PCS | Cumplimiento de la promoción del control social (CS.PCS) | — |
| PDM | Plan de Desarrollo Municipal 2024-2027 (Acuerdo 017 de 2024) | — |
| PE / PEC | Proyectos con entrega exigible / proyectos con entrega completa y verificada (variables del CEP) | — |
| POT | Plan de Ordenamiento Territorial del Municipio de Envigado | — |
| PP | Presupuesto participativo | — |
| PPSS | Política de Participación Social en Salud | — |
| ROCA | Nombre de una hoja del plan de acción con rubros, fuente, proyecto, BPIN, apropiación y obligación | El significado de la sigla no consta en el repositorio |
| RPO | Respuesta pública a observaciones ciudadanas (CS.RPO) | — |
| RS | Régimen subsidiado de salud | — |
| S1 | Primer semestre | — |
| SECOP | Sistema Electrónico para la Contratación Pública | Significado general; no está escrito en el repositorio |
| SGSSS | Sistema General de Seguridad Social en Salud | — |
| SisMaster | Sistema de información de aseguramiento citado como fuente de la línea base de salud | El significado no consta en el repositorio |
| TerriData | Plataforma de datos territoriales del DNP | No es sigla |
| TLS | Protocolo que cifra la conexión con un sitio web (el "candado" del navegador) | — |
| UTF-8 | Codificación de texto que permite tildes y eñes | — |
| VRC | Presencia de veedurías en la rendición de cuentas (CS.VRC) | — |
| YAML | Formato de archivo de texto para configuración (el de los supuestos) | — |

**Otros términos:** *Script* (programa corto que hace los cálculos), *GitHub* (servicio donde se guarda el historial del proyecto), *commit* (versión guardada con fecha y descripción), *Claude Code* (asistente de programación que ejecutó el trabajo), *Agente Analizador* (planea y revisa), *Agente Ejecutor* (programa y ejecuta), *LaTeX/Overleaf* (sistema para componer documentos formales y su editor en línea), *dominio espejo* (otra dirección web que publica los mismos archivos), *Índer* (entidad de deporte de Envigado; su nombre completo no consta en el repositorio).
