# Propuesta aprobada – sistema de indicadores simplificado (10-nov-2026)

Acordado con el profesor Francisco Correa (reunión) y aprobado por Isabella. Todo PROVISIONAL; nada validado.
Regla general: cada indicador lleva un indicador de respaldo o una variable proxy. Los resultados ya calculados (commit d723c4c, outputs/indicadores/) NO cambian; lo nuevo se marca "provisional".

## Decisiones de la usuaria sobre la propuesta (paso 2)
1. Fidelidad 2024: denominador 51 (proyectos votados) → 39/51 = 76,47. El 39/50 (solo proyectos cruzados con Hacienda) queda como nota.
2. Primera infancia (ILE.A): usar el rango de edad que reproduzca la línea base del PDM (16,3) en su año base. Calcular el año base con 0-5 y 0-6 y quedarse con el que dé 16,3; explicar en una línea por qué. Si ninguno la reproduce, usar 0-5 y anotarlo. (Dato previo: calibración 2023 con TerriData 2.059 atendidos: 0-6 = 14,56; 0-5 = 17,18; verificar año base en el PDM p. 652.)
3. Cobertura de salud: serie con UNA sola proyección de población DANE para todos los años (comparable). Nota sobre la diferencia con el corte de agosto de 2026 (82,59 con la población de MinSalud actualizada).
4. ICPD-6 con promedio parcial: solo como respaldo, rotulado "referencial: no cumple la regla de la Guía de no calcular índices con entradas faltantes".
5. Respaldo de I1: anotar que usa los mismos datos de la fidelidad de PP; para evitar doble conteo entre componentes.
6. Resumen: contar por separado indicadores principales y respaldos (en PP los principales son fidelidad y EPP; los otros 4 son respaldos: versión monetaria de fidelidad, tasa de compromiso, proyectos bien ejecutados, equidad entre zonas, avance físico — revisar: son 5 medidas; la usuaria dijo "los otros 4": tratar fidelidad monetaria como complemento de la fidelidad, no como respaldo aparte).
7. Buscar en TerriData e ICFES las series de Saber 11 y de cobertura de salud.
8. Semáforo de fidelidad: umbrales de PP (Verde ≥ 90, Amarillo 75 a < 90, Naranja 50 a < 75, Rojo < 50), rotulado "propuesta".

## Tablas propuestas (estado al paso 1)

### Control social (CS) – se mantienen los indicadores
| Indicador | Qué mide | Fórmula | Estado hoy | Respaldo o proxy | Fuente |
|---|---|---|---|---|---|
| PCS – Promoción del control social | Estrategias de promoción ejecutadas vs. programadas | 100 × ejecutadas / programadas | 100 (2024 y 2025), Verde | Número de estrategias ejecutadas (1 por año) | Plan Indicativo (DAP) |
| DSI – Disponibilidad de información de seguimiento | Metas del plan con valor publicado a tiempo | 100 × metas con valor / metas seleccionadas | 0 (2024 y 2025), Rojo | Con todo lo publicado hasta hoy: 28,57 (2024), Rojo | Plan Indicativo, informes de gestión, fuentes oficiales |
| VRC – Veedurías en la rendición de cuentas | Eventos de rendición con presencia de veedurías | 100 × eventos con veeduría / eventos | N/D (2024 y 2025) | Evento sin mención = ausencia: 66,67 (2024) y 20 (2025), provisional | Actas y cuadros de rendición |
| RPO – Respuesta a observaciones ciudadanas | Observaciones con respuesta de fondo publicada | 100 × respondidas / observaciones | 94,12 (2024), Verde; N/D (2025) | 2025 con el cuadro publicado: 50, Rojo | Cuadros de preguntas y respuestas |
| ICRCS – Índice de Condiciones y Respuesta para el Control Social | Resumen del componente | (DSI + PCS + VRC + RPO) / 4 | N/D | Con los respaldos de VRC y RPO: 65,20 (2024), Amarillo; 42,50 (2025), Rojo, provisional | Los cuatro anteriores |

### Participación ciudadana (PC) – se eliminan I5 e I6
Nota: el ISPC original ya usa solo I1 a I4 (0,30·I1 + 0,25·I2 + 0,20·I3 + 0,25·I4); eliminar I5 e I6 no cambia sus pesos. Opción para los profesores mientras I1 e I4 no sean viables: redistribuir en proporción → 0,556·I2 + 0,444·I3 (0,25/0,45 y 0,20/0,45).
| Indicador | Qué mide | Fórmula | Estado hoy | Respaldo o proxy | Fuente |
|---|---|---|---|---|---|
| I1 – Incorporación de aportes | Procesos en que un aporte ciudadano entró a una decisión formal | 100 × procesos con aporte incorporado / procesos cerrados | No disponible (sin inventario) | Proyectos votados con presupuesto asignado / proyectos votados: 50/51 (2024), 61/61 (2025), provisional; usa los mismos datos de la fidelidad de PP | Listado EFPC, Hacienda/plan de acción |
| I2 – Cobertura territorial | Zonas con al menos una actividad participativa | 100 × zonas con actividad / 13 | 100 (2024), 92,31 (2025), Verde | Asambleas de PP: 12/13 | Informe de gestión 2025, PDM |
| I3 – Oferta formativa | Acciones formativas ejecutadas vs. programadas | 100 × ejecutadas / programadas | 83,33 (2025), Verde; N/D (2024) | Número de acciones culminadas: 7 (2024), 21 (2025) | Informe de gestión 2025 |
| I4 – Publicación de resultados | Procesos con resultados publicados | 100 × con resultados publicados / procesos cerrados | No disponible | Universo exploratorio de procesos con documento de resultados: 3 (2024), 5 (2025), construible | Portal de la Alcaldía; Subdirección de Participación |
| ISPC – Índice | Resumen | Original 0,30·I1 + 0,25·I2 + 0,20·I3 + 0,25·I4; opción 0,556·I2 + 0,444·I3 | N/D (original) | Opción redistribuida: 88,32 (2025), Verde, provisional; 2024 N/D | I2, I3 |

### Cumplimiento del Plan de Desarrollo (PDM)
| Indicador | Qué mide | Fórmula | Estado hoy | Respaldo o proxy | Fuente |
|---|---|---|---|---|---|
| Cobertura de salud (nuevo; reemplaza aseguramiento subsidiado) | Población afiliada al sistema de salud | 100 × afiliados (contributivo + subsidiado + especiales) / población DANE | 87,65 (dic-2025, MinSalud); 82,59 (ago-2026, cambio de proyección); serie larga construible | Discapacidad o inclusión (p. ej. personas en el registro de discapacidad), por verificar | MinSalud, TerriData |
| CM – Mortalidad vial | Muertes en vías por 100.000 hab. vs. meta | Logro vs. meta (Matriz) | 100 (2024; 7,25 vs. 8), Verde | Policía: 8,05 → 99,32, Amarillo | Forensis (INMLCF), Policía |
| CA – Espacio público verde | m² por habitante vs. meta | Logro vs. meta | 100 (2024; 5,75 vs. 3,88), Verde | EPV total del POT 69,41 (otra serie) | POT |
| CP – Procesos participativos | Implementación oficial de procesos de participación | Logro vs. meta | No disponible | Índice situacional 3,23 (2022) | PDM; ficha al DAP |
| CD – Servicios deportivos | Cobertura deportiva | Logro vs. meta | No disponible | Disciplinas atendidas 38/40 = 95 (2025), provisional | Informe de gestión (Índer) |
| ILE – Logro educativo (evolución) | Avance de la serie hacia la meta 2027 | Serie anual; variación anual; avance = 100 × (valor − línea base)/(meta − línea base), tope 0-100 | Primera infancia: 0-6 años 15,58 (2024), 15,40 (2025) → avance 0 (LB 16,3, meta 19); 0-5 años 18,43, 18,23 → 78,9 y 71,5. Serie 2019-2025 construible | Saber 11: si no hay clasificación de planteles, puntaje promedio Saber 11 del municipio | TerriData, ICFES, Secretaría de Educación |
| ICPD-6 – Índice | Resumen | IBS = (salud + ILE + CD)/3; ITG = (CM + CA + CP)/3; (IBS + ITG)/2 | N/D | Promedio parcial, REFERENCIAL (no cumple la regla de la Guía) | Los anteriores |
| Logros de políticas públicas (nueva dimensión) | Seguimiento de políticas públicas municipales | a) 100 × políticas con informe de seguimiento publicado / políticas vigentes; b) con plan de acción vigente / vigentes; c) metas cumplidas / metas | No disponible (sin inventario oficial); a) construible armando la lista con los acuerdos municipales | a) hoy; b) y c) futuro | Concejo (acuerdos), Planeación |

### Presupuesto participativo (PP)
| Indicador | Qué mide | Fórmula | Estado hoy | Respaldo o proxy | Fuente |
|---|---|---|---|---|---|
| Fidelidad a lo votado (principal) | Proyectos cuya apropiación coincide con lo votado | 100 × proyectos con apropiación = valor votado (±1 %, parámetro) / proyectos votados | 76,47 (2024 aprox.; 39/51), 95,08 (2025; 58/61), provisional; 2026 construible (falta valor votado) | Complemento monetario: apropiación total / valor votado total = 103,55 (2024), 100,74 (2025) | Listado EFPC, Hacienda, plan de acción |
| EPP – Ejecución presupuestal (principal) | Dinero del PP convertido en obligaciones | 100 × OPP / APP | 98,69 (2024 aprox.), 98,83 (2025), Verde; 23,26 (2026 S1) sin semáforo | — | Hacienda, plan de acción |
| Tasa de compromiso (respaldo) | Dinero comprometido | 100 × compromisos / apropiación | 99,56 (2025), provisional; 2024 Hacienda no separa compromisos | — | Igual |
| Proyectos bien ejecutados (respaldo) | Proyectos con ejecución ≥ 90 % | proyectos ≥ 90 % / proyectos | 48/50 (2024), 57/61 (2025) | — | Igual |
| Equidad entre zonas (respaldo) | Brecha de ejecución entre zonas | zona más alta − zona más baja (puntos) | 17,41 (2024: 82,58-99,99), 2,51 (2025: 97,32-99,83) | — | Igual |
| Avance físico (respaldo de entrega) | Proyectos reportados terminados | proyectos con avance físico 100 % / proyectos | 49/61 (2025, autorreporte) | — | Plan de acción 2025 |
CEP, CMA e IREPP salen por ahora.
