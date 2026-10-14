# Notas Fase 2 - Sistema PP (presupuesto participativo)

Corte: 14-oct-2026. Ejecutor PP. No se calcula ningún índice; todos los valores quedan en estado "Pendiente". Las cifras salen de scripts temporales del Ejecutor (no versionados) y deben revisarse antes de usarlas.

## 1. Resultado por variable

| Variable | Fase 1 | Fase 2 | Nivel municipal | Nivel zona |
|---|---|---|---|---|
| PP.EPP.OPP | Construible | **Construible** | Suma de obligaciones de los proyectos PP identificados | Construible zonas 2-13 |
| PP.EPP.APP | Construible | **Construible** | Suma de apropiación definitiva | Construible zonas 2-13 |
| PP.CEP.PEC | Parcial | **Proxy** | Avance físico autorreportado, sin acta de recibo | Proxy zonas 2-13 (2025) |
| PP.CEP.PE | Construible | **Construible** | 50 / 61 / 66 proyectos (2024/2025/2026) | Construible zonas 2-13 |
| PP.CMA.NA | Parcial | **Parcial** | Sin meta de personas verificable | Parcial (2024, INDER y algunos operadores) |

Conteo: Construible 3, Proxy 1, Parcial 1, Disponible 0, No disponible 0.

## 2. Hallazgos clave

1. **Los proyectos PP se pueden identificar en el plan de acción.** Los proyectos con "ZONA n" en el nombre del plan de acción IV trimestre 2025 son 61 y coinciden uno a uno con el listado EFPC 2025 (56 con valor idéntico a la apropiación definitiva; 5 con diferencias por adiciones). Cada uno trae código de proyecto, BPIN (hoja ROCA), apropiación definitiva, compromiso y obligación. Esto permite OPP y APP sin marcador PP en CUIPO.
2. **2026:** el plan de acción II trimestre 2026 trae 66 proyectos zonales con apropiación de $24.519.738.737, que coincide con los "66 proyectos por $24.519 millones" de la Rendición de Cuentas 2025 p. 61 (fuente verificada por el profesor Francisco). Es apropiación del año de ejecución, no ejecución.
3. **Vigencia de ejecución fijada:** el catálogo de ejecución 2024 ("Proyectos ejecutados PP 2024", 50 proyectos) corresponde a la vigencia 2024; "Proyectos en ejecución PP 2025" (61) a 2025; los 66 del plan de acción 2026 a 2026 (corte 30-jun-2026).
4. **Fuente financiera:** en 2025 y 2026 toda la apropiación de esos proyectos es "Ingresos corrientes de libre destinación": no se detecta cofinanciación (la Guía 13 pide excluirla). El proyecto 0553 "Implementación del presupuesto participativo" ($776,96 M apropiados en 2025; $799,37 M en 2026) es el costo del proceso (asambleas) y no se sumó; falta decisión metodológica.
5. **Hacienda:** el PDF de ejecución de egresos no menciona "participativo", pero el código de rubro termina en el código de proyecto (4 dígitos). Con ese sufijo, 58 de 61 proyectos de 2025 coinciden al peso con el plan de acción; 3 difieren por filas que mi lectura no recuperó (sin explicar). También sirve para 2024 (código de rubro con sufijo de proyecto en la última posición; verificado con el proyecto 0230).
6. **CUIPO:** el cruce por los 62 BPIN del ROCA (no por el puente invalidado) da $20.016.692.636 de obligaciones acumuladas a dic-2025, 3 % menos que el ROCA; no se explica. A jun-2026 CUIPO da $6,60 mil millones frente a $11,56 mil millones de Hacienda. CUIPO sirve solo como contraste.
7. **SECOP II:** los 13 números de contrato citados en los PDF de PP 2024 (por ejemplo ENV-04-09-0298-24, de INDER, $4.753.266.676) aparecen en el CSV de SECOP II con estado, liquidación y valor pagado. No hay BPIN ni código de proyecto, y el estado es poco fiable (ENV-07-09-0939-24 con fin de contrato 15-dic-2024 figura "En ejecución"). No se identificaron los contratos de 2025 y 2026.

## 3. Valores encontrados (todos "Pendiente")

### Municipal

| Vigencia | Proyectos | Apropiación definitiva (COP) | Obligaciones (COP) | Corte y fuente |
|---|---|---|---|---|
| 2024 | 50 en catálogo ($18.061.257.369); 47 con correspondencia en Hacienda | 17.332.821.550 | 17.199.119.162 | 31-dic-2024, Hacienda; aprox. (34 exactos: 12.201.643.882 / 12.132.617.344) |
| 2025 | 61 | 20.928.690.218 | 20.684.427.720 | 7-ene-2026, plan de acción/ROCA |
| 2026 | 66 | 24.519.738.737 | 11.555.041.247 | 30-jun-2026, Hacienda (el XLSX trae un corte anterior: 5.702.528.148) |

### Por zona (2-13); zona 1 no tiene proyectos en ningún año

| Zona | PE 2024 | PE 2025 | PE 2026 | APP 2025 | OPP 2025 | APP 2026 | OPP 30-jun-2026 | APP 2024 (aprox.) | OPP 2024 (aprox.) | PEC proxy 2025 |
|---|---|---|---|---|---|---|---|---|---|---|
| 2 | 6 | 7 | 8 | 1.475.002.700 | 1.436.299.907 | 1.708.264.698 | 794.356.051 | 1.386.358.258 | 1.349.272.698 | 6/7 |
| 3 | 5 | 6 | 7 | 2.025.400.785 | 2.021.909.532 | 2.688.858.071 | 1.225.911.779 | 1.933.201.859 | 1.916.144.412 | 5/6 |
| 4 | 3 | 4 | 5 | 1.115.536.510 | 1.101.017.167 | 1.215.177.388 | 729.754.168 | 712.019.088 | 694.056.043 | 2/4 |
| 5 | 3 | 3 | 4 | 1.118.572.760 | 1.108.917.164 | 1.751.378.240 | 1.056.833.192 | 1.022.991.857 | 1.013.457.629 | 2/3 |
| 6 | 5 | 6 | 6 | 3.294.857.896 | 3.251.989.801 | 4.170.462.376 | 2.208.933.152 | 3.012.867.886 | 3.008.392.862 | 5/6 |
| 7 | 4 | 5 | 5 | 2.977.933.111 | 2.971.171.891 | 3.487.947.388 | 1.557.363.232 | 2.797.037.774 | 2.796.886.585 | 3/5 |
| 8 | 4 | 5 | 5 | 1.855.252.954 | 1.846.066.028 | 1.883.657.092 | 990.431.975 | 1.280.085.916 | 1.279.783.430 | 4/5 |
| 9 | 5 | 6 | 6 | 2.809.659.342 | 2.734.256.520 | 3.162.025.221 | 1.036.617.867 | 2.210.935.232 (4 de 5) | 2.210.784.166 | 5/6 |
| 10 | 4 | 5 | 6 | 658.027.155 | 647.489.065 | 515.862.905 | 151.705.782 | 271.710.200 (3 de 4) | 251.344.179 | 5/5 |
| 11 | 3 | 5 | 5 | 1.116.527.774 | 1.095.411.285 | 1.244.453.235 | 497.054.299 | 923.341.503 | 907.958.905 | 3/5 |
| 12 | 5 | 5 | 4 | 1.082.975.350 | 1.073.434.698 | 1.146.685.305 | 522.443.097 | 990.855.184 | 979.621.460 | 5/5 |
| 13 | 3 | 4 | 5 | 1.398.943.881 | 1.396.464.662 | 1.544.966.818 | 783.636.653 | 791.416.793 (2 de 3) | 791.416.793 | 4/4 |

Notas de la tabla: PEC proxy = proyectos con todos los indicadores de producto al 100 % de ejecución física neta al 31-dic-2025 sobre proyectos de la zona (autorreporte, no entrega verificada). Los valores 2024 vienen de una correspondencia por valor y zona entre el catálogo y los códigos de Hacienda; 34 de 50 son exactos y 13 aproximados (p. ej. zona 13 y zona 9); 3 proyectos no tienen correspondencia. Puede haber errores de asignación entre proyectos de la misma zona.

### Disponibilidad por zona

| Zona | OPP/APP | PE | PEC | NA |
|---|---|---|---|---|
| 1 | N/A: sin proyectos PP | N/A | N/A | N/A |
| 2 a 13 (cada una) | Construible | Construible | Proxy | Parcial |
| Municipal consolidado | Construible | Construible | Proxy | Parcial |

La categoría no difiere entre niveles.

## 4. PA y MP (hoja PP_Proyectos)

- **PA (personas atendidas): Parcial (2024), No disponible consolidado en 2025-2026.** INDER informa población beneficiada 2024 por zona y componente: deportes 4.586, actividad física 5.454, ludotecas 2.149, apoyos a instituciones educativas 3.857 y salud 1.153 (rendición INDER pp. 5-9, leído de texto). Son atenciones por componente, sin deduplicar, y no distinguen proyecto. Otros operadores reportan "inscritos" o "asistentes" (Comfenalco, Fundarte, cuidadores). No hay personas únicas por proyecto, que es lo que exige la Matriz.
- **MP (meta anual de personas): Parcial.** Algunos operadores declaran población objeto (zona 9 Fundarte: 70 objeto frente a 188 beneficiarios; ecozona zona 10: 50 familias; escuelas de música zona 7: 115 y 30 en el objetivo del catálogo 2025). El plan de acción mide actividades, porcentajes y paquetes, no personas. No hay MP positiva sistemática. Búsqueda web: sin fichas públicas; SUIFP/MapaInversiones por BPIN sin verificar.
- Los PDF de PP 2024 contienen listados con nombres de beneficiarios: no se reprodujeron.

## 5. Diferencias con la Fase 1

- PP.CEP.PEC pasa de Parcial a **Proxy**: ni 2024 ni 2025 tienen acta de recibo o informe final; solo avance físico autorreportado.
- "Zonas en movimiento" no es un proyecto multi-zona: son 12 proyectos con código propio; no hay proyectos compartidos que deduplicar.
- Los archivos `PP-*.pdf` de `concejoenvigado_gov_co/media_/` (bienestar animal, discapacidad, salud mental, movilidad, turismo) son **políticas públicas**, no presupuesto participativo. La Fase 1 los listó como candidatos: descartar.
- El plan de acción/ROCA 2025 y 2026 resuelve la identificación de proyectos PP; para 2024 no hay equivalente en el repo (se usó Hacienda y `proyectosdeinversion2024.xlsx`).
- No se usó el puente CODIGO_PROYECTO-BPIN por sufijo (invalidado). El BPIN sale de la misma fila del ROCA.
- La fecha de la rendición PP 2024 (12-ene o 12-feb-2025) no se resolvió.

## 6. URLs y accesos que fallaron

- `https://www.envigado.gov.co/` falla por certificado TLS (verify result 20); no se eludió.
- `https://efpc.envigado.gov.co/observatorio/` y `/presupuesto-participativo/`: 404. `.../uploads/2025/07/Proyectos_PP_2026.pdf` y `.../2026/07/Proyectos_PP_2026.pdf`: 404 (rutas supuestas por el Ejecutor; no existen o no se publicaron).
- MapaInversiones: la consulta por BPIN redirige a una página de error.
- WebSearch general (ejecución, actas de recibo, 66 proyectos 2026) no devolvió páginas de Envigado.

## 7. Documentos que no pude descargar o revisar

- Rendición de Cuentas 2025 (`.../content/files/002947/147324_rc2025.pdf`, p. 61) y Acta 096 de 2026 (`https://www.concejoenvigado.gov.co/wp-content/uploads/2026/09/Acta-096-de-2026.pdf`, p. 3) y Acuerdo 007 de 2025: se registran como verificados por el profesor Francisco; el Ejecutor no los abrió (el Acuerdo responde HTTP 200).
- Listado oficial de proyectos PP 2026 (no publicado en EFPC).
- Actas de recibo, informes finales, informes de supervisión y actas de liquidación por proyecto: no se encontraron publicados.
- Plan de acción 2024 por proyecto (apropiación/obligación): no está en el repo ni lo localicé en fuentes públicas.
- `AV/ejecucionppzona*inder.pdf` (INDER), diapositiva "Ejecución presupuestal a 17 de septiembre de 2024": es imagen y no se leyó; se usó la tabla de texto del consolidado a 31-dic-2024.
- Contratos SECOP II de 2025 y 2026 asociados a PP: sin número de contrato en el repo.

## 8. Pendiente y recomendaciones

- Confirmar con Planeación la lista de códigos PP (en especial el proyecto 0553 y los 16 proyectos 2024 con correspondencia aproximada o ausente).
- Reconciliar OPP con CUIPO (3 % de diferencia en 2025) antes de publicar.
- Definir "entrega exigible en la vigencia" y qué evidencia acepta el equipo para PEC; si se acepta el proxy, rotularlo como autorreporte.
