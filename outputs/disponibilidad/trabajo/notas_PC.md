# Notas Fase 2 - Sistema PC (participación ciudadana)

Ejecutor, 14-oct-2026. No se calculó ningún índice. Todos los valores quedan en estado "Pendiente".

## Conteo por categoría (12 variables)
- Disponible: 1 (PC.I2.V2 = 13 zonas).
- Construible: 1 (PC.I2.V1).
- Parcial: 10 (I1.V1, I1.V2, I3.V1, I3.V2, I4.V1, I4.V2, I5.V1, I5.V2, I6.V1, I6.V2).
- Proxy: 0. No disponible: 0.

## Cambios frente a la Fase 1
| Variable | Fase 1 | Fase 2 | Motivo |
|---|---|---|---|
| PC.I3.V1 y V2 | Proxy | Parcial | El informe de gestión 2025 (pp. 87-88) publica el conteo de acciones: 3 diplomados, 9 seminarios y 9 cursos culminados; indica además 1 diplomado y 3 seminarios no ejecutados (permite inferir 25 programadas). Para 2024 se reconstruyen 7 acciones desde el informe de gestión 2024 (pp. 62-63) y el plan indicativo. |
| PC.I6.V1 | No disponible | Parcial | Rendición 2024 con preguntas de participantes virtuales (QA24 pp. 6-8) y votación del PP en plataforma web, ago-2025 (G25 p. 92). |
| PC.I2.V2 | Disponible con duda 12 vs 13 | Disponible = 13 | PDM PDF p. 558: 13 zonas, 12 comités, zona 1 pendiente (menos de 1.000 habitantes). Cronograma POT 2025 incluye zona 1. |
| PC.I5 | Parcial | Parcial con cifra | PPSS 2025: COPACO 4/4 y CMSSS 4/4 sesiones. Se corrige un error de la Fase 1: el PDF "20251125_instancia-copaco-1.pdf" no es un acta, es una ficha descriptiva. |
| Resto | igual | igual | I1, I4 y I6.V2 siguen Parcial; I2.V1 sigue Construible. |

## Hallazgos clave
1. Zonas: el municipio tiene 13 zonas, pero la zona 1 no tiene comité zonal ni asambleas de presupuesto participativo (se conformaría al superar 1.000 habitantes). Por eso el informe de gestión 2025 habla de "12 asambleas que cubren la totalidad de las zonas" (p. 92). Para PC.I2.V1 el numerador del PP tiene tope 12 con denominador 13; las asambleas de revisión del POT (2025) sí incluyeron la zona 1. Decisión pendiente para los docentes: qué eventos cuentan como "actividad participativa".
2. Formación: el reporte oficial (plan indicativo, productos 520, 521, 524, 525) cuenta personas (95 en 2024; 186 en 2025 en producto 521). El conteo de acciones solo aparece en las narrativas de los informes de gestión. En 2025 hay mezcla de formación a ciudadanía, servidores (47 funcionarios certificados; diplomado Gestión Pública) y semilleros. En 2026 la Escuela figura "en planeación" (observación al 1.er trimestre) y la oferta web 2026 lista 4 ofertas (3 diplomados y el curso Control Social) sin fechas.
3. Procesos participativos (I1 e I4): no existe inventario oficial. Se propuso un universo exploratorio de 3 procesos (2024) y 5 (2025), descrito en `variables_PC.csv` (nota de PC.I1.V1). Es un insumo para discusión, no una cifra.
4. Instancias (I5): el Consejo Municipal de Participación, el CTP y los comités zonales reciben "asistencia técnica por demanda" (indicador en porcentaje, 100 % cada año), que no equivale a sesiones previstas/realizadas.
5. Canal virtual (I6): las transmisiones por YouTube y Facebook (rendición 12-feb-2025: 776 vistas; "Conversando con los Comités Zonales", 11 programas en 2024) no se cuentan como canal para intervenir. Los formatos de consulta de rendición 2025/2026 solo preguntan el medio de participación preferido.
6. Observatorio de Participación: el producto 536 reporta informes (incluido el del semestre II-2025), pero la sección Observatorio Ciudadano del sitio EFPC está vacía ("No se encontró una URL válida para mostrar") y /observatorio-ciudadano/ devuelve 404.

## Inconsistencias de la matriz y de las fuentes
- La matriz exige "zonas oficiales" sin catálogo; el Registro no valida la columna Territorio (ya señalado en la Fase 1). Resuelto en este sistema con 13, pero falta el acto oficial en el repo.
- PC.I3 pide acciones "programadas con finalización prevista en el periodo"; la única programación oficial está en personas.
- PC.I5 e PC.I6 son complementarios y admiten semestre, pero las fuentes solo permiten anual.
- Plan indicativo, producto 531: el campo de observaciones 2025 dice "12 zonas del municipio" aunque el catálogo tiene 13 (ver hallazgo 1).
- PPSS 2025: la actividad E5Ld-M02-A02 está rotulada "VIGENCIA 2024" dentro del archivo de 2025.
- Plan indicativo, producto 520 (personas en capacitación): el valor ejecutado 2025 del plan indicativo es 65, mientras el informe de gestión 2025 (p. 88) reporta 170 personas capacitadas por demanda. Diferencia sin explicar (posible corte distinto).

## URLs y fuentes que fallaron
- https://www.envigado.gov.co (cualquier ruta): WebFetch "unable to verify the first certificate"; curl error 60 (cadena incompleta). No se eludió la verificación TLS. Todo lo del portal salió de las copias del repo (`url_origen` del inventario) o de títulos de búsqueda, marcado sin verificar.
- https://mi.envigado.gov.co: tiempo de espera agotado.
- https://efpc.envigado.gov.co/observatorio-ciudadano/: 404. /informes/, /reportes/ y /diagnostico/: entradas sin URL ("Informes 1 a 3", "Reportes 1 a 3", "Diagnóstico 1 a 3"). /semilleros/: sin contenido listado.
- Búsquedas sin resultado útil: Acuerdo 007 de 2025 de Envigado; actas del Consejo Municipal de Participación; datos.gov.co (no hay conjunto de Envigado sobre procesos participativos; el de veedurías jju5-br3w pertenece a CS).
- No se leyó ninguna cifra de imágenes (no hizo falta renderizar páginas); el cronograma de asambleas del PDM (anteproyecto p. 53, figura 3) es una imagen no leída. Las memorias de la EFPC son imágenes no leídas.
- Las respuestas de WebFetch son resúmenes automáticos: las listas de oferta EFPC (27 ofertas 2025; 4 ofertas 2026) deben contrastarse a mano antes de validar.

## Datos personales
Los cuadros de preguntas y respuestas de rendición, el calendario 2024 y la programación de asambleas 2026 contienen nombres, teléfonos y funcionarios. Solo se contaron filas; no se transcribió ningún dato personal.

## Fuera de alcance
El informe de auditoría de participación ciudadana de la Contraloría (2026) mide la participación en el control fiscal, no los procesos de la Alcaldía; no se usó.
