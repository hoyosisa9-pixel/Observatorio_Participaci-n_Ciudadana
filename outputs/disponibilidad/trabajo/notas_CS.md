# Notas Fase 2 - Sistema CS (control social)

Todos los valores son exploratorios y quedan en estado "Pendiente". No se calculo ningun indicador. Zonas: el Plan Indicativo trae columnas ZONA 1 a ZONA 13, coherente con las 13 zonas; ninguna fuente de CS dice otra cosa. Todas las variables de CS se publican solo a nivel Municipio (no hay desagregacion por zona/comuna).

## Resultado por categoria (8 variables)
- Disponible (2): CS.PCS.EE, CS.PCS.EP.
- Construible (5): CS.DSI.MD, CS.DSI.MS, CS.VRC.ET, CS.RPO.OR, CS.RPO.OT.
- Parcial (1): CS.VRC.EV.
- Proxy: 0. No disponible: 0.

## Hallazgos principales
1. PCS: el Plan Indicativo 2T-2026 (hoja Plan-Indicativo-Productos, fila 547, consecutivo 544, producto 5.1.2.1.2) da meta programada 1 por anio (2024-2027, total 4) y avance ejecutado 2024 = 1, 2025 = 1, 2026 = 0 al 2T (T1-2026 "en planeacion"). El DAP califico las evidencias 2024 y 2025 como "Si, pero con observaciones".
2. Otros productos de CS en el mismo archivo: consecutivo 545 (5.1.2.1.3, eventos de rendicion con veedurias: 1/1/0, obs. 2024: evento de las veedurias el 12-dic-2024 en la Casa de la Cultura) y consecutivo 546 (5.1.2.1.4, eventos de rendicion publica de cuentas: 1/1/0, obs. 2024: 26-nov-2024 en la Escuela Fernando Gonzalez; 2026: se hara en el ultimo trimestre). Fase 1 los citaba como 544 y 545 mas "el de rendicion"; el tercero es el 546.
3. DSI: en el Plan Indicativo 477 (2024), 528 (2025) y 539 (2026, al 2T) productos con meta > 0 tienen avance numerico publicado. Eso lo vuelve poco discriminante si el universo son productos; las evaluaciones PDM (IV trim. 2024 y 2025, PDF) solo evaluan indicadores de producto. Para metas de resultado no se hallo valor observado publicado. No fue necesario renderizar imagenes de PDF (no se crearon PNG; tmp_cs se borro).
4. VRC: las actas de Salud (21-mar-2024, 26-nov-2024, 19-mar-2025, 18-mar-2026) mencionan veedurias entre asistentes/participantes; la de nov-2024 muestra a tres veedurias presentando informes. Los eventos del Alcalde (Q&A 2024 y 2025, encuesta de la jornada 2024 con 167 respuestas) no mencionan veedurias. Una carta del Comite de Discapacidad (dic-2024) confirma que un integrante de la Veeduria 284 formulo preguntas el 17-dic-2024. Las listas de asistencia anexas no estan publicadas.
5. Registro de veedurias de la Personeria (datos.gov.co jju5-br3w, 45 registros, metadato actualizado ~7-oct-2026): vigentes por fecha de fin >= anio e inscripcion <= anio: 25 (2024), 27 (2025), 25 (2026); 16 registros con duracion en texto no fechado.
6. RPO: conteos exploratorios en variables_CS.csv (Alcalde nov-2024: 22 intervenciones, 5 felicitaciones sin respuesta; 100 dias: 14; Alcalde 2025: 3; Salud 2025: 6; Salud 2026: 3). Universo y criterio "respuesta de fondo" sin definir.

## Eventos de rendicion identificados (fechas segun el propio documento)
2024: Salud 2024-01 (21-mar); Alcalde 100 dias (10-abr); Comite Discapacidad S1 (30-jul); Victimas (solicitud 24-oct-2024; fecha del evento sin verificar; justificacion 10-dic); Alcalde audiencia (ver inconsistencia 1); Salud 2024-02 (26-nov); evento de veedurias (12-dic); Comite Discapacidad S2 (17-dic).
2025: PP vigencia 2024 (acta dice 12-ene-2025); Salud 2025-01 (19-mar); Discapacidad S1 (18-jul); Victimas (30-oct); Discapacidad S2 (10-dic); Alcalde 2025 (sin fecha en el repo; el Q&A cita el 10-dic-2025 como "ayer").
2026: Salud 2026-01 (18-mar); Discapacidad S1 (22-jul); Grupos Etnicos vigencia 2025 (3-sep-2026). Pendientes en el ultimo trimestre: Alcalde, PP, evento de veedurias.

## Inconsistencias y diferencias
1. Fecha de la audiencia del Alcalde 2024: la encuesta de la jornada esta fechada en su totalidad 20-nov-2024; el Plan Indicativo dice 26-nov-2024 (y la acta de Salud 2024-02 tambien 26-nov). Pregunta abierta.
2. PP: la justificacion dice rendicion realizada "12 de enero de 2025"; una noticia municipal encontrada por buscador sugiere 12-feb-2025 (no abierta).
3. Matriz PCS ("no asumir una por año") vs programacion oficial de 1 por anio: se resuelve con el Plan Indicativo (coincide con Fase 1).
4. Matriz: RPO exige OT "sobre el PDM"; los cuadros mezclan temas y felicitaciones. VRC/RPO dependen de definir "evento municipal". El registro tiene filas 2027 que quedan fuera del periodo pedido.
5. La rendicion de Grupos Etnicos de la vigencia 2025 se hizo el 3-sep-2026, despues del corte del 31-mar-2026 que fija la Guia (punto 17).
6. Cambios respecto a Fase 1: ninguna categoria cambio. Se concretaron valores (PCS), se identifico el consecutivo 546, se encontraron actas con mencion de veedurias (Salud, Discapacidad) y se confirmo que el Q&A del Alcalde 2025 es mucho mas corto que el de 2024.

## Busqueda web y URLs con problemas
- https://www.envigado.gov.co/rendicion-de-cuentas y paginas de noticias: WebFetch falla ("unable to verify the first certificate"); con curl -k se obtiene solo el esqueleto de una aplicacion JavaScript. Los archivos /content/files/... si descargan (verificados el Q&A 2025 y el xlsx del Plan Indicativo). Brecha: no se pudo ver el listado oficial de eventos ni los informes de rendicion 2025.
- WebSearch sobre el portal solo devolvio titulos: boletin del 24-oct-2025 que invita a la audiencia 2025 (https://www.envigado.gov.co/noticias/envigado-invita-a-la-audiencia-publica-de-rendicion-990814, no abierto) y la pagina "6.2.5 Rendicion de cuentas" (sin abrir).
- No se hallaron informes de la Personeria o Contraloria Municipal de Envigado ni del Concejo sobre control social 2024-2026 (sin resultados).
- https://efpc.envigado.gov.co/ abre; no lista cursos de control social 2024-2026.
- https://www.datos.gov.co/api/views/jju5-br3w.json abre (45 registros).
- Datos personales: los cuadros de P&R, la encuesta de la jornada y el CSV de veedurias los contienen; solo se contaron y no se reprodujeron en los entregables.
