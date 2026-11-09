# Avance 1 – Cálculo de indicadores viables (tanda rápida)

**Fecha:** 9-nov-2026 · **Estado de todos los resultados:** PROVISIONAL · **Variables:** Pendiente (nada validado)

## Resultados del escenario por defecto (Municipio)

| Indicador | Año / periodo | Variables | Puntaje | Semáforo |
|---|---|---|---:|---|
| CS.PCS – Promoción del control social | 2024 | 1 ejecutada / 1 programada | 100 | Verde |
| CS.PCS | 2025 | 1 / 1 | 100 | Verde |
| PC.I5 – Sesiones de instancias | 2025 | 8 realizadas / 8 previstas (COPACO y CMSSS) | 100 | Verde |
| PC.I5 | 2024 y semestres 2025 | sin programación ni fechas por sesión | N/D | — |
| PDM.CA – Espacio público verde | 2024 | 5,75 m²/hab (POT) frente a meta 3,88 | 100 | Verde |
| PP.EPP – Ejecución presupuestal PP | 2025 | OPP $20.684.427.720 / APP $20.928.690.218 | **98,83** | Verde |
| PP.EPP | 2026, semestre 1 | OPP $5.702.528.148 / APP $24.519.738.737 | 23,26 | (ver nota) |

- **PP.EPP 2025 por zona (2 a 13):** de 97,32 % (zona 9) a 99,83 % (zona 3). Zona 1: NE, porque no tiene comité ni proyectos (PDM p. 558); la Guía exige confirmarlo.
- **Las 7 metas del PDM** se leyeron directamente de las páginas del PDM aprobado y coinciden con la Guía: 99,5 (p. 602), 8 (p. 706), 3,88 (p. 717), 100 (p. 742), 6 (p. 609), 19 (p. 650) y 14 (p. 652). El documento técnico de soporte de diciembre de 2025 da las mismas cifras.

## Supuestos del escenario por defecto (PROVISIONAL, en `config/supuestos.yaml`)
- **Metas:** se evalúan contra la meta 2027.
- **Fechas:** la fecha de corte se registra al 31-dic de cada año, y el corte real va en Observaciones.
- **2026:** solo primer semestre, como provisional.
- **PC.I5:** instancias COPACO y CMSSS, según lo que reporta la dependencia en el seguimiento del plan de acción en salud. Las actas están en SharePoint institucional, sin acceso.
- **PDM.CA:** usa el espacio verde urbano del diagnóstico del POT. Es una serie distinta de la línea base del plan (3,88 en 2023).
- **PP.EPP:**
  - proyectos con "ZONA n" en el plan de acción, contrastados con el listado de la EFPC;
  - sin el proyecto 0553 (con él, 98,86 %);
  - municipal = suma de obligaciones sobre suma de apropiaciones, nunca un promedio de zonas.

## Verificación contra la matriz de los profesores
- Llené una copia temporal marcada como "Validado", fuera de `outputs/`, y la recalculé con LibreOffice. Comparé 35 resultados.
- **21 coinciden** con los scripts, entre ellos CS.PCS, PC.I5, PDM.CA y PP.EPP 2025 municipal y por zona.
- **14 difieren, todos de PP.EPP 2026 primer semestre, por diseño:**
  - la hoja Resultados solo calcula el presupuesto participativo con periodo anual y devuelve N/D para el semestre;
  - en la zona 1, la matriz da N/D donde el escenario por defecto da NE.
- La copia temporal se borró.
- La entrega, `matriz_diligenciada.xlsx`, queda en "Pendiente" y no contiene ningún "Validado". Por eso su hoja Resultados muestra N/D.

## Correcciones y alertas
- **Obligaciones de PP 2026: había un error en la versión 1.1.**
  - La cifra $11.555.041.247 que se citó en octubre no se reproduce. El informe de Hacienda al 30-jun-2026 y la hoja ROCA del plan de acción coinciden al peso en $5.702.528.148.
  - Ya está corregida en los supuestos y en el plan. Falta decidir si se corrige en el Excel de disponibilidad ya entregado.
- **El semáforo de PP.EPP 2026 S1 (Rojo) no es comparable:** compara una ejecución de medio año con umbrales anuales. Debe leerse como avance, no como incumplimiento.
- **Forensis 2024:** sigue pendiente de que lo suba la usuaria. PDM.CM se calculará en la próxima tanda con la cifra citada (7,25 por 100.000, p. 440), en estado Pendiente.

## Qué falta (próximas tandas)
1. **Tanda 2:**
   - PDM.CM 2024;
   - PC.I2 2024-2025;
   - PC.I3 (acciones 2025 / personas 2024-2025);
   - CS.DSI 2024-2025 (7 metas del ICPD-6 por defecto; la matriz sugería unas 10 y la lista la definen los profesores);
   - variables de los indicadores no viables de PP, PDM y PC.
2. **Tanda 3:** CS.VRC y CS.RPO.
   - Tabla de clasificación con una fila por evento u observación.
   - Muestra aleatoria del 20 % para su segunda lectura.
   - PP.EPP 2024, aproximado.
3. **Tanda 4:**
   - índice CS.ICRCS 2024;
   - revisión crítica del Analizador, que recalculará una muestra a mano;
   - `resumen_calculos.md` y la entrega final.
