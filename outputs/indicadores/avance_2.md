# Avance 2 – Cálculo de indicadores viables (tanda 2)

**Fecha:** 9-nov-2026 · **Estado de todos los resultados:** PROVISIONAL · **Variables:** Pendiente (nada validado)

## Resultados del escenario por defecto (Municipio, anual)

| Indicador | Año | Variables | Puntaje | Semáforo |
|---|---|---|---:|---|
| CS.DSI – Disponibilidad de información de seguimiento | 2024 | 0 de 7 metas con valor publicado a tiempo | 0 | Rojo |
| CS.DSI | 2025 | 0 de 7 | 0 | Rojo |
| PC.I2 – Cobertura territorial | 2024 | 13 zonas con asamblea / 13 | 100 | Verde |
| PC.I2 | 2025 | 12 / 13 | 92,3 | Verde |
| PC.I3 – Oferta formativa (acciones) | 2025 | 20 ejecutadas / 24 programadas | 83,3 | Verde |
| PC.I3 | 2024 | 7 ejecutadas; programadas sin dato | N/D | — |
| PDM.CM – Mortalidad vial | 2024 | 7,25 por 100.000 (Forensis, cifra citada) frente a meta 8 | 100 | Verde |

Con los de la tanda 1 ya hay **10 resultados municipales calculados** (CS.PCS, CS.DSI, PC.I2, PC.I3, PC.I5, PDM.CA, PDM.CM y PP.EPP).

## Lo que más cambia con los escenarios

- **CS.DSI = 0 depende de la ventana documental.**
  - El valor por defecto (Guía p. 17) solo cuenta un dato si se publicó a más tardar el 31 de marzo del año siguiente.
  - Forensis (sin fecha de publicación verificada) y el POT (publicado en oct-2025) quedan por fuera.
  - Si se acepta todo lo publicado hasta hoy, 2024 sube a 2 de 7 (28,6, Rojo).
  - Si se aceptan proxies, 2025 sube a 2 de 7 (28,6, Rojo).
  - Con la lista de 10 metas sugerida por la matriz: 0 de 10. Tres de esas metas adicionales (claridad de la información, IDI y Gobierno Digital) no tienen evidencia en el repositorio.
  - La lista de metas la definen los profesores.
- **PDM.CM:** todos los escenarios dan Verde (100), salvo uno: con datos de la Policía (20 casos) y población DANE post-COVID da 99,3 (Amarillo). La Policía usa otra definición: homicidio en accidente de tránsito.
- **PC.I3:** si se incluye el diplomado en Gestión Pública, 2025 da 21/25 = 84 (Verde). Medido en personas (producto 521), las ejecutadas superan a las programadas (95 frente a 10 en 2024 y 186 frente a 15 en 2025), y la matriz devuelve "Revisar".
- **PC.I2:** si se cuenta la programación de las asambleas POT, 2025 sube a 13/13. Si el universo es de 12 zonas con comité, ambos años dan 100.

## Supuestos que conviene confirmar
1. **PC.I2:** se asignaron las 12 asambleas de PP a las zonas 2 a 13. Las fuentes solo dan el total y la zona 1 no tiene comité zonal.
2. **PC.I3:**
   - Se excluyó el diplomado en Gestión Pública por considerarlo formación para servidores. El informe no acredita su público.
   - Las 24 programadas de 2025 se reconstruyeron como 21 culminadas más 4 no ejecutadas (1 diplomado y 3 seminarios); en el escenario por defecto quedan 20 ejecutadas y 24 programadas.
3. **PDM.CM:** Forensis 7,25 y las 18 muertes (p. 37) son cifras citadas. Se verificarán cuando suba el PDF a `data_raw_IIC/06_descargas_indicadores/inmlcf/`.
4. **PP, proyectos con entrega exigible en 2024: 51, no 50.**
   - El listado de la Escuela de Formación Política y Ciudadana tiene 51 registros. Uno de la zona 8 viene sin el signo "$" en el PDF y antes no se contaba.
   - Se incluyó marcado como dudoso. Hay que confirmarlo con Planeación.
5. **PP, proyectos con entrega cumplida en 2025:**
   - Según la regla de conteo, son 49 (por defecto), 45 o 50.
   - La cifra de 55 citada en el plan no se reproduce con ninguna de las cerca de 40 definiciones probadas.
   - Es un proxy de autorreporte, sin acta de recibo.

## Variables construidas de los indicadores no viables (sin calcular el indicador)
- **PC.I1, PC.I4 y PC.I6:**
  - Universo exploratorio de procesos: 3 en 2024 y 5 en 2025.
  - Participación virtual: 1 por año, como cota inferior.
  - Todo Pendiente; el universo lo definen los profesores.
- **PDM.CS (MinSalud):**
  - Afiliados al régimen subsidiado: 29.794 (dic-2025) y 31.997 (ago-2026).
  - El método del plan (99,5 %) no es reproducible con las fuentes públicas.
- **PDM.ILE.A (atención integral a la primera infancia):**
  - 15,58 (2024) y 15,40 (2025) con el rango de 0 a 6 años, que es el valor por defecto.
  - 18,43 y 18,23 con el rango de 0 a 5 años. El informe de gestión dice "≤5 años", y este rango es el que más se acerca a la línea base.
  - La decisión es de los profesores.
- **PDM.CP, PDM.CD y PDM.ILE.S:** quedan variables auxiliares. El índice situacional es 3,23 (2022); las disciplinas deportivas son 38 de 40; hay 1 institución educativa que mejora, como cota inferior.
- **PP.CEP y PP.CMA:**
  - Hay una tabla de clasificación por proyecto (`clasificacion_pp_entregas.csv`, 180 filas).
  - PP.CMA queda N/D mientras no se decida a qué proyectos aplican las metas de atención.

## Verificación contra la matriz de los profesores
- Se compararon 42 resultados: **28 coinciden**.
- Las **14 diferencias** son las mismas esperadas de la tanda 1: PP.EPP 2026 del primer semestre, que la matriz solo calcula con periodo anual, y la zona 1 (NE frente a N/D).
- La matriz calcula los mismos puntajes de CS.DSI (0), PC.I2, PC.I3 y PDM.CM.
- La copia temporal "Validado" se borró. `matriz_diligenciada.xlsx` tiene 0 celdas "Validado".
- No hay datos personales. Los 12 aparentes "teléfonos" que marcó el filtro son montos en pesos de 10 dígitos.

## Próxima tanda (3)
- CS.VRC y CS.RPO:
  - tabla de clasificación con una fila por evento u observación;
  - muestra aleatoria del 20 % para su segunda lectura.
- PP.EPP 2024, aproximado.
