# Avance 3 – Cálculo de indicadores viables (tanda 3)

**Fecha:** 9-nov-2026 · **Estado de todos los resultados:** PROVISIONAL · **Variables:** Pendiente (nada validado)

## Resultados del escenario por defecto (Municipio, anual)

| Indicador | Año | Variables | Puntaje | Semáforo |
|---|---|---|---:|---|
| CS.RPO – Respuesta pública a observaciones | 2024 | 16 respondidas de fondo / 17 observaciones (audiencia del Alcalde, sin felicitaciones) | 94,1 | Verde |
| CS.RPO | 2025 | cuadro publicado incompleto (3 intervenciones) | N/D | — |
| CS.VRC – Veedurías en la rendición de cuentas | 2024 | 3 eventos; presencia desconocida en el del Alcalde | N/D | — |
| CS.VRC | 2025 | 5 eventos; presencia desconocida en el del Alcalde | N/D | — |
| PP.EPP – Ejecución del presupuesto participativo | 2024 (aproximado) | OPP $18.357.265.731 / APP $18.601.236.386 | 98,69 | Verde |

- **PP.EPP 2024 por zona:** de 97,32 a 99,99 en Verde, salvo la zona 10 (82,58, Amarillo). Zona 1: NE.

## Por qué VRC queda N/D y cómo cambia con los escenarios
- **Por qué es N/D.**
  - La Matriz dice que, si hay eventos con presencia desconocida de veedurías, VRC es N/D.
  - Para la audiencia del Alcalde solo se publicó el cuadro de preguntas, sin acta ni lista de asistentes, así que no se sabe si hubo veedurías.
  - Por eso el valor por defecto (`cero_si_documento_integral`) da N/D en ambos años.
- **Cómo cambia con los escenarios.**
  - Si un evento sin mención de veedurías se cuenta como ausencia (`cero_siempre`): 2024 da 2/3 = 66,7 (Amarillo) y 2025 da 1/5 = 20 (Rojo).
  - Con el universo institucional del plan (consecutivos 545 y 546) y todo lo publicado hasta hoy: 1/2 = 50 (Amarillo) en ambos años.
- **Consecuencia.** El índice CS.ICRCS 2024 por defecto será N/D, porque le falta VRC. La Guía no permite calcular un índice sin una de sus entradas.

## RPO por escenario (2024)
| Escenario | OR/OT | Puntaje | Semáforo |
|---|---|---:|---|
| Defecto (Alcalde, sin felicitaciones, rúbrica estricta) | 16/17 | 94,1 | Verde |
| Criterio laxo de "respuesta de fondo" | 17/17 | 100 | Verde |
| Solo observaciones temáticas del PDM | 15/16 | 93,8 | Verde |
| Todas las intervenciones (con felicitaciones) | 16/22 | 72,7 | Amarillo |
| Alcalde y 100 días | 27/29 | 93,1 | Verde |
| Todos los cuadros (Alcalde, 100 días, Salud, Comité de Discapacidad) | 34/38 | 89,5 | Amarillo |

- **2025:** con el cuadro incompleto calculado tal como está publicado, da 1/2 = 50 (Rojo). Por defecto queda N/D, porque la Matriz dice que un listado parcial no permite conocer el total.

## Segunda lectura (para usted)
- **Dónde están:** en `outputs/indicadores/segunda_lectura/`, con instrucciones en `LEEME.md`.
- **Qué hay:**
  - `muestra_vrc.csv`: 4 eventos de 17.
  - `muestra_rpo.csv`: 12 observaciones de 56.
  - Es el 20 % redondeado hacia arriba, elegido con semilla fija (20261109).
- **Qué llenar:** las columnas `clasificacion_lectora_2`, `acuerdo_si_no` y `comentario`. Si vuelvo a correr los scripts, sus respuestas no se sobrescriben.
- **Tablas completas:** `clasificacion_vrc.csv` (17 eventos) y `clasificacion_rpo.csv` (56 observaciones). Tienen una fila por evento u observación, con documento, página, enlace, una cita breve sin datos personales, la clasificación y el criterio aplicado.

## Puntos dudosos que conviene confirmar
1. **VRC:**
   - El evento de veedurías del 12-dic-2024 (consecutivo 545) queda fuera de la ventana por defecto. Su única evidencia es el Plan Indicativo publicado en jul-2026.
   - En el Comité de Discapacidad de 2025, las veedurías PAE y PIC figuran como integrantes, pero no consta que asistieran. Quedó como presencia desconocida.
   - Hay fechas contradictorias entre documentos para cuatro eventos.
2. **RPO:**
   - El ejercicio de los 100 días tiene 12 intervenciones, no las 14 del plan.
   - Las remisiones sin contenido no cuentan como respuesta de fondo con la rúbrica estricta y sí con la laxa.
   - Los 15 casos dudosos están marcados con `dudoso=1`.
3. **PP.EPP 2024:**
   - **Cruce:** se hizo proyecto por proyecto entre el listado de la Escuela de Formación Política y Ciudadana y el informe de Hacienda al 31-dic-2024. Hay 35 coincidencias exactas, 15 aproximadas y 1 sin correspondencia (`clasificacion_pp_2024.csv`).
   - **Diferencia con la estimación previa:** la nota anterior hablaba de 47 de 50 proyectos y $17,33 mil millones. Ahora se cruzan 50 de 51 y la apropiación es de $18,6 mil millones; la cifra anterior no se reproduce.
   - **Registro dudoso de la zona 8:** su valor coincide al peso con Hacienda (proyecto 0726), así que se incluye.
   - **Zona 10 en Amarillo:** se debe a un cruce aproximado con 32 % de diferencia en valor.
   - **Por confirmar con Planeación:** dos cruces débiles, 13-2 ↔ 0714 y 12-3 ↔ 0733.
   - **Solo con cruces exactos:** 99,28 (Verde).

## Verificación contra la matriz de los profesores
- Se compararon 60 resultados: **46 coinciden**.
- Las **14 diferencias** son las mismas esperadas de las tandas anteriores: PP.EPP 2026 del primer semestre y la zona 1.
- PP.EPP 2024 municipal y por zona coincide con la matriz.
- La copia temporal "Validado" se borró. La entrega no tiene ningún "Validado" ni datos personales.

## Próxima tanda (4, final)
- **CS.ICRCS 2024:** N/D por defecto. Se calcularán los escenarios en que las cuatro entradas tienen número.
- **Revisión crítica del Analizador:** recálculo a mano de una muestra y revisión de fórmulas, semáforos y soporte.
- **Cierre:** `resumen_calculos.md` y la entrega final.
