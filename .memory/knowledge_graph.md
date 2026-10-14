# Grafo de Conocimiento del Proyecto - OCE & OPCEN

## 1. Documentos Base y Directrices (`Documentos guias/`, aún sin versionar en git)
- `Documentos guias/02032026_Propuesta  Observatorio Ciudadano.docx` (Propuesta metodológica y conceptual del OCE). Nota: el nombre lleva doble espacio.
- `Documentos guias/Bases_Datos_IIC_Envigado_Completo.zip` (fuentes crudas del IIC). **ESTADO: incompleto/corrupto** (59 MB, sin directorio central; no se puede listar ni extraer). Pendiente volver a subirlo.
- **PENDIENTE:** `12022026_Observatorio_Participación_Ciudadana_OPCEN.docx` (Diagnóstico base institucional OPCEN). No está en el repositorio.

## 2. Datos y Matriz de Indicadores (IIC) (`data_raw_IIC/`)
- `data_raw_IIC/inventario_detallado_IIC.csv` (versión estándar para Python, separador `,`).
- `data_raw_IIC/inventario_detallado_IIC_excel.csv` (versión para Excel, separador `;`, UTF-8 con BOM; generada por `scripts/csv_para_excel.py`).
- Ambas: 2.787 registros y 11 campos de metadata: `subcarpeta`, `archivo`, `nombre_original`, `ruta_local`, `url_origen`, `pagina_fuente`, `titulo_publicacion`, `ventana_temporal`, `emisor`, `fecha_publicacion`, `bytes`. Es un inventario de archivos fuente, no la tabla de indicadores.
- `inventario_IIC_relaciones.md` (raíz): trazabilidad y fórmula matemática del IIC con llaves de cruce `BPIN`, `CODIGO_PROYECTO`, `zona`, `vigencia`. Las llaves solo se pueden verificar contra los datos crudos del ZIP.

## 3. Fuentes Crudas (EXTRAÍDAS en `data_raw_IIC/`, 2026-10-14)
- Categorías esperadas (valores de `subcarpeta`): `01 Propuestas y acuerdos/`, `02 Ejecucion y presupuesto/`, `03 Actas y comites tecnicos/`, `04 Datos abiertos y secop/`, `05 Fuentes externas/`.

## 4. Scripts (`scripts/`)
- `scripts/csv_para_excel.py`: genera `inventario_detallado_IIC_excel.csv` a partir del CSV estándar.
- `scripts/md_a_latex_iic.py`: convierte `inventario_IIC_relaciones.md` en `informe_IIC_overleaf.tex` y su copia `.txt`.

## 5. Informes (raíz)
- `informe_IIC_overleaf.tex`: informe LaTeX para Overleaf.
- `informe_IIC_overleaf.txt`: copia idéntica del `.tex`.

## 6. Agentes (`.claude/agents/`)
- `analizador.md`: estrategia y auditoría metodológica (Opus). Diseña planes; no escribe código definitivo.
- `ejecutor.md`: implementación de scripts y datos (Sonnet). Ejecuta los planes del Analizador.

## 7. Estado de los ZIP en `Documentos guias/` (actualizado 2026-10-14)
- Verificados (testzip OK, entradas y bytes coinciden 1:1 con `inventario_detallado_IIC.csv`: ruta_local sin prefijo `data_raw_IIC/` = nombre de entrada del ZIP): `01_propuestas_y_acuerdos.zip`, `03_actas_y_comites_tecnicos.zip`, `04_datos_abiertos_y_secop.zip`.
- `02_ejecucion_y_presupuesto.zip` REPUESTO desde Google Drive (carpeta con ZIP 01-05 + 3 archivos de inventario): 2.840.613.915 B, 1.067 entradas, 3.322.769.994 B descomprimidos, 0 faltantes/sobrantes/diferencias. SHA-256 `194a5dfc7164889cbb8868e1c5a67ce1a21473da149e800aff47098b2307eda4`.
- `05_fuentes_externas_y_pot.zip` REPUESTO desde Drive: 1.835.200.497 B, 571 entradas, 2.626.440.084 B descomprimidos, 0 diferencias. SHA-256 `8110571c695a851bd4735369b62992105df80310ad269d49d974b36f8a84f4d6`.
- Copia de descarga en `/tmp/iic_drive/` (4,4 GB, fuera del repo).
- PENDIENTES DE BORRADO (esperan confirmacion explicita de la usuaria): `02_ejecucion_y_presupuesto.zip.truncado`, `05_fuentes_externas_y_pot.zip.truncado`, `Bases_Datos_IIC_Envigado_Completo.zip` (corrupto, ya innecesario), `inventario_detallado_IIC (1).csv` (duplicado), y duplicados de inventario/relaciones dentro de `Documentos guias/`.
- Pendiente: extraer a `data_raw_IIC/` (no hecho; espacio libre en /workspaces ~9,6 GB vs ~9,4 GB descomprimidos de 02+05, borrar los .truncado liberaria 3 GB). Pendiente: DOCX OPCEN (`12022026_Observatorio_Participación_Ciudadana_OPCEN.docx`).

## 8. Extracción de datos crudos (2026-10-14)
- Los 5 ZIP se extrajeron a `data_raw_IIC/` (script en scratchpad, `python -I`, con protección contra path traversal y chequeo de espacio). Entradas: 01=389, 02=1.067, 03=754, 04=6, 05=571 (total 2.787). 0 colisiones; inventarios CSV intactos.
- Verificación contra `inventario_detallado_IIC.csv`: 2.787/2.787 `ruta_local` existen con el tamaño exacto en `bytes`; 0 faltantes, 0 discrepancias.
- Git: `data_raw_IIC/*` y `*.zip` siguen ignorados por `.gitignore`; sin commits.
- Espacio libre en /workspaces tras extraer: ~5,2 GB (83% usado).
- Los ZIP de `Documentos guias/` se conservan como respaldo; borrarlos liberaría ~6,3 GB (decisión pendiente de la usuaria). Los .truncado, el ZIP corrupto y duplicados ya fueron borrados (con autorización); la sección 7 sobre pendientes de borrado y extracción queda superada.
- Pendiente: DOCX OPCEN (`12022026_Observatorio_Participación_Ciudadana_OPCEN.docx`).

## 9. Llave SECOP I `codigo_bpin` y perfilado de llaves (2026-10-14)
- `inventario_IIC_relaciones.md` B.1: añadida la llave puente SECOP I Procesos `codigo_bpin` = CUIPO `bpin` (contrato ↔ BPIN ↔ ejecución) y la fila "Ausencia de BPIN": SECOP II Contratos e Integrado NO tienen BPIN (solo `origen_de_los_recursos` / `origen`). `informe_IIC_overleaf.tex/.txt` regenerados con `scripts/md_a_latex_iic.py` (+2 líneas). No hay compilador LaTeX en el Codespace (sin pdflatex/xelatex/tectonic): sin compilar.
- `scripts/perfil_llaves_iic.py` (solo stdlib; pypdf opcional vía `--lib` para leer el PDF PPP 2022, que NO es tabular). Ejecutar: `python -I scripts/perfil_llaves_iic.py [--lib RUTA_PYPDF]`. Salidas sin datos personales en `outputs/`: `perfil_llaves_iic_perfil.csv`, `perfil_llaves_iic_cruces.csv`, `perfil_llaves_iic.md` (incluye la normalización).
- Resultados (Envigado): SECOP I 2.531 filas, `codigo_bpin` 91,15 % nulo (placeholder "No definido"), 63 BPIN únicos; CUIPO `bpin` 62.413 filas, 31,56 % nulo (valor `0`), 450 únicos; PPP 2022: 14 contratos únicos (52 apariciones).
- Cruces: PPP↔SECOP II `referencia_del_contrato` 78,6 % (85,7 % ignorando sufijo `-1`); PPP↔SECOP I 7,1 %; PPP↔`id_contrato` e Integrado 0 %. SECOP I `codigo_bpin`↔CUIPO `bpin` solo 5/63 (7,9 %): BPIN de SECOP I son de 2018-2021, CUIPO es mayoritariamente 2021-2026. `CODIGO_PROYECTO` (3-4 dígitos locales) ↔ BPIN completo = 0 %; vía sufijo local `bpin[-4:]` (BPIN = AAAA+05266+NNNN) ≈ 25 % (exploratorio). SECOP II `id_contrato` ↔ Integrado `numero_del_contrato` 90,8 %.
- Calidad: CUIPO `bpin` trae valores malformados tipo `19.0`, `22.1` (longitudes 6-15); SECOP I tiene un BPIN de 11 dígitos; SECOP II `referencia_del_contrato` repetida (25.549 únicos en 25.854 filas); `nombre_entidad` varía (p. ej. `MUNICIPIO DE ENVIGADO*`).
- `outputs/` no está en `.gitignore` (archivos pequeños).

## 10. Limpieza de `bpin` CUIPO y validación del puente por nombre (2026-10-14)
- `scripts/limpieza_bpin_cuipo.py` (stdlib; reutiliza `perfil_llaves_iic.py`). Salidas en `outputs/`: `limpieza_bpin_cuipo.md`, `limpieza_bpin_cuipo_categorias.csv`, `limpieza_bpin_cuipo_valores.csv`, `puente_codigo_proyecto_bpin.csv`.
- CUIPO Ejecución (62.413 filas): válido 42.319; nulo (`0`) 19.700; recuperado 191 (corrección inequívoca verificada contra BPIN válidos); malformado 165 (esquema `2025000000NNNNN`, 15 dígitos; marcado, no cruzado); código sectorial `NN.NN.NN` 38 (no es BPIN).
- Cobertura: `periodo` solo 2023-2026 (sin 2018-2022). BPIN de SECOP I son 2018-2021; solo 4 tienen prefijo >=2021 y los 4 están en CUIPO. SECOP I↔CUIPO con bpin limpio: sigue 5/63 (7,9 %): el límite es estructural (cobertura temporal), no de formato.
- Puente `CODIGO_PROYECTO` ↔ BPIN por sufijo de 4 dígitos validado por nombre: NO válido (2023: 55 candidatos, 0 confirmados, 2 dudosos, 53 rechazados; 2024: 51, 0, 2, 49). Sufijo ambiguo entre vigencias (47 y 43 códigos con varios BPIN). Los nombres MGA son genéricos de producto. B.1 actualizado y `.tex` regenerado. Pendiente: tabla oficial de equivalencias proyecto-BPIN.

## 11. Control de versiones (2026-10-14)
- Commit `9aca547` en `main`, subido a GitHub (origin/main): matriz B.1, informe LaTeX/TXT, `scripts/perfil_llaves_iic.py`, `scripts/limpieza_bpin_cuipo.py`, `outputs/` y `.gitignore` (`__pycache__/`).
- Fuera de git: `.claude/`, `.memory/`, `CLAUDE.md`, `Documentos guias/` (ZIP + DOCX), `data_raw_IIC/` (datos crudos).

## 12. Tarea "Disponibilidad y brecha del sistema de indicadores" (2026-10-14, TERMINADA)
- Encargo: para la reunión de los profesores Francisco Correa y Eber (entrega a la Alcaldía 15-oct), clasificar 39 variables y 23 indicadores/índices (4 sistemas: CS, PC, PDM, PP) por disponibilidad. NO calcula el IIC ni índices. **Punto de control autosuficiente: `outputs/disponibilidad/ESTADO.md` (leerlo primero; sección "Cómo retomar").**
- Matriz de los profesores (no modificar): `Documentos guias/Matriz de informacion de indicadores del observatorio ciudadano.xlsx` (Guía, Matriz, Registro, PP_Proyectos, Resultados). Copia y volcados CSV/TXT en `outputs/disponibilidad/trabajo/`.
- Escala: Disponible / Construible / Parcial / Proxy / No disponible. Valores siempre "Pendiente" (nunca "Validado"). Zonas = 13 (PDM p. 558; zona 1 sin comité).
- Fases: 1 y 2 TERMINADAS (39/39; `trabajo/{variables,fuentes,solicitudes}_<SIS>.csv` + `notas_<SIS>.md` para CS, PC, PDM, PP). PP usó la info del profesor Francisco (4 fuentes verificadas) y reconstruyó APP/OPP 2025 cruzando 61 proyectos "ZONA n" del plan de acción con el listado EFPC. Fases 3 y 4 pendientes.
- Conteo tras Fase 2: Disponible 10 (7 metas), Construible 11, Parcial 15, Proxy 2 (PDM.CD.O, PP.CEP.PEC), No disponible 1 (PDM.CP.O). Repo GitHub PRIVADO (matriz_copia.xlsx versionada sin problema). Se omitió un nombre de persona en un nombre de archivo citado en variables_CS.csv.
- Hallazgos: las 7 metas de la Guía (99,5; 8; 3,88; 100; 6; 19; 14) coinciden en PDM, DTS 22-dic-2025 y Apéndice 9.1.2 (meta 2027, sin metas anuales); páginas de la Matriz no corresponden. Forensis 2024: 18 muertes viales, 7,25/100.000. MinSalud RS 29.794 afiliados (dic-2025). POT: 5,75 m²/hab verde urbano 2024 (no cuadra con 3,88). No existe inventario oficial de procesos participativos (PC.I1/I4/I6).
- Acceso: `www.envigado.gov.co` falla por certificado TLS (no eludir); solo bajan `/content/files/`. También fallan mi.envigado.gov.co, ANSV, Tableau MinSalud, ICFES por plantel, EFPC /observatorio-ciudadano/ (404).
- Herramientas instaladas (pip, usuario): openpyxl, pdfplumber, pypdf, pymupdf. Sin pdftotext/pdftoppm.
- Inventario: 5 evaluaciones del PDM 2024-2026 catalogadas con ventana "2022" (buscar por nombre).
- Fases 3 y 4 TERMINADAS. Entregables en `outputs/disponibilidad/`: `disponibilidad_brecha_indicadores.xlsx` (Léeme, Variables 39, Indicadores 23, Fuentes 66, Solicitudes 32), `disponibilidad_variables.csv` y `_excel.csv`, `resumen_reunion.md`, `documentos_por_descargar.md`. Se regenera con `python3 -I scripts/generar_disponibilidad.py` (aplica las 20 correcciones de `trabajo/revision_fase3.md`; se detiene si el conteo no cuadra).
- Conteo FINAL de las 39: Disponible 10 (7 metas), Construible 7, Parcial 19, Proxy 2 (PDM.CS.O, PDM.CD.O), No disponible 1 (PDM.CP.O). Indicadores: 1 Viable (CS.PCS), 10 Viable con ajuste, 12 No viable por ahora. Solo CS.ICRCS calculable (2024, tras definir universos); ISPC, ICPD-6 e IREPP bloqueados (Guía 15).
- PP: EPP 2025 reconstruido (APP 20.928.690.218; OPP 20.684.427.720; ≈98,8 %), complemento a la revisión del profesor Francisco (que lo dejó N/D). Cálculo hecho con scripts temporales: falta script reproducible. Proxy PEC inestable (45-55 de 61 según agregación).
- EPP 2025 reproducible: `python3 -I scripts/epp_pp_2025.py` (salidas `outputs/disponibilidad/epp_pp_2025*`): 61/61 proyectos, 56 iguales al peso, EPP 98,83 % confirmado; marcado "PRELIMINAR, POR VALIDAR" en Excel y resumen. `resumen_reunion.docx` con `python3 scripts/md_a_docx.py` (SIN -I: python-docx, typing_extensions y charset_normalizer están en ~/.local, que -I ignora; pdfplumber falla con -I por eso; usar pymupdf).
- Informe LaTeX/Overleaf: `python3 -I scripts/generar_informe_latex.py` → `outputs/disponibilidad/informe_latex/informe_disponibilidad.tex` + `informe_disponibilidad_overleaf.zip` (excepción en .gitignore). Compilador: tectonic 0.15 descargado al scratchpad de sesión (efímero; volver a bajar el binario musl de GitHub releases). `generar_disponibilidad.py` ahora unifica `sistema` al código (CS en vez de "Control social").
- Excel para profesores (presentación pulida, reproducible en `generar_disponibilidad.py`: JERGA, TILDES, URL desde el inventario, hoja Enlaces). Encabezados nuevos en español ("ID de la variable", "Disponibilidad", "Veredicto"…); `generar_informe_latex.py` los traduce con ALIAS. Las notas históricas de PP.CEP.PEC aún dicen "se reclasifica como Proxy" aunque la categoría final es Parcial (texto heredado de la Fase 2).
- Informe Word: `python3 scripts/generar_informe_docx.py` → `outputs/disponibilidad/informe_disponibilidad.docx` (traduce el .tex de `generar_informe_latex.construir()`; fuente única). LibreOffice (`soffice`, instalado con apt) + fuente Caladea (métrica de Cambria) sirven para previsualizar en PDF. Tectonic vive en el scratchpad y se pierde entre sesiones: volver a descargarlo antes de compilar (si no, el PDF queda viejo sin aviso).
