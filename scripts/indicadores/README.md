# scripts/indicadores - etapa "organizar datos y calcular indicadores viables"

Todo es PROVISIONAL: estado "Pendiente" en las variables, "Provisional" en los resultados, nada "Validado".
Ejecutar siempre desde la raíz del repositorio y **sin `-I`** (PyYAML y openpyxl están en el site-packages del usuario):
`python3 scripts/indicadores/<script>.py`.

## Orden de ejecución
1. `config/supuestos.yaml` (36 parámetros PROVISIONAL; lo mantiene el coordinador).
2. `descargas.py` descarga a `data_raw_IIC/06_descargas_indicadores/` (DANE, MinSalud, Policía) y escribe `manifiesto.csv`.
   Forensis 2024 se deja a mano en `.../inmlcf/Forensis_2024.pdf`; sin el PDF se usa la cifra citada (7,25; p. 440) en estado Pendiente.
3. Scripts por sistema (`cs_*.py`, `pc_*.py`, `pdm_*.py`, `pp_*.py`), que usan `comun.py`. `cs_icrcs.py` va al final.
   Orden probado: `cs_pcs cs_dsi pdm_metas pdm_ca pdm_cm pdm_no_viables pc_i5 pc_i2 pc_i3 pc_no_viables pp_epp pp_no_viables`.
   En PDM el orden importa: `pdm_metas.py` reescribe las filas de metas por código de indicador, y si se corre después de
   `pdm_cm.py` o `pdm_no_viables.py` borra sus filas; en ese caso, volver a correr esos dos. `pp_epp.py` va antes que
   `pp_no_viables.py` (este no toca `pp_proyectos`).
4. `ensamblar.py` consolida los parciales en `outputs/indicadores/` (CSV `;` UTF-8 con BOM, `base_indicadores.xlsx`, `diccionario_datos`).
5. `llenar_matriz.py` crea `outputs/indicadores/matriz_diligenciada.xlsx` (hojas Registro y PP_Proyectos, escenario por defecto).
   `--verificar` recalcula en LibreOffice una copia temporal "Validado" y compara Resultados con `resultados_indicadores`.
   Requiere `libreoffice-calc` (no solo writer). `--zonas` añade los resultados por zona.

## Dónde escribe cada sistema
Solo en `outputs/indicadores/parciales/<SIS>/` (SIS = CS, PC, PDM, PP) mediante `comun.escribir_parcial`:
`valores_variables.csv`, `resultados_indicadores.csv`, `soporte_conteos.csv`, `pp_proyectos.csv` (PP) y `clasificacion_<tema>.csv`.
Pasar `codigos=[...]` reemplaza solo las filas de esos indicadores (varios scripts del mismo sistema conviven en un archivo).
Insumos manuales versionados: `outputs/indicadores/insumos_manual/<SIS>_<tema>.csv`. Nadie edita `comun.py` ni `supuestos.yaml`.

## Cambiar un supuesto y recalcular
1. Editar `defecto` del parámetro en `config/supuestos.yaml` (debe estar entre sus `opciones`; se valida al cargar).
2. Volver a ejecutar los scripts por sistema afectados, luego `ensamblar.py` y `llenar_matriz.py`.
   Los escenarios (`comun.escenarios(codigo)`) se regeneran solos; el escenario "defecto" siempre es el del YAML vigente.
3. Pruebas del módulo común: `python3 scripts/indicadores/comun.py --pruebas`.
