# Ejecución presupuestal del Presupuesto Participativo 2025 (PP.EPP)

**Estado: PRELIMINAR, POR VALIDAR.** La lista de proyectos PP se reconstruye; debe validarla el DAP. Generado por `scripts/epp_pp_2025.py`.

- Fuente financiera: `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/seguimiento-a-proyectos-2025/evaluaciondefinitivaplanaccionivtri.xlsx`, hoja `PlanAcción27.01.2026` (corte 07-01-2026).
- Contraste: `data_raw_IIC/01_propuestas_y_acuerdos/efpc_envigado_gov_co/efpc_2025-07/Proyectos_PP_2025.pdf`.
- Proyectos PP (nombre con "ZONA n", sin el 0553): **61**; en el listado EFPC: **61**. Apropiación igual al valor EFPC al peso: 56 de 61 (el resto difiere por adiciones).
- Proyectos por zona iguales en ambas fuentes: sí.
- **APP = $20.928.690.218** · **OPP = $20.684.427.720** · **EPP = 98,83 %**.
- Con el proyecto 0553 (costo del proceso: APP $776.957.999, OPP $773.246.790) el EPP sería 98,86 %. Se excluye por defecto; decisión metodológica pendiente.

Detalle por zona en `epp_pp_2025_zonas.csv` y por proyecto en `epp_pp_2025_proyectos.csv`.
