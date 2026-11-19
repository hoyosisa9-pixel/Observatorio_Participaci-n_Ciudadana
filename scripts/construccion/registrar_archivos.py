#!/usr/bin/env python3
"""Registrador general de archivos nuevos en data_raw_IIC/07_revision_nov2026/<carpeta>/ con manifiesto.csv.

Para cada (ruta, carpeta destino, url_origen, nota):
  - localiza el archivo: una ruta absoluta existente, la subcarpeta `zip/<ruta>` del scratchpad de extracciones, o, si no esta,
    lo vuelve a extraer del ZIP de Google Drive con `ArchivoRemoto` (scripts/construccion/cs_extraer_de_drive.py; HTTP Range,
    CRC-32 verificado por zipfile). La ruta de un archivo del ZIP es `<NN_subcarpeta>/<ruta dentro del ZIP>`.
  - si la ruta esta en data_raw_IIC/inventario_detallado_IIC.csv, comprueba que los bytes coincidan;
  - copia a data_raw_IIC/07_revision_nov2026/<carpeta>/ (no sobrescribe un archivo distinto), calcula el sha256 y
    agrega la fila al manifiesto (archivo, url, fecha_consulta, bytes, sha256, nota; respeta el separador del manifiesto existente).
Los IDs de Drive NO se escriben aqui (el repositorio es publico): se leen de drive_zips.json por cs_extraer_de_drive.
Nada de lo descargado se ejecuta; son solo datos.

Uso: python3 scripts/construccion/registrar_archivos.py [carpeta_de_extracciones]   (desde la raiz; registra el lote PP/PC/CS del 19-nov-2026)
"""
import csv
import hashlib
import io
import os
import shutil
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import cs_extraer_de_drive as D  # noqa: E402

RAIZ = D.RAIZ
RAW = D.RAW
INV = RAIZ / "data_raw_IIC/inventario_detallado_IIC.csv"
# Carpeta de extracciones del Drive (subcarpetas zip/ y externos/). Se pasa como primer argumento opcional; si no, la variable de
# entorno OCE_SCRATCH_DRIVE o, por defecto, data_raw_IIC/07_revision_nov2026/_extracciones_drive (no depende de la sesion).
SCR = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(os.environ.get("OCE_SCRATCH_DRIVE", RAIZ / "data_raw_IIC/07_revision_nov2026/_extracciones_drive"))
FECHA = "2026-11-19"
NOTA_ZIP = "extraído del ZIP de Google Drive (D85)"
NOTA_CONCEJO = "descargado de concejoenvigado.gov.co"
COLS = ["archivo", "url", "fecha_consulta", "bytes", "sha256", "nota"]


def _inventario():
    with open(INV, encoding="utf-8-sig", newline="") as f:
        return {r["ruta_local"].replace("data_raw_IIC/", "", 1): int(r["bytes"]) for r in csv.DictReader(f)}


def _manifiesto_agregar(carpeta, fila):
    m = RAW / carpeta / "manifiesto.csv"
    if m.exists():
        txt = m.read_text(encoding="utf-8-sig")
        sep = ";" if txt.splitlines()[0].count(";") >= 5 else ","
        bom = m.read_bytes().startswith(b"\xef\xbb\xbf")
        filas = list(csv.reader(io.StringIO(txt), delimiter=sep))
        if any(r and r[0] == fila[0] for r in filas[1:]):
            return "ya estaba en el manifiesto"
    else:
        sep, bom, filas = ",", False, [COLS]
    filas.append(fila)
    with open(m, "w", encoding="utf-8-sig" if bom else "utf-8", newline="") as f:
        csv.writer(f, delimiter=sep).writerows(filas)
    return "fila agregada"


def registrar(items):
    """items: lista de (ruta, carpeta, url_origen, nota). Devuelve la lista de resultados."""
    inv = _inventario()
    por_zip, res = {}, []
    for it in items:
        ruta = it[0]
        if not Path(ruta).is_absolute() and not (SCR / "zip" / ruta).exists():
            por_zip.setdefault(ruta.split("/")[0] + ".zip", []).append(ruta)
    remotos = {}
    for nz, rutas in por_zip.items():
        fid, tam = D.ZIPS[nz]
        rem = D.ArchivoRemoto(D.URL.format(fid), tam)
        with zipfile.ZipFile(io.BufferedReader(rem, buffer_size=1 << 20)) as z:
            ent = {i.filename: i for i in z.infolist()}
            for r in rutas:
                remotos[r] = z.read(ent[r])           # zipfile verifica el CRC-32
    for ruta, carpeta, url, nota in items:
        p = Path(ruta)
        if p.is_absolute():
            datos, nombre, clave = p.read_bytes(), p.name, None
        elif ruta in remotos:
            datos, nombre, clave = remotos[ruta], p.name, ruta
        else:
            datos, nombre, clave = (SCR / "zip" / ruta).read_bytes(), p.name, ruta
        esperado = inv.get(clave) if clave else None
        if esperado is not None and esperado != len(datos):
            raise SystemExit(f"{ruta}: {len(datos)} bytes != inventario {esperado}")
        destino = RAW / carpeta / nombre
        destino.parent.mkdir(parents=True, exist_ok=True)
        if destino.exists() and destino.read_bytes() != datos:
            raise SystemExit(f"{destino} ya existe y es distinto; no se sobrescribe")
        destino.write_bytes(datos)
        h = hashlib.sha256(datos).hexdigest()
        est = _manifiesto_agregar(carpeta, [nombre, url, FECHA, len(datos), h, nota])
        res.append((carpeta, nombre, len(datos), "si" if esperado is not None else ("sin inventario" if clave is None else "no esta en el inventario"), h, est))
    return res


def lote_19nov():
    inv = list(csv.DictReader(open(INV, encoding="utf-8-sig", newline="")))
    U = {r["ruta_local"].replace("data_raw_IIC/", "", 1): r["url_origen"] for r in inv}
    por_nombre = {}
    for k in U:
        por_nombre.setdefault(Path(k).name, []).append(k)

    def z(nombre, carpeta, sub=""):
        ks = [k for k in por_nombre[nombre] if sub in k]
        assert len(ks) == 1, (nombre, ks)
        return (ks[0], carpeta, U[ks[0]], NOTA_ZIP)
    L = [z("2021proyectosppp.pdf", "pp_listados_oficiales"), z("2022proyectosppp.pdf", "pp_listados_oficiales"),
         z("proyectos-pp2020-para-pagina.pdf", "pp_listados_oficiales"), z("2022saludz42021.pdf", "pp_listados_oficiales"),
         z("proyectos-priorizados-2021.pdf", "pp_votaciones")]
    L += [z(n, "planes_zonales_2023") for n in sorted(k for k in por_nombre if k.startswith("pze_z"))]
    L += [z("decreto-242-de-2017.pdf", "normativa_pp"), z("2020periodicorendicioncuentas.pdf", "rendicion_cuentas"),
          z("pdf-espacios-zonales.pdf", "pc_cobertura_zonal"), z("veedurias-envi2022v2.pdf", "personeria_jju5br3w")]
    L += [(str(SCR / "externos/Acuerdo-007-de-2025.pdf"), "normativa_pp", "https://www.concejoenvigado.gov.co/wp-content/uploads/2025/02/Acuerdo-007-de-2025.pdf", NOTA_CONCEJO),
          (str(SCR / "externos/Acta-096-de-2026.pdf"), "normativa_pp", "https://www.concejoenvigado.gov.co/wp-content/uploads/2026/09/Acta-096-de-2026.pdf", NOTA_CONCEJO),
          (str(SCR / "externos/147324_rc2025.pdf"), "rendicion_cuentas",
           "https://alcaldia-municipal-de-envigado-en-antioquia.micolombiadigital.gov.co/sites/alcaldia-municipal-de-envigado-en-antioquia/content/files/002947/147324_rc2025.pdf",
           "descargado del portal de la Alcaldía (micolombiadigital); tomado del scratchpad del Analizador")]
    # lista CSV de las 93 fichas de espacios zonales 2020 (solo la lista; los PDF no se copian)
    fichas = [r for r in inv if "espacios-zonales-de-deliberacion-2020/" in r["ruta_local"]]
    assert len(fichas) == 93, len(fichas)
    out = RAW / "pc_cobertura_zonal" / "fichas_espacios_zonales_2020_lista.csv"
    out.parent.mkdir(parents=True, exist_ok=True)
    with open(out, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["archivo", "ruta_dentro_del_zip_01", "bytes", "url_origen", "zona_en_el_nombre"])
        import re
        for r in sorted(fichas, key=lambda r: r["archivo"]):
            m = re.search(r"(?:zona-|-z-?)(\d{1,2})(?!\d)", r["archivo"])
            w.writerow([r["archivo"], r["ruta_local"].replace("data_raw_IIC/", "", 1), r["bytes"], r["url_origen"], m.group(1) if m else ""])
    L.append((str(out), "pc_cobertura_zonal", "lista construida desde data_raw_IIC/inventario_detallado_IIC.csv",
              "lista de las 93 fichas de los espacios zonales de 2020 (ZIP 01, espacios-zonales-de-deliberacion-2020/); generada, no es un documento fuente"))
    return L


if __name__ == "__main__":
    res = registrar(lote_19nov())
    print("| carpeta | archivo | bytes | coincide con inventario | manifiesto |")
    for c, n, b, i, h, e in res:
        print(f"| {c} | {n} | {b} | {i} | {e} |")
