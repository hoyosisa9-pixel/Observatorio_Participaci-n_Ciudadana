#!/usr/bin/env python3
"""Descarga las fuentes obligatorias de la etapa a data_raw_IIC/06_descargas_indicadores/<emisor>/.

Fuentes (plan §6.3): DANE (población municipal post-COVID, serie VP y edad simple), MinSalud (cifras de afiliación
dic-2025 y ago-2026) y Policía/datos.gov.co (homicidios en accidente de tránsito, filtrado a Envigado).
Forensis 2024 (INMLCF) responde 403 a clientes automáticos: la usuaria lo descarga a mano en
data_raw_IIC/06_descargas_indicadores/inmlcf/. Aquí solo se registra su estado en el manifiesto.

Verificación TLS siempre activa. Si una descarga falla se registra y se sigue.
Uso (desde la raíz, sin -I): python3 scripts/indicadores/descargas.py [--forzar]
"""
from __future__ import annotations

import csv
import datetime as dt
import hashlib
import ssl
import sys
import urllib.error
import urllib.request
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import comun as C  # noqa: E402

DANE = "https://www.dane.gov.co/files/censo2018/proyecciones-de-poblacion/Municipal/"
MINSALUD = "https://minsalud.gov.co/sites/rid/Lists/BibliotecaDigital/RIDE/VP/"
FUENTES = [
    # (emisor, archivo, url, para qué sirve)
    ("dane", "DCD-area-proypoblacion-Mun-2020-2035-ActPostCOVID-19.xlsx",
     DANE + "DCD-area-proypoblacion-Mun-2020-2035-ActPostCOVID-19.xlsx", "Población municipal por área 2020-2035 post-COVID (PDM.CM, auxiliares PDM.CS)"),
    ("dane", "PPED-AreaMun-2018-2042_VP.xlsx", DANE + "PPED-AreaMun-2018-2042_VP.xlsx",
     "Población municipal por área 2018-2042 (serie VP; escenario poblacion_dane)"),
    ("dane", "DCD-area-sexo-edad-proypoblacion-Mun-2020-2035-ActPostCOVID-19.xlsx",
     DANE + "DCD-area-sexo-edad-proypoblacion-Mun-2020-2035-ActPostCOVID-19.xlsx", "Población por edad simple post-COVID (PDM.ILE.A, 0-6 años)"),
    ("minsalud", "cifras-afiliacion-salud-dic-2025.zip", MINSALUD + "DOA/cifras-afiliacion-salud-dic-2025.zip",
     "Afiliación a salud, diciembre 2025 (auxiliares PDM.CS)"),
    ("minsalud", "cifras-afiliacion-salud-agosto2026.zip", MINSALUD + "RBC/cifras-afiliacion-salud-agosto2026.zip",
     "Afiliación a salud, agosto 2026 (auxiliares PDM.CS)"),
    ("policia_datosgov", "uav5-b85g_homicidio_accidente_transito_envigado.csv",
     "https://www.datos.gov.co/resource/uav5-b85g.csv?cod_muni=05266&$limit=50000",
     "Homicidios en accidente de tránsito, Envigado (escenario pdm_cm_fuente=policia_datosgov)"),
]
CAMPOS = ["archivo", "url", "fecha_descarga", "bytes", "sha256", "estado", "observaciones"]
MANIFIESTO = C.DIR_DESCARGAS / "manifiesto.csv"
TAM_MAX = 250 * 1024 * 1024


def descargar(url: str, destino: Path) -> tuple[int, str]:
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (OCE-indicadores; investigacion academica)"})
    ctx = ssl.create_default_context()     # verificación TLS activa
    h, n = hashlib.sha256(), 0
    destino.parent.mkdir(parents=True, exist_ok=True)
    tmp = destino.with_suffix(destino.suffix + ".part")
    with urllib.request.urlopen(req, timeout=120, context=ctx) as r, open(tmp, "wb") as f:
        while chunk := r.read(1 << 20):
            n += len(chunk)
            if n > TAM_MAX:
                raise RuntimeError("archivo mayor al límite de 250 MB")
            h.update(chunk)
            f.write(chunk)
    tmp.replace(destino)
    return n, h.hexdigest()


def main() -> int:
    forzar = "--forzar" in sys.argv
    C.DIR_DESCARGAS.mkdir(parents=True, exist_ok=True)
    previo = {r["archivo"]: r for r in C.leer_csv(MANIFIESTO)} if MANIFIESTO.exists() else {}
    filas = {}
    fallos = 0
    for emisor, archivo, url, para in FUENTES:
        rel = f"{emisor}/{archivo}"
        destino = C.DIR_DESCARGAS / emisor / archivo
        if destino.exists() and rel in previo and not forzar:
            print(f"ya descargado: {rel}")
            filas[rel] = previo[rel]
            continue
        try:
            n, sha = descargar(url, destino)
            filas[rel] = {"archivo": rel, "url": url, "fecha_descarga": dt.date.today().isoformat(), "bytes": n,
                          "sha256": sha, "estado": "ok", "observaciones": para}
            print(f"OK    {rel}  {n/1e6:.2f} MB")
        except (urllib.error.URLError, urllib.error.HTTPError, TimeoutError, OSError, RuntimeError, ssl.SSLError) as e:
            fallos += 1
            filas[rel] = {"archivo": rel, "url": url, "fecha_descarga": dt.date.today().isoformat(), "bytes": "",
                          "sha256": "", "estado": "fallo", "observaciones": f"{type(e).__name__}: {e}. {para}"}
            print(f"FALLO {rel}: {type(e).__name__}: {e}")
            (destino.with_suffix(destino.suffix + ".part")).unlink(missing_ok=True)
    # Forensis 2024: descarga manual de la usuaria; se usa el PDF local si existe, si no la cifra citada
    forensis = C.RAIZ / C.cargar_supuestos()["fuentes_referencia"]["forensis_2024"]["ruta_pdf_local"]
    rel = "inmlcf/Forensis_2024.pdf"
    url_f = C.cargar_supuestos()["fuentes_referencia"]["forensis_2024"]["url"]
    if forensis.exists():
        b = forensis.read_bytes()
        filas[rel] = {"archivo": rel, "url": url_f, "fecha_descarga": dt.date.fromtimestamp(forensis.stat().st_mtime).isoformat(),
                      "bytes": len(b), "sha256": hashlib.sha256(b).hexdigest(), "estado": "ok",
                      "observaciones": "Descarga manual de la usuaria; PDM.CM usa este PDF"}
    else:
        filas[rel] = {"archivo": rel, "url": url_f, "fecha_descarga": "", "bytes": "", "sha256": "",
                      "estado": "pendiente_manual",
                      "observaciones": "El sitio responde 403 a clientes automáticos: descargar a mano en "
                                       "data_raw_IIC/06_descargas_indicadores/inmlcf/. Mientras tanto PDM.CM usa la cifra "
                                       "citada (7,25 por 100.000; Forensis 2024 p. 440) en estado Pendiente."}
    C.escribir_csv(MANIFIESTO, list(filas.values()), CAMPOS)
    print(f"manifiesto: {MANIFIESTO}  ({fallos} fallo(s))")
    return 0


if __name__ == "__main__":
    sys.exit(main())
