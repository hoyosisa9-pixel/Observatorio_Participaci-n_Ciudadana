#!/usr/bin/env python3
"""Descargas nuevas del sistema simplificado (provisional). Guarda cada archivo en
data_raw_IIC/06_descargas_indicadores/<fuente>/ y lo registra en manifiesto.csv (sin tocar las filas existentes).
Verificación TLS activa; User-Agent genérico. Si una fuente falla se anota y se sigue.
Uso (desde la raíz, sin -I): python3 scripts/sistema_simplificado/descargas_nuevas.py
"""
from __future__ import annotations
import csv, datetime as dt, hashlib, ssl, sys, urllib.error, urllib.parse, urllib.request
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[2]
DESC = RAIZ / "data_raw_IIC" / "06_descargas_indicadores"
MANIFIESTO = DESC / "manifiesto.csv"
CAMPOS = ["archivo", "url", "fecha_descarga", "bytes", "sha256", "estado", "observaciones"]
UA = "Mozilla/5.0 (observatorio-indicadores; investigacion academica)"
SOC = "https://www.datos.gov.co/resource/"
DANE = "https://www.dane.gov.co/files/censo2018/proyecciones-de-poblacion/Municipal/"


def soql(recurso: str, **kw) -> str:
    return SOC + recurso + ".csv?" + urllib.parse.urlencode({"$" + k: v for k, v in kw.items()}, quote_via=urllib.parse.quote)


FUENTES = [
    ("icfes", "saber11_envigado_agregado_kgxf-xxbe.csv",
     soql("kgxf-xxbe", select="periodo, cole_naturaleza, cole_area_ubicacion, avg(punt_global::number) as puntaje_promedio, count(*) as n_estudiantes",
          where="cole_mcpio_ubicacion='ENVIGADO' and periodo>='20191'", group="periodo, cole_naturaleza, cole_area_ubicacion",
          order="periodo", limit="1000"),
     "Saber 11 (Resultados únicos, ICFES), puntaje global promedio de Envigado por periodo, naturaleza y área (consulta agregada SoQL)"),
    ("icfes", "saber11_envigado_agregado_rnvb-vnyh_2020-2.csv",
     soql("rnvb-vnyh", select="periodo, cole_naturaleza, cole_area_ubicacion, avg(punt_global::number) as puntaje_promedio, count(*) as n_estudiantes",
          where="cole_mcpio_ubicacion like '%NVIGADO%'", group="periodo, cole_naturaleza, cole_area_ubicacion", order="periodo", limit="1000"),
     "Saber 11 calendario A 2020 (periodo 20204), puntaje promedio de Envigado (consulta agregada SoQL)"),
    ("minsalud_datosgov", "rlcpd_envigado_por_anio_registro_kjje-h58h.csv",
     soql("kjje-h58h", select="date_extract_y(fecharegistro) as anio_registro, sum(numpersonas::number) as personas_registradas",
          where="municipioregistro='05266 - Envigado'", group="anio_registro", order="anio_registro", limit="200"),
     "Registro de Localización y Caracterización de Personas con Discapacidad (MinSalud), Envigado, personas registradas por año de registro"),
    ("dane", "PPED-AreaSexoEdadMun-2018-2042_VP.xlsx", DANE + "PPED-AreaSexoEdadMun-2018-2042_VP.xlsx",
     "Población por área, sexo y edad simple, proyección 2018-2042 (misma serie PPED que la población total usada para salud)"),
]


def descargar(url: str, destino: Path) -> tuple[int, str]:
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    h, n = hashlib.sha256(), 0
    destino.parent.mkdir(parents=True, exist_ok=True)
    tmp = destino.with_suffix(destino.suffix + ".part")
    with urllib.request.urlopen(req, timeout=180, context=ssl.create_default_context()) as r, open(tmp, "wb") as f:
        while chunk := r.read(1 << 20):
            n += len(chunk)
            if n > 250 * 1024 * 1024:
                raise RuntimeError("archivo mayor al límite de 250 MB")
            h.update(chunk); f.write(chunk)
    tmp.replace(destino)
    return n, h.hexdigest()


def leer_manifiesto() -> list[dict]:
    if not MANIFIESTO.exists():
        return []
    with open(MANIFIESTO, encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f, delimiter=";"))


def main() -> int:
    filas = leer_manifiesto()
    por_arch = {r["archivo"]: r for r in filas}
    for emisor, archivo, url, para in FUENTES:
        rel = f"{emisor}/{archivo}"
        destino = DESC / emisor / archivo
        if destino.exists() and por_arch.get(rel, {}).get("estado") == "ok":
            print("ya descargado:", rel); continue
        try:
            n, sha = descargar(url, destino)
            reg = {"archivo": rel, "url": url, "fecha_descarga": dt.date.today().isoformat(), "bytes": n, "sha256": sha,
                   "estado": "ok", "observaciones": para + " [sistema simplificado, provisional]"}
            print(f"OK    {rel} {n/1e6:.2f} MB")
        except (urllib.error.URLError, TimeoutError, OSError, RuntimeError, ssl.SSLError) as e:
            reg = {"archivo": rel, "url": url, "fecha_descarga": dt.date.today().isoformat(), "bytes": "", "sha256": "",
                   "estado": "fallo", "observaciones": f"{type(e).__name__}: {e}. {para}"}
            print(f"FALLO {rel}: {e}")
            destino.with_suffix(destino.suffix + ".part").unlink(missing_ok=True)
        if rel in por_arch:
            filas[filas.index(por_arch[rel])] = reg
        else:
            filas.append(reg)
        por_arch[rel] = reg
    with open(MANIFIESTO, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=CAMPOS, delimiter=";"); w.writeheader(); w.writerows(filas)
    return 0


if __name__ == "__main__":
    sys.exit(main())
