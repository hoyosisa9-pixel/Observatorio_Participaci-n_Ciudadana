"""Extrae del ZIP original de Google Drive solo los archivos que faltan (PQRSDF y actas), sin bajar los 2,8 GB.

Contexto: desde el 19-nov-2026 el servidor de archivos de www.envigado.gov.co devuelve HTTP 500 para todos los PDF.
Los mismos archivos estan en 02_ejecucion_y_presupuesto.zip (Google Drive de la usuaria, carpeta Bases_Datos_IIC_Envigado),
que ella autorizo leer. El ZIP se lee con peticiones HTTP Range: indice central + cada entrada pedida.

Entrada:  data_raw_IIC/07_revision_nov2026/LISTA_DESCARGA_MANUAL.csv (columna ruta_dentro_del_zip y bytes_esperados),
          generada por `python3 scripts/construccion/cs_descargas.py --lista`.
Salida:   data_raw_IIC/07_revision_nov2026/manual/<ruta_dentro_del_zip>. Luego se registran con
          `python3 scripts/construccion/cs_descargas.py --registrar-manuales` (verifica bytes contra el inventario y escribe el manifiesto).
Controles: zipfile verifica el CRC-32 de cada entrada al leerla; se compara el tamano con el inventario; nada se ejecuta.
"""
import csv
import io
import json
import sys
import urllib.request
import zipfile
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[2]
RAW = RAIZ / "data_raw_IIC/07_revision_nov2026"
# Los IDs de Google Drive de los ZIP de la usuaria NO se versionan (el repositorio es publico):
# viven en data_raw_IIC/07_revision_nov2026/drive_zips.json (data_raw_IIC/ esta en .gitignore), con el formato
# {"02_ejecucion_y_presupuesto.zip": {"id": "<id de Drive>", "bytes": 2840613915}, ...}
_CFG = RAW / "drive_zips.json"
ZIPS = {k: (v["id"], v["bytes"]) for k, v in json.loads(_CFG.read_text(encoding="utf-8")).items()} if _CFG.exists() else {}
URL = "https://drive.usercontent.google.com/download?id={}&export=download&confirm=t"


class ArchivoRemoto(io.RawIOBase):
    """Archivo de solo lectura sobre HTTP Range (lo justo para zipfile)."""

    def __init__(self, url, tamano):
        self.url, self.tamano, self.pos, self.peticiones = url, tamano, 0, 0

    def readable(self):
        return True

    def seekable(self):
        return True

    def tell(self):
        return self.pos

    def seek(self, off, whence=0):
        self.pos = off if whence == 0 else (self.pos + off if whence == 1 else self.tamano + off)
        return self.pos

    def read(self, n=-1):
        if n is None or n < 0:
            n = self.tamano - self.pos
        if n == 0 or self.pos >= self.tamano:
            return b""
        fin = min(self.pos + n, self.tamano) - 1
        req = urllib.request.Request(self.url, headers={"Range": f"bytes={self.pos}-{fin}", "User-Agent": "oce-extractor/1.0"})
        with urllib.request.urlopen(req, timeout=120) as r:
            if r.status != 206:
                raise IOError(f"Drive no respondio con 206 (respondio {r.status})")
            datos = r.read()
        self.peticiones += 1
        self.pos += len(datos)
        return datos

    def readinto(self, b):
        d = self.read(len(b))
        b[:len(d)] = d
        return len(d)


def main():
    lista = list(csv.DictReader((RAW / "LISTA_DESCARGA_MANUAL.csv").open(encoding="utf-8-sig"), delimiter=";"))
    destino = RAW / "manual"
    ok, mal = [], []
    for nombre_zip, (fid, tamano) in ZIPS.items():
        pedidos = [x for x in lista if x["zip_de_google_drive"] == nombre_zip]
        if not pedidos:
            continue
        remoto = ArchivoRemoto(URL.format(fid), tamano)
        with zipfile.ZipFile(io.BufferedReader(remoto, buffer_size=1 << 20)) as z:
            entradas = {i.filename: i for i in z.infolist()}
            print(f"{nombre_zip}: {len(entradas)} entradas en el indice")
            for x in pedidos:
                ruta = x["ruta_dentro_del_zip"]
                info = entradas.get(ruta)
                if info is None:
                    mal.append((ruta, "no esta en el ZIP"))
                    continue
                if x["bytes_esperados"] and str(info.file_size) != x["bytes_esperados"]:
                    mal.append((ruta, f"tamano {info.file_size} != inventario {x['bytes_esperados']}"))
                    continue
                datos = z.read(info)  # zipfile verifica el CRC-32
                out = destino / ruta
                out.parent.mkdir(parents=True, exist_ok=True)
                if out.exists() and out.stat().st_size == len(datos):
                    ok.append(ruta)
                    continue
                out.write_bytes(datos)
                ok.append(ruta)
        print(f"  peticiones HTTP: {remoto.peticiones}")
    for nombre_zip in {x["zip_de_google_drive"] for x in lista} - set(ZIPS):
        for x in lista:
            if x["zip_de_google_drive"] == nombre_zip:
                mal.append((x["ruta_dentro_del_zip"], f"{nombre_zip}: falta su id de Google Drive en ZIPS"))
    print(f"Extraidos (CRC y tamano OK): {len(ok)} / con problema: {len(mal)}")
    for r, m in mal:
        print("  PROBLEMA:", r, "-", m)
    return 1 if mal else 0


if __name__ == "__main__":
    sys.exit(main())
