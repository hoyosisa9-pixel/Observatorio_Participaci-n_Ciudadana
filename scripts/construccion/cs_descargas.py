#!/usr/bin/env python3
"""Descargas de la dimension Control social (CS): informes trimestrales de PQRSDF 2020-2025 y actas de rendicion de cuentas de Salud 2021-2023.

Lista de URL: data_raw_IIC/inventario_detallado_IIC.csv (columna url_origen).
  - PQRSDF: archivos PDF cuya URL o nombre contiene 'pqr', periodos 2020-2025 (se omiten los xls/xlsx de 'analisis de calidad' y el ano 2026).
  - Rendicion de cuentas (VRC, D77 y D81): documentos 6384 (acta de feb-2021), 6387, 24141, 34258, 55874, 111982, 111981 (actas de Salud 2021-2023) y 40433.
Destino: data_raw_IIC/07_revision_nov2026/pqrs/ y .../rendicion_cuentas/, cada carpeta con manifiesto.csv
  (archivo, url, fecha_consulta, bytes, sha256, nota con periodo y sistema segun el inventario; el periodo real lo confirma el lector de cs_icrcs.py).
Idempotente: lo que ya esta en el manifiesto con el sha256 correcto no se baja de nuevo; nunca se sobrescribe un archivo existente.
Servidores: www.envigado.gov.co (curl -k: cadena de certificado incompleta, autorizado por la usuaria) y su espejo micolombiadigital (curl normal).
Si el portal no responde, lo dice y termina sin error (los pendientes se bajan al volver a correr el script).
Los archivos descargados se tratan solo como datos: aqui nunca se abren ni se ejecutan; solo se calcula su sha256 y se comprueba su firma (%PDF o PK).
Uso: python3 scripts/construccion/cs_descargas.py   (desde la raiz)
"""
import csv
import datetime as dt
import hashlib
import re
import subprocess
import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[2]
RAW = RAIZ / "data_raw_IIC/07_revision_nov2026"
INVENTARIO = RAIZ / "data_raw_IIC/inventario_detallado_IIC.csv"
HOY = dt.date.today().isoformat()
HOST_PRINCIPAL = "www.envigado.gov.co"
HOST_ESPEJO = "alcaldia-municipal-de-envigado-en-antioquia.micolombiadigital.gov.co"
CONNECT_S, MAX_S = 8, 25
IDS_RC = ["6384", "6387", "24141", "34258", "55874", "111982", "111981", "40433"]
COLS_MAN = ["archivo", "url", "fecha_consulta", "bytes", "sha256", "nota"]
SUFIJO = {"pdf": b"%PDF", "docx": b"PK", "xlsx": b"PK"}


def sha256(p):
    return hashlib.sha256(Path(p).read_bytes()).hexdigest()


def leer_manifiesto(carpeta):
    p = carpeta / "manifiesto.csv"
    if not p.exists():
        return []
    with open(p, encoding="utf-8-sig", newline="") as f:
        sep = ";" if ";" in f.readline() else ","
    with open(p, encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f, delimiter=sep))


def escribir_manifiesto(carpeta, filas):
    with open(carpeta / "manifiesto.csv", "w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=COLS_MAN)
        w.writeheader()
        for r in filas:
            w.writerow({k: r.get(k, "") for k in COLS_MAN})


def adivinar_periodo(fila):
    """(anio, periodo) segun el inventario: periodo = 'T1'..'T4', 'S1', 'S2' o None. Primero el nombre del archivo y despues la pagina de origen.
    Es solo una pista (el lector de cs_icrcs.py confirma el periodo real leyendo el informe)."""
    def de(texto):
        t = texto.lower()
        m = re.search(r"trimestre[-_]?([1-4])(?=20\d\d)", t) or re.search(r"(?:trimestre|trim)[-_ ]?([1-4])(?!\d)", t)
        if m:
            return "T" + m.group(1)
        if "semestr" in t:
            return "S2" if ("segundo" in t or "ii" in t.split("semestre")[-1][:3]) else "S1"
        for pat, q in (("primer", "T1"), ("segundo", "T2"), ("sgundo", "T2"), ("tercer", "T3"), ("cuarto", "T4")):
            if pat in t and "trimestre" in t:
                return q
        return None
    nombre = fila["archivo"]
    per = de(nombre) or de(fila["pagina_fuente"])
    a = re.search(r"(20[12]\d)", nombre) or re.search(r"(20[12]\d)", fila["pagina_fuente"]) or re.search(r"(20[12]\d)", fila.get("ventana_temporal", ""))
    return (int(a.group(1)) if a else None), per


def periodo_inventario(fila):
    anio, per = adivinar_periodo(fila)
    sis = "Educacion (SAC)" if "educacion" in (fila["archivo"] + fila["pagina_fuente"]).lower() else "BPMS y SAC (por confirmar en el informe)"
    return f"{anio or '?'} {per or 'periodo por confirmar'}", sis


def seleccionar(filas):
    """(carpeta, nombre_destino, url, nota) por cada documento a bajar; sin duplicados de URL."""
    out, vistos = [], set()
    for r in filas:
        u = r["url_origen"]
        ul = (u + " " + r["archivo"]).lower()
        if u in vistos or u.rsplit("/", 1)[-1] in vistos:
            continue
        if "pqr" in ul and u.lower().endswith(".pdf") and "calidad" not in ul and "2026" not in r["pagina_fuente"]:
            per, sis = periodo_inventario(r)
            vistos.update((u, u.rsplit("/", 1)[-1]))
            out.append(("pqrs", u.rsplit("/", 1)[-1], u, f"Informe PQRSDF; periodo segun inventario: {per}; sistema: {sis}; pagina de origen: {r['pagina_fuente']}"))
    for idd in IDS_RC:
        for r in filas:
            if re.search(r"/" + idd + r"_", r["url_origen"]) and r["url_origen"] not in vistos:
                vistos.add(r["url_origen"])
                uso = "documento de RPO 2021 (respuestas escritas de jul-2022; NO es un evento de VRC, D81)" if idd == "40433" else "VRC (D77, D81)"
                out.append(("rendicion_cuentas", r["url_origen"].rsplit("/", 1)[-1], r["url_origen"],
                            f"Rendicion de cuentas ({uso}): {r['titulo_publicacion']}; vigencia {r['ventana_temporal']}; sistema: no aplica; pagina de origen: {r['pagina_fuente']}"))
                break
        else:
            print(f"AVISO: el documento {idd} no esta en el inventario")
    return out


def curl(url, destino, inseguro):
    cmd = ["curl", "-sS", "-L", "--connect-timeout", str(CONNECT_S), "-m", str(MAX_S), "-o", str(destino), "-w", "%{http_code}"]
    if inseguro:
        cmd.insert(1, "-k")
    try:
        p = subprocess.run(cmd + [url], capture_output=True, text=True, timeout=MAX_S + 10)
        return p.stdout.strip(), p.stderr.strip()[:120]
    except subprocess.TimeoutExpired:
        return "000", "timeout"


def host_vivo(host):
    tmp = RAW / "_sonda.tmp"
    cod, err = curl(f"https://{host}/", tmp, inseguro=(host == HOST_PRINCIPAL))
    tmp.unlink(missing_ok=True)
    return cod.startswith(("2", "3")), (cod + " " + err).strip()


def main():
    with open(INVENTARIO, encoding="utf-8-sig", newline="") as f:
        filas = list(csv.DictReader(f))
    lista = seleccionar(filas)
    print(f"Documentos a gestionar: {len(lista)} ({sum(1 for x in lista if x[0] == 'pqrs')} de PQRSDF, {sum(1 for x in lista if x[0] != 'pqrs')} de rendicion de cuentas)")
    vivos = {}
    for h in (HOST_PRINCIPAL, HOST_ESPEJO):
        vivos[h], det = host_vivo(h)
        print(f"  sonda {h}: {'responde' if vivos[h] else 'NO responde'} ({det})")
    mans = {c: leer_manifiesto(RAW / c) for c in ("pqrs", "rendicion_cuentas")}
    desc, ya, fall = [], [], []
    for carpeta, nombre, url, nota in lista:
        d = RAW / carpeta
        d.mkdir(parents=True, exist_ok=True)
        # ya esta: la URL figura en el manifiesto con el archivo presente y el sha256 correcto
        previo = [r for r in mans[carpeta] if r["url"].rsplit("/", 1)[-1] == nombre]
        if previo:
            r = previo[0]
            p = d / r["archivo"]
            if p.exists() and sha256(p) == r["sha256"]:
                ya.append(r["archivo"])
                continue
            fall.append((nombre, "esta en el manifiesto pero el archivo falta o su sha256 no coincide: no se sobrescribe, revisar a mano"))
            continue
        dest = d / nombre
        if dest.exists():
            fall.append((nombre, "existe un archivo con ese nombre que no esta en el manifiesto: no se sobrescribe, revisar a mano"))
            continue
        if not any(vivos.values()):
            fall.append((nombre, "portal sin respuesta (no se intento)"))
            continue
        ok, motivo = False, ""
        ruta = url.split("/sites/", 1)[1]
        for host in (HOST_PRINCIPAL, HOST_ESPEJO):
            if not vivos[host]:
                continue
            u = f"https://{host}/sites/{ruta}"
            tmp = d / (nombre + ".part")
            cod, err = curl(u, tmp, inseguro=(host == HOST_PRINCIPAL))
            ext = nombre.rsplit(".", 1)[-1].lower()
            if cod == "200" and tmp.exists() and tmp.stat().st_size > 0 and tmp.read_bytes()[:4].startswith(SUFIJO.get(ext, b"")):
                tmp.rename(dest)
                mans[carpeta].append({"archivo": nombre, "url": u, "fecha_consulta": HOY, "bytes": dest.stat().st_size, "sha256": sha256(dest),
                                      "nota": nota + ("; envigado.gov.co descargado con curl -k (cadena de certificado incompleta; autorizado por la usuaria)" if host == HOST_PRINCIPAL else "")})
                ok = True
                break
            motivo = f"{host}: HTTP {cod} {err}".strip()
            tmp.unlink(missing_ok=True)
        (desc if ok else fall).append(nombre if ok else (nombre, motivo))
        if ok:
            escribir_manifiesto(d, mans[carpeta])       # se guarda tras cada descarga
    print("\n=== Resultado ===")
    print(f"descargados: {len(desc)} / ya estaban: {len(ya)} / fallidos: {len(fall)}")
    for n in desc:
        print("  descargado:", n)
    for n, m in fall[:60]:
        print("  fallido:", n, "-", m)
    if fall and not any(vivos.values()):
        print("\nEl portal (www.envigado.gov.co y su espejo micolombiadigital) no responde hoy. Vuelve a correr este script cuando responda;"
              " solo bajara lo que falta.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
