#!/usr/bin/env python3
"""Descargas de la dimension Control social (CS): informes trimestrales de PQRSDF 2020-2025 y actas de rendicion de cuentas de Salud 2021-2023.

Lista de URL: data_raw_IIC/inventario_detallado_IIC.csv (columna url_origen).
  - PQRSDF: archivos PDF cuya URL o nombre contiene 'pqr', periodos 2020-2025 (se omiten los xls/xlsx de 'analisis de calidad' y el ano 2026).
  - Rendicion de cuentas (VRC, D77 y D81): documentos 6384 (acta de feb-2021), 6387, 24141, 34258, 55874, 111982, 111981 (actas de Salud 2021-2023) y 40433.
Destino: data_raw_IIC/07_revision_nov2026/pqrs/ y .../rendicion_cuentas/, cada carpeta con manifiesto.csv
Opciones: --lista (escribe LISTA_DESCARGA_MANUAL.csv con lo que falta) y --registrar-manuales (anota en el manifiesto los archivos bajados a mano)
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
IDS_RC = ["6384", "6387", "24141", "34258", "55874", "111982", "111981", "40433", "40434"]
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
                uso = "encuesta (no respuestas escritas) del grupo 'Informe de respuesta a las solicitudes' de jul-2022: " + ("encuesta de la rendicion de PP 2021" if idd == "40433" else "encuesta de evaluacion de la audiencia 2021") + "; verificado, no responde las preguntas 4 y 9 de RPO 2021; NO es un evento de VRC (D81)" if idd in ("40433", "40434") else "VRC (D77, D81)"
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


def lista_manual(lista):
    """Escribe la lista de lo que falta, para descargar a mano desde el navegador."""
    mans = {c: leer_manifiesto(RAW / c) for c in ("pqrs", "rendicion_cuentas")}
    faltan = []
    for carpeta, nombre, url, nota in lista:
        if any(r["url"].rsplit("/", 1)[-1] == nombre for r in mans[carpeta]):
            continue
        faltan.append((carpeta, nombre, url, nota))
    salida = RAW / "LISTA_DESCARGA_MANUAL.csv"
    with open(salida, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";")
        inv = inventario_por_url()
        w.writerow(["n", "carpeta_destino", "nombre_exacto_del_archivo", "enlace", "que_es", "zip_de_google_drive", "ruta_dentro_del_zip", "bytes_esperados"])
        for i, (c, n, u, nota) in enumerate(faltan, 1):
            ruta = inv.get(u, {}).get("ruta_local", "").replace("data_raw_IIC/", "")
            w.writerow([i, f"data_raw_IIC/07_revision_nov2026/{c}/", n, u, nota.split(";")[0] + ("; " + nota.split(";")[1] if ";" in nota else ""),
                        (ruta.split("/")[0] + ".zip") if ruta else "", ruta, inv.get(u, {}).get("bytes", "")])
    print(f"Faltan {len(faltan)} archivos. Lista: {salida.relative_to(RAIZ)}")
    return faltan


def inventario_por_url():
    with open(INVENTARIO, encoding="utf-8-sig", newline="") as f:
        return {r["url_origen"]: r for r in csv.DictReader(f)}


def registrar_manuales(lista):
    """Registra en el manifiesto los archivos que la usuaria trajo a mano.

    Acepta dos formas: (a) el archivo ya esta en su carpeta con el nombre exacto de la lista; (b) el archivo esta en
    data_raw_IIC/07_revision_nov2026/manual/ con el nombre que tiene dentro del ZIP de Google Drive (nombre_original
    del inventario) o con el nombre exacto. En (b) se verifica que los bytes coincidan con el inventario y se copia a su carpeta.
    """
    import shutil
    inv = inventario_por_url()
    manual = RAW / "manual"
    mans = {c: leer_manifiesto(RAW / c) for c in ("pqrs", "rendicion_cuentas")}
    reg, malos = [], []
    for carpeta, nombre, url, nota in lista:
        d = RAW / carpeta
        p = d / nombre
        if any(r["archivo"] == nombre for r in mans[carpeta]):
            continue
        fila_inv = inv.get(url, {})
        origen_zip = ""
        if not p.exists() and manual.exists():
            nombres = {nombre, fila_inv.get("nombre_original", ""), fila_inv.get("archivo", "")} - {""}
            for cand in sorted(x for x in manual.rglob("*") if x.is_file() and x.name in nombres):
                if True:
                    esperado = fila_inv.get("bytes", "")
                    if esperado and str(cand.stat().st_size) != esperado:
                        malos.append(f"{cand.name} (tamano {cand.stat().st_size} B; el inventario dice {esperado} B: no es el mismo archivo)")
                        break
                    d.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(cand, p)
                    origen_zip = f"; extraido del ZIP de Google Drive ({fila_inv.get('ruta_local', '').replace('data_raw_IIC/', '')}), tamano igual al inventario ({esperado} B)"
                    break
        if not p.exists():
            continue
        ext = nombre.rsplit(".", 1)[-1].lower()
        if p.stat().st_size == 0 or not p.read_bytes()[:4].startswith(SUFIJO.get(ext, b"")):
            malos.append(nombre)
            continue
        mans[carpeta].append({"archivo": nombre, "url": url, "fecha_consulta": HOY, "bytes": p.stat().st_size, "sha256": sha256(p),
                              "nota": nota + (origen_zip or "; descargado a mano por la usuaria desde el navegador (el servidor de archivos de la Alcaldia fallaba)")})
        reg.append(nombre)
    for c in ("pqrs", "rendicion_cuentas"):
        escribir_manifiesto(RAW / c, mans[c])
    print(f"Registrados: {len(reg)}")
    for n in reg:
        print("  registrado:", n)
    for n in malos:
        print("  NO registrado (no parece PDF/Word o esta vacio; vuelve a bajarlo):", n)
    return 0


def main():
    with open(INVENTARIO, encoding="utf-8-sig", newline="") as f:
        filas = list(csv.DictReader(f))
    lista = seleccionar(filas)
    if "--lista" in sys.argv:
        lista_manual(lista)
        return 0
    if "--registrar-manuales" in sys.argv:
        return registrar_manuales(lista)
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
