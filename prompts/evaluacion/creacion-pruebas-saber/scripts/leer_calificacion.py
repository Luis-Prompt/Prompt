#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
leer_calificacion.py — Lee la plantilla de calificación ya diligenciada y la convierte en JSON.

Uso:  python3 leer_calificacion.py Calificacion.xlsx > respuestas.json
Lee las respuestas CRUDAS que escribió el docente (no los valores calculados), así funciona aunque el archivo
nunca se haya abierto en Excel. Devuelve:
  { "parametros": { exigencia, notaMin, notaMax, basico, alto, superior, preguntasMC, puntosAbiertas, hayVersionB },
    "estudiantes": [ { "fila", "nombre", "version": "A"|"B", "respuestas": ["A", null, "C", ...], "puntosAbiertas": 3.0|null } ],
    "prueba": { meta, preguntas, versionB, abiertasInfo }   (los datos de la prueba que lleva la hoja oculta «DatosPrueba»; null si no existe) }
Solo se incluyen las filas con nombre. Una celda vacía o con una letra no válida queda como null (no contestó).
"""
import json, sys
from openpyxl import load_workbook

def numero(v, defecto):
    try: return float(v)
    except (TypeError, ValueError): return defecto

def main():
    if len(sys.argv) < 2: sys.exit("Uso: leer_calificacion.py Calificacion.xlsx")
    wb = load_workbook(sys.argv[1], data_only=False)
    for hoja in ("Estudiantes", "Parametros"):
        if hoja not in wb.sheetnames: sys.exit(f"❌ El libro no tiene la hoja «{hoja}». ¿Es la plantilla generada por build_calificacion.py?")
    P, E = wb["Parametros"], wb["Estudiantes"]
    n = int(numero(P["C9"].value, 0)); pts_ab = numero(P["C10"].value, 0.0)
    if n < 1: sys.exit("❌ La hoja Parametros no indica cuántas preguntas tiene la prueba (celda C9).")
    par = {"exigencia": numero(P["C3"].value, .6), "notaMin": numero(P["C4"].value, 1), "notaMax": numero(P["C5"].value, 5),
           "basico": numero(P["C6"].value, 3), "alto": numero(P["C7"].value, 4), "superior": numero(P["C8"].value, 4.6),
           "preguntasMC": n, "puntosAbiertas": pts_ab, "hayVersionB": bool(int(numero(P["C12"].value, 0)))}
    H, c0 = 4, 4                                          # fila de encabezado y columna de P1 (igual que build_calificacion.py)
    col_ab = c0 + n + 2 if pts_ab > 0 else None           # tras «Aciertos» y «% aciertos»
    est = []
    for r in range(H + 1, E.max_row + 1):
        nombre = E.cell(row=r, column=2).value
        if nombre is None or not str(nombre).strip(): continue
        ver = str(E.cell(row=r, column=3).value or "A").strip().upper()
        ver = "B" if ver == "B" and par["hayVersionB"] else "A"
        resp = []
        for q in range(n):
            v = E.cell(row=r, column=c0 + q).value
            v = str(v).strip().upper() if v is not None else ""
            resp.append(v if v in ("A", "B", "C", "D") else None)
        pa = E.cell(row=r, column=col_ab).value if col_ab else None
        est.append({"fila": r, "nombre": str(nombre).strip(), "version": ver, "respuestas": resp,
                    "puntosAbiertas": numero(pa, None) if pa not in (None, "") else None})
    prueba = None                                         # datos de la prueba guardados en la hoja oculta «DatosPrueba»
    if "DatosPrueba" in wb.sheetnames:
        DP = wb["DatosPrueba"]
        texto = "".join(str(DP.cell(row=r, column=1).value or "") for r in range(2, DP.max_row + 1))
        try: prueba = json.loads(texto) if texto else None
        except ValueError: prueba = None
    json.dump({"parametros": par, "estudiantes": est, "prueba": prueba}, sys.stdout, ensure_ascii=False, indent=2)

if __name__ == "__main__":
    main()
