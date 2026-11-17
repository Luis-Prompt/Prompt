#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Prueba automática de la plantilla de calificación: genera una prueba con versiones A y B, llena 12 estudiantes
simulados, recalcula el libro y compara nota, desempeño, aciertos y conteos por pregunta con un cálculo independiente.
Uso: python3 scripts/probar_calificacion.py      (sale con código 1 si hay una sola diferencia)"""
import json, os, random, shutil, subprocess, sys, tempfile
from decimal import Decimal, ROUND_HALF_UP
from openpyxl import load_workbook

AQUI = os.path.dirname(os.path.abspath(__file__)); RAIZ = os.path.dirname(AQUI)
tmp = tempfile.mkdtemp(prefix="cal_")
def run(cmd, **k): return subprocess.run(cmd, capture_output=True, text=True, **k)

r = run(["node", os.path.join(AQUI, "build_exam.js"), os.path.join(RAIZ, "examples", "exam_ejemplo.json"), "--out", tmp, "--rapido", "--versiones", "2"])
if r.returncode: sys.exit("❌ No se pudo generar la prueba:\n" + r.stdout[-600:] + r.stderr[-300:])
res = os.path.join(tmp, "resumen.json"); xlsx = os.path.join(tmp, "Plantilla.xlsx")
r = run(["python3", os.path.join(AQUI, "build_calificacion.py"), res, "--out", xlsx, "--estudiantes", "30"])
if r.returncode: sys.exit("❌ build_calificacion.py falló:\n" + r.stderr[-600:])

R = json.load(open(res, encoding="utf-8")); A = sorted(R["preguntas"], key=lambda q: q["num"]); B = sorted(R["versionB"]["preguntas"], key=lambda q: q["num"]); n = len(A)
total_ab = sum(x["puntos"] for x in R["abiertasInfo"]); random.seed(11)
wb = load_workbook(xlsx); E = wb["Estudiantes"]; col = {c.value: c.column for c in E[4] if c.value}
cP1, cAc, cNota, cDes = col["P1"], col["Aciertos"], col["NOTA"], col["Desempeño"]; cAb = [v for k, v in col.items() if str(k).startswith("Puntos abiertas")][0]
esperado = []
for i in range(12):
    r_ = 5 + i; ver = "B" if i % 2 else "A"; hab = [.95, .8, .7, .55, .4, .2][i % 6]; seq = B if ver == "B" else A; resp = []
    for q in seq:
        if random.random() < 0.08: resp.append(None)
        elif random.random() < hab: resp.append(q["clave"])
        else: resp.append(random.choice([x for x in "ABCD" if x != q["clave"]]))
    E.cell(row=r_, column=2, value=f"Estudiante {i + 1}"); E.cell(row=r_, column=3, value=(ver.lower() if (ver == "B" and i % 5 == 0) else ver))
    for j, l in enumerate(resp):
        if l: E.cell(row=r_, column=cP1 + j, value=l.lower() if (i + j) % 7 == 0 else l)
    ab = [0, 1, 2, 3, 4, 2, 1, 0, 3, 4, 2, 1][i] % (total_ab + 1); E.cell(row=r_, column=cAb, value=ab)
    ok = {q["id"]: (None if resp[k] is None else int(resp[k] == q["clave"])) for k, q in enumerate(seq)}
    ac = sum(v for v in ok.values() if v); p = (ac + ab) / (n + total_ab)
    nota = 3 + 2 * (p - .6) / .4 if p >= .6 else 1 + 2 * p / .6
    nota = float(Decimal(str(nota)).quantize(Decimal("0.1"), rounding=ROUND_HALF_UP))
    esperado.append((ac, nota, "Superior" if nota >= 4.6 else "Alto" if nota >= 4 else "Básico" if nota >= 3 else "Bajo", ok))
wb.save(xlsx)
rec = "/mnt/skills/public/xlsx/scripts/recalc.py"
if os.path.exists(rec): out = run(["python3", rec, xlsx, "120"]).stdout
else:
    run(["soffice", "--headless", "--convert-to", "xlsx", "--outdir", os.path.join(tmp, "rc"), xlsx], timeout=180); shutil.copy(os.path.join(tmp, "rc", "Plantilla.xlsx"), xlsx); out = ""
v = load_workbook(xlsx, data_only=True); E = v["Estudiantes"]; Z = v["Analisis"]; fallos = 0
for i, (ac, nota, des, ok) in enumerate(esperado):
    got = (E.cell(row=5 + i, column=cAc).value, E.cell(row=5 + i, column=cNota).value, E.cell(row=5 + i, column=cDes).value)
    if got != (ac, nota, des): fallos += 1; print(f"✗ estudiante {i + 1}: esperado {(ac, nota, des)} obtenido {got}")
for k, q in enumerate(A):
    aci = sum(1 for e in esperado if e[3].get(q["id"]) == 1); rsp = sum(1 for e in esperado if e[3].get(q["id"]) is not None)
    got = (Z.cell(row=20 + k, column=5).value, Z.cell(row=20 + k, column=6).value)
    if got != (aci, rsp): fallos += 1; print(f"✗ pregunta {k + 1}: esperado {(aci, rsp)} obtenido {got}")
errores = sum(1 for ws in v.worksheets for row in ws.iter_rows() for c in row if isinstance(c.value, str) and c.value.startswith("#"))
print(f"{'✅' if not fallos and not errores else '❌'} Calificación: 12 estudiantes, {n} preguntas, versiones A y B · diferencias: {fallos} · celdas con error: {errores}")
shutil.rmtree(tmp, ignore_errors=True); sys.exit(1 if fallos or errores else 0)
