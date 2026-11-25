#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Prueba automática de la retroalimentación por estudiante.

Genera una prueba con versiones A y B, llena 14 estudiantes simulados (incluye uno sin respuestas y uno perfecto),
genera la retroalimentación a partir del Excel SIN recalcular (como lo entregaría un docente) y compara aciertos,
nota y desempeño con los que calcula el propio Excel (recalculado con LibreOffice). Además comprueba que:
  · cada estudiante ocupa una página del PDF y su nombre aparece en ella,
  · se explica el error conceptual de lo que marcó (versión con opciones),
  · con un resumen antiguo (sin opciones) el generador sigue funcionando y avisa,
  · --individual crea un Word por estudiante.
Uso: python3 scripts/probar_retroalimentacion.py      (sale con código 1 si algo falla)"""
import json, os, random, shutil, subprocess, sys, tempfile
from decimal import Decimal, ROUND_HALF_UP
from openpyxl import load_workbook

AQUI = os.path.dirname(os.path.abspath(__file__)); RAIZ = os.path.dirname(AQUI)
tmp = tempfile.mkdtemp(prefix="retro_")
def run(cmd, **k): return subprocess.run(cmd, capture_output=True, text=True, **k)
errores = []
def chequea(ok, msg):
    if not ok: errores.append(msg)

r = run(["node", os.path.join(AQUI, "build_exam.js"), os.path.join(RAIZ, "examples", "exam_ejemplo.json"), "--out", tmp, "--rapido", "--versiones", "2"])
if r.returncode: sys.exit("❌ No se pudo generar la prueba:\n" + r.stdout[-600:] + r.stderr[-300:])
res = os.path.join(tmp, "resumen.json"); xlsx = os.path.join(tmp, "Plantilla.xlsx")
r = run(["python3", os.path.join(AQUI, "build_calificacion.py"), res, "--out", xlsx, "--estudiantes", "30"])
if r.returncode: sys.exit("❌ build_calificacion.py falló:\n" + r.stderr[-600:])

R = json.load(open(res, encoding="utf-8")); A = sorted(R["preguntas"], key=lambda q: q["num"]); B = sorted(R["versionB"]["preguntas"], key=lambda q: q["num"]); n = len(A)
chequea(all("opciones" in q and len(q["opciones"]) == 4 and "justificacion" in q for q in A + B), "resumen.json: cada pregunta debe traer opciones y justificación")
total_ab = sum(x["puntos"] for x in R["abiertasInfo"]); random.seed(5)
wb = load_workbook(xlsx); E = wb["Estudiantes"]; col = {c.value: c.column for c in E[4] if c.value}
cP1 = col["P1"]; cAb = [v for k, v in col.items() if str(k).startswith("Puntos abiertas")][0]
esperado = []
for i in range(14):
    fila = 5 + i; ver = "B" if i % 2 else "A"; seq = B if ver == "B" else A; resp = []
    if i == 12: resp = [None] * n                                    # no contestó nada
    elif i == 13: resp = [q["clave"] for q in seq]                   # perfecto
    else:
        hab = [.9, .75, .6, .45, .3, .15][i % 6]
        for q in seq:
            if random.random() < .06: resp.append(None)
            elif random.random() < hab: resp.append(q["clave"])
            else: resp.append(random.choice([x for x in "ABCD" if x != q["clave"]]))
    E.cell(row=fila, column=2, value=f"Estudiante Número {i + 1}"); E.cell(row=fila, column=3, value=ver.lower() if (ver == "B" and i % 4 == 1) else ver)
    for j, l in enumerate(resp):
        if l: E.cell(row=fila, column=cP1 + j, value=l.lower() if (i + j) % 6 == 0 else l)
    ab = (i * 3) % (total_ab + 1); E.cell(row=fila, column=cAb, value=ab)
    ac = sum(1 for k, q in enumerate(seq) if resp[k] == q["clave"]); p = (ac + ab) / (n + total_ab)
    nota = 3 + 2 * (p - .6) / .4 if p >= .6 else 1 + 2 * p / .6
    nota = float(Decimal(str(nota)).quantize(Decimal("0.1"), rounding=ROUND_HALF_UP))
    esperado.append((ac, nota, "Superior" if nota >= 4.6 else "Alto" if nota >= 4 else "Básico" if nota >= 3 else "Bajo"))
wb.save(xlsx)
llena = os.path.join(tmp, "Llena.xlsx"); shutil.copy(xlsx, llena)         # copia SIN recalcular: sin valores en caché

# ── 1. Retroalimentación desde el Excel sin recalcular ────────────────────────
out = os.path.join(tmp, "retro")
r = run(["node", os.path.join(AQUI, "build_retroalimentacion.js"), res, llena, "--out", out, "--individual", "--docente", "Docente de prueba"], timeout=300)
chequea(r.returncode == 0, "build_retroalimentacion.js falló:\n" + (r.stdout + r.stderr)[-800:])
if r.returncode == 0:
    D = json.load(open(os.path.join(out, "retroalimentacion.json"), encoding="utf-8"))["estudiantes"]
    chequea(len(D) == 14, f"debía haber 14 estudiantes y hay {len(D)}")
    # ── 2. Comparar con el Excel recalculado por LibreOffice ──────────────────
    rec = "/mnt/skills/public/xlsx/scripts/recalc.py"
    if os.path.exists(rec): run(["python3", rec, xlsx, "120"])
    else:
        run(["soffice", "--headless", "--convert-to", "xlsx", "--outdir", os.path.join(tmp, "rc"), xlsx], timeout=180); shutil.copy(os.path.join(tmp, "rc", "Plantilla.xlsx"), xlsx)
    Ev = load_workbook(xlsx, data_only=True)["Estudiantes"]; cAc, cNota, cDes = col["Aciertos"], col["NOTA"], col["Desempeño"]
    for i, d in enumerate(D):
        excel = (Ev.cell(row=5 + i, column=cAc).value, Ev.cell(row=5 + i, column=cNota).value, Ev.cell(row=5 + i, column=cDes).value)
        chequea((d["aciertos"], d["nota"], d["desempeno"]) == excel, f"estudiante {i + 1}: la retroalimentación dice {(d['aciertos'], d['nota'], d['desempeno'])} y el Excel {excel}")
        chequea((d["aciertos"], d["nota"], d["desempeno"]) == esperado[i], f"estudiante {i + 1}: cálculo independiente {esperado[i]} ≠ {(d['aciertos'], d['nota'], d['desempeno'])}")
        chequea(len(d["falladas"]) == n - d["aciertos"], f"estudiante {i + 1}: el número de preguntas falladas no coincide")
    chequea(D[13]["falladas"] == [] and D[13]["aciertos"] == n, "el estudiante perfecto no debe tener preguntas falladas")
    chequea(D[12]["aciertos"] == 0 and len(D[12]["falladas"]) == n, "el estudiante sin respuestas debe tener todas falladas")
    # ── 3. PDF: una página por estudiante, con su nombre y con el error explicado ─
    pdf = [f for f in os.listdir(out) if f.endswith(".pdf")]
    chequea(len(pdf) == 1, "no se generó el PDF")
    if pdf:
        ruta = os.path.join(out, pdf[0]); info = run(["pdfinfo", ruta]).stdout
        paginas = int([l for l in info.splitlines() if l.startswith("Pages")][0].split()[1]) if "Pages" in info else -1
        chequea(paginas == 14, f"el PDF debía tener 14 páginas (una por estudiante) y tiene {paginas}")
        for i in range(min(paginas, 14)):
            txt = run(["pdftotext", "-f", str(i + 1), "-l", str(i + 1), ruta, "-"]).stdout
            chequea(f"Estudiante Número {i + 1}" in txt and "RETROALIMENTACIÓN DE LA PRUEBA" in txt, f"la página {i + 1} no corresponde al estudiante {i + 1}")
            chequea(i == 13 or "Marcaste:" in txt, f"la página {i + 1} no explica lo que marcó el estudiante")
        todo = run(["pdftotext", ruta, "-"]).stdout
        chequea("→" in todo, "no aparece ninguna explicación del error conceptual (→)")
        chequea("Respondiste correctamente todas" in todo, "el estudiante perfecto debe recibir el mensaje de felicitación")
        chequea("no marcaste ninguna respuesta" in todo, "debe decir cuando una pregunta quedó sin respuesta")
    ind = os.path.join(out, "individuales")
    chequea(os.path.isdir(ind) and len([f for f in os.listdir(ind) if f.endswith(".docx")]) == 14, "--individual debía crear 14 Word")

# ── 3b. Flujo real: SOLO el Excel (ya recalculado), sin resumen.json → debe dar exactamente lo mismo ──
chequea(load_workbook(xlsx)["DatosPrueba"].sheet_state == "hidden", "la hoja «DatosPrueba» debe existir y estar oculta")
out2 = os.path.join(tmp, "retro_solo_excel")
r = run(["node", os.path.join(AQUI, "build_retroalimentacion.js"), xlsx, "--out", out2, "--sin-pdf", "--docente", "Docente de prueba"], timeout=300)
chequea(r.returncode == 0, "con solo el Excel debía funcionar:\n" + (r.stdout + r.stderr)[-600:])
if r.returncode == 0 and os.path.exists(os.path.join(out, "retroalimentacion.json")):
    a_ = json.load(open(os.path.join(out, "retroalimentacion.json"), encoding="utf-8"))["estudiantes"]; b_ = json.load(open(os.path.join(out2, "retroalimentacion.json"), encoding="utf-8"))["estudiantes"]
    chequea(a_ == b_, "con solo el Excel los resultados deben ser idénticos a los obtenidos con resumen.json")
    chequea("versión anterior" not in r.stderr, "con solo el Excel debían salir las opciones y los errores de los distractores (la hoja oculta quedó incompleta)")
sin = load_workbook(llena); del sin["DatosPrueba"]; sinp = os.path.join(tmp, "SinDatos.xlsx"); sin.save(sinp)
r = run(["node", os.path.join(AQUI, "build_retroalimentacion.js"), sinp, "--out", os.path.join(tmp, "z")])
chequea(r.returncode == 1 and "DatosPrueba" in r.stderr, "un Excel sin la hoja oculta y sin resumen.json debía fallar explicando qué hacer")

# ── 4. Resumen antiguo (sin opciones): debe seguir funcionando y avisar ───────
viejo = json.load(open(res, encoding="utf-8"))
for q in viejo["preguntas"] + viejo["versionB"]["preguntas"]:
    for k in ("opciones", "pregunta", "justificacion"): q.pop(k, None)
rv = os.path.join(tmp, "resumen_viejo.json"); json.dump(viejo, open(rv, "w", encoding="utf-8"))
r = run(["node", os.path.join(AQUI, "build_retroalimentacion.js"), rv, llena, "--out", os.path.join(tmp, "retro_viejo"), "--sin-pdf"], timeout=120)
chequea(r.returncode == 0 and "versión anterior" in r.stderr, "con un resumen antiguo debía funcionar y avisar:\n" + (r.stdout + r.stderr)[-400:])

# ── 5. Errores claros ─────────────────────────────────────────────────────────
r = run(["node", os.path.join(AQUI, "build_retroalimentacion.js"), res, os.path.join(tmp, "no_existe.xlsx"), "--out", os.path.join(tmp, "x")])
chequea(r.returncode == 1, "con un Excel inexistente debía fallar con código 1")
vacia = os.path.join(tmp, "Vacia.xlsx"); run(["python3", os.path.join(AQUI, "build_calificacion.py"), res, "--out", vacia, "--estudiantes", "5"])
r = run(["node", os.path.join(AQUI, "build_retroalimentacion.js"), res, vacia, "--out", os.path.join(tmp, "y")])
chequea(r.returncode == 1 and "sin nombre" not in r.stderr and "no tiene estudiantes" in r.stderr, "con una plantilla vacía debía decir que no hay estudiantes")

if os.environ.get("RETRO_CONSERVAR"): shutil.copytree(os.path.join(tmp, "retro"), os.environ["RETRO_CONSERVAR"], dirs_exist_ok=True)   # para inspeccionar los resultados
shutil.rmtree(tmp, ignore_errors=True)
if errores: print("\n".join("❌ " + e for e in errores)); sys.exit(1)
print("✅ Retroalimentación: 14 estudiantes · aciertos, nota y desempeño coinciden con el Excel · una página por estudiante · error conceptual explicado · funciona con solo el Excel · compatible con resúmenes antiguos")
