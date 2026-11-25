#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_calificacion.py — PLANTILLA DE CALIFICACIÓN Y ANÁLISIS EN EXCEL para la prueba generada.

El docente solo escribe el nombre de cada estudiante y las letras que marcó (A, B, C o D);
el libro calcula solo: aciertos, nota (escala 1,0-5,0), desempeño, análisis por pregunta
(índice de dificultad y distractores), por competencia y por nivel de desempeño. Entiende las versiones A y B.

Uso:  python3 build_calificacion.py resumen.json [--out Calificacion.xlsx] [--estudiantes 45]
Datos: resumen.json lo escribe build_exam.js (claves, competencia y nivel de cada pregunta, versión B).
"""
import argparse, json, os, sys
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.properties import PageSetupProperties

FUENTE = "Arial"
ROJO = "C00000"
f = lambda **k: Font(name=FUENTE, size=k.pop("size", 10), **k)
RELLENO_ENC = PatternFill("solid", fgColor="E7E6E6")
RELLENO_IN = PatternFill("solid", fgColor="FFF2CC")       # celdas que el docente llena
RELLENO_OK = PatternFill("solid", fgColor="E2F0D9")
RELLENO_MAL = PatternFill("solid", fgColor="F8CBAD")
RELLENO_AV = PatternFill("solid", fgColor="FFE699")
lado = Side(style="thin", color="BFBFBF")
BORDE = Border(left=lado, right=lado, top=lado, bottom=lado)
CENTRO = Alignment(horizontal="center", vertical="center", wrap_text=True)
IZQ = Alignment(horizontal="left", vertical="center", wrap_text=True)


def enc(ws, fila, col, texto, ancho=None):
    c = ws.cell(row=fila, column=col, value=texto)
    c.font, c.fill, c.alignment, c.border = f(bold=True), RELLENO_ENC, CENTRO, BORDE
    if ancho: ws.column_dimensions[L(col)].width = ancho
    return c


def celda(ws, fila, col, valor=None, entrada=False, formato=None, alinea=CENTRO, negrita=False):
    c = ws.cell(row=fila, column=col, value=valor)
    c.font = f(color="0000FF", bold=negrita) if entrada else f(bold=negrita)
    c.alignment, c.border = alinea, BORDE
    if entrada: c.fill = RELLENO_IN
    if formato: c.number_format = formato
    return c


def titulo(ws, texto, sub=None):
    ws["A1"] = texto; ws["A1"].font = f(size=14, bold=True, color=ROJO)
    if sub: ws["A2"] = sub; ws["A2"].font = f(size=9, italic=True, color="595959")


def main():
    ap = argparse.ArgumentParser(description="Plantilla de calificación en Excel")
    ap.add_argument("resumen"); ap.add_argument("--out", default=None); ap.add_argument("--estudiantes", type=int, default=45)
    a = ap.parse_args()
    R = json.load(open(a.resumen, encoding="utf-8"))
    A = R["preguntas"]; B = (R.get("versionB") or {}).get("preguntas"); tiene_b = bool(B); n = len(A)
    if not n: sys.exit("El resumen no trae preguntas de selección múltiple.")
    meta = R.get("meta", {}); ab = R.get("abiertasInfo", []); pts_ab = sum(x.get("puntos", 0) for x in ab); N = a.estudiantes
    posB = {q["id"]: q for q in (B or [])}
    wb = Workbook()

    # ── Hoja Parametros ───────────────────────────────────────────────────────────
    P = wb.active; P.title = "Parametros"; titulo(P, "Parámetros de calificación", "Celdas amarillas = las puedes cambiar. Ajusta la escala al SIEE del colegio.")
    filas = [("Exigencia para alcanzar la nota 3,0", 0.60, "0%", True, "Porcentaje de aciertos que equivale a 3,0 (se propone 60 %)."),
             ("Nota mínima", 1.0, "0.0", True, "Escala nacional: 1,0 a 5,0."),
             ("Nota máxima", 5.0, "0.0", True, ""),
             ("Desempeño Básico desde", 3.0, "0.0", True, "Rangos propuestos (Decreto 1290 de 2009, compilado en el Decreto 1075 de 2015): el SIEE de cada colegio define los suyos."),
             ("Desempeño Alto desde", 4.0, "0.0", True, ""),
             ("Desempeño Superior desde", 4.6, "0.0", True, ""),
             ("Preguntas de selección múltiple", n, "0", False, "Cada una vale 1 punto."),
             ("Puntos de las preguntas abiertas", pts_ab, "0", False, "Se califican a mano con los criterios del documento del docente."),
             ("Puntaje máximo total", "=C9+C10", "0", False, ""),
             ("¿La prueba tiene versión B? (1 = sí)", 1 if tiene_b else 0, "0", False, "")]
    for i, (et, v, fmt, ent, nota) in enumerate(filas, start=3):
        P.cell(row=i, column=2, value=et).font = f(); P.cell(row=i, column=2).alignment = IZQ
        celda(P, i, 3, v, entrada=ent, formato=fmt); P.cell(row=i, column=4, value=nota).font = f(size=9, italic=True, color="595959")
    P["C9"] = n; P["C10"] = pts_ab; P["C11"] = "=C9+C10"; P["C12"] = 1 if tiene_b else 0
    P.column_dimensions["A"].width = 3; P.column_dimensions["B"].width = 42; P.column_dimensions["C"].width = 12; P.column_dimensions["D"].width = 90
    EXIG, NMIN, NMAX, BAS, ALT, SUP, TOTAL, HAYB = "Parametros!$C$3", "Parametros!$C$4", "Parametros!$C$5", "Parametros!$C$6", "Parametros!$C$7", "Parametros!$C$8", "Parametros!$C$11", "Parametros!$C$12"

    # ── Hoja Claves ───────────────────────────────────────────────────────────────
    K = wb.create_sheet("Claves"); titulo(K, "Claves de respuesta", "Una fila por pregunta (en el orden de la versión A). No la cambies: el análisis depende de ella.")
    cab = ["Pregunta (versión A)", "Clave A", "Pregunta en versión B", "Clave B", "Competencia", "Componente", "Nivel", "DBA / evidencia"]
    for j, t in enumerate(cab, start=1): enc(K, 4, j, t, [11, 9, 12, 9, 44, 26, 8, 40][j - 1])
    K.row_dimensions[4].height = 30
    for i, q in enumerate(sorted(A, key=lambda x: x["num"])):
        r = 5 + i; b = posB.get(q["id"], {})
        celda(K, r, 1, q["num"]); celda(K, r, 2, q["clave"], negrita=True); celda(K, r, 3, b.get("num", "—")); celda(K, r, 4, b.get("clave", "—"), negrita=True)
        celda(K, r, 5, q.get("competencia") or "Sin clasificar", alinea=IZQ); celda(K, r, 6, q.get("componente") or "—", alinea=IZQ)
        celda(K, r, 7, q.get("nivel") or "—"); celda(K, r, 8, q.get("dba") or "—", alinea=IZQ)
    K.freeze_panes = "A5"
    R0, R1 = 5, 4 + n                                                  # filas de las preguntas en Claves

    # ── Hoja Estudiantes ──────────────────────────────────────────────────────────
    E = wb.create_sheet("Estudiantes", 0)
    titulo(E, f"Calificación — {meta.get('area') or meta.get('asignatura') or 'Prueba'} · {meta.get('tema', '')} · Grado {meta.get('grado', '')}",
           f"{meta.get('institucion', '')}  ·  Escribe SOLO en las celdas amarillas: nombre, versión (A o B) y la letra que marcó cada estudiante. Deja vacía la pregunta que no contestó.")
    H = 4; c0 = 4                                                      # fila de encabezado y columna de la P1
    enc(E, H, 1, "N.º", 5); enc(E, H, 2, "Nombre completo del estudiante", 34); enc(E, H, 3, "Versión", 9)
    for q in range(n): enc(E, H, c0 + q, f"P{q + 1}", 5.5)
    cA = c0 + n; cP = cA + 1; cAb = cP + 1 if ab else None; cN = (cAb or cP) + 1; cD = cN + 1
    enc(E, H, cA, "Aciertos", 9); enc(E, H, cP, "% aciertos", 9)
    if cAb: enc(E, H, cAb, f"Puntos abiertas (máx. {pts_ab})", 14)
    enc(E, H, cN, "NOTA", 8); enc(E, H, cD, "Desempeño", 12); E.row_dimensions[H].height = 42
    dv_l = DataValidation(type="list", formula1='"A,B,C,D,a,b,c,d"', allow_blank=True, showErrorMessage=True, errorTitle="Respuesta no válida", error="Escribe A, B, C o D (o deja la celda vacía).")
    dv_v = DataValidation(type="list", formula1='"A,B"', allow_blank=True)
    E.add_data_validation(dv_l); E.add_data_validation(dv_v)
    f0, f1 = H + 1, H + N
    ans_rng = lambda r: f"${L(c0)}{r}:${L(c0 + n - 1)}{r}"
    for i in range(N):
        r = f0 + i
        celda(E, r, 1, i + 1); celda(E, r, 2, None, entrada=True, alinea=IZQ); celda(E, r, 3, None, entrada=True)
        dv_v.add(E.cell(row=r, column=3))
        for q in range(n):
            c = celda(E, r, c0 + q, None, entrada=True); dv_l.add(c)
        det_ok = f"Detalle!${L(4 + n)}{r - 1}:${L(3 + 2 * n)}{r - 1}"      # bloque de aciertos (1/0) de la hoja Detalle (fila alineada)
        celda(E, r, cA, f'=IF($B{r}="","",SUM({det_ok}))', formato="0")
        pextra = f"+IF(ISNUMBER({L(cAb)}{r}),{L(cAb)}{r},0)" if cAb else ""
        celda(E, r, cP, f'=IF($B{r}="","",({L(cA)}{r}{pextra})/{TOTAL})', formato="0%")
        if cAb: celda(E, r, cAb, None, entrada=True, formato="0.0")
        p = f"{L(cP)}{r}"
        celda(E, r, cN, f'=IF($B{r}="","",ROUND(IF({p}>={EXIG},3+({NMAX}-3)*({p}-{EXIG})/(1-{EXIG}),{NMIN}+(3-{NMIN})*{p}/{EXIG}),1))', formato="0.0", negrita=True)
        nt = f"{L(cN)}{r}"
        celda(E, r, cD, f'=IF($B{r}="","",IF({nt}>={SUP},"Superior",IF({nt}>={ALT},"Alto",IF({nt}>={BAS},"Básico","Bajo"))))')
    rd = f"{L(cD)}{f0}:{L(cD)}{f1}"
    for txt, rell in (("Superior", PatternFill("solid", fgColor="C6E0B4")), ("Alto", RELLENO_OK), ("Básico", RELLENO_AV), ("Bajo", RELLENO_MAL)):
        E.conditional_formatting.add(rd, CellIsRule(operator="equal", formula=[f'"{txt}"'], fill=rell))
    E.freeze_panes = E.cell(row=f0, column=c0)
    E.sheet_view.zoomScale = 90

    # ── Hoja Detalle (letra elegida por pregunta + acierto 1/0, unidas las versiones) ──
    D = wb.create_sheet("Detalle"); titulo(D, "Detalle por estudiante y pregunta", "Calculado: no escribas aquí. ✔ = acierto, ✘ = error. Las preguntas aparecen en el orden de la versión A.")
    enc(D, 3, 1, "N.º", 5); enc(D, 3, 2, "Estudiante", 30); enc(D, 3, 3, "Versión", 9)
    for q in range(n):
        enc(D, 3, 4 + q, f"Letra P{q + 1}(A)", 7); enc(D, 3, 4 + n + q, f"✔✘ P{q + 1}(A)", 7)
    D.row_dimensions[3].height = 40
    for i in range(N):
        r = 4 + i; er = f0 + i
        celda(D, r, 1, f"=Estudiantes!A{er}"); celda(D, r, 2, f'=IF(Estudiantes!B{er}="","",Estudiantes!B{er})', alinea=IZQ)
        celda(D, r, 3, f'=IF(Estudiantes!B{er}="","",IF(UPPER(Estudiantes!C{er})="B","B","A"))')
        for q in range(n):
            kr = R0 + q                                                # fila de la pregunta en Claves
            rango = f"Estudiantes!{ans_rng(er)}"
            pos = f'IF(AND($C{r}="B",{HAYB}=1),Claves!$C${kr},Claves!$A${kr})'
            idx = f"INDEX({rango},1,{pos})"
            celda(D, r, 4 + q, f'=IF($B{r}="","",IF({idx}="","",UPPER({idx})))')
            letra = f"{L(4 + q)}{r}"
            clave = f'IF(AND($C{r}="B",{HAYB}=1),Claves!$D${kr},Claves!$B${kr})'
            celda(D, r, 4 + n + q, f'=IF({letra}="","",IF({letra}={clave},1,0))', formato='[=1]"✔";[=0]"✘";""')
    rc = f"{L(4 + n)}4:{L(3 + 2 * n)}{3 + N}"
    D.conditional_formatting.add(rc, CellIsRule(operator="equal", formula=["1"], fill=RELLENO_OK))
    D.conditional_formatting.add(rc, CellIsRule(operator="equal", formula=["0"], fill=RELLENO_MAL))
    D.freeze_panes = "D4"

    # ── Hoja Análisis ─────────────────────────────────────────────────────────────
    Z = wb.create_sheet("Analisis", 1); titulo(Z, "Análisis de resultados", "Todo se calcula solo a partir de la hoja «Estudiantes». Lee primero el resumen del grupo y luego las preguntas con la lectura «Revisar».")
    notas = f"Estudiantes!${L(cN)}${f0}:${L(cN)}${f1}"; desemp = f"Estudiantes!${L(cD)}${f0}:${L(cD)}${f1}"; nombres = f"Estudiantes!$B${f0}:$B${f1}"
    enc(Z, 4, 1, "Resumen del grupo"); Z.merge_cells("A4:C4")
    res = [("Estudiantes evaluados", f'=SUMPRODUCT(--({nombres}<>""))', "0"), ("Nota promedio", f'=IFERROR(AVERAGE({notas}),"")', "0.00"),
           ("Nota más alta", f'=IF(COUNT({notas})=0,"",MAX({notas}))', "0.0"), ("Nota más baja", f'=IF(COUNT({notas})=0,"",MIN({notas}))', "0.0"),
           ("Aprobaron (nota ≥ 3,0)", f'=IFERROR(COUNTIF({notas},">=3")/COUNT({notas}),"")', "0%")]
    for i, (et, fo, fm) in enumerate(res):
        celda(Z, 5 + i, 1, et, alinea=IZQ); Z.merge_cells(start_row=5 + i, start_column=1, end_row=5 + i, end_column=2); celda(Z, 5 + i, 3, fo, formato=fm, negrita=True)
    enc(Z, 11, 1, "Desempeño"); Z.merge_cells("A11:B11"); enc(Z, 11, 3, "Estudiantes")
    for i, d in enumerate(("Superior", "Alto", "Básico", "Bajo")):
        celda(Z, 12 + i, 1, d, alinea=IZQ); Z.merge_cells(start_row=12 + i, start_column=1, end_row=12 + i, end_column=2); celda(Z, 12 + i, 3, f'=COUNTIF({desemp},"{d}")', formato="0")
    g1 = BarChart(); g1.type = "col"; g1.title = "Estudiantes por desempeño"; g1.style = 10; g1.legend = None
    g1.add_data(Reference(Z, min_col=3, min_row=11, max_row=15), titles_from_data=True); g1.set_categories(Reference(Z, min_col=1, min_row=12, max_row=15))
    g1.height, g1.width = 6.5, 11; g1.dataLabels = DataLabelList(); g1.dataLabels.showVal = True; g1.dataLabels.showSerName = False; g1.dataLabels.showCatName = False; g1.dataLabels.showLegendKey = False; Z.add_chart(g1, "E4")

    t0 = 19                                                            # tabla por pregunta
    Z.cell(row=t0 - 1, column=1, value="Análisis por pregunta").font = f(size=12, bold=True, color=ROJO)
    cols = ["Pregunta (versión A)", "Competencia", "Componente", "Nivel", "Aciertos", "Respondieron", "% de acierto", "Lectura", "Clave A"]
    anch = [11, 40, 24, 14, 9, 12, 11, 22, 8]
    for j, t in enumerate(cols, start=1): enc(Z, t0, j, t, anch[j - 1])
    colA = 10                                                          # distribución de respuestas versión A
    for j, d in enumerate("ABCD"): enc(Z, t0, colA + j, f"Eligieron {d} (A)")
    colB = colA + 4
    if tiene_b:
        enc(Z, t0, colB - 0, "Clave B")
        for j, d in enumerate("ABCD"): enc(Z, t0, colB + 1 + j, f"Eligieron {d} (B)")
    Z.row_dimensions[t0].height = 42
    vers = f"Detalle!$C$4:$C${3 + N}"
    for q in range(n):
        r = t0 + 1 + q; kr = R0 + q; ok = f"Detalle!${L(4 + n + q)}$4:${L(4 + n + q)}${3 + N}"; let = f"Detalle!${L(4 + q)}$4:${L(4 + q)}${3 + N}"
        celda(Z, r, 1, f"=Claves!A{kr}"); celda(Z, r, 2, f"=Claves!E{kr}", alinea=IZQ); celda(Z, r, 3, f"=Claves!F{kr}", alinea=IZQ); celda(Z, r, 4, f"=Claves!G{kr}")
        celda(Z, r, 5, f"=SUM({ok})", formato="0"); celda(Z, r, 6, f"=COUNT({ok})", formato="0")
        celda(Z, r, 7, f'=IF(F{r}=0,"",E{r}/F{r})', formato="0%", negrita=True)
        celda(Z, r, 8, f'=IF(G{r}="","",IF(G{r}>=0.85,"Muy fácil",IF(G{r}>=0.7,"Fácil",IF(G{r}>=0.4,"Adecuada",IF(G{r}>=0.25,"Difícil","Revisar la pregunta")))))')
        celda(Z, r, 9, f"=Claves!B{kr}", negrita=True)
        for j, d in enumerate("ABCD"):
            celda(Z, r, colA + j, f'=IF(COUNTIFS({let},"?*",{vers},"A")=0,"",COUNTIFS({let},"{d}",{vers},"A")/COUNTIFS({let},"?*",{vers},"A"))', formato="0%")
        if tiene_b:
            celda(Z, r, colB, f"=Claves!D{kr}", negrita=True)
            for j, d in enumerate("ABCD"):
                celda(Z, r, colB + 1 + j, f'=IF(COUNTIFS({let},"?*",{vers},"B")=0,"",COUNTIFS({let},"{d}",{vers},"B")/COUNTIFS({let},"?*",{vers},"B"))', formato="0%")
    q0, q1 = t0 + 1, t0 + n
    Z.conditional_formatting.add(f"H{q0}:H{q1}", CellIsRule(operator="equal", formula=['"Revisar la pregunta"'], fill=RELLENO_MAL))
    Z.conditional_formatting.add(f"H{q0}:H{q1}", CellIsRule(operator="equal", formula=['"Muy fácil"'], fill=RELLENO_AV))
    for j, d in enumerate("ABCD"):                                     # resalta la opción correcta en cada distribución
        cl = L(colA + j); Z.conditional_formatting.add(f"{cl}{q0}:{cl}{q1}", FormulaRule(formula=[f'$I{q0}="{d}"'], fill=RELLENO_OK))
        if tiene_b: cl = L(colB + 1 + j); Z.conditional_formatting.add(f"{cl}{q0}:{cl}{q1}", FormulaRule(formula=[f'${L(colB)}{q0}="{d}"'], fill=RELLENO_OK))
    nota_fin = t0 + n + 1
    Z.cell(row=nota_fin, column=1, value="Cómo leerlo: «Muy fácil» (≥ 85 %) puede no discriminar; «Revisar la pregunta» (< 25 %) suele indicar un error de clave, una redacción ambigua o un tema no enseñado. Si un distractor lo elige mucho más que la clave, revisa esa pregunta.").font = f(size=9, italic=True, color="595959")

    # resumen por competencia y por nivel
    cs = sorted({(q.get("competencia") or "Sin clasificar") for q in A}); c0r = nota_fin + 3
    Z.cell(row=c0r - 1, column=1, value="Resultado por competencia").font = f(size=12, bold=True, color=ROJO)
    enc(Z, c0r, 1, "N.º de preguntas"); enc(Z, c0r, 2, "Competencia"); enc(Z, c0r, 3, "% de acierto"); enc(Z, c0r, 4, "Lectura")
    for i, c in enumerate(cs):
        r = c0r + 1 + i
        celda(Z, r, 2, c, alinea=IZQ); celda(Z, r, 1, f'=COUNTIF($B${q0}:$B${q1},B{r})', formato="0")
        celda(Z, r, 3, f'=IFERROR(AVERAGEIF($B${q0}:$B${q1},B{r},$G${q0}:$G${q1}),"")', formato="0%", negrita=True)
        celda(Z, r, 4, f'=IF(C{r}="","",IF(C{r}>=0.7,"Fortaleza",IF(C{r}>=0.5,"En proceso","A reforzar")))')
    ce = c0r + len(cs)
    Z.conditional_formatting.add(f"D{c0r + 1}:D{ce}", CellIsRule(operator="equal", formula=['"A reforzar"'], fill=RELLENO_MAL))
    Z.conditional_formatting.add(f"D{c0r + 1}:D{ce}", CellIsRule(operator="equal", formula=['"Fortaleza"'], fill=RELLENO_OK))
    g2 = BarChart(); g2.type = "bar"; g2.title = "% de acierto por competencia"; g2.style = 10; g2.legend = None; g2.y_axis.scaling.min = 0; g2.y_axis.scaling.max = 1; g2.y_axis.number_format = "0%"
    g2.add_data(Reference(Z, min_col=3, min_row=c0r, max_row=ce), titles_from_data=True); g2.set_categories(Reference(Z, min_col=2, min_row=c0r + 1, max_row=ce))
    g2.height, g2.width = 7, 16; g2.dataLabels = DataLabelList(); g2.dataLabels.showVal = True; g2.dataLabels.showSerName = False; g2.dataLabels.showCatName = False; g2.dataLabels.showLegendKey = False; Z.add_chart(g2, f"F{c0r - 1}")
    nv = sorted({q.get("nivel") for q in A if q.get("nivel")}); n0 = max(ce, c0r + 14) + 3
    if nv:
        Z.cell(row=n0 - 1, column=1, value="Resultado por nivel de desempeño").font = f(size=12, bold=True, color=ROJO)
        enc(Z, n0, 1, "N.º de preguntas"); enc(Z, n0, 2, "Nivel"); enc(Z, n0, 3, "% de acierto")
        for i, v in enumerate(nv):
            r = n0 + 1 + i
            celda(Z, r, 2, v, formato="0"); celda(Z, r, 1, f'=COUNTIF($D${q0}:$D${q1},B{r})', formato="0")
            celda(Z, r, 3, f'=IFERROR(AVERAGEIF($D${q0}:$D${q1},B{r},$G${q0}:$G${q1}),"")', formato="0%", negrita=True)
    Z.freeze_panes = "A4"; Z.sheet_view.zoomScale = 90

    # ── Hoja Instrucciones ────────────────────────────────────────────────────────
    I = wb.create_sheet("Instrucciones", 0); titulo(I, "Plantilla de calificación y análisis de la prueba", f"{meta.get('area') or meta.get('asignatura') or ''} · {meta.get('tema', '')} · Grado {meta.get('grado', '')} · {meta.get('institucion', '')}")
    I.column_dimensions["A"].width = 4; I.column_dimensions["B"].width = 110
    pasos = ["CÓMO SE USA",
             "1. Ve a la hoja «Estudiantes». Escribe el nombre de cada estudiante en la columna B.",
             "2. En «Versión» escribe A o B (si la prueba tiene una sola versión, déjalo vacío o escribe A).",
             "3. En las columnas P1, P2… escribe la letra que marcó el estudiante (A, B, C o D). Si no contestó una pregunta, déjala vacía.",
             f"4. {'Si hay preguntas abiertas, califícalas a mano con los criterios del documento del docente y escribe los puntos en «Puntos abiertas».' if ab else 'Esta prueba no tiene preguntas abiertas.'}",
             "5. Mira la nota y el desempeño en «Estudiantes» y el análisis en «Analisis». No escribas en las demás celdas: se calculan solas.",
             "",
             "QUÉ CALCULA",
             "• Aciertos, porcentaje, NOTA (escala 1,0 a 5,0) y desempeño de cada estudiante.",
             "• Por pregunta: porcentaje de acierto (índice de dificultad), lectura (muy fácil, fácil, adecuada, difícil, revisar) y qué letra eligieron en cada versión.",
             "• Por competencia y por nivel de desempeño: dónde está el grupo fuerte y dónde hay que reforzar.",
             "• RETROALIMENTACIÓN POR ESTUDIANTE: cuando termines de llenar la hoja, devuélvele este mismo archivo a Claude y pídele «la retroalimentación por estudiante»: entrega una hoja por estudiante (Word y PDF) con su nota, sus fortalezas y, en cada pregunta fallada, qué marcó, qué error representa y por qué la correcta lo es.",
             "",
             "IMPORTANTE",
             "• La nota usa la exigencia de la hoja «Parametros» (60 % de aciertos = 3,0). Ajusta esos valores al Sistema Institucional de Evaluación (SIEE) del colegio: son una propuesta.",
             f"• Capacidad: {N} estudiantes. Si necesitas más, copia el archivo y vuelve a generar con --estudiantes.",
             "",
             "EJEMPLO DE LLENADO (solo ilustrativo, no hace parte de los datos)",
             "   N.º 1 · Ana Pérez · Versión A · P1: B · P2: D · P3: A · P4: (vacía, no contestó) · P5: C …"]
    for i, t in enumerate(pasos, start=4):
        c = I.cell(row=i, column=2, value=t); c.alignment = Alignment(wrap_text=True, vertical="top"); c.font = f(bold=t.isupper() and bool(t.strip()), size=10 if not t.isupper() else 11)
    # ── Hoja oculta «DatosPrueba»: lleva las preguntas, opciones y errores para la retroalimentación por estudiante ──
    # (así el docente solo devuelve este Excel; no hace falta conservar resumen.json)
    util = {k: R[k] for k in ("meta", "preguntas", "abiertasInfo") if k in R}
    if tiene_b: util["versionB"] = {"preguntas": B}
    texto = json.dumps(util, ensure_ascii=False, separators=(",", ":"))
    DP = wb.create_sheet("DatosPrueba"); DP["A1"] = "Datos de la prueba para la retroalimentación (hoja oculta: no la borres ni la cambies)."
    for i in range(0, len(texto), 30000): DP.cell(row=2 + i // 30000, column=1, value=texto[i:i + 30000])   # límite de Excel: 32 767 caracteres por celda
    DP.sheet_state = "hidden"
    for ws in wb.worksheets:
        ws.page_setup.orientation = "landscape"; ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True); ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    wb.move_sheet("Instrucciones", offset=-wb.index(wb["Instrucciones"]))
    wb["Detalle"].sheet_properties.tabColor = "A6A6A6"; wb["Claves"].sheet_properties.tabColor = "A6A6A6"; wb["Parametros"].sheet_properties.tabColor = "FFC000"; wb["Estudiantes"].sheet_properties.tabColor = ROJO
    wb.active = wb.index(wb["Estudiantes"])
    salida = a.out or os.path.join(os.path.dirname(os.path.abspath(a.resumen)), "Plantilla_Calificacion.xlsx")
    wb.save(salida); print(f"📊 {salida} · {n} preguntas · {N} estudiantes · {'versiones A y B' if tiene_b else 'versión única'}")


if __name__ == "__main__":
    main()
