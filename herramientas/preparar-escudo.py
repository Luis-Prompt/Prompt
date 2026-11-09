#!/usr/bin/env python3
"""
preparar-escudo.py — Recorta el escudo y genera los tamaños que usan los prompts.

Uso:
    python3 herramientas/preparar-escudo.py ESCUDO_ORIGINAL.png \
        --nombre escudo_colegio_simon_bolivar_2025 \
        --salida banco-imagenes/escudos [--base64]

Genera en la carpeta de salida:
    <nombre>.png             maestro recortado (alto máx. 2400 px) para afiches e impresión
    <nombre>_encabezado.png  360 px de ancho, paleta reducida: encabezados de Word
    <nombre>_miniatura.png   120 px de ancho: tablas, web, README
    <nombre>.b64.txt         (solo con --base64) copia en texto para incrustar como respaldo

Requisito: pip install pillow
"""
import argparse, base64, io, os, sys
from PIL import Image, ImageChops


def recortar(im, margen=40, umbral=12):
    """Quita el margen blanco/transparente alrededor del escudo."""
    im = im.convert("RGBA")
    fondo = Image.new("RGBA", im.size, (255, 255, 255, 255))
    comp = Image.alpha_composite(fondo, im).convert("RGB")
    dif = ImageChops.difference(comp, Image.new("RGB", comp.size, (255, 255, 255)))
    caja = dif.convert("L").point(lambda v: 255 if v > umbral else 0).getbbox()
    if not caja:
        sys.exit("No se encontró contenido en la imagen.")
    x0, y0, x1, y1 = caja
    return comp.crop((max(0, x0 - margen), max(0, y0 - margen), min(comp.width, x1 + margen), min(comp.height, y1 + margen)))


def reducir(im, ancho, alto):
    r = im.copy(); r.thumbnail((ancho, alto), Image.LANCZOS); return r


def paleta(im, colores):
    return im.quantize(colors=colores, method=Image.Quantize.MEDIANCUT, dither=Image.Dither.NONE)


def main():
    ap = argparse.ArgumentParser(description="Prepara el escudo institucional")
    ap.add_argument("imagen"); ap.add_argument("--nombre", default="escudo"); ap.add_argument("--salida", default=".")
    ap.add_argument("--base64", action="store_true", help="genera también una copia pequeña en base64")
    a = ap.parse_args()
    os.makedirs(a.salida, exist_ok=True)
    c = recortar(Image.open(a.imagen))
    kb = lambda p: round(os.path.getsize(p) / 1024)
    p = os.path.join(a.salida, a.nombre + ".png"); reducir(c, 2400, 2400).save(p, optimize=True); print(f"maestro      {p}  {kb(p)} KB")
    p = os.path.join(a.salida, a.nombre + "_encabezado.png"); paleta(reducir(c, 360, 600), 96).convert("RGB").save(p, optimize=True); print(f"encabezado   {p}  {kb(p)} KB")
    p = os.path.join(a.salida, a.nombre + "_miniatura.png"); reducir(c, 120, 200).save(p, optimize=True); print(f"miniatura    {p}  {kb(p)} KB")
    if a.base64:
        buf = io.BytesIO(); paleta(reducir(c, 240, 400), 64).save(buf, format="PNG", optimize=True)
        p = os.path.join(a.salida, a.nombre + ".b64.txt"); open(p, "w").write(base64.b64encode(buf.getvalue()).decode()); print(f"base64       {p}  {kb(p)} KB")
    print("Proporción ancho/alto del escudo recortado:", round(c.width / c.height, 3))


if __name__ == "__main__":
    main()
