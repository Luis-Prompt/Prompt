# Banco de imágenes

Recursos visuales reutilizables por todos los prompts.

```
banco-imagenes/
├─ escudos/        Escudo institucional en tres tamaños (maestro, encabezado, miniatura)
├─ plantillas-svg/ Gráficos reutilizables (figuras geométricas, mapas, esquemas)   ← por crear
└─ por-area/       matematicas/  ciencias/  lenguaje/  sociales/                    ← por crear
```

## Escudo del colegio

| Archivo | Para qué |
|---|---|
| `escudo_colegio_simon_bolivar_2025.png` | Maestro recortado (alto 2400 px): afiches, diplomas, impresión grande |
| `escudo_colegio_simon_bolivar_2025_encabezado.png` | 360 px de ancho: encabezado de los documentos Word |
| `escudo_colegio_simon_bolivar_2025_miniatura.png` | 120 px de ancho: tablas, web, README |

**Si cambia el escudo:** `python3 herramientas/preparar-escudo.py NUEVO.png --nombre escudo_colegio_simon_bolivar_2026 --salida banco-imagenes/escudos`, luego actualiza las rutas en `config/institucion.json` y copia el de encabezado a `assets/escudo_header.png` de cada prompt.

Reglas: nombres sin tildes ni espacios; preferir **SVG** para gráficos (se ve nítido al imprimir); no subir imágenes con derechos de autor ajenos ni fotos de estudiantes.
