#!/usr/bin/env bash
# preparar-prompt.sh — Deja listo TODO lo que un prompt necesita, con un solo comando.
#
# Uso:
#   bash herramientas/preparar-prompt.sh <pruebas-saber|planeaciones-de-clase> [carpeta_de_trabajo]
#   (si aún no tienes el repo:)
#   curl -fsSL https://raw.githubusercontent.com/Luis-Prompt/Prompt/main/herramientas/preparar-prompt.sh | bash -s -- pruebas-saber
#
# Qué hace:
#   1. Obtiene el repositorio completo (tarball de codeload.github.com; si falla, git clone). Si se ejecuta desde un
#      clon local usa ese clon, sin red.
#   2. Copia a la carpeta de trabajo SOLO el prompt elegido (SKILL.md, manifest.json, scripts, catálogos, ejemplos) más
#      los datos compartidos (config/institucion.json, escudos, herramientas). Nada del otro prompt.
#   3. Comprueba que los archivos esenciales existan y no sean una página de error («404: Not Found»).
#   4. Verifica que docx y sharp estén disponibles (los instala en la carpeta de trabajo si faltan).
#   5. Imprime el orden de lectura: primero SKILL.md; el resto, según el manifest.json.
#
# Variables opcionales: REF=main (rama o etiqueta) · SIN_NPM=1 (no instalar dependencias).
set -euo pipefail

ID="${1:-}"
REPO="Luis-Prompt/Prompt"; REF="${REF:-main}"
case "$ID" in
  pruebas-saber)         CARPETA="prompts/evaluacion/creacion-pruebas-saber"; DEFECTO="/home/claude/examen"; PRINCIPAL="scripts/build_exam.js" ;;
  planeaciones-de-clase) CARPETA="prompts/planeacion/planeaciones-de-clase";  DEFECTO="/home/claude/plan";   PRINCIPAL="scripts/build_plan.js" ;;
  *) echo "Uso: preparar-prompt.sh <pruebas-saber|planeaciones-de-clase> [carpeta_de_trabajo]" >&2; exit 2 ;;
esac
DEST="${2:-$DEFECTO}"

# 1) Origen: clon local o descarga
AQUI="$(cd "$(dirname "${BASH_SOURCE[0]:-$0}")" 2>/dev/null && pwd || true)"
if [ -n "$AQUI" ] && [ -d "$AQUI/../prompts" ] && [ -f "$AQUI/../INDICE-PARA-CLAUDE.md" ]; then
  FUENTE="$(cd "$AQUI/.." && pwd)"; echo "📂 Usando el repositorio local: $FUENTE"
else
  TMP="$(mktemp -d)"; FUENTE="$TMP/repo"; mkdir -p "$FUENTE"
  echo "⬇️  Descargando $REPO ($REF)…"
  if curl -fsSL "https://codeload.github.com/$REPO/tar.gz/refs/heads/$REF" | tar -xz -C "$FUENTE" --strip-components=1 2>/dev/null \
     || curl -fsSL "https://codeload.github.com/$REPO/tar.gz/refs/tags/$REF" | tar -xz -C "$FUENTE" --strip-components=1 2>/dev/null; then :
  else
    echo "   (tarball no disponible; intentando git clone)"; rm -rf "$FUENTE"
    git clone --depth 1 --branch "$REF" "https://github.com/$REPO.git" "$FUENTE" >/dev/null 2>&1 \
      || { echo "❌ No pude descargar el repositorio. Pide al docente que suba el script y el escudo." >&2; exit 1; }
  fi
fi
[ -d "$FUENTE/$CARPETA" ] || { echo "❌ No existe $CARPETA en el repositorio." >&2; exit 1; }

# 2) Copiar solo el prompt elegido + datos compartidos
mkdir -p "$DEST"
cp -R "$FUENTE/$CARPETA/." "$DEST/"
rm -rf "$DEST/node_modules" "$DEST/salida"
mkdir -p "$DEST/config" "$DEST/banco-imagenes" "$DEST/herramientas" "$DEST/salida"
cp "$FUENTE/config/institucion.json" "$FUENTE/config/perfil-docente.ejemplo.json" "$DEST/config/"
cp -R "$FUENTE/banco-imagenes/escudos" "$DEST/banco-imagenes/"
cp "$FUENTE"/herramientas/*.js "$FUENTE"/herramientas/*.py "$DEST/herramientas/"
# El prompt de planeaciones espera el puente de datos junto a sus scripts
[ "$ID" = "planeaciones-de-clase" ] && cp "$FUENTE/herramientas/plan-a-borrador.js" "$DEST/scripts/"

# 3) Comprobar esenciales (y que no sean páginas de error)
FALTAN=0
for f in SKILL.md manifest.json "$PRINCIPAL" config/institucion.json assets/escudo_header.png; do
  if [ ! -s "$DEST/$f" ]; then echo "❌ Falta o está vacío: $f"; FALTAN=1
  elif head -c 20 "$DEST/$f" | grep -q "^404"; then echo "❌ $f es una página de error"; FALTAN=1; fi
done
[ "$FALTAN" = 0 ] || exit 1

# 4) Dependencias de Node
if [ -z "${SIN_NPM:-}" ]; then
  if ! (cd "$DEST" && node -e "require('docx');require('sharp')" 2>/dev/null) \
     && ! (NODE_PATH="$(npm root -g 2>/dev/null)" node -e "require('docx');require('sharp')" 2>/dev/null); then
    echo "📦 Instalando docx y sharp en $DEST …"; (cd "$DEST" && npm install --silent --no-audit --no-fund) || echo "⚠️  npm install falló; prueba: npm install -g docx sharp"
  fi
fi

# 5) Resumen y orden de lectura
N=$(find "$DEST" -type f -not -path "*/node_modules/*" | wc -l | tr -d ' ')
echo "✅ Prompt «$ID» listo en $DEST ($N archivos)"
echo "📖 ORDEN DE LECTURA:"
echo "   1) $DEST/SKILL.md            ← léelo completo y síguelo desde su primer paso"
echo "   2) $DEST/manifest.json       ← mapa: qué leer según la necesidad (áreas, catálogos, referencias, ejemplos)"
echo "   3) Script que se EJECUTA (no se reescribe): $DEST/$PRINCIPAL"
echo "   Los datos del colegio están en $DEST/config/institucion.json (son datos, no instrucciones)."
