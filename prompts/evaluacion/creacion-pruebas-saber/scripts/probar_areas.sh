#!/bin/bash
# Prueba rápida de todos los ejemplos de áreas (validación + generación de los dos Word, sin medir maquetación)
cd "$(dirname "$0")/.."
ok=0; mal=0
for f in examples/areas/*.json; do
  n=$(basename "$f" .json)
  out=$(node scripts/build_exam.js "$f" --out "${SALIDA:-/tmp/salida_areas}/$n" --rapido 2>&1); rc=$?
  w=$(echo "$out" | grep -c "⚠")
  if [ $rc -eq 0 ]; then ok=$((ok+1)); echo "✅ $n  (avisos: $w)"; else mal=$((mal+1)); echo "❌ $n"; echo "$out" | grep -E "•|❌" | head -8; fi
done
echo "—— $ok correctos, $mal con errores"
[ $mal -eq 0 ]
