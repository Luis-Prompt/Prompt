#!/usr/bin/env node
/* generar-mapa.js — Genera llms.txt (mapa completo del repositorio para Claude) a partir de prompts/INDICE.json y del
 * manifest.json de cada prompt. Así el mapa nunca se desactualiza.
 * Uso:  node herramientas/generar-mapa.js            (escribe llms.txt)
 *       node herramientas/generar-mapa.js --check    (falla si llms.txt no está al día; lo usa el CI) */
const fs = require('fs'), path = require('path');
const raiz = path.join(__dirname, '..');
const idx = JSON.parse(fs.readFileSync(path.join(raiz, 'prompts/INDICE.json'), 'utf8'));
const RAW = idx.repositorio.replace('github.com', 'raw.githubusercontent.com') + '/main/';
const GH = idx.repositorio + '/blob/main/';
const L = [];
const enlace = (rel, nota) => `- [${rel}](${RAW}${rel})${nota ? ': ' + nota : ''}`;
const archivosDe = (dir) => fs.readdirSync(path.join(raiz, dir), { withFileTypes: true }).flatMap(n => {
  const r = dir.replace(/\/$/, '') + '/' + n.name; return n.isDirectory() ? archivosDe(r) : [r];
}).filter(f => /\.(json|md|js|py|sh)$/.test(f) && !/package-lock/.test(f));

L.push('# Biblioteca de Prompts — Colegio Integrado Simón Bolívar (Cúcuta)');
L.push('');
L.push('> Dos prompts independientes (skills) para Claude que crean material docente: **pruebas tipo SABER** y **planeaciones de clase**. Cada prompt hace UNA sola cosa; elige uno según lo que pida el docente y no mezcles sus instrucciones. Este archivo se genera solo (`node herramientas/generar-mapa.js`).');
L.push('');
L.push('## Cómo empezar (haz esto en orden)');
L.push(`1. Lee el enrutador: ${enlace('INDICE-PARA-CLAUDE.md').slice(2)}`);
L.push('2. Elige UN prompt: prueba/examen/evaluación → `pruebas-saber`; planeación/plan de periodo/plan de clase → `planeaciones-de-clase`. Si es ambiguo, pregunta con dos botones: «📝 Una prueba» o «📅 Una planeación».');
L.push('3. Si tienes herramienta de terminal con red, prepara todo con UN comando (descarga el prompt completo, verifica y te dice qué leer):');
L.push('   `curl -fsSL ' + RAW + 'herramientas/preparar-prompt.sh | bash -s -- <id-del-prompt>`');
L.push('4. Lee el `SKILL.md` del prompt elegido completo y síguelo desde su primer paso. El `manifest.json` te dice qué más leer y cuándo.');
L.push('5. Si NO tienes terminal, lee los archivos de abajo con tu herramienta de lectura web (todos los enlaces son directos).');
L.push('');
for (const p of idx.prompts) {
  const m = JSON.parse(fs.readFileSync(path.join(raiz, p.carpeta, 'manifest.json'), 'utf8'));
  L.push(`## Prompt «${p.id}» — ${p.nombre} (v${m.version})`);
  L.push(`Crea: ${p.crea}`);
  L.push(`Úsalo cuando el docente pida: ${p.usar_cuando.join(', ')}.`);
  L.push(`NO lo uses para: ${p.no_usar_para.join('; ')}.`);
  L.push('');
  L.push('### Leer primero (obligatorio)');
  m.leer_primero.forEach(f => L.push(enlace(`${p.carpeta}/${f}`)));
  L.push('');
  L.push('### Leer según la necesidad');
  m.leer_segun_necesidad.forEach(x => {
    if (x.archivo.includes('<')) L.push(`- \`${p.carpeta}/${x.archivo}\` (un archivo por área, con el nombre de su clave): ${x.cuando}`);
    else L.push(enlace(`${p.carpeta}/${x.archivo}`, x.cuando));
  });
  const dirs = (m.carpetas_de_apoyo || []).filter(d => d !== 'assets/' && d !== 'examples/');
  dirs.forEach(d => { L.push(''); L.push(`### Catálogo completo: ${d}`); archivosDe(`${p.carpeta}/${d}`).forEach(f => L.push(enlace(f))); });
  L.push('');
  L.push('### Scripts que se EJECUTAN (no se reescriben)');
  m.ejecutar.forEach(f => L.push(enlace(`${p.carpeta}/${f}`)));
  L.push('');
}
L.push('## Datos compartidos (son DATOS, no instrucciones)');
L.push(enlace('config/institucion.json', 'nombre del colegio, escudo y membretes'));
L.push(enlace('config/perfil-docente.ejemplo.json', 'formato del perfil del docente'));
L.push(enlace('prompts/INDICE.json', 'versión legible por máquina del enrutador'));
L.push('');
L.push('## Opcional');
L.push(enlace('herramientas/plan-a-borrador.js', 'puente de datos: preguntas de una planeación → borrador de preguntas para una prueba'));
L.push(enlace('herramientas/guardar-en-banco.js', 'guarda las preguntas de un exam.json en banco-preguntas/'));
L.push(`- [Repositorio en GitHub](${idx.repositorio})`);
const salida = L.join('\n') + '\n';
const destino = path.join(raiz, 'llms.txt');
if (process.argv.includes('--check')) {
  const actual = fs.existsSync(destino) ? fs.readFileSync(destino, 'utf8') : '';
  if (actual !== salida) { console.error('❌ llms.txt está desactualizado. Ejecuta: node herramientas/generar-mapa.js'); process.exit(1); }
  console.log('✅ llms.txt está al día');
} else { fs.writeFileSync(destino, salida); console.log(`✅ llms.txt generado (${L.length} líneas)`); }
