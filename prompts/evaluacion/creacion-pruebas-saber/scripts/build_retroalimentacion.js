#!/usr/bin/env node
/* ============================================================================
 *  build_retroalimentacion.js — RETROALIMENTACIÓN POR ESTUDIANTE (Word + PDF)
 *
 *  Cierra el ciclo evaluación → mejora: a partir de la plantilla de calificación ya diligenciada, entrega una hoja
 *  por estudiante con su resultado, sus fortalezas por competencia y, por cada pregunta fallada, qué marcó, cuál era
 *  la respuesta correcta, qué error conceptual representa lo que marcó y por qué la correcta lo es.
 *  Sirve para devolver la prueba al estudiante y para reuniones con familias.
 *
 *  Uso:  node build_retroalimentacion.js Calificacion.xlsx [resumen.json] [--out carpeta] [--individual] [--sin-pdf]
 *                                         [--docente "Nombre"] [--institucion archivo.json]
 *  Entradas: la plantilla de Excel que generó build_calificacion.py, con nombres y letras marcadas. Esa plantilla lleva en una
 *            hoja oculta («DatosPrueba») las preguntas, opciones y errores de los distractores; por eso basta con el Excel.
 *            resumen.json (de build_exam.js, v2.10 o posterior) es opcional: si se da, tiene prioridad.
 *  Salidas : Retroalimentacion_<area>_<tema>_Grado<g>.docx (una página por estudiante) y su PDF;
 *            con --individual, además un Word por estudiante; y retroalimentacion.json con los resultados calculados.
 *  La nota se calcula con los mismos parámetros y la misma fórmula del Excel (hoja «Parametros»).
 * ========================================================================== */
const fs = require('fs'), path = require('path'), { spawnSync } = require('child_process');
const lib = (m) => { try { return require(m); } catch (e) { return require(path.join(process.env.NODE_PATH || '/home/claude/.npm-global/lib/node_modules', m)); } };
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, AlignmentType, BorderStyle, WidthType, ShadingType, ImageRun, VerticalAlign, Footer, PageNumber } = lib('docx');

const args = process.argv.slice(2);
const pos = args.filter((a, i) => !a.startsWith('--') && !(i > 0 && ['--out', '--docente', '--institucion'].includes(args[i - 1])));
const val = (k, d) => args.includes(k) ? args[args.indexOf(k) + 1] : d;
const fail = (m) => { console.error('❌ ' + m); process.exit(1); };
const xlsxPath = pos.find(f => /\.xlsx$/i.test(f)), resumenPath = pos.find(f => /\.json$/i.test(f)), outDir = path.resolve(val('--out', 'salida'));
if (!xlsxPath) { console.error('Uso: node build_retroalimentacion.js Calificacion.xlsx [resumen.json] [--out carpeta] [--individual] [--sin-pdf] [--docente "Nombre"] [--institucion archivo.json]'); process.exit(2); }
if (!fs.existsSync(xlsxPath)) fail(`No encuentro el Excel: ${xlsxPath}`);
fs.mkdirSync(outDir, { recursive: true });

// ═══ 1. DATOS ════════════════════════════════════════════════════════════════
const lector = spawnSync('python3', [path.join(__dirname, 'leer_calificacion.py'), xlsxPath], { encoding: 'utf8', maxBuffer: 256 * 1024 * 1024 });
if (lector.error || lector.status !== 0) fail('No pude leer el Excel:\n' + ((lector.stderr || '') + (lector.error ? lector.error.message : '')).trim());
const { parametros: P, estudiantes: ALUMNOS, prueba: PRUEBA } = JSON.parse(lector.stdout);
let R;
if (resumenPath) { try { R = JSON.parse(fs.readFileSync(resumenPath, 'utf8')); } catch (e) { fail(`No pude leer ${resumenPath}: ${e.message}`); } }
else if (PRUEBA) R = PRUEBA;
else fail('Este Excel no lleva los datos de la prueba (hoja «DatosPrueba»). Pasa también el resumen.json de la prueba: node build_retroalimentacion.js Calificacion.xlsx resumen.json');
const A = (R.preguntas || []).slice().sort((a, b) => a.num - b.num), B = ((R.versionB || {}).preguntas || []).slice().sort((a, b) => a.num - b.num);
if (!A.length) fail('El resumen no trae preguntas de selección múltiple.');
const conOpciones = A.every(q => Array.isArray(q.opciones) && q.opciones.length);
if (!conOpciones) console.warn('⚠  Este resumen.json es de una versión anterior (sin opciones ni errores de los distractores): la retroalimentación dirá cuál era la respuesta correcta, pero no explicará el error. Vuelve a generar la prueba con build_exam.js para obtener el detalle.');
if (P.preguntasMC !== A.length) fail(`El Excel tiene ${P.preguntasMC} preguntas y el resumen ${A.length}: no corresponden a la misma prueba.`);
if (!ALUMNOS.length) fail('El Excel no tiene estudiantes con nombre. Escribe los nombres y las letras marcadas en la hoja «Estudiantes».');
const meta = R.meta || {}, docente = val('--docente', meta.docente || '');

// ═══ 2. INSTITUCIÓN Y ESCUDO (mismo orden de búsqueda que build_exam.js) ═════
let INST = {}, INST_DIR = null;
for (const f of [val('--institucion', null), process.env.INSTITUCION_JSON, path.join(__dirname, '..', 'config', 'institucion.json'),
  path.join(__dirname, '..', '..', '..', '..', 'recursos', 'colegio', 'institucion.json'), path.join(process.cwd(), 'config', 'institucion.json')]) {
  try { if (f && fs.existsSync(f)) { INST = JSON.parse(fs.readFileSync(f, 'utf8')); INST_DIR = path.dirname(path.resolve(f)); break; } } catch (e) { /* se prueba el siguiente */ }
}
const institucion = INST.nombre || meta.institucion || 'Institución educativa';
let escudo = null;
for (const f of [process.env.ESCUDO_PATH, INST.escudo && INST.escudo.encabezado && INST_DIR ? path.resolve(INST_DIR, INST.escudo.encabezado) : null, path.join(__dirname, '..', 'assets', 'escudo_header.png')]) {
  if (f && fs.existsSync(f)) { escudo = fs.readFileSync(f); break; }
}

// ═══ 3. CÁLCULO (misma fórmula de la hoja «Estudiantes») ═════════════════════
const redondea1 = (x) => Math.round((x + 1e-9) * 10) / 10;                     // ROUND de Excel: mitad hacia arriba
const total = P.preguntasMC + P.puntosAbiertas;
function nota(aciertos, abiertas) {
  const p = (aciertos + (abiertas || 0)) / total;
  return redondea1(p >= P.exigencia ? 3 + (P.notaMax - 3) * (p - P.exigencia) / (1 - P.exigencia) : P.notaMin + (3 - P.notaMin) * p / P.exigencia);
}
const desempeno = (n) => n >= P.superior ? 'Superior' : n >= P.alto ? 'Alto' : n >= P.basico ? 'Básico' : 'Bajo';
const FORT = 0.75, PROC = 0.5;                                                  // ≥75 % fortaleza · 50-74 % en proceso · <50 % por reforzar
const estadoComp = (p) => p >= FORT ? 'Fortaleza' : p >= PROC ? 'En proceso' : 'Por reforzar';

function evaluar(al) {
  const seq = al.version === 'B' && B.length ? B : A;
  const items = seq.map((q, k) => { const marco = al.respuestas[k], op = (q.opciones || []).find(o => o.letra === marco), correcta = (q.opciones || []).find(o => o.ok);
    return { num: q.num, q, marco, ok: marco === q.clave, opMarcada: op || null, opCorrecta: correcta || null }; });
  const aciertos = items.filter(i => i.ok).length;
  const comps = {}; items.forEach(i => { const c = i.q.competencia || 'Sin clasificar'; (comps[c] = comps[c] || { n: 0, ok: 0 }); comps[c].n++; if (i.ok) comps[c].ok++; });
  const lista = Object.entries(comps).map(([nombre, v]) => ({ nombre, aciertos: v.ok, total: v.n, pct: v.ok / v.n, estado: estadoComp(v.ok / v.n) }));
  const n = nota(aciertos, al.puntosAbiertas);
  return { ...al, items, aciertos, competencias: lista, nota: n, desempeno: desempeno(n), falladas: items.filter(i => !i.ok) };
}
const RES = ALUMNOS.map(evaluar);

// ═══ 4. DOCUMENTO ═══════════════════════════════════════════════════════════
const F = 'Arial', VERDE = '2E7D32', GRIS = '595959';
const COL = { Superior: 'C6E0B4', Alto: 'E2F0D9', 'Básico': 'FFE699', Bajo: 'F8CBAD', Fortaleza: 'E2F0D9', 'En proceso': 'FFF2CC', 'Por reforzar': 'F8CBAD' };
const run = (t, o = {}) => new TextRun({ text: String(t), font: F, size: 19, ...o });
const para = (ch, o = {}) => new Paragraph({ children: Array.isArray(ch) ? ch : [ch], spacing: { before: 0, after: 50 }, ...o });
const fmt = (n) => n.toFixed(1).replace('.', ',');
const espacio = (after = 30) => new Paragraph({ children: [], run: { size: 4 }, spacing: { before: 0, after } });   // separador casi sin altura
const lado = { style: BorderStyle.SINGLE, size: 4, color: 'BFBFBF' }, BORDES = { top: lado, bottom: lado, left: lado, right: lado };
const NONE = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' }, SIN = { top: NONE, bottom: NONE, left: NONE, right: NONE };
const celda = (ch, w, o = {}) => new TableCell({ children: Array.isArray(ch) ? ch : [ch], width: { size: w, type: WidthType.DXA }, margins: { top: 50, bottom: 50, left: 100, right: 100 }, borders: BORDES, verticalAlign: VerticalAlign.CENTER, ...o });
const sombra = (c) => ({ type: ShadingType.CLEAR, fill: c, color: 'auto' });
const W = 10000;                                                                  // ancho útil (Carta, márgenes de 0,6")

const MENSAJE = {
  Superior: 'Excelente trabajo: demuestras un dominio sólido de lo evaluado. Sigue retándote con problemas más complejos.',
  Alto: 'Muy buen resultado: tienes bases firmes. Repasar los puntos de abajo te acerca al nivel superior.',
  'Básico': 'Vas por buen camino: ya logras lo esencial. Con práctica en los puntos de abajo puedes subir de nivel.',
  Bajo: 'Esta prueba nos muestra en qué necesitas más apoyo. Con práctica acompañada y repasando los puntos de abajo puedes mejorar.'
};

let MAXF = 5;                                                                     // preguntas falladas que se detallan por hoja (se reduce solo si la hoja no cabe)
function paginaEstudiante(r) {
  const ch = [];
  // Encabezado
  ch.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [1500, 8500], rows: [new TableRow({ children: [
    celda(escudo ? para(new ImageRun({ type: 'png', data: escudo, transformation: { width: 58, height: 70 }, altText: { name: 'Escudo', title: 'Escudo', description: 'Escudo institucional' } }), { alignment: AlignmentType.CENTER }) : para(run('')), 1500),
    celda([para(run(institucion.toUpperCase(), { bold: true, size: 22 }), { alignment: AlignmentType.CENTER, spacing: { after: 20 } }), para(run('RETROALIMENTACIÓN DE LA PRUEBA', { bold: true, size: 26, color: VERDE }), { alignment: AlignmentType.CENTER })], 8500)] })] }));
  ch.push(espacio(40));
  // Datos
  const dato = (k, v) => [run(k + ': ', { bold: true }), run(v || '—'), run('     ')];
  ch.push(para([...dato('Estudiante', r.nombre), ...(R.versionB && B.length ? dato('Versión', r.version) : [])], { spacing: { after: 20 } }));
  ch.push(para([...dato('Área / asignatura', meta.area || meta.asignatura), ...dato('Tema', meta.tema), ...dato('Grado', meta.grado ? meta.grado + '°' : '')], { spacing: { after: 20 } }));
  if (docente) ch.push(para(dato('Docente', docente), { spacing: { after: 60 } }));
  // Resultado
  const caja = (titulo, valor, w, relleno) => celda([para(run(titulo, { size: 16, color: GRIS }), { alignment: AlignmentType.CENTER, spacing: { after: 0 } }), para(run(valor, { bold: true, size: 30 }), { alignment: AlignmentType.CENTER, spacing: { after: 0 } })], w, { shading: relleno ? sombra(relleno) : undefined });
  ch.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [3300, 3400, 3300], rows: [new TableRow({ children: [
    caja('Aciertos (selección múltiple)', `${r.aciertos} de ${P.preguntasMC}`, 3300), caja('Nota', fmt(r.nota), 3400), caja('Desempeño', r.desempeno, 3300, COL[r.desempeno])] })] }));
  if (P.puntosAbiertas > 0) ch.push(para(run(r.puntosAbiertas == null ? `Preguntas abiertas (${P.puntosAbiertas} puntos): aún sin calificar; la nota de arriba no las incluye todavía.` : `Preguntas abiertas: ${String(r.puntosAbiertas).replace('.', ',')} de ${P.puntosAbiertas} puntos (calificadas por el docente).`, { size: 16, italics: true, color: GRIS }), { spacing: { before: 40, after: 40 } }));
  ch.push(para(run(MENSAJE[r.desempeno], { italics: true }), { spacing: { before: 80, after: 100 } }));
  // Competencias
  ch.push(para(run('Resultado por competencia', { bold: true, size: 22, color: VERDE }), { spacing: { after: 40 } }));
  ch.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [5600, 1400, 1100, 1900], rows: [
    new TableRow({ tableHeader: true, children: ['Competencia', 'Aciertos', '%', 'Estado'].map((t, i) => celda(para(run(t, { bold: true, size: 17 }), { alignment: i ? AlignmentType.CENTER : AlignmentType.LEFT, spacing: { after: 0 } }), [5600, 1400, 1100, 1900][i], { shading: sombra('E7E6E6') })) }),
    ...r.competencias.map(c => new TableRow({ children: [celda(para(run(c.nombre, { size: 17 }), { spacing: { after: 0 } }), 5600), celda(para(run(`${c.aciertos} de ${c.total}`, { size: 17 }), { alignment: AlignmentType.CENTER, spacing: { after: 0 } }), 1400),
      celda(para(run(Math.round(c.pct * 100) + ' %', { size: 17 }), { alignment: AlignmentType.CENTER, spacing: { after: 0 } }), 1100), celda(para(run(c.estado, { bold: true, size: 17 }), { alignment: AlignmentType.CENTER, spacing: { after: 0 } }), 1900, { shading: sombra(COL[c.estado]) })] }))] }));
  const fuertes = r.competencias.filter(c => c.estado === 'Fortaleza').map(c => c.nombre), debiles = r.competencias.filter(c => c.estado === 'Por reforzar').map(c => c.nombre);
  const mejor = r.competencias.slice().sort((x, y) => y.pct - x.pct)[0];
  const textoFuerte = fuertes.length ? fuertes.join('; ') + '.' : (mejor && mejor.pct >= PROC ? `tu mejor competencia en esta prueba fue «${mejor.nombre}» (${Math.round(mejor.pct * 100)} %); sigue por ahí.` : 'todavía estás construyendo las bases; cada práctica cuenta.');
  ch.push(para([run(fuertes.length ? 'Tus fortalezas: ' : 'Lo que más lograste: ', { bold: true }), run(textoFuerte)], { spacing: { before: 80, after: 20 } }));
  if (debiles.length) ch.push(para([run('Para reforzar primero: ', { bold: true }), run(debiles.join('; ') + '.')], { spacing: { after: 80 } }));
  // Preguntas falladas
  const rank = Object.fromEntries(r.competencias.map(c => [c.nombre, c.pct]));
  const falladas = r.falladas.slice().sort((a, b) => (rank[a.q.competencia || 'Sin clasificar'] - rank[b.q.competencia || 'Sin clasificar']) || a.num - b.num);
  const MAX = MAXF, mostradas = falladas.slice(0, MAX);
  if (!falladas.length) ch.push(para(run('¡Respondiste correctamente todas las preguntas de selección múltiple!', { bold: true, color: VERDE }), { spacing: { before: 100 } }));
  else {
    ch.push(para(run(`Para revisar con tu docente (${falladas.length} ${falladas.length === 1 ? 'pregunta' : 'preguntas'})`, { bold: true, size: 22, color: VERDE }), { spacing: { before: 100, after: 40 } }));
    mostradas.forEach(f => {
      const q = f.q, marcada = f.opMarcada, ok = f.opCorrecta;
      const bloque = [para([run(`Pregunta ${f.num}`, { bold: true, size: 18 }), run(`  ·  ${q.competencia || 'Sin clasificar'}${q.dba ? '  ·  ' + q.dba : ''}`, { size: 16, color: GRIS })], { spacing: { after: 10 } })];
      if (q.pregunta) bloque.push(para(run(q.pregunta, { size: 17, italics: true }), { spacing: { after: 10 } }));
      bloque.push(para([run('Marcaste: ', { bold: true, size: 17 }), run(f.marco ? (marcada ? `${f.marco}) ${marcada.texto}` : f.marco) : 'no marcaste ninguna respuesta', { size: 17 }),
        ...(marcada && marcada.error ? [run('  →  ' + marcada.error, { size: 17, color: 'C00000' })] : [])], { spacing: { after: 10 } }));
      bloque.push(para([run('Correcta: ', { bold: true, size: 17 }), run(ok ? `${ok.letra}) ${ok.texto}` : q.clave, { size: 17 })], { spacing: { after: 10 } }));
      if (q.justificacion) bloque.push(para([run('Por qué: ', { bold: true, size: 17 }), run(q.justificacion, { size: 17 })], { spacing: { after: 10 } }));
      ch.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W], rows: [new TableRow({ cantSplit: true, children: [celda(bloque, W, { borders: { ...BORDES, left: { style: BorderStyle.SINGLE, size: 18, color: 'C00000' } } })] })] }));
      ch.push(espacio(40));
    });
    if (falladas.length > MAX) ch.push(para(run(`Y ${falladas.length - MAX} más para revisar con tu docente: preguntas ${falladas.slice(MAX).map(f => f.num).sort((a, b) => a - b).join(', ')}.`, { size: 17, italics: true }), { spacing: { after: 40 } }));
  }
  // Cierre con familias
  const linea = (t) => para([run(t + ' ', { bold: true, size: 17 }), run('_'.repeat(Math.max(10, 86 - t.length)), { size: 17, color: 'A6A6A6' })], { spacing: { before: 110, after: 0 } });
  ch.push(linea('Mi compromiso para mejorar:')); ch.push(linea('Observaciones del docente:')); ch.push(linea('Firma del acudiente:'));
  return ch;
}

const sectionProps = { page: { size: { width: 12240, height: 15840 }, margin: { top: 600, bottom: 600, left: 1120, right: 1120 } } };
const pie = new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: `${institucion} · Retroalimentación · ${meta.area || ''} ${meta.tema ? '· ' + meta.tema : ''} · Grado ${meta.grado || ''}`, font: F, size: 14, color: '808080' })] })] });
const nuevoDoc = (rs) => new Document({ creator: docente || 'Docente', title: 'Retroalimentación de la prueba', styles: { default: { document: { run: { font: F, size: 19 } } } }, compatibility: { version: 15 },
  sections: rs.map(r => ({ properties: sectionProps, footers: { default: pie }, children: paginaEstudiante(r) })) });
const slug = (t) => String(t || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^A-Za-z0-9]+/g, '_').replace(/^_+|_+$/g, '');

const paginasPdf = (f) => (fs.readFileSync(f).toString('latin1').match(/\/Type\s*\/Page[^s]/g) || []).length;
(async () => {
  const base = `Retroalimentacion_${slug(meta.area || meta.asignatura) || 'Prueba'}${meta.tema ? '_' + slug(meta.tema) : ''}_Grado${slug(meta.grado)}`;
  const archivo = path.join(outDir, base + '.docx'), fPdf = path.join(outDir, base + '.pdf');
  const salidas = [archivo]; let pdf = null;
  // Se intenta con 5, 4 y 3 preguntas detalladas por hoja hasta que cada estudiante ocupe UNA página
  for (const m of [5, 4, 3]) {
    MAXF = m; fs.writeFileSync(archivo, await Packer.toBuffer(nuevoDoc(RES)));
    if (args.includes('--sin-pdf')) break;
    const c = spawnSync('soffice', ['--headless', '--convert-to', 'pdf', '--outdir', outDir, archivo], { encoding: 'utf8', timeout: 180000 });
    if (c.status !== 0 || !fs.existsSync(fPdf)) { console.warn('⚠  No se pudo generar el PDF (¿está instalado LibreOffice?). Abre el Word y expórtalo a PDF.'); break; }
    pdf = fPdf; const pg = paginasPdf(fPdf);
    if (pg <= RES.length) { console.log(`   Ajuste: ${m} preguntas detalladas por hoja · ${pg} ${pg === 1 ? 'página' : 'páginas'} para ${RES.length} ${RES.length === 1 ? 'estudiante' : 'estudiantes'}`); break; }
    if (m === 3) console.warn(`⚠  Algunas hojas ocupan más de una página (${pg} páginas para ${RES.length} estudiantes). Revisa el PDF.`);
  }
  if (pdf) salidas.push(pdf);
  if (args.includes('--individual')) {
    const dirInd = path.join(outDir, 'individuales'); fs.mkdirSync(dirInd, { recursive: true });
    for (let i = 0; i < RES.length; i++) { const f = path.join(dirInd, `${String(i + 1).padStart(2, '0')}_${slug(RES[i].nombre)}.docx`); fs.writeFileSync(f, await Packer.toBuffer(nuevoDoc([RES[i]]))); }
  }
  const resumen = { prueba: { area: meta.area || '', tema: meta.tema || '', grado: meta.grado || '' }, parametros: P, estudiantes: RES.map(r => ({ nombre: r.nombre, version: r.version, aciertos: r.aciertos, nota: r.nota, desempeno: r.desempeno, puntosAbiertas: r.puntosAbiertas,
    competencias: Object.fromEntries(r.competencias.map(c => [c.nombre, { aciertos: c.aciertos, total: c.total, estado: c.estado }])), falladas: r.falladas.map(f => f.num) })) };
  fs.writeFileSync(path.join(outDir, 'retroalimentacion.json'), JSON.stringify(resumen, null, 2));
  salidas.forEach(f => console.log('📄 ' + f));
  const dist = RES.reduce((a, r) => (a[r.desempeno] = (a[r.desempeno] || 0) + 1, a), {});
  console.log(`✅ Retroalimentación de ${RES.length} ${RES.length === 1 ? 'estudiante' : 'estudiantes'} · ${Object.entries(dist).map(([k, v]) => `${k}: ${v}`).join(' · ')}${pdf ? '' : ' · (sin PDF)'}`);
})().catch(e => fail(e.stack || e.message));
