#!/usr/bin/env node
/* ============================================================================
 *  build_exam.js  —  GENERADOR DE EXAMEN TIPO SABER v2.0
 *  Entrada : exam.json  (meta + config + preguntas)
 *  Salida  : <base>_ESTUDIANTE.docx   y   <base>_DOCENTE_CLAVES.docx
 *  Uso     : node build_exam.js exam.json [--nivel 0|1|2|3] [--out ./salida]
 *  Autor del flujo pedagógico: Mg. Luis Alberto Avellaneda C.
 * ========================================================================== */
const fs = require('fs');
const path = require('path');
const lib = (m) => { try { return require(m); } catch (e) { return require(path.join(process.env.NODE_PATH || '/home/claude/.npm-global/lib/node_modules', m)); } };
const D = lib('docx');
const sharp = lib('sharp');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, AlignmentType, BorderStyle,
        WidthType, ShadingType, SectionType, ImageRun, VerticalAlign, Header, PageNumber, Tab,
        TabStopType, TableLayoutType, HeightRule } = D;

// ═══ 1. ESCUDO INSTITUCIONAL ════════════════════════════════════════════════
// Escudo optimizado para encabezado (PNG 200×242, ~9 KB) incrustado en base64.
// Es el ÚLTIMO recurso: solo se usa si no se encuentra el archivo en disco.
const ESCUDO_B64 = "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";

function localizarEscudo() {
  if (process.env.ESCUDO_USAR_BASE64 === '1' && ESCUDO_B64 && !ESCUDO_B64.startsWith('__')) return { buf: Buffer.from(ESCUDO_B64, 'base64'), origen: 'base64 incrustado (forzado)' };
  const dirs = [process.env.ESCUDO_PATH,
    path.join(__dirname, '..', 'assets', 'escudo_header.png'),
    path.join(__dirname, 'assets', 'escudo_header.png'),
    path.join(process.cwd(), 'assets', 'escudo_header.png'),
    path.join(process.cwd(), 'escudo_header.png'),
    '/mnt/skills/user/generador-examen-tipo-saber/assets/escudo_header.png'];
  try { for (const f of fs.readdirSync('/mnt/user-data/uploads')) if (/escudo.*\.png$/i.test(f)) dirs.push('/mnt/user-data/uploads/' + f); } catch (e) {}
  for (const p of dirs) { try { if (p && fs.existsSync(p)) return { buf: fs.readFileSync(p), origen: p }; } catch (e) {} }
  if (ESCUDO_B64 && !ESCUDO_B64.startsWith('__')) return { buf: Buffer.from(ESCUDO_B64, 'base64'), origen: 'base64 incrustado' };
  return null;
}
async function prepararEscudo() {
  const f = localizarEscudo();
  if (!f) throw new Error('ESCUDO NO ENCONTRADO: ni archivo ni base64. Abortando para no entregar un examen sin logo.');
  const sig = f.buf.slice(0, 8).toString('hex');
  if (sig !== '89504e470d0a1a0a') throw new Error('El escudo (' + f.origen + ') no es un PNG válido.');
  // Normaliza: fondo blanco, ancho máx. 300 px, PNG (calidad de impresión ≈ 400 dpi en 0.75")
  const buf = await sharp(f.buf).flatten({ background: '#ffffff' }).resize({ width: 300, withoutEnlargement: true }).png().toBuffer();
  const m = await sharp(buf).metadata();
  if (m.width < 60 || m.height < 60) throw new Error('El escudo es demasiado pequeño (' + m.width + '×' + m.height + ').');
  console.log(`🛡️  Escudo OK  → origen: ${f.origen} · ${m.width}×${m.height}px`);
  return { buf, w: m.width, h: m.height, origen: f.origen };
}

// ═══ 2. CONSTANTES DE PÁGINA (una sola fuente de verdad → todo queda alineado) ═
const PAGE_W = 12240, PAGE_H = 15840;
const MARG = { top: 1750, bottom: 720, left: 864, right: 864, header: 340, footer: 300 };
const TEXT_W = PAGE_W - MARG.left - MARG.right;           // 10512 DXA = ancho útil REAL
const COL_GAP = 720;
const COL_W = Math.floor((TEXT_W - COL_GAP) / 2);          // 4896 DXA
const F = 'Arial', SZ = 20, SZ_S = 18, SZ_XS = 16;
const none = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const bNone = { top: none, bottom: none, left: none, right: none };
const line = (c = '808080', s = 4) => ({ style: BorderStyle.SINGLE, size: s, color: c });
const bAll = (c = '808080', s = 4) => ({ top: line(c, s), bottom: line(c, s), left: line(c, s), right: line(c, s) });
const T0 = { top: 0, bottom: 0, left: 0, right: 0 };

// ═══ 3. ARITMÉTICA RACIONAL (verificación automática de respuestas) ══════════
const bg = (a, b) => { a = a < 0n ? -a : a; b = b < 0n ? -b : b; while (b) [a, b] = [b, a % b]; return a || 1n; };
const Q = (n, d = 1n) => { if (d < 0n) { n = -n; d = -d; } const g = bg(n, d); return { n: n / g, d: d / g }; };
const qAdd = (a, b) => Q(a.n * b.d + b.n * a.d, a.d * b.d), qSub = (a, b) => Q(a.n * b.d - b.n * a.d, a.d * b.d);
const qMul = (a, b) => Q(a.n * b.n, a.d * b.d), qDiv = (a, b) => { if (b.n === 0n) throw new Error('división por cero'); return Q(a.n * b.d, a.d * b.n); };
const qStr = (q) => q.d === 1n ? String(q.n) : `${q.n}/${q.d}`;
const qCmp = (a, b) => { const l = a.n * b.d, r = b.n * a.d; return l < r ? -1 : l > r ? 1 : 0; };
const qDec = (s) => { s = s.replace(',', '.'); const [i, f = ''] = s.split('.'); return Q(BigInt(i + f), 10n ** BigInt(f.length)); };
function evalExpr(src) {                       // + - * / ( ) y decimales; sin eval()
  const t = src.replace(/×|·/g, '*').replace(/÷/g, '/').match(/\d+(?:[.,]\d+)?|[()+\-*/]/g) || [];
  let i = 0;
  const prim = () => { const k = t[i++]; if (k === '(') { const v = sum(); i++; return v; } if (k === '-') return qSub(Q(0n), prim()); if (!/^\d/.test(k || '')) throw new Error('expresión inválida: ' + src); return qDec(k); };
  const prod = () => { let v = prim(); while (t[i] === '*' || t[i] === '/') { const o = t[i++]; v = o === '*' ? qMul(v, prim()) : qDiv(v, prim()); } return v; };
  const sum = () => { let v = prod(); while (t[i] === '+' || t[i] === '-') { const o = t[i++]; v = o === '+' ? qAdd(v, prod()) : qSub(v, prod()); } return v; };
  const r = sum(); if (i < t.length) throw new Error('expresión inválida: ' + src); return r;
}
const NUM_RE = /(\d{1,3}(?:\.\d{3})+(?!\d))|(\d+)\s+(\d+)\s*\/\s*(\d+)|(\d+)\s*\/\s*(\d+)|(\d+(?:[.,]\d+)?)(\s*%)?/g;
function numerosDe(txt) {
  const out = [];
  for (const m of String(txt).matchAll(NUM_RE)) {
    let q;
    if (m[1]) q = Q(BigInt(m[1].replace(/\./g, '')));
    else if (m[2]) q = qAdd(Q(BigInt(m[2])), Q(BigInt(m[3]), BigInt(m[4])));
    else if (m[5]) q = Q(BigInt(m[5]), BigInt(m[6]));
    else { q = qDec(m[7]); if (m[8]) q = qDiv(q, Q(100n)); }
    out.push(q);
  }
  return out;
}
const claveOpcion = (txt) => String(txt).replace(NUM_RE, (...a) => { const q = numerosDe(a[0])[0]; return '#' + qStr(q) + '#'; })
  .toLowerCase().replace(/\s+/g, ' ').trim();
const valorUnico = (txt) => { const n = numerosDe(txt); return n.length === 1 ? n[0] : null; };

// ═══ 4. VALIDACIÓN DE LA PRUEBA ═════════════════════════════════════════════
const RANGOS = { corto: [80, 180], medio: [250, 420], largo: [550, 800] };
function validar(ex) {
  const E = [], W = [];
  const mc = ex.preguntas.filter(p => (p.tipo || 'multiple') === 'multiple');
  const ab = ex.preguntas.filter(p => p.tipo === 'abierta');
  const rango = RANGOS[(ex.config.longitud || '').toLowerCase()];
  mc.forEach((p, k) => {
    const id = `P${k + 1}`;
    if (!p.enunciado || !p.pregunta) E.push(`${id}: falta enunciado o pregunta`);
    if (rango && p.enunciado && (p.enunciado.length < rango[0] || p.enunciado.length > rango[1]))
      E.push(`${id}: enunciado de ${p.enunciado.length} caracteres; para "${ex.config.longitud}" debe estar entre ${rango[0]} y ${rango[1]}`);
    if (!p.correcta || !Array.isArray(p.distractores) || p.distractores.length !== 3) { E.push(`${id}: requiere "correcta" (texto) y 3 "distractores"`); return; }
    const ops = [p.correcta, ...p.distractores].map(s => String(s).trim());
    if (ops.some(o => !o)) E.push(`${id}: hay una opción vacía`);
    if (ops.some(o => /(ninguna|todas)\s+(de\s+)?las\s+(anteriores|opciones)/i.test(o))) E.push(`${id}: prohibido "ninguna/todas las anteriores"`);
    if (/\bNO\b/.test(p.pregunta) && !/\bNO\b.*\bNO\b/.test(p.pregunta)) W.push(`${id}: la pregunta usa "NO" (negación); mejor formularla en positivo`);
    const claves = ops.map(claveOpcion);
    claves.forEach((c, i) => claves.forEach((d, j) => { if (j > i && c === d) E.push(`${id}: las opciones "${ops[i]}" y "${ops[j]}" son EQUIVALENTES (dos respuestas verdaderas)`); }));
    if (p.calculo) {
      try {
        const r = evalExpr(p.calculo);
        const hits = ops.map((o, i) => { const v = valorUnico(o); return v && qCmp(v, r) === 0 ? i : -1; }).filter(i => i >= 0);
        if (!hits.length) E.push(`${id}: el cálculo "${p.calculo}" = ${qStr(r)} NO aparece entre las opciones`);
        else if (hits.length > 1) E.push(`${id}: más de una opción vale ${qStr(r)}`);
        else if (hits[0] !== 0) E.push(`${id}: el cálculo da ${qStr(r)} pero la "correcta" declarada es "${p.correcta}"`);
      } catch (e) { E.push(`${id}: "calculo" no se pudo evaluar (${e.message})`); }
    } else if (ops.every(o => valorUnico(o))) W.push(`${id}: opciones numéricas sin campo "calculo" (no se pudo verificar automáticamente)`);
    if (!p.justificacion) W.push(`${id}: falta "justificacion" para la clave docente`);
    if (!Array.isArray(p.errores) || p.errores.length !== 3) W.push(`${id}: falta "errores" (3) — qué error conceptual representa cada distractor`);
    const v = p.visual;
    if (v && v.tipo === 1 && !v.plantilla) E.push(`${id}: visual tipo 1 sin "plantilla"`);
  });
  ab.forEach((p, k) => { if (!p.enunciado) E.push(`A${k + 1}: falta enunciado`); if (!p.respuestaModelo) W.push(`A${k + 1}: falta "respuestaModelo"`); });
  mc.forEach((p, k) => { if (k && p.visual && mc[k - 1].visual) W.push(`P${k + 1}: dos preguntas con elemento visual consecutivas (P${k} y P${k + 1}); espaciarlas`); });
  mc.forEach((p, k) => { if (p.nivel !== undefined && ![1, 2, 3, 4].includes(Number(p.nivel))) E.push(`P${k + 1}: "nivel" debe ser 1, 2, 3 o 4 (nivel de desempeño SABER)`); });
  const conComp = mc.filter(p => p.competencia).length;
  if (!conComp) W.push('Ninguna pregunta tiene "competencia": no se generará la matriz de especificaciones');
  else if (conComp < mc.length) W.push(`Solo ${conComp} de ${mc.length} preguntas tienen "competencia"; el resto saldrá como "Sin clasificar" en la matriz`);
  if (!mc.length) E.push('No hay preguntas de selección múltiple');
  return { E, W };
}

// Asignación de letras: ordena numéricas, baraja el resto y balancea A-D (máx. 2 iguales seguidas)
function mulberry(seed) { let a = seed >>> 0; return () => { a |= 0; a = a + 0x6D2B79F5 | 0; let t = Math.imul(a ^ a >>> 15, 1 | a); t = t + Math.imul(t ^ t >>> 7, 61 | t) ^ t; return ((t ^ t >>> 14) >>> 0) / 4294967296; }; }
function asignarLetras(mc, seed) {
  const rnd = mulberry(seed), L = 'ABCD';
  const items = mc.map(p => ({ p, ops: [{ t: p.correcta, ok: true, e: null }, ...p.distractores.map((t, i) => ({ t, ok: false, e: (p.errores || [])[i] || '' }))] }));
  const letras = new Array(items.length).fill(null);
  const viola = (i, l) => { const a = letras.slice(); a[i] = l; for (let s = Math.max(0, i - 2); s <= i; s++) if (a[s] && a[s] === a[s + 1] && a[s] === a[s + 2]) return true; return false; };
  items.forEach((it, i) => {            // 1) numéricas ordenadas
    const vs = it.ops.map(o => valorUnico(o.t));
    if (vs.every(Boolean) && it.p.ordenar !== false && new Set(vs.map(qStr)).size === 4) {
      let asc = [...it.ops].sort((a, b) => qCmp(valorUnico(a.t), valorUnico(b.t)));
      let l = L[asc.findIndex(o => o.ok)];
      if (viola(i, l)) { asc = asc.reverse(); l = L[asc.findIndex(o => o.ok)]; }
      it.orden = asc; letras[i] = l;
    }
  });
  const cnt = { A: 0, B: 0, C: 0, D: 0 }; letras.forEach(l => l && cnt[l]++);
  items.forEach((it, i) => {            // 2) el resto: letra objetivo balanceada
    if (it.orden) return;
    const cand = [...L].filter(l => !viola(i, l)); const pool = cand.length ? cand : [...L];
    const min = Math.min(...pool.map(l => cnt[l])); const best = pool.filter(l => cnt[l] === min);
    const l = best[Math.floor(rnd() * best.length)];
    const dis = it.ops.filter(o => !o.ok).map(o => [rnd(), o]).sort((a, b) => a[0] - b[0]).map(x => x[1]);
    const arr = []; let di = 0; for (let k = 0; k < 4; k++) arr.push(L[k] === l ? it.ops[0] : dis[di++]);
    it.orden = arr; letras[i] = l; cnt[l]++;
  });
  items.forEach(it => { it.p._ops = it.orden.map((o, k) => ({ letra: L[k], texto: o.t, ok: o.ok, error: o.e })); it.p._clave = L[it.orden.findIndex(o => o.ok)]; });
  const dist = { A: 0, B: 0, C: 0, D: 0 }; mc.forEach(p => dist[p._clave]++);
  return dist;
}

// ═══ 5. IMÁGENES PNG DE ALTA RESOLUCIÓN (plantillas ilustradas) ═════════════
const FONT = `font-family="DejaVu Sans, Arial, Helvetica, sans-serif"`;
const DEFS = `<defs>
<filter id="blur" x="-30%" y="-30%" width="160%" height="160%"><feGaussianBlur stdDeviation="9"/></filter>
<filter id="blur2" x="-30%" y="-30%" width="160%" height="160%"><feGaussianBlur stdDeviation="3"/></filter>
<linearGradient id="gCrust" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#F0BE78"/><stop offset="1" stop-color="#C98A3A"/></linearGradient>
<radialGradient id="gCheese" cx=".4" cy=".35" r=".8"><stop offset="0" stop-color="#FFE27A"/><stop offset="1" stop-color="#F5B82E"/></radialGradient>
<radialGradient id="gPep" cx=".35" cy=".3" r=".8"><stop offset="0" stop-color="#E0523F"/><stop offset="1" stop-color="#A52A1E"/></radialGradient>
<linearGradient id="gChoc" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#8A5236"/><stop offset="1" stop-color="#4E2B18"/></linearGradient>
<linearGradient id="gWater" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#7CC8F2"/><stop offset="1" stop-color="#2B8FD6"/></linearGradient>
<linearGradient id="gGlass" x1="0" y1="0" x2="1" y2="0"><stop offset="0" stop-color="#EAF6FD"/><stop offset=".5" stop-color="#FFFFFF"/><stop offset="1" stop-color="#DCEFFA"/></linearGradient>
<linearGradient id="gGrass" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#7BC96F"/><stop offset="1" stop-color="#3E9B4F"/></linearGradient>
<linearGradient id="gSoil" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#A9744A"/><stop offset="1" stop-color="#7B4B2A"/></linearGradient>
<linearGradient id="gPool" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#8ED8F8"/><stop offset="1" stop-color="#3AA6E0"/></linearGradient>
<linearGradient id="gB1" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#5B9BF0"/><stop offset="1" stop-color="#2E62C0"/></linearGradient>
<linearGradient id="gB2" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#F4A261"/><stop offset="1" stop-color="#D4691F"/></linearGradient>
<linearGradient id="gB3" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#6FCF97"/><stop offset="1" stop-color="#2E9D63"/></linearGradient>
<linearGradient id="gB4" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#C48BE8"/><stop offset="1" stop-color="#8A4FC0"/></linearGradient>
<linearGradient id="gShade" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#5B9BF0"/><stop offset="1" stop-color="#2E62C0"/></linearGradient>
</defs>`;
const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const wrapSVG = (inner, W = 900, H = 600) => `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}">${DEFS}<rect width="${W}" height="${H}" fill="#FFFFFF"/>${inner}</svg>`;
const txt = (x, y, s, size = 34, o = {}) => `<text x="${x}" y="${y}" ${FONT} font-size="${size}" font-weight="${o.w || 'bold'}" text-anchor="${o.a || 'middle'}" fill="${o.c || '#1F2937'}">${esc(s)}</text>`;
function leyenda(items, y = 560) {                 // items: [{k:'ghost'|'fill', fill, t}]
  items = items.filter(i => i && i.t); if (!items.length) return '';
  const w = items.map(i => 70 + i.t.length * 17); const tot = w.reduce((a, b) => a + b, 0) + 40 * (items.length - 1);
  let x = (900 - tot) / 2, out = '';
  items.forEach((i, k) => {
    out += i.k === 'ghost'
      ? `<rect x="${x}" y="${y - 24}" width="40" height="40" rx="8" fill="#F3F4F6" stroke="#6B7280" stroke-width="3" stroke-dasharray="7 5"/>`
      : `<rect x="${x}" y="${y - 24}" width="40" height="40" rx="8" fill="${i.fill || 'url(#gChoc)'}" stroke="#1F2937" stroke-width="2"/>`;
    out += txt(x + 54, y + 8, i.t, 30, { a: 'start' }); x += w[k] + 40;
  });
  return out;
}
const polar = (cx, cy, r, deg) => [cx + r * Math.cos((deg - 90) * Math.PI / 180), cy + r * Math.sin((deg - 90) * Math.PI / 180)];
function wedge(cx, cy, r, a0, a1) { const [x0, y0] = polar(cx, cy, r, a0), [x1, y1] = polar(cx, cy, r, a1); const lg = (a1 - a0) > 180 ? 1 : 0; return `M${cx},${cy} L${x0.toFixed(1)},${y0.toFixed(1)} A${r},${r} 0 ${lg} 1 ${x1.toFixed(1)},${y1.toFixed(1)} Z`; }

const PLANTILLAS = {
  // pizza / torta / círculo dividido en partes iguales; "marcadas" = partes que ya no están (comidas)
  pizza(p) {
    const n = p.partes, m = p.marcadas || 0, cx = 450, cy = 262, R = 225, st = 360 / n; let s = `<ellipse cx="450" cy="508" rx="230" ry="26" fill="#000" opacity=".22" filter="url(#blur)"/>`;
    const torta = p.estilo === 'torta';
    for (let i = 0; i < n; i++) {
      const a0 = i * st, a1 = (i + 1) * st, am = (a0 + a1) / 2;
      if (i < m) { s += `<path d="${wedge(cx, cy, R, a0, a1)}" fill="#F3F4F6" stroke="#6B7280" stroke-width="3.5" stroke-dasharray="9 7" stroke-linejoin="round"/>`; continue; }
      if (torta) {
        s += `<path d="${wedge(cx, cy, R, a0, a1)}" fill="${i % 2 ? '#F8B6C8' : '#F6A5BB'}" stroke="#B5456A" stroke-width="3" stroke-linejoin="round"/>`;
        s += `<path d="${wedge(cx, cy, R * .62, a0 + 3, a1 - 3)}" fill="#FFE9F0" opacity=".7"/>`;
        const [fx, fy] = polar(cx, cy, R * .78, am); s += `<circle cx="${fx}" cy="${fy}" r="20" fill="url(#gPep)" stroke="#7A1B12" stroke-width="2"/><circle cx="${fx - 6}" cy="${fy - 7}" r="5" fill="#fff" opacity=".7"/>`;
      } else {
        s += `<path d="${wedge(cx, cy, R, a0, a1)}" fill="url(#gCrust)" stroke="#8A5A1E" stroke-width="3" stroke-linejoin="round"/>`;
        s += `<path d="${wedge(cx, cy, R - 24, a0 + .8, a1 - .8)}" fill="url(#gCheese)" stroke="#D99A1E" stroke-width="2" stroke-linejoin="round"/>`;
        [[.5, 0], [.78, 0]].forEach(([f]) => { const [px, py] = polar(cx, cy, (R - 24) * f, am); const rr = n > 10 ? 14 : 22; s += `<circle cx="${px.toFixed(1)}" cy="${py.toFixed(1)}" r="${rr}" fill="url(#gPep)" stroke="#7A1B12" stroke-width="2"/><circle cx="${(px - rr * .3).toFixed(1)}" cy="${(py - rr * .3).toFixed(1)}" r="${rr * .22}" fill="#fff" opacity=".6"/>`; });
      }
    }
    return s + leyenda([{ k: 'ghost', t: p.etiquetaMarcadas }, p.etiquetaRestantes && { k: 'fill', fill: torta ? '#F6A5BB' : 'url(#gCheese)', t: p.etiquetaRestantes }], 570);
  },
  barra(p) {                                           // barra de chocolate
    const c = p.cols, f = p.filas || 1, total = c * f, m = p.marcadas || 0;
    const cw = Math.min(150, 700 / c), ch = Math.min(150, 330 / f), W = cw * c, H = ch * f, x0 = (900 - W) / 2, y0 = 250 - H / 2;
    let s = `<rect x="${x0 + 10}" y="${y0 + H + 6}" width="${W}" height="26" rx="13" fill="#000" opacity=".2" filter="url(#blur)"/>`;
    for (let i = 0; i < total; i++) {
      const x = x0 + (i % c) * cw, y = y0 + Math.floor(i / c) * ch;
      if (i < m) s += `<rect x="${x + 4}" y="${y + 4}" width="${cw - 8}" height="${ch - 8}" rx="12" fill="#F3F4F6" stroke="#6B7280" stroke-width="3.5" stroke-dasharray="9 7"/>`;
      else s += `<rect x="${x + 3}" y="${y + 3}" width="${cw - 6}" height="${ch - 6}" rx="12" fill="url(#gChoc)" stroke="#2E1709" stroke-width="3"/><rect x="${x + 16}" y="${y + 16}" width="${cw - 32}" height="${ch - 32}" rx="8" fill="#9B5F3F" opacity=".55"/><rect x="${x + 16}" y="${y + 16}" width="${cw - 32}" height="${(ch - 32) / 2}" rx="8" fill="#fff" opacity=".12"/>`;
    }
    return s + leyenda([p.etiquetaMarcadas && { k: 'ghost', t: p.etiquetaMarcadas }, p.etiquetaRestantes && { k: 'fill', t: p.etiquetaRestantes }], 540);
  },
  cuadricula(p) {                                      // figura dividida en celdas, n sombreadas
    const c = p.cols, f = p.filas, m = p.sombreadas, cw = Math.min(130, 720 / c), ch = Math.min(130, 360 / f), W = cw * c, H = ch * f, x0 = (900 - W) / 2, y0 = 255 - H / 2;
    let s = `<rect x="${x0 + 8}" y="${y0 + H + 4}" width="${W}" height="22" rx="11" fill="#000" opacity=".18" filter="url(#blur)"/>`;
    for (let i = 0; i < c * f; i++) { const x = x0 + (i % c) * cw, y = y0 + Math.floor(i / c) * ch; s += `<rect x="${x}" y="${y}" width="${cw}" height="${ch}" fill="${i < m ? 'url(#gShade)' : '#FFFFFF'}" stroke="#1F2937" stroke-width="4"/>`; if (i < m) s += `<rect x="${x + 6}" y="${y + 6}" width="${cw - 12}" height="${(ch - 12) / 2}" fill="#fff" opacity=".16"/>`; }
    s += `<rect x="${x0}" y="${y0}" width="${W}" height="${H}" fill="none" stroke="#111827" stroke-width="7" stroke-linejoin="round"/>`;
    return s + leyenda([p.etiquetaSombreadas && { k: 'fill', fill: 'url(#gShade)', t: p.etiquetaSombreadas }, p.etiquetaRestantes && { k: 'ghost', t: p.etiquetaRestantes }], 550);
  },
  recta(p) {                                           // recta numérica
    const x0 = 80, x1 = 820, y = 300, d = p.divisiones, step = (x1 - x0) / d; let s = '';
    s += `<line x1="${x0 - 30}" y1="${y}" x2="${x1 + 30}" y2="${y}" stroke="#111827" stroke-width="7" stroke-linecap="round"/><path d="M${x1 + 44},${y} l-24,-15 v30 z" fill="#111827"/><path d="M${x0 - 44},${y} l24,-15 v30 z" fill="#111827"/>`;
    for (let k = 0; k <= d; k++) { const x = x0 + k * step, big = (k === 0 || k === d); s += `<line x1="${x}" y1="${y - (big ? 34 : 22)}" x2="${x}" y2="${y + (big ? 34 : 22)}" stroke="#111827" stroke-width="${big ? 7 : 5}"/>`; }
    s += txt(x0, y + 90, p.etiquetaInicio ?? '0', 46) + txt(x1, y + 90, p.etiquetaFin ?? '1', 46);
    (p.marcas || []).forEach(m => { const x = x0 + m.pos * step; s += `<rect x="${x - 62}" y="${y - 118}" width="124" height="60" rx="14" fill="#EEF4FF" stroke="#2E62C0" stroke-width="3"/>` + txt(x, y - 76, m.etiqueta, 38, { c: '#1E3A8A' }) + `<line x1="${x}" y1="${y - 56}" x2="${x}" y2="${y - 36}" stroke="#2E62C0" stroke-width="4"/>`; });
    if (p.punto) { const x = x0 + p.punto.pos * step; s += `<circle cx="${x}" cy="${y}" r="21" fill="#E11D48" stroke="#7F1029" stroke-width="4"/><circle cx="${x - 6}" cy="${y - 7}" r="6" fill="#fff" opacity=".6"/><path d="M${x},${y + 26} l-13,34 h26 z" fill="#E11D48" opacity="0"/>` + (p.punto.etiqueta ? txt(x, y + 140, p.punto.etiqueta, 40, { c: '#BE123C' }) : ''); }
    return s + (p.nota ? txt(450, 520, p.nota, 34, { w: 'normal', c: '#374151' }) : '');
  },
  vaso(p) {                                            // vaso / recipiente graduado
    const [nu, de] = p.nivel, top = 80, bot = 500, f = nu / de, yL = bot - (bot - top) * f; const CX = 450, tw = (y) => 150 + (y - top) / (bot - top) * -20;  // semiancho
    const L = (y) => CX - tw(y), R = (y) => CX + tw(y);
    let s = `<ellipse cx="450" cy="515" rx="150" ry="22" fill="#000" opacity=".22" filter="url(#blur)"/>`;
    s += `<path d="M${L(top)},${top} L${R(top)},${top} L${R(bot)},${bot} L${L(bot)},${bot} Z" fill="url(#gGlass)" stroke="#5B7083" stroke-width="7" stroke-linejoin="round"/>`;
    s += `<path d="M${L(yL)},${yL} L${R(yL)},${yL} L${R(bot) - 4},${bot - 4} L${L(bot) + 4},${bot - 4} Z" fill="url(#gWater)" opacity=".92"/><ellipse cx="450" cy="${yL}" rx="${tw(yL) - 2}" ry="9" fill="#BFE6FA" stroke="#2B8FD6" stroke-width="3"/>`;
    for (let k = 1; k < de; k++) { const y = bot - (bot - top) * k / de; s += `<line x1="${R(y) - 36}" y1="${y}" x2="${R(y)}" y2="${y}" stroke="#334155" stroke-width="5"/>`; }
    s += `<path d="M${L(top) + 22},${top + 30} L${L(bot) + 34},${bot - 30}" stroke="#fff" stroke-width="12" opacity=".7" stroke-linecap="round"/>`;
    if (p.etiquetaCapacidad) s += `<rect x="20" y="70" width="250" height="120" rx="18" fill="#F1F5F9" stroke="#475569" stroke-width="3"/>` + txt(145, 114, 'Capacidad total', 28, { w: 'normal' }) + txt(145, 164, p.etiquetaCapacidad, 40);
    if (p.etiquetaNivel) s += `<line x1="${R(yL) + 6}" y1="${yL}" x2="640" y2="${yL}" stroke="#2B8FD6" stroke-width="5" stroke-dasharray="10 7"/><rect x="640" y="${yL - 38}" width="250" height="76" rx="16" fill="#E0F2FE" stroke="#2B8FD6" stroke-width="3"/>` + txt(765, yL + 14, p.etiquetaNivel, 36, { c: '#075985' });
    return s;
  },
  barras(p) {                                          // gráfico de barras
    const cats = p.categorias, x0 = 150, x1 = 840, y0 = 470, y1 = 120, max = p.max || Math.max(...cats.map(c => c.valor)), step = p.paso || (max <= 10 ? 1 : max <= 20 ? 2 : max <= 50 ? 5 : 10);
    let s = txt(450, 70, p.titulo || '', 38);
    for (let v = 0; v <= max; v += step) { const y = y0 - (y0 - y1) * v / max; s += `<line x1="${x0}" y1="${y}" x2="${x1}" y2="${y}" stroke="${v === 0 ? '#111827' : '#D1D5DB'}" stroke-width="${v === 0 ? 5 : 2}"/>` + txt(x0 - 14, y + 11, v, 30, { a: 'end', w: 'normal', c: '#374151' }); }
    s += `<line x1="${x0}" y1="${y1 - 10}" x2="${x0}" y2="${y0}" stroke="#111827" stroke-width="5"/>`;
    const bw = Math.min(110, (x1 - x0) / cats.length * .62), gap = (x1 - x0) / cats.length;
    cats.forEach((c, i) => { const cx = x0 + gap * (i + .5), h = (y0 - y1) * (c.oculto ? .5 * max : c.valor) / max, yb = y0 - h; s += (c.oculto ? `<rect x="${cx - bw / 2}" y="${yb}" width="${bw}" height="${h}" rx="6" fill="#F3F4F6" stroke="#6B7280" stroke-width="3.5" stroke-dasharray="10 8"/>` : `<rect x="${cx - bw / 2 + 6}" y="${yb + 6}" width="${bw}" height="${h}" rx="6" fill="#000" opacity=".12"/><rect x="${cx - bw / 2}" y="${yb}" width="${bw}" height="${h}" rx="6" fill="url(#gB${i % 4 + 1})" stroke="#1F2937" stroke-width="2.5"/>`) + (c.oculto ? txt(cx, yb + h / 2 + 18, '?', 60, { c: '#BE123C' }) : (p.mostrarValores === false ? '' : txt(cx, yb - 14, c.etiquetaValor ?? c.valor, 32))) + txt(cx, y0 + 44, c.nombre, 30, { w: 'normal' }); });
    return s + (p.ejeY ? `<text x="38" y="300" ${FONT} font-size="28" fill="#374151" text-anchor="middle" transform="rotate(-90 38 300)">${esc(p.ejeY)}</text>` : '');
  },
  cancha(p) {                                          // cancha / piscina / lote con medidas
    const x0 = 140, y0 = 130, W = 620, H = 300, est = p.estilo || 'cancha'; let s = `<rect x="${x0 + 10}" y="${y0 + H + 4}" width="${W}" height="30" rx="15" fill="#000" opacity=".2" filter="url(#blur)"/>`;
    if (est === 'piscina') { s += `<rect x="${x0 - 12}" y="${y0 - 12}" width="${W + 24}" height="${H + 24}" rx="14" fill="#E5E7EB" stroke="#9CA3AF" stroke-width="4"/><rect x="${x0}" y="${y0}" width="${W}" height="${H}" rx="6" fill="url(#gPool)" stroke="#1D6FA5" stroke-width="5"/>`; for (let k = 1; k < 5; k++) s += `<line x1="${x0}" y1="${y0 + H * k / 5}" x2="${x0 + W}" y2="${y0 + H * k / 5}" stroke="#fff" stroke-width="5" stroke-dasharray="26 14" opacity=".85"/>`; }
    else if (est === 'lote') s += `<rect x="${x0}" y="${y0}" width="${W}" height="${H}" rx="6" fill="url(#gSoil)" stroke="#4B2E16" stroke-width="5"/>`;
    else { s += `<rect x="${x0}" y="${y0}" width="${W}" height="${H}" rx="6" fill="url(#gGrass)" stroke="#fff" stroke-width="7"/><line x1="${x0 + W / 2}" y1="${y0}" x2="${x0 + W / 2}" y2="${y0 + H}" stroke="#fff" stroke-width="6"/><circle cx="${x0 + W / 2}" cy="${y0 + H / 2}" r="56" fill="none" stroke="#fff" stroke-width="6"/><rect x="${x0}" y="${y0 + H / 2 - 70}" width="70" height="140" fill="none" stroke="#fff" stroke-width="6"/><rect x="${x0 + W - 70}" y="${y0 + H / 2 - 70}" width="70" height="140" fill="none" stroke="#fff" stroke-width="6"/>`; }
    const arrow = (xa, ya, xb, yb) => `<line x1="${xa}" y1="${ya}" x2="${xb}" y2="${yb}" stroke="#111827" stroke-width="5"/>`;
    s += arrow(x0, y0 + H + 60, x0 + W, y0 + H + 60) + `<path d="M${x0},${y0 + H + 60} l22,-13 v26 z M${x0 + W},${y0 + H + 60} l-22,-13 v26 z" fill="#111827"/>` + txt(x0 + W / 2, y0 + H + 112, p.largo || '', 40);
    s += arrow(x0 + W + 52, y0, x0 + W + 52, y0 + H) + `<path d="M${x0 + W + 52},${y0} l-13,22 h26 z M${x0 + W + 52},${y0 + H} l-13,-22 h26 z" fill="#111827"/>` + `<text x="${x0 + W + 100}" y="${y0 + H / 2}" ${FONT} font-size="40" font-weight="bold" fill="#1F2937" text-anchor="middle" transform="rotate(90 ${x0 + W + 100} ${y0 + H / 2})">${esc(p.ancho || '')}</text>`;
    return s;
  },
  huerta(p) {                                          // surcos de huerta
    const n = p.surcos, k = p.conPlantas, x0 = 110, W = 680, H = Math.min(64, 330 / n), y0 = 250 - (H * n) / 2; let s = `<rect x="${x0 - 14}" y="${y0 - 14}" width="${W + 28}" height="${H * n + 28}" rx="14" fill="#7B4B2A" opacity=".25" filter="url(#blur2)"/>`;
    for (let i = 0; i < n; i++) {
      const y = y0 + i * H; s += `<rect x="${x0}" y="${y + 3}" width="${W}" height="${H - 6}" rx="12" fill="url(#gSoil)" stroke="#4B2E16" stroke-width="3"/>`;
      if (i < k) for (let j = 0; j < 9; j++) { const px = x0 + 50 + j * ((W - 100) / 8), py = y + H / 2; s += `<line x1="${px}" y1="${py + 12}" x2="${px}" y2="${py - 4}" stroke="#2F7A3C" stroke-width="5"/><ellipse cx="${px - 11}" cy="${py - 6}" rx="12" ry="7" fill="#58B35F" transform="rotate(-30 ${px - 11} ${py - 6})"/><ellipse cx="${px + 11}" cy="${py - 6}" rx="12" ry="7" fill="#6FCF77" transform="rotate(30 ${px + 11} ${py - 6})"/>`; }
    }
    return s + leyenda([p.etiquetaPlantas && { k: 'fill', fill: 'url(#gGrass)', t: p.etiquetaPlantas }, p.etiquetaVacios && { k: 'fill', fill: 'url(#gSoil)', t: p.etiquetaVacios }], 545);
  },
  svg(p) { return p.svg; }                             // SVG propio (solo si ninguna plantilla sirve)
};
async function renderPNG(inner, file, W = 900, H = 600) {
  const svg = inner.startsWith('<svg') ? inner : wrapSVG(inner, W, H);
  await sharp(Buffer.from(svg), { density: 96 * 900 / W * 1.8 }).resize({ width: 1400 }).flatten({ background: '#fff' }).png({ compressionLevel: 9 }).toFile(file);
}
async function bubble(letra, relleno, file) {
  const s = `<svg xmlns="http://www.w3.org/2000/svg" width="120" height="120"><circle cx="60" cy="60" r="52" fill="${relleno ? '#111827' : '#FFFFFF'}" stroke="#111827" stroke-width="7"/><text x="60" y="78" ${FONT} font-size="52" font-weight="bold" text-anchor="middle" fill="${relleno ? '#FFFFFF' : '#6B7280'}">${letra}</text></svg>`;
  await sharp(Buffer.from(s), { density: 288 }).resize(240, 240).png().toFile(file);
}

// ═══ 6. BLOQUES DE DOCUMENTO ═════════════════════════════════════════════════
const NIVELES = [                       // 0 normal · 1 compacto · 2 muy compacto · 3 permite cortar
  { enun: 60, opt: 20, fin: 170, img: 1.00, keep: true },
  { enun: 40, opt: 0,  fin: 110, img: .88,  keep: true },
  { enun: 30, opt: 0,  fin: 80,  img: .75,  keep: true },
  { enun: 30, opt: 0,  fin: 80,  img: .75,  keep: false }];
const run = (t, o = {}) => new TextRun({ text: t, font: F, size: SZ, ...o });
const para = (children, o = {}) => new Paragraph({ children, ...o });
const celda = (children, w, o = {}) => new TableCell({ children, width: { size: w, type: WidthType.DXA }, margins: { top: 40, bottom: 40, left: 90, right: 90 }, ...o });
const tabla = (rows, widths, o = {}) => new Table({ rows, columnWidths: widths, width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
  layout: TableLayoutType.FIXED, indent: { size: 0, type: WidthType.DXA }, margins: T0, ...o });

function makeHeader(escudo, cfg) {       // Encabezado real de Word, mismo ancho que el texto (TEXT_W)
  const wLogo = 1300, wCod = 2300, wTit = TEXT_W - wLogo - wCod, H = 300;
  const hPx = 76, wPx = Math.round(hPx * escudo.w / escudo.h);
  const cod = (children) => celda([para(children, { alignment: AlignmentType.CENTER })], wCod, { borders: bAll('000000', 6), verticalAlign: VerticalAlign.CENTER, margins: { top: 20, bottom: 20, left: 60, right: 60 } });
  const t = tabla([
    new TableRow({ height: { value: H, rule: HeightRule.ATLEAST }, children: [
      celda([para([new ImageRun({ type: 'png', data: escudo.buf, transformation: { width: wPx, height: hPx }, altText: { name: 'Escudo', title: 'Escudo institucional', description: 'Escudo del Colegio Integrado Simón Bolívar' } })], { alignment: AlignmentType.CENTER })],
        wLogo, { rowSpan: 4, borders: bAll('000000', 6), verticalAlign: VerticalAlign.CENTER, margins: { top: 30, bottom: 30, left: 40, right: 40 } }),
      celda(cfg.tituloMembrete.map(x => para([run(x, { bold: true, size: 22 })], { alignment: AlignmentType.CENTER, spacing: { before: 10, after: 10 } })),
        wTit, { rowSpan: 4, borders: bAll('000000', 6), verticalAlign: VerticalAlign.CENTER }),
      cod([run(cfg.codFormato, { bold: true, size: SZ_S })])] }),
    new TableRow({ height: { value: H, rule: HeightRule.ATLEAST }, children: [cod([run(cfg.codVersion, { size: SZ_XS })])] }),
    new TableRow({ height: { value: H, rule: HeightRule.ATLEAST }, children: [cod([run(cfg.codFecha, { size: SZ_XS })])] }),
    new TableRow({ height: { value: H, rule: HeightRule.ATLEAST }, children: [cod([run('Página ', { size: SZ_XS }), new TextRun({ children: [PageNumber.CURRENT], font: F, size: SZ_XS }), run(' de ', { size: SZ_XS }), new TextRun({ children: [PageNumber.TOTAL_PAGES], font: F, size: SZ_XS })])] })
  ], [wLogo, wTit, wCod]);
  return new Header({ children: [t, new Paragraph({ spacing: { before: 0, after: 0, line: 20, lineRule: 'exact' }, children: [] })] });
}

function tablaDatos(meta, docente) {     // SOLO en la primera página (cuerpo)
  const w = [3000, 3000, 2256, 2256];
  const c = (t, ww, span = 1) => celda([para([run(t, { bold: true, size: SZ_S })])], ww, { columnSpan: span, borders: bAll('808080', 4), shading: { type: ShadingType.CLEAR, fill: 'F2F2F2', color: 'auto' } });
  return tabla([
    new TableRow({ children: [c(`ASIGNATURA:  ${meta.asignatura || '________________'}`, w[0]), c(`DOCENTE:  ${meta.docente || '________________'}`, w[1]), c('PERÍODO:  ______', w[2]), c('VALORACIÓN:  ______', w[3])] }),
    new TableRow({ children: [c('ESTUDIANTE:  ______________________________________', w[0] + w[1], 2), c(`GRADO:  ${meta.grado}°`, w[2]), c('FECHA:  ___/___/_____', w[3])] }),
    new TableRow({ children: [c('TIEMPO:  ____________________', w[0] + w[1], 2), c('CURSO:  ________', w[2]), c('PUNTAJE:  ____ / ____', w[3])] })
  ], w);
}

const keepAll = (st, last) => ({ keepNext: !last && st.keep, keepLines: true });
function bloqueParte(titulo) {
  return para([run(titulo, { bold: true, size: SZ_S })], { shading: { type: ShadingType.CLEAR, fill: 'D9D9D9', color: 'auto' }, spacing: { before: 60, after: 100 }, keepNext: true, alignment: AlignmentType.CENTER, border: { top: line('000000', 6), bottom: line('000000', 6) } });
}
function bloqueVisual(p, st, imgs) {
  const v = p.visual, out = [];
  if (!v) return out;
  if (v.tipo === 1 || v.tipo === 3) {
    const L = p.enunciado.length, base = L <= 180 ? [216, 144] : L <= 420 ? [189, 126] : [162, 108];
    out.push(para([new ImageRun({ type: 'png', data: fs.readFileSync(imgs[p.num]), transformation: { width: Math.round(base[0] * st.img), height: Math.round(base[1] * st.img) }, altText: { name: 'img' + p.num, title: 'Imagen pregunta ' + p.num, description: v.leyenda || 'Imagen de apoyo' } })],
      { alignment: AlignmentType.CENTER, spacing: { before: 20, after: v.leyenda ? 10 : 50 }, keepNext: st.keep, keepLines: true }));
    if (v.leyenda) out.push(para([run(v.leyenda, { size: SZ_XS, italics: true, color: '555555' })], { alignment: AlignmentType.CENTER, spacing: { after: 50 }, keepNext: st.keep, keepLines: true }));
  } else if (v.tipo === 2) {
    const hs = v.headers, cw = Math.floor(COL_W / hs.length), ws = hs.map((_, i) => i === hs.length - 1 ? COL_W - cw * (hs.length - 1) : cw);
    if (v.titulo) out.push(para([run(v.titulo, { bold: true, italics: true, size: SZ_S })], { alignment: AlignmentType.CENTER, spacing: { after: 30 }, keepNext: true }));
    const cel = (t, w, h) => celda([para([run(t, { size: SZ_S, bold: h })], { alignment: AlignmentType.CENTER })], w, { borders: bAll(), margins: { top: 30, bottom: 30, left: 70, right: 70 }, ...(h ? { shading: { type: ShadingType.CLEAR, fill: 'F2F2F2', color: 'auto' } } : {}) });
    out.push(tabla([new TableRow({ cantSplit: true, children: hs.map((h, i) => cel(h, ws[i], true)) }), ...v.rows.map(r => new TableRow({ cantSplit: true, children: r.map((c, i) => cel(String(c), ws[i], false)) }))], ws, { alignment: AlignmentType.CENTER }));
    out.push(para([], { spacing: { before: 0, after: 20, line: 60, lineRule: 'exact' }, keepNext: true }));
  } else if (v.tipo === 4) {
    out.push(tabla([new TableRow({ cantSplit: true, children: [celda([
      ...(v.titulo ? [para([run(v.titulo, { bold: true, size: SZ_S })])] : []),
      para([run(v.texto, { size: SZ_S })], { alignment: AlignmentType.JUSTIFIED }),
      ...(v.fuente ? [para([run(v.fuente, { italics: true, size: SZ_XS, color: '666666' })])] : [])], COL_W,
      { borders: bAll(), shading: { type: ShadingType.CLEAR, fill: 'F5F5F5', color: 'auto' }, margins: { top: 60, bottom: 60, left: 120, right: 120 } })] })], [COL_W]));
    out.push(para([], { spacing: { before: 0, after: 20, line: 60, lineRule: 'exact' }, keepNext: true }));
  }
  return out;
}
function bloqueMC(p, st, imgs, docente) {
  const el = [];
  el.push(para([run(`${p.num}. `, { bold: true }), run(p.enunciado)], { alignment: AlignmentType.JUSTIFIED, spacing: { before: 0, after: st.enun }, keepNext: true, keepLines: true }));
  el.push(...bloqueVisual(p, st, imgs));
  el.push(para([run(p.pregunta, { bold: true })], { alignment: AlignmentType.JUSTIFIED, spacing: { before: 30, after: st.opt + 40 }, keepNext: true, keepLines: true }));
  const ops = p._ops, corto = ops.every(o => o.texto.length <= 16);
  const mark = (o) => docente && o.ok ? { bold: true, shading: { type: ShadingType.CLEAR, fill: 'CDEFD0', color: 'auto' } } : {};
  const optRuns = (o) => [run(`${o.letra})`, { bold: true, ...mark(o) }), new TextRun({ children: [new Tab(), o.texto + (docente && o.ok ? (corto ? '  ←' : '   ← CORRECTA') : '')], font: F, size: SZ, ...mark(o) })];
  if (corto) {
    for (let r = 0; r < 2; r++) {
      const a = ops[r * 2], b = ops[r * 2 + 1];
      el.push(para([...optRuns(a), new TextRun({ children: [new Tab()], font: F, size: SZ }), ...optRuns(b).map(x => x)], { tabStops: [{ type: TabStopType.LEFT, position: 360 }, { type: TabStopType.LEFT, position: 2448 }, { type: TabStopType.LEFT, position: 2448 + 360 }], spacing: { after: r === 1 ? st.fin : st.opt }, keepNext: r === 0 && st.keep, keepLines: true }));
    }
  } else {
    ops.forEach((o, i) => el.push(para(optRuns(o), { indent: { left: 360, hanging: 360 }, tabStops: [{ type: TabStopType.LEFT, position: 360 }], alignment: AlignmentType.LEFT, spacing: { after: i === 3 ? st.fin : st.opt }, keepNext: i < 3 && st.keep, keepLines: true })));
  }
  return el;
}
function bloqueAbierta(p, st, docente) {
  const el = [para([run(`${p.num}. `, { bold: true }), run(p.enunciado + (p.enunciado && p.pregunta ? ' ' : '')), run(p.pregunta || '', { bold: true })], { alignment: AlignmentType.JUSTIFIED, spacing: { after: 40 }, keepNext: true, keepLines: true })];
  if (p.visual && p.visual.tipo) el.push(...bloqueVisual(p, st, p._imgs || {}));
  const n = p.lineas || 4;
  if (docente) {
    el.push(tabla([new TableRow({ cantSplit: true, children: [celda([
      para([run('Respuesta modelo: ', { bold: true, size: SZ_S }), run(p.respuestaModelo || '', { size: SZ_S })], { alignment: AlignmentType.JUSTIFIED }),
      ...(p.criterios || []).map(c => para([run('• ' + c, { size: SZ_XS })])),
      para([run(`Valor: ${p.puntos || 1} punto(s)`, { size: SZ_XS, bold: true })])], COL_W, { borders: bAll('2E7D32', 6), shading: { type: ShadingType.CLEAR, fill: 'E8F5E9', color: 'auto' } })] })], [COL_W]));
    el.push(para([], { spacing: { after: st.fin } }));
  } else {
    const lb = { top: none, left: none, right: none, bottom: line('7F7F7F', 4) };
    el.push(tabla(Array.from({ length: n }, () => new TableRow({ cantSplit: true, height: { value: 400, rule: HeightRule.EXACT }, children: [celda([para([])], COL_W, { borders: lb, margins: T0 })] })), [COL_W]));
    el.push(para([], { spacing: { before: 0, after: st.fin - 40, line: 20, lineRule: 'exact' } }));
  }
  return el;
}

// Hoja de respuestas: UNA sola tabla (sin anidar), óvalos PNG grandes con la letra dentro
function hojaRespuestas(N, bubs, claves) {
  const cols = N <= 6 ? 2 : N <= 12 ? 3 : 4, rows = Math.ceil(N / cols);
  const bw = Math.min(620, Math.floor((Math.floor(TEXT_W / cols) - 460) / 4)), nw = 460, block = nw + 4 * bw, Wt = block * cols, PX = 31;
  const heavy = line('000000', 10), thin = line('9CA3AF', 4);
  const filas = [];
  for (let r = 0; r < rows; r++) {
    const cells = [];
    for (let b = 0; b < cols; b++) {
      const n = b * rows + r + 1, ok = n <= N, left = b === 0 ? heavy : thin, bt = r === 0 ? heavy : none, bb = r === rows - 1 ? heavy : none;
      cells.push(celda([para([run(ok ? String(n) : '', { bold: true, size: SZ_S })], { alignment: AlignmentType.CENTER, spacing: { before: 0, after: 0 } })], nw,
        { borders: { top: bt, bottom: bb, left, right: none }, shading: { type: ShadingType.CLEAR, fill: 'E5E7EB', color: 'auto' }, verticalAlign: VerticalAlign.CENTER, margins: T0 }));
      ['A', 'B', 'C', 'D'].forEach((L, k) => {
        const lleno = claves && ok && claves[n - 1] === L;
        cells.push(celda([para(ok ? [new ImageRun({ type: 'png', data: fs.readFileSync(bubs[lleno ? 'F' + L : L]), transformation: { width: PX, height: PX }, altText: { name: 'ovalo', title: 'Óvalo ' + L, description: 'Óvalo de respuesta ' + L } })] : [], { alignment: AlignmentType.CENTER, spacing: { before: 0, after: 0 } })], bw,
          { borders: { top: bt, bottom: bb, left: none, right: (b === cols - 1 && k === 3) ? heavy : none }, verticalAlign: VerticalAlign.CENTER, margins: T0 }));
      });
    }
    filas.push(new TableRow({ cantSplit: true, height: { value: 520, rule: HeightRule.ATLEAST }, children: cells }));
  }
  const widths = []; for (let b = 0; b < cols; b++) widths.push(nw, bw, bw, bw, bw);
  return tabla(filas, widths, { alignment: AlignmentType.CENTER });
}
function tablaIdentificacion() {
  const wl = 1300, w1 = TEXT_W - wl, w2 = Math.floor((TEXT_W - 2 * wl) / 2), w3 = TEXT_W - 2 * wl - w2, b = bAll('000000', 6);
  const et = (t, w) => celda([para([run(t, { bold: true, size: SZ_XS })], { keepNext: true })], w, { borders: b, shading: { type: ShadingType.CLEAR, fill: 'D9D9D9', color: 'auto' }, verticalAlign: VerticalAlign.CENTER, margins: { top: 20, bottom: 20, left: 80, right: 40 } });
  const ca = (w, span = 1) => celda([para([], { keepNext: true })], w, { borders: b, columnSpan: span, margins: { top: 20, bottom: 20, left: 60, right: 60 } });
  return tabla([
    new TableRow({ cantSplit: true, height: { value: 360, rule: HeightRule.EXACT }, children: [et('NOMBRE', wl), ca(w1, 3)] }),
    new TableRow({ cantSplit: true, height: { value: 360, rule: HeightRule.EXACT }, children: [et('FECHA', wl), ca(w2), et('CÓDIGO', wl), ca(w3)] })
  ], [wl, w2, wl, w3]);
}

// ═══ 7. MEDICIÓN, EMPAQUETADO Y VERIFICACIÓN DE MAQUETACIÓN ════════════════
const cp = require('child_process');
function aPDF(docx, outDir) {
  const wrapper = '/mnt/skills/public/docx/scripts/office/soffice.py', args = ['--headless', '--convert-to', 'pdf', '--outdir', outDir, docx];
  const pdf = path.join(outDir, path.basename(docx).replace(/\.docx$/i, '.pdf'));
  for (let intento = 1; intento <= 3; intento++) {
    try { if (fs.existsSync(pdf)) fs.unlinkSync(pdf); } catch (e) {}
    try { if (fs.existsSync(wrapper)) cp.execFileSync('python3', [wrapper, ...args], { stdio: 'ignore', timeout: 180000 }); else cp.execFileSync('soffice', args, { stdio: 'ignore', timeout: 180000 }); } catch (e) {}
    if (fs.existsSync(pdf) && fs.statSync(pdf).size > 500) return pdf;
    try { cp.execSync('pkill -x soffice.bin || true', { stdio: 'ignore' }); } catch (e) {}
    cp.execSync('sleep 2');
  }
  console.warn('⚠  LibreOffice no pudo convertir ' + path.basename(docx)); return null;
}
async function paginasPDF(pdf) {                      // → páginas en gris a 72 dpi (1 px = 1 pt)
  const pref = pdf.replace(/\.pdf$/i, '_pg'); cp.execFileSync('pdftoppm', ['-png', '-gray', '-r', '72', pdf, pref]);
  const dir = path.dirname(pdf), b = path.basename(pref), out = [];
  for (const f of fs.readdirSync(dir).filter(f => f.startsWith(b) && f.endsWith('.png')).sort()) {
    const fp = path.join(dir, f), { data, info } = await sharp(fp).raw().toBuffer({ resolveWithObject: true });
    out.push({ w: info.width, h: info.height, ch: info.channels, data }); fs.unlinkSync(fp);
  }
  return out;
}
function tinta(pg, x0, x1, y0, y1) {                  // primera y última fila con tinta en una región
  let first = -1, last = -1; y1 = Math.min(y1, pg.h);
  for (let y = Math.floor(y0); y < y1; y++) for (let x = Math.floor(x0); x < x1; x++) if (pg.data[(y * pg.w + x) * pg.ch] < 225) { if (first < 0) first = y; last = y; break; }
  return [first, last];
}
const estilosDoc = { default: { document: { run: { font: F, size: SZ } } } };
async function medirBloques(builders, outDir, tag, fin) {  // alto real (pt) de cada bloque en una columna de ancho COL_W
  const children = [];
  builders.forEach((b, i) => { children.push(new Paragraph({ children: [], pageBreakBefore: i > 0, spacing: { before: 0, after: 0, line: 20, lineRule: 'exact' } })); children.push(...b()); });
  const d = new Document({ styles: estilosDoc, compatibility: { version: 15 }, sections: [{ properties: { page: { size: { width: PAGE_W, height: 31000 }, margin: { top: 360, bottom: 360, left: MARG.left, right: PAGE_W - MARG.left - COL_W } } }, children }] });
  const f = path.join(outDir, `_medida_${tag}.docx`); fs.writeFileSync(f, await Packer.toBuffer(d));
  const pdf = aPDF(f, outDir); if (!pdf) return null;
  const pgs = await paginasPDF(pdf); if (pgs.length !== builders.length) return null;
  const hs = pgs.map(pg => { const [, l] = tinta(pg, 0, pg.w, 0, pg.h); return (l - 18) + 3 + fin / 20; });
  try { fs.unlinkSync(f); fs.unlinkSync(pdf); } catch (e) {} return hs;
}
function estimarAlto(builderInfo, fin) {              // respaldo si no hay LibreOffice: estimación por caracteres
  const lin = (n, ancho = COL_W / 20) => Math.ceil(n * 5.1 / ancho) * 11.8;
  return builderInfo.map(b => lin(b.chars) + b.extra + fin / 20);
}
function empaquetar(mc, ab, hP1, hP2, capFirst, capFull, SAFE = 8) {   // reordena ventanas de ≤6 preguntas para llenar columnas
  let k = 0, u = 0; const capOf = (j) => (j < 2 ? capFirst : capFull) - SAFE, res = { mc: [], ab: [] };
  const grupo = (items, hHead, out) => {
    if (!items.length) return;
    if (hHead) { if (capOf(k) - u < hHead + Math.min(...items.map(i => i.h))) { k++; u = 0; } u += hHead; }
    let rest = items.slice();
    while (rest.length) {
      const free = capOf(k) - u, win = rest.slice(0, 6); let best = null;
      for (let mask = 1; mask < (1 << win.length); mask++) {
        let sum = 0; const idx = []; for (let i = 0; i < win.length; i++) if (mask >> i & 1) { sum += win[i].h; idx.push(i); }
        if (sum > free) continue; if (win[0].h <= free && !(mask & 1)) continue;   // no se salta una pregunta que sí cabe
        const disp = idx.reduce((a, v) => a + v, 0) - idx.reduce((a, _, j) => a + j, 0), score = sum - 14 * disp;
        if (!best || score > best.score) best = { score, idx, sum };
      }
      if (!best) { if (u === 0) best = { idx: [0], sum: win[0].h, score: 0 }; else { k++; u = 0; continue; } }
      best.idx.forEach(i => out.push(win[i])); u += best.sum; rest = rest.filter((_, i) => !(i < win.length && best.idx.includes(i)));
    }
  };
  grupo(mc, hP1, res.mc); grupo(ab, hP2, res.ab); return res;
}
async function analizarLayout(pdf, claveUlt) {
  const pgs = await paginasPDF(pdf), norm = (t) => t.replace(/\s+/g, ' ');
  const textos = pgs.map((_, i) => norm(cp.execFileSync('pdftotext', ['-f', String(i + 1), '-l', String(i + 1), pdf, '-']).toString()));
  let ult = -1; textos.forEach((t, i) => { if (t.includes(claveUlt)) ult = i; });
  const hoja = textos.findIndex(t => t.includes('Marca un solo'));
  const top = 92, bot = PAGE_H / 20 - MARG.bottom / 20, lx0 = MARG.left / 20, lx1 = lx0 + COL_W / 20, rx0 = lx1 + COL_GAP / 20, rx1 = rx0 + COL_W / 20 + 4;
  const gaps = [];
  for (let p = 0; p < Math.max(ult, 0); p++) { const [, l] = tinta(pgs[p], lx0, lx1, top, bot + 10), [, r] = tinta(pgs[p], rx0, rx1, top, bot + 10); gaps.push({ pagina: p + 1, izq: Math.round(bot - l), der: Math.round(bot - r) }); }
  const maxGap = gaps.length ? Math.max(...gaps.flatMap(g => [g.izq, g.der])) : 0;
  return { paginas: pgs.length, paginaUltimaPregunta: ult + 1, paginaHoja: hoja + 1, hojaMismaPagina: hoja === ult, gaps, maxGap };
}

// ═══ 8. PROGRAMA PRINCIPAL ═══════════════════════════════════════════════════
// Clasificación pedagógica de cada pregunta (competencia · componente · nivel · DBA)
function metaPregunta(p) {
  const t = [p.competencia && `Competencia: ${p.competencia}`, p.componente && `Componente: ${p.componente}`, p.nivel && `Nivel ${p.nivel}`, p.dba && `DBA/evidencia: ${p.dba}`].filter(Boolean);
  return t.length ? [para([run(t.join('  ·  '), { size: SZ_XS, italics: true, color: '1F4E79' })], { spacing: { after: 30 } })] : [];
}
// Matriz de especificaciones: competencias × niveles de desempeño (solo documento docente)
function matrizEspecificaciones(mc) {
  if (!mc.some(p => p.competencia)) return [];
  const grupos = new Map();
  mc.forEach(p => {
    const c = p.competencia || 'Sin clasificar';
    if (!grupos.has(c)) grupos.set(c, { comps: new Set(), nums: [], niv: { 1: 0, 2: 0, 3: 0, 4: 0, s: 0 } });
    const g = grupos.get(c); if (p.componente) g.comps.add(p.componente); g.nums.push(p.num); g.niv[p.nivel ? Number(p.nivel) : 's']++;
  });
  const wN = 600, wT = 700, wP = 2000, wComp = 2400, wC = TEXT_W - wComp - wP - wT - 4 * wN;
  const h = (t, w) => celda([para([run(t, { bold: true, size: SZ_XS })], { alignment: AlignmentType.CENTER })], w, { borders: bAll('000000', 6), shading: { type: ShadingType.CLEAR, fill: 'D9D9D9', color: 'auto' }, verticalAlign: VerticalAlign.CENTER });
  const c = (t, w, o = {}) => celda([para([run(String(t), { size: SZ_XS, ...(o.bold ? { bold: true } : {}) })], { alignment: o.left ? AlignmentType.LEFT : AlignmentType.CENTER })], w, { borders: bAll(), verticalAlign: VerticalAlign.CENTER });
  const tot = { 1: 0, 2: 0, 3: 0, 4: 0 };
  const filas = [new TableRow({ tableHeader: true, children: [h('Competencia', wC), h('Componente(s)', wComp), h('Preguntas', wP), h('Total', wT), h('N1', wN), h('N2', wN), h('N3', wN), h('N4', wN)] })];
  grupos.forEach((g, nombre) => {
    [1, 2, 3, 4].forEach(k => tot[k] += g.niv[k]);
    filas.push(new TableRow({ cantSplit: true, children: [c(nombre, wC, { left: true, bold: true }), c([...g.comps].join('; ') || '—', wComp, { left: true }), c(g.nums.join(', '), wP), c(g.nums.length, wT, { bold: true }), ...[1, 2, 3, 4].map(k => c(g.niv[k] || '·', wN))] }));
  });
  filas.push(new TableRow({ cantSplit: true, children: [celda([para([run('TOTAL', { size: SZ_XS, bold: true })], { alignment: AlignmentType.CENTER })], wC + wComp + wP, { columnSpan: 3, borders: bAll(), verticalAlign: VerticalAlign.CENTER }), c(mc.length, wT, { bold: true }), ...[1, 2, 3, 4].map(k => c(tot[k], wN, { bold: true }))] }));
  return [
    para([run('MATRIZ DE ESPECIFICACIONES', { bold: true })], { alignment: AlignmentType.CENTER, spacing: { before: 240, after: 60 }, keepNext: true }),
    para([run('N1–N4 = nivel de desempeño SABER de cada pregunta de selección múltiple. Úsala para comprobar que la prueba cubre todas las competencias y no concentra la dificultad en un solo nivel.', { size: SZ_XS, italics: true })], { alignment: AlignmentType.JUSTIFIED, spacing: { after: 80 }, keepNext: true }),
    tabla(filas, [wC, wComp, wP, wT, wN, wN, wN, wN])
  ];
}
async function armar(ctx, nivel, hojaPropia) {
  const { ex, imgs, bubs, escudo, outDir } = ctx, st = NIVELES[nivel], m = ex.meta;
  const mc = ctx.mc.slice(), ab = ctx.ab.slice();
  mc.forEach(p => p._ops = [p.correcta, ...p.distractores].map((t, i) => ({ letra: 'ABCD'[i], texto: String(t), ok: i === 0, error: '' })));
  const hdrCfg = Object.assign({ tituloMembrete: ['GESTION ACADEMICA', 'SEGUIMIENTO ACADEMICO', 'EVALUACION DE COMPRENSION'], codFormato: 'GA-F16', codVersion: 'Versión 1.0', codFecha: '04-04-2016' }, ex.config.membrete || {});
  const TITULO = (m.titulo || `EVALUACIÓN ESCRITA — ${(m.area || '').toUpperCase()}: ${(m.tema || '').toUpperCase()} — GRADO ${m.grado}°`) + (ctx.etiquetaVersion ? ` — VERSIÓN ${ctx.etiquetaVersion}` : '');
  const sizeOf = (a) => ({ page: { size: { width: PAGE_W, height: PAGE_H }, margin: MARG }, ...a });
  const instr = 'INSTRUCCIONES: Lee con atención cada pregunta. ' + (ab.length ? 'En la PARTE I selecciona la única respuesta correcta y márcala en la HOJA DE RESPUESTAS rellenando por completo el óvalo de la letra (A, B, C o D). En la PARTE II responde con tus propias palabras en las líneas dispuestas. ' : 'Selecciona la única respuesta correcta y márcala en la HOJA DE RESPUESTAS rellenando por completo el óvalo de la letra (A, B, C o D). ') + 'No se admiten tachones ni enmendaduras.';
  const tituloP = (t) => para([run(t, { bold: true, size: 24 })], { alignment: AlignmentType.CENTER, spacing: { before: 140, after: 80 } });
  const sec1 = (extra) => [tablaDatos(m), tituloP(TITULO + (extra ? ' — CLAVES DOCENTE' : '')), para([run(instr, { italics: true, size: SZ_S })], { alignment: AlignmentType.JUSTIFIED, spacing: { after: 100 } })];
  const mkDoc = (secciones) => new Document({ creator: 'Generador de Examen Tipo SABER', title: TITULO, styles: estilosDoc, compatibility: { version: 15 }, sections: secciones });

  // ── medir y empaquetar (solo si hay LibreOffice; si no, se conserva el orden) ──
  let medido = false, movidas = 0;
  if (ctx.empaquetar) {
    mc.forEach((p, i) => p.num = i + 1); ab.forEach((p, i) => p.num = mc.length + i + 1);
    const bl = [...mc.map(p => () => bloqueMC(p, st, imgs, false)), () => [bloqueParte('PARTE I')], ...(ab.length ? [() => [bloqueParte('PARTE II')], ...ab.map(p => () => bloqueAbierta(p, st, false))] : [])];
    const hs = await medirBloques(bl, outDir, 'b' + nivel, st.fin);
    const hdoc = new Document({ styles: estilosDoc, compatibility: { version: 15 }, sections: [{ properties: sizeOf({}), headers: { default: makeHeader(escudo, hdrCfg) }, children: sec1(false) }] });
    const fS = path.join(outDir, `_sec1_${nivel}.docx`); fs.writeFileSync(fS, await Packer.toBuffer(hdoc)); const pS = aPDF(fS, outDir);
    if (hs && pS) {
      const pg = (await paginasPDF(pS))[0], [, l] = tinta(pg, 0, pg.w, 92, pg.h), capFirst = (PAGE_H / 20 - MARG.bottom / 20) - (l + 4 + 5), capFull = (PAGE_H - MARG.top - MARG.bottom) / 20;
      const itM = mc.map((p, i) => ({ p, h: hs[i] })), hP1 = hs[mc.length], hP2 = ab.length ? hs[mc.length + 1] : 0, itA = ab.map((p, i) => ({ p, h: hs[mc.length + 2 + i] }));
      const r = empaquetar(itM, itA, hP1, hP2, capFirst, capFull);
      r.mc.forEach((x, i) => { if (x.p !== mc[i]) movidas++; mc[i] = x.p; }); r.ab.forEach((x, i) => { ab[i] = x.p; });
      medido = true; try { fs.unlinkSync(fS); fs.unlinkSync(pS); } catch (e) {}
    } else console.warn('⚠  No se pudo medir con LibreOffice; se mantiene el orden original.');
  }
  mc.forEach((p, i) => p.num = i + 1); ab.forEach((p, i) => p.num = mc.length + i + 1);

  const seed = ([...(m.tema || 'examen')].reduce((a, c) => a * 31 + c.charCodeAt(0) >>> 0, 7) + (ctx.seedExtra || 0)) >>> 0;
  const dist = asignarLetras(mc, seed), claves = mc.map(p => p._clave);
  const pc = (l) => Math.round(100 * dist[l] / mc.length);
  const FRASES = { matematicas: '¡Cada problema tiene solución! Confía en ti.', ciencias: '¡El conocimiento es tu superpoder! Da lo mejor.', lenguaje: '¡Las palabras son tu fuerza! Lee, piensa y triunfa.', sociales: '¡Conocer el mundo comienza contigo! Éxito.', tecnologia: '¡La tecnología está en tus manos! Demuéstralo.', general: '¡Tú puedes! Respira, enfócate y demuestra todo lo que sabes.' };
  const frase = FRASES[(m.areaClave || 'general')] || FRASES.general;
  const cuerpo = (docente) => { const els = [bloqueParte(`PARTE I — SELECCIÓN MÚLTIPLE CON ÚNICA RESPUESTA (${mc.length})`)]; mc.forEach(p => els.push(...bloqueMC(p, st, imgs, docente))); if (ab.length) { els.push(bloqueParte(`PARTE II — PREGUNTAS ABIERTAS (${ab.length})`)); ab.forEach(p => els.push(...bloqueAbierta(p, st, docente))); } return els; };
  const cols = { column: { count: 2, space: COL_GAP, separate: true, equalWidth: true } };
  const sec = (children, extra = {}) => ({ properties: sizeOf({ type: SectionType.CONTINUOUS, ...extra }), headers: { default: makeHeader(escudo, hdrCfg) }, children });
  const tipoHoja = hojaPropia ? { type: SectionType.NEXT_PAGE } : {};

  const dEst = mkDoc([
    { properties: sizeOf({}), headers: { default: makeHeader(escudo, hdrCfg) }, children: sec1(false) },
    sec(cuerpo(false), cols),
    sec([para([], { spacing: { before: 120, after: 0 }, border: { bottom: line('000000', 6) }, keepNext: true }),
      para([run('HOJA DE RESPUESTAS', { bold: true })], { alignment: AlignmentType.CENTER, spacing: { before: 80, after: 60 }, keepNext: true }),
      tablaIdentificacion(),
      para([run('Rellena por completo el óvalo de la letra elegida (●). Marca un solo óvalo por pregunta.' + (ab.length ? ' Las preguntas abiertas se responden en la hoja de preguntas.' : ''), { size: SZ_XS, italics: true })], { alignment: AlignmentType.CENTER, spacing: { before: 70, after: 70 }, keepNext: true }),
      hojaRespuestas(mc.length, bubs, null),
      para([run(frase, { italics: true, color: '444444', size: SZ_S })], { alignment: AlignmentType.CENTER, spacing: { before: 100, after: 0 } })], tipoHoja)
  ]);
  const colW2 = TEXT_W - 1650, hc = (t, w) => celda([para([run(t, { bold: true, size: SZ_S })], { alignment: AlignmentType.CENTER })], w, { borders: bAll('000000', 6), shading: { type: ShadingType.CLEAR, fill: 'D9D9D9', color: 'auto' } });
  const filasClave = [
    new TableRow({ tableHeader: true, children: [hc('N.º', 700), hc('Clave', 950), hc('Justificación (por qué es correcta) y error que representa cada distractor', colW2)] }),
    ...mc.map(p => new TableRow({ cantSplit: true, children: [
      celda([para([run(String(p.num), { bold: true, size: SZ_S })], { alignment: AlignmentType.CENTER })], 700, { borders: bAll(), verticalAlign: VerticalAlign.CENTER }),
      celda([para([run(p._clave, { bold: true, size: 24 })], { alignment: AlignmentType.CENTER })], 950, { borders: bAll(), verticalAlign: VerticalAlign.CENTER, shading: { type: ShadingType.CLEAR, fill: 'CDEFD0', color: 'auto' } }),
      celda([...metaPregunta(p), para([run('Correcta: ', { size: SZ_S, bold: true }), run(p.justificacion || '', { size: SZ_S })]), ...p._ops.filter(o => !o.ok).map(o => para([run(`${o.letra}) `, { bold: true, size: SZ_XS }), run(o.error || '', { size: SZ_XS, color: '444444' })]))], colW2, { borders: bAll() })] })),
    ...ab.map(p => new TableRow({ cantSplit: true, children: [
      celda([para([run(String(p.num), { bold: true, size: SZ_S })], { alignment: AlignmentType.CENTER })], 700, { borders: bAll(), verticalAlign: VerticalAlign.CENTER }),
      celda([para([run('ABIERTA', { bold: true, size: SZ_XS })], { alignment: AlignmentType.CENTER })], 950, { borders: bAll(), verticalAlign: VerticalAlign.CENTER, shading: { type: ShadingType.CLEAR, fill: 'FFF3CD', color: 'auto' } }),
      celda([para([run('Respuesta modelo: ', { bold: true, size: SZ_S }), run(p.respuestaModelo || '', { size: SZ_S })]), ...(p.criterios || []).map(c => para([run('• ' + c, { size: SZ_XS })])), para([run(`Valor: ${p.puntos || 1} punto(s)`, { size: SZ_XS, bold: true })])], colW2, { borders: bAll() })] }))];
  const dDoc = mkDoc([
    { properties: sizeOf({}), headers: { default: makeHeader(escudo, hdrCfg) }, children: [...sec1(true).slice(0, 2), para([run('USO EXCLUSIVO DEL DOCENTE. ', { bold: true, size: SZ_S, color: 'B91C1C' }), run(`Respuesta correcta resaltada en verde, plantilla de calificación y justificación pedagógica. Claves: ${claves.join(' ')}`, { italics: true, size: SZ_S })], { alignment: AlignmentType.JUSTIFIED, spacing: { after: 100 } })] },
    sec(cuerpo(true), cols),
    sec([para([], { spacing: { before: 120, after: 0 }, border: { bottom: line('000000', 6) }, keepNext: true }),
      para([run('PLANTILLA DE CALIFICACIÓN (CLAVE) — óvalos llenos = respuesta correcta', { bold: true })], { alignment: AlignmentType.CENTER, spacing: { before: 80, after: 80 }, keepNext: true }),
      hojaRespuestas(mc.length, bubs, claves),
      para([run('CLAVE DE RESPUESTAS Y JUSTIFICACIÓN', { bold: true })], { alignment: AlignmentType.CENTER, spacing: { before: 200, after: 80 }, keepNext: true }),
      tabla(filasClave, [700, 950, colW2]),
      ...matrizEspecificaciones(mc)], tipoHoja)
  ]);
  const base = (m.archivo || `Examen_${(m.area || 'Area')}_${(m.tema || 'Tema')}_Grado${m.grado}`).normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^A-Za-z0-9_\-]+/g, '_').replace(/_+/g, '_');
  const sfx = ctx.sufijo || '';
  const fE = path.join(outDir, base + sfx + '_ESTUDIANTE.docx'), fD = path.join(outDir, base + sfx + '_DOCENTE_CLAVES.docx');
  fs.writeFileSync(fE, await Packer.toBuffer(dEst)); fs.writeFileSync(fD, await Packer.toBuffer(dDoc));
  const ult = mc.length && !ab.length ? mc[mc.length - 1].enunciado : (ab.length ? ab[ab.length - 1].enunciado : '');
  return { fE, fD, claves, dist, pc, movidas, medido, claveUlt: ult.replace(/\s+/g, ' ').split(' ').slice(0, 4).join(' ') };
}

(async () => {
  const args = process.argv.slice(2), jsonPath = args.find(a => a.endsWith('.json'));
  if (!jsonPath) { console.error('Uso: node build_exam.js exam.json [--nivel 0-3] [--out carpeta] [--sin-ajuste]'); process.exit(2); }
  const nivelFijo = args.includes('--nivel') ? Number(args[args.indexOf('--nivel') + 1]) : null;
  const outDir = args.includes('--out') ? args[args.indexOf('--out') + 1] : './salida';
  fs.mkdirSync(path.join(outDir, 'img'), { recursive: true });
  const ex = JSON.parse(fs.readFileSync(jsonPath, 'utf8')); ex.config = ex.config || {}; ex.meta = ex.meta || {};

  const { E, W } = validar(ex); W.forEach(w => console.warn('⚠  ' + w));
  if (E.length) { console.error('\n❌ LA PRUEBA NO PASA LA VALIDACIÓN. Corrige y vuelve a ejecutar:\n' + E.map(e => '   • ' + e).join('\n')); process.exit(1); }
  console.log('✅ Validación OK: opciones únicas, respuesta presente, longitudes en rango.');

  const mc = ex.preguntas.filter(p => (p.tipo || 'multiple') === 'multiple'), ab = ex.preguntas.filter(p => p.tipo === 'abierta');
  mc.forEach((p, i) => p.num = i + 1); ab.forEach((p, i) => p.num = mc.length + i + 1);
  const imgs = {}; let nImg = 0;
  for (const p of ex.preguntas) {                      // PNG de alta resolución, una vez por pregunta
    const v = p.visual; if (!v || !(v.tipo === 1 || v.tipo === 3)) continue;
    const plantilla = PLANTILLAS[v.plantilla]; if (!plantilla) { console.error(`❌ P${p.num}: plantilla "${v.plantilla}" no existe. Disponibles: ${Object.keys(PLANTILLAS).join(', ')}`); process.exit(1); }
    const f = path.join(outDir, 'img', `pregunta_${String(p.num).padStart(2, '0')}.png`); await renderPNG(plantilla(v.params || {}), f); imgs[p.num] = f; nImg++;
  }
  // La imagen se asocia al OBJETO pregunta (no al número), así sobrevive al reordenamiento
  const imgPorPregunta = new Map(); for (const p of ex.preguntas) if (imgs[p.num]) imgPorPregunta.set(p, imgs[p.num]);
  const imgsProxy = new Proxy({}, { get: (_, n) => { for (const [p, f] of imgPorPregunta) if (String(p.num) === String(n)) return f; } });
  const bubs = {}; for (const L of 'ABCD') { bubs[L] = path.join(outDir, 'img', `ovalo_${L}.png`); bubs['F' + L] = path.join(outDir, 'img', `ovalo_${L}_lleno.png`); await bubble(L, false, bubs[L]); await bubble(L, true, bubs['F' + L]); }
  const escudo = await prepararEscudo();

  const ctx = { ex, mc, ab, imgs: imgsProxy, bubs, escudo, outDir, empaquetar: !args.includes('--sin-ajuste') };
  const niveles = nivelFijo !== null ? [nivelFijo] : [0, 1, 2];
  const optimizar = async (cx) => {
    let mejor = null;
  for (const nv of niveles) {
    for (const hojaPropia of [false, true]) {
      if (hojaPropia && mejor && mejor.an && mejor.an.hojaMismaPagina) continue;
      console.log(`🧩 Maquetando (nivel de compactación ${nv}${hojaPropia ? ', hoja de respuestas en página propia' : ''})…`);
      const r = await armar(cx, nv, hojaPropia); const pdf = aPDF(r.fE, outDir);
      r.nivel = nv; r.hojaPropia = hojaPropia;
      if (pdf) { r.an = await analizarLayout(pdf, r.claveUlt); try { fs.unlinkSync(pdf); } catch (e) {} console.log(`   páginas: ${r.an.paginas} · hueco máx. al pie de columna: ${r.an.maxGap} pt · hoja en la misma página que las preguntas: ${r.an.hojaMismaPagina ? 'sí' : 'no'}`); }
      const score = r.an ? r.an.maxGap + (r.an.hojaMismaPagina || hojaPropia ? 0 : 200) : 999;
      r.score = score; if (!mejor || score < mejor.score) mejor = r;
      if (r.an && r.an.maxGap <= 55) break;
    }
    if (mejor && mejor.an && mejor.an.maxGap <= 55) break;
  }
    await armar(cx, mejor.nivel, mejor.hojaPropia);   // deja en disco la mejor versión
    return mejor;
  };
  const versiones = args.includes('--versiones') ? Number(args[args.indexOf('--versiones') + 1]) : Number(ex.config.versiones || 1);
  if (versiones >= 2) { ctx.sufijo = '_VersionA'; ctx.etiquetaVersion = 'A'; }
  const mejor = await optimizar(ctx);
  let mejorB = null;
  if (versiones >= 2) {                                // VERSIÓN B: otro orden de preguntas y otras letras
    const rnd = mulberry(20260508), barajar = (a) => { const x = a.slice(); for (let q = x.length - 1; q > 0; q--) { const r = Math.floor(rnd() * (q + 1)); [x[q], x[r]] = [x[r], x[q]]; } return x; };
    console.log('🔀 Armando la VERSIÓN B (otro orden de preguntas y otras letras)…');
    mejorB = await optimizar({ ...ctx, mc: barajar(ctx.mc), ab: barajar(ctx.ab), sufijo: '_VersionB', etiquetaVersion: 'B', seedExtra: 7919 });
    const igual = mejor.claves.filter((l, q) => l === mejorB.claves[q]).length;
    console.log(`🎯 Versión B claves: ${mejorB.claves.join('')} (misma letra que A en la misma posición: ${igual} de ${mc.length})`);
  }
  console.log(`🎯 Claves: A:${mejor.pc('A')}% B:${mejor.pc('B')}% C:${mejor.pc('C')}% D:${mejor.pc('D')}%  → ${mejor.claves.join('')}`);
  if (mc.length >= 8 && Math.max(...Object.values(mejor.dist)) / mc.length > .4) console.warn('⚠  Distribución de claves desbalanceada');
  const resumen = { estudiante: mejor.fE, docente: mejor.fD, preguntasMC: mc.length, abiertas: ab.length, imagenes: nImg, claves: mejor.claves.join(''), distribucion: mejor.dist, escudo: escudo.origen, nivelCompactacion: mejor.nivel, hojaEnPaginaPropia: mejor.hojaPropia, preguntasReubicadas: mejor.movidas, layout: mejor.an || 'sin verificar (LibreOffice no disponible)' };
  if (mejorB) resumen.versionB = { estudiante: mejorB.fE, docente: mejorB.fD, claves: mejorB.claves.join(''), distribucion: mejorB.dist, layout: mejorB.an || 'sin verificar' };
  fs.writeFileSync(path.join(outDir, 'resumen.json'), JSON.stringify(resumen, null, 2));
  console.log('📄 ' + mejor.fE + '\n📄 ' + mejor.fD + (mejorB ? '\n📄 ' + mejorB.fE + '\n📄 ' + mejorB.fD : '') + '\n' + JSON.stringify(resumen));
  if (mejor.an && mejor.an.maxGap > 55) console.warn(`⚠  Queda un hueco de ${mejor.an.maxGap} pt al pie de una columna; revisa y, si es necesario, ejecuta con --nivel 3.`);
})().catch(e => { console.error('❌ ' + (e.stack || e)); process.exit(1); });
