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
        TabStopType, TableLayoutType, HeightRule, HorizontalPositionRelativeFrom, HorizontalPositionAlign, VerticalPositionRelativeFrom, TextWrappingType, TextWrappingSide } = D;

// ═══ 1. ESCUDO INSTITUCIONAL ════════════════════════════════════════════════
// Escudo 2025 optimizado para encabezado (PNG 240×290, 64 colores) incrustado en base64.
// Es el ÚLTIMO recurso: solo se usa si no se encuentra el archivo en disco.
const ESCUDO_B64 = "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";

// ═══ 1b. CONFIGURACIÓN DE LA INSTITUCIÓN (config/institucion.json) ═══════════
// Se busca (en orden): --institucion <archivo>, variable INSTITUCION_JSON, ../config/, raíz del repositorio, ./config/
let INST = {}, INST_DIR = null;
function cargarInstitucion(args) {
  const i = args.indexOf('--institucion');
  const cand = [i >= 0 ? args[i + 1] : null, process.env.INSTITUCION_JSON,
    path.join(__dirname, '..', 'config', 'institucion.json'), path.join(__dirname, '..', '..', '..', '..', 'config', 'institucion.json'),
    path.join(process.cwd(), 'config', 'institucion.json')];
  for (const f of cand) {
    try { if (f && fs.existsSync(f)) { INST = JSON.parse(fs.readFileSync(f, 'utf8')); INST_DIR = path.dirname(path.resolve(f)); console.log(`🏫 Institución: ${INST.nombre || '(sin nombre)'}  ← ${f}`); return; } }
    catch (e) { console.warn(`⚠  No se pudo leer ${f}: ${e.message}`); }
  }
  console.warn('⚠  No se encontró config/institucion.json: se usan los datos por defecto del script.');
}
// ═══ 1c. CATÁLOGO DE ÁREAS (areas/<areaClave>.json): competencias, componentes, niveles y frase final ═══
let AREA = null;
function cargarArea(clave, grado) {
  if (!clave) return;
  const cand = [process.env.AREAS_DIR && path.join(process.env.AREAS_DIR, clave + '.json'), path.join(__dirname, '..', 'areas', clave + '.json')];
  for (const f of cand) {
    try { if (f && fs.existsSync(f)) {
      AREA = JSON.parse(fs.readFileSync(f, 'utf8'));
      if (AREA.marcos && AREA.marcos.length) {            // varios marcos (p. ej. Saber 11.° y Saber 3.°-5.°-9.°): se elige por grado
        const g = Number(grado), mk = AREA.marcos.find(m => (m.grados || []).includes(g)) || AREA.marcos[0];
        AREA._marco = mk; AREA.niveles_cortos = mk.niveles_cortos || AREA.niveles_cortos;
        AREA.competencias = AREA.marcos.flatMap(m => m.competencias || []); AREA.componentes = AREA.marcos.flatMap(m => m.componentes || []);
      }
      console.log(`📚 Área: ${AREA.nombre}${AREA._marco ? ' · marco: ' + AREA._marco.nombre : ''}  ← ${f}`); return;
    } }
    catch (e) { console.warn(`⚠  No se pudo leer ${f}: ${e.message}`); }
  }
  if (clave !== 'general') console.warn(`⚠  No hay catálogo para el área "${clave}" (areas/${clave}.json): no se validarán competencias ni componentes.`);
}
const nz = (t) => String(t).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
const escudoDeConfig = () => (INST.escudo && INST.escudo.encabezado && INST_DIR) ? path.resolve(INST_DIR, INST.escudo.encabezado) : null;

function localizarEscudo() {
  if (process.env.ESCUDO_USAR_BASE64 === '1' && ESCUDO_B64 && !ESCUDO_B64.startsWith('__')) return { buf: Buffer.from(ESCUDO_B64, 'base64'), origen: 'base64 incrustado (forzado)' };
  const dirs = [process.env.ESCUDO_PATH, escudoDeConfig(),
    path.join(__dirname, '..', 'assets', 'escudo_header.png'),
    path.join(__dirname, 'assets', 'escudo_header.png'),
    path.join(process.cwd(), 'assets', 'escudo_header.png'),
    path.join(process.cwd(), 'escudo_header.png'),
    '/mnt/skills/user/generador-examen-tipo-saber/assets/escudo_header.png'];
  try { for (const f of fs.readdirSync('/mnt/user-data/uploads')) if (/escudo.*\.(png|jpe?g|webp)$/i.test(f)) dirs.push('/mnt/user-data/uploads/' + f); } catch (e) {}
  for (const p of dirs) { try { if (p && fs.existsSync(p)) return { buf: fs.readFileSync(p), origen: p }; } catch (e) {} }
  if (ESCUDO_B64 && !ESCUDO_B64.startsWith('__')) return { buf: Buffer.from(ESCUDO_B64, 'base64'), origen: 'base64 incrustado' };
  return null;
}
async function prepararEscudo() {
  const f = localizarEscudo();
  if (!f) throw new Error('ESCUDO NO ENCONTRADO: ni archivo ni base64. Abortando para no entregar un examen sin logo.');
  // Acepta PNG, JPEG o WebP aunque la extensión no coincida (p. ej. un .jpg renombrado a .png): sharp decide por el contenido real.
  let fmt; try { fmt = (await sharp(f.buf).metadata()).format; } catch (e) { fmt = null; }
  if (!['png', 'jpeg', 'webp'].includes(fmt)) throw new Error('El escudo (' + f.origen + ') no es una imagen PNG, JPG o WebP válida.');
  // Normaliza: recorta márgenes blancos, fondo blanco, ancho máx. 300 px, PNG con paleta (liviano, buena calidad de impresión)
  const buf = await sharp(f.buf).flatten({ background: '#ffffff' }).trim({ threshold: 12 }).resize({ width: 300, withoutEnlargement: true }).png({ palette: true, quality: 95, colours: 128 }).toBuffer();
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
    if (!p.pregunta || (!p.enunciado && !p.estimulo)) E.push(`${id}: falta ${p.estimulo ? 'la pregunta' : 'enunciado o pregunta'}`);
    if (rango && !p.estimulo && p.enunciado && (p.enunciado.length < rango[0] || p.enunciado.length > rango[1]))
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
  // Textos compartidos (estímulos): un texto, tabla o gráfico para varias preguntas
  const ests = ex.estimulos || [], idsEst = new Set();
  ests.forEach((e, k) => {
    const id = `Estímulo ${k + 1}`;
    if (!e.id) E.push(`${id}: falta "id"`); else if (idsEst.has(e.id)) E.push(`${id}: id repetido "${e.id}"`); idsEst.add(e.id);
    if (!e.texto && !e.visual) E.push(`${id}: requiere "texto" o "visual"`);
    if (e.texto && e.texto.length > 3300) W.push(`${id}: texto muy largo (${e.texto.length} caracteres); la prueba SABER usa textos de hasta ~500 palabras`);
  });
  mc.forEach((p, k) => { if (p.estimulo && !idsEst.has(p.estimulo)) E.push(`P${k + 1}: el estímulo "${p.estimulo}" no existe en "estimulos"`); });
  ests.forEach(e => { const n = mc.filter(p => p.estimulo === e.id).length; if (e.id && !n) W.push(`El estímulo "${e.id}" no tiene preguntas`); else if (n > 6) W.push(`El estímulo "${e.id}" tiene ${n} preguntas (recomendado: hasta 6)`); });
  if (AREA) {
    const comps = (AREA.competencias || []).map(c => nz(c.nombre)), cmps = (AREA.componentes || []).map(c => nz(typeof c === 'string' ? c : c.nombre));
    mc.forEach((p, k) => {
      if (p.competencia && comps.length && !comps.includes(nz(p.competencia))) W.push(`P${k + 1}: la competencia "${p.competencia}" no está en el catálogo de ${AREA.nombre}`);
      if (p.componente && cmps.length && !cmps.includes(nz(p.componente))) W.push(`P${k + 1}: el componente "${p.componente}" no está en el catálogo de ${AREA.nombre}`);
    });
  }
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

// ── Ayudas para las plantillas nuevas ───────────────────────────────────────
const PAL = ['#3B82F6', '#F59E0B', '#10B981', '#EF4444', '#8B5CF6', '#EC4899', '#14B8A6', '#F97316'];
function envolver(t, max) {                           // parte un texto en líneas de ≤ max caracteres
  const out = []; let l = '';
  for (const x of String(t).split(/\s+/)) { if (l && (l + ' ' + x).length > max) { out.push(l); l = x; } else l = (l ? l + ' ' : '') + x; }
  if (l) out.push(l); return out;
}
const txtH = (x, y, t, size = 30, o = {}) => `<text x="${x}" y="${y}" ${FONT} font-size="${size}" font-weight="${o.w || 'bold'}" text-anchor="${o.a || 'middle'}" fill="${o.c || '#1F2937'}" stroke="#fff" stroke-width="8" paint-order="stroke" stroke-linejoin="round">${esc(t)}</text>`;
function flecha(x1, y1, x2, y2, c = '#111827', w = 5) {   // línea con punta de flecha en (x2,y2)
  const a = Math.atan2(y2 - y1, x2 - x1), L = 24, bx = x2 - L * 0.8 * Math.cos(a), by = y2 - L * 0.8 * Math.sin(a);
  const p1 = [x2 - L * Math.cos(a - .42), y2 - L * Math.sin(a - .42)], p2 = [x2 - L * Math.cos(a + .42), y2 - L * Math.sin(a + .42)];
  return `<line x1="${x1.toFixed(1)}" y1="${y1.toFixed(1)}" x2="${bx.toFixed(1)}" y2="${by.toFixed(1)}" stroke="${c}" stroke-width="${w}" stroke-linecap="round"/><path d="M${x2.toFixed(1)},${y2.toFixed(1)} L${p1[0].toFixed(1)},${p1[1].toFixed(1)} L${p2[0].toFixed(1)},${p2[1].toFixed(1)} Z" fill="${c}"/>`;
}
const ejesXY = (x0, x1, y0, y1, max, paso, min = 0) => {   // cuadrícula horizontal con valores para gráficos de barras/líneas
  let s = '';
  for (let v = min; v <= max + 1e-9; v += paso) { const y = y0 - (y0 - y1) * (v - min) / (max - min); s += `<line x1="${x0}" y1="${y.toFixed(1)}" x2="${x1}" y2="${y.toFixed(1)}" stroke="${v === min ? '#111827' : '#D1D5DB'}" stroke-width="${v === min ? 5 : 2}"/>` + txt(x0 - 14, y + 9, String(Math.round(v * 100) / 100), 25, { a: 'end', w: 'normal' }); }
  return s + `<line x1="${x0}" y1="${y1 - 10}" x2="${x0}" y2="${y0}" stroke="#111827" stroke-width="5"/>`;
};
const pasoAuto = (r) => r <= 10 ? 1 : r <= 20 ? 2 : r <= 50 ? 5 : r <= 100 ? 10 : r <= 200 ? 20 : r <= 500 ? 50 : 100;

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
  circular(p) {                                        // gráfico circular (diagrama de torta)
    const C = p.categorias, tot = C.reduce((a, c) => a + c.valor, 0), cx = 290, cy = 330, R = 215; let s = txt(450, 58, p.titulo || '', 36), a0 = 0;
    C.forEach((c, i) => {
      const a1 = a0 + 360 * c.valor / tot, col = c.color || PAL[i % PAL.length];
      s += C.length === 1 ? `<circle cx="${cx}" cy="${cy}" r="${R}" fill="${col}" stroke="#fff" stroke-width="5"/>` : `<path d="${wedge(cx, cy, R, a0, a1)}" fill="${col}" stroke="#fff" stroke-width="5" stroke-linejoin="round"/>`;
      if ((p.mostrar === 'porcentaje' || p.mostrar === 'valor') && a1 - a0 > 20) { const [lx, ly] = polar(cx, cy, R * .64, (a0 + a1) / 2); s += txtH(lx, ly + 11, c.oculto ? '?' : p.mostrar === 'porcentaje' ? Math.round(100 * c.valor / tot) + '%' : String(c.valor), 32); }
      a0 = a1;
    });
    C.forEach((c, i) => { const y = 150 + i * 60, col = c.color || PAL[i % PAL.length]; s += `<rect x="560" y="${y - 26}" width="38" height="38" rx="8" fill="${col}" stroke="#1F2937" stroke-width="2"/>` + txt(612, y + 4, c.nombre + (p.leyendaValores ? ' — ' + (c.oculto ? '?' : c.valor) : ''), 29, { a: 'start', w: 'normal' }); });
    return s;
  },
  lineas(p) {                                          // gráfico de líneas (una o varias series)
    const x0 = 150, x1 = 850, y0 = 450, y1 = 125, X = p.x, S = p.series, vals = S.flatMap(r => r.valores.filter(v => v != null)), min = p.min || 0, max = p.max || Math.max(...vals), paso = p.paso || pasoAuto(max - min);
    const gap = X.length > 1 ? (x1 - x0) / (X.length - 1) : 0, px = (i) => X.length > 1 ? x0 + gap * i : (x0 + x1) / 2, py = (v) => y0 - (y0 - y1) * (v - min) / (max - min);
    let s = txt(450, 58, p.titulo || '', 36) + ejesXY(x0, x1, y0, y1, max, paso, min);
    X.forEach((lab, i) => { s += `<line x1="${px(i)}" y1="${y0}" x2="${px(i)}" y2="${y0 + 12}" stroke="#111827" stroke-width="4"/>` + txt(px(i), y0 + 46, String(lab), 26, { w: 'normal' }); });
    S.forEach((r, k) => {
      const col = r.color || PAL[k % PAL.length], pts = r.valores.map((v, i) => v == null ? null : [px(i), py(v)]); let d = '';
      pts.forEach(q => { if (q) d += (d ? 'L' : 'M') + q[0].toFixed(1) + ',' + q[1].toFixed(1); });
      s += `<path d="${d}" fill="none" stroke="${col}" stroke-width="7" stroke-linejoin="round" stroke-linecap="round"/>` + pts.filter(Boolean).map(q => `<circle cx="${q[0].toFixed(1)}" cy="${q[1].toFixed(1)}" r="10" fill="#fff" stroke="${col}" stroke-width="5"/>`).join('');
    });
    if (S.length > 1) { let lx = 160; S.forEach((r, k) => { const col = r.color || PAL[k % PAL.length]; s += `<line x1="${lx}" y1="96" x2="${lx + 44}" y2="96" stroke="${col}" stroke-width="7" stroke-linecap="round"/>` + txt(lx + 54, 106, r.nombre, 26, { a: 'start', w: 'normal' }); lx += 80 + r.nombre.length * 15; }); }
    if (p.ejeX) s += txt(450, 575, p.ejeX, 28, { w: 'normal', c: '#374151' });
    return s + (p.ejeY ? `<text x="40" y="290" ${FONT} font-size="28" fill="#374151" text-anchor="middle" transform="rotate(-90 40 290)">${esc(p.ejeY)}</text>` : '');
  },
  histograma(p) {                                      // histograma: barras pegadas por intervalos
    const I = p.intervalos, x0 = 150, x1 = 850, y0 = 450, y1 = 125, max = p.max || Math.max(...I.map(i => i.frecuencia)), paso = p.paso || pasoAuto(max), w = (x1 - x0) / I.length;
    let s = txt(450, 58, p.titulo || '', 36) + ejesXY(x0, x1, y0, y1, max, paso);
    I.forEach((it, i) => {
      const x = x0 + w * i, h = (y0 - y1) * (it.oculto ? .5 * max : it.frecuencia) / max;
      s += it.oculto ? `<rect x="${x}" y="${y0 - h}" width="${w}" height="${h}" fill="#F3F4F6" stroke="#6B7280" stroke-width="4" stroke-dasharray="10 7"/>` + txt(x + w / 2, y0 - h / 2 + 18, '?', 52, { c: '#6B7280' })
        : `<rect x="${x}" y="${y0 - h}" width="${w}" height="${h}" fill="#60A5FA" stroke="#1E3A8A" stroke-width="3"/>` + (p.mostrarValores ? txt(x + w / 2, y0 - h - 10, String(it.frecuencia), 26) : '');
      s += txt(x + w / 2, y0 + 44, it.etiqueta, 24, { w: 'normal' });
    });
    if (p.ejeX) s += txt(450, 575, p.ejeX, 28, { w: 'normal', c: '#374151' });
    return s + (p.ejeY ? `<text x="40" y="290" ${FONT} font-size="28" fill="#374151" text-anchor="middle" transform="rotate(-90 40 290)">${esc(p.ejeY)}</text>` : '');
  },
  plano(p) {                                           // plano cartesiano con puntos, segmentos y polígonos
    const xmin = p.xmin ?? -6, xmax = p.xmax ?? 6, ymin = p.ymin ?? -6, ymax = p.ymax ?? 6, u = Math.min(740 / (xmax - xmin), 450 / (ymax - ymin));
    const left = 450 - u * (xmax - xmin) / 2, top = 325 - u * (ymax - ymin) / 2, X = (x) => left + (x - xmin) * u, Y = (y) => top + (ymax - y) * u, paso = p.paso || ((xmax - xmin) > 14 ? 2 : 1);
    let s = txt(450, 40, p.titulo || '', 30);
    for (let x = Math.ceil(xmin); x <= xmax; x++) s += `<line x1="${X(x)}" y1="${Y(ymin)}" x2="${X(x)}" y2="${Y(ymax)}" stroke="#E5E7EB" stroke-width="2"/>`;
    for (let y = Math.ceil(ymin); y <= ymax; y++) s += `<line x1="${X(xmin)}" y1="${Y(y)}" x2="${X(xmax)}" y2="${Y(y)}" stroke="#E5E7EB" stroke-width="2"/>`;
    s += `<rect x="${X(xmin)}" y="${Y(ymax)}" width="${u * (xmax - xmin)}" height="${u * (ymax - ymin)}" fill="none" stroke="#9CA3AF" stroke-width="3"/>`;
    const ax = (xmin <= 0 && xmax >= 0) ? X(0) : X(xmin), ay = (ymin <= 0 && ymax >= 0) ? Y(0) : Y(ymin);
    s += flecha(X(xmin), ay, X(xmax) + 22, ay, '#111827', 5) + flecha(ax, Y(ymin), ax, Y(ymax) - 20, '#111827', 5) + txt(X(xmax) + 38, ay + 10, 'x', 30) + txt(ax + 4, Y(ymax) - 24, 'y', 30);
    for (let x = Math.ceil(xmin); x <= xmax; x++) if (x !== 0 && x % paso === 0) s += `<line x1="${X(x)}" y1="${ay - 7}" x2="${X(x)}" y2="${ay + 7}" stroke="#111827" stroke-width="3"/>` + txt(X(x), ay + 32, String(x), 20, { w: 'normal' });
    for (let y = Math.ceil(ymin); y <= ymax; y++) if (y !== 0 && y % paso === 0) s += `<line x1="${ax - 7}" y1="${Y(y)}" x2="${ax + 7}" y2="${Y(y)}" stroke="#111827" stroke-width="3"/>` + txt(ax - 14, Y(y) + 7, String(y), 20, { a: 'end', w: 'normal' });
    const P = p.puntos || [];
    if (p.poligono) s += `<polygon points="${p.poligono.map(i => `${X(P[i].x)},${Y(P[i].y)}`).join(' ')}" fill="#BFDBFE" fill-opacity=".55" stroke="#1E3A8A" stroke-width="6" stroke-linejoin="round"/>`;
    (p.segmentos || []).forEach(([a, b]) => { s += `<line x1="${X(P[a].x)}" y1="${Y(P[a].y)}" x2="${X(P[b].x)}" y2="${Y(P[b].y)}" stroke="#1E3A8A" stroke-width="6" stroke-linecap="round"/>`; });
    P.forEach(q => { s += `<circle cx="${X(q.x)}" cy="${Y(q.y)}" r="10" fill="${q.color || '#E11D48'}" stroke="#7F1029" stroke-width="3"/>` + (q.etiqueta ? txtH(X(q.x) + 16, Y(q.y) - 14, q.etiqueta, 28, { a: 'start' }) : ''); });
    return s;
  },
  figura(p) {                                          // figuras geométricas con medidas (usa "?" para la medida que se pide)
    const e = p.etiquetas || {}, F = '#DBEAFE', K = '#1E3A8A', poly = (pts) => `<polygon points="${pts.map(q => q.join(',')).join(' ')}" fill="${F}" stroke="${K}" stroke-width="7" stroke-linejoin="round"/>`;
    const recto = (x, y, dx, dy) => `<path d="M${x + dx * 34},${y} L${x + dx * 34},${y + dy * 34} L${x},${y + dy * 34}" fill="none" stroke="${K}" stroke-width="5"/>`, disc = (x1, y1, x2, y2) => `<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" stroke="#DC2626" stroke-width="5" stroke-dasharray="12 8"/>`;
    let s = '';
    switch (p.tipo) {
      case 'triangulo_rectangulo': s = poly([[220, 450], [690, 450], [220, 150]]) + recto(220, 450, 1, -1) + (e.cateto1 ? txtH(455, 505, e.cateto1, 40) : '') + (e.cateto2 ? txtH(190, 310, e.cateto2, 40, { a: 'end' }) : '') + (e.hipotenusa ? txtH(500, 285, e.hipotenusa, 40, { a: 'start' }) : ''); break;
      case 'triangulo': s = poly([[150, 450], [750, 450], [430, 150]]) + disc(430, 150, 430, 450) + recto(430, 450, 1, -1) + (e.base ? txtH(450, 505, e.base, 40) : '') + (e.altura ? txtH(444, 320, e.altura, 40, { a: 'start' }) : '') + (e.lado1 ? txtH(250, 290, e.lado1, 38, { a: 'end' }) : '') + (e.lado2 ? txtH(640, 290, e.lado2, 38, { a: 'start' }) : ''); break;
      case 'rectangulo': s = poly([[170, 170], [730, 170], [730, 430], [170, 430]]) + recto(170, 430, 1, -1) + (e.largo ? txtH(450, 495, e.largo, 40) : '') + (e.ancho ? txtH(760, 315, e.ancho, 40, { a: 'start' }) : ''); break;
      case 'cuadrado': s = poly([[290, 130], [610, 130], [610, 450], [290, 450]]) + recto(290, 450, 1, -1) + (e.lado ? txtH(450, 508, e.lado, 40) : '') + (e.lado2 ? txtH(640, 300, e.lado2, 40, { a: 'start' }) : ''); break;
      case 'circulo': s = `<circle cx="450" cy="300" r="195" fill="${F}" stroke="${K}" stroke-width="7"/><circle cx="450" cy="300" r="9" fill="${K}"/>` + (e.radio ? `<line x1="450" y1="300" x2="645" y2="300" stroke="#DC2626" stroke-width="6"/>` + txtH(548, 280, e.radio, 40) : '') + (e.diametro ? `<line x1="255" y1="${e.radio ? 380 : 300}" x2="645" y2="${e.radio ? 380 : 300}" stroke="#16A34A" stroke-width="6"/>` + txtH(450, e.radio ? 440 : 280, e.diametro, 40) : '') + (e.circunferencia ? txtH(450, 550, e.circunferencia, 38) : ''); break;
      case 'trapecio': s = poly([[270, 170], [630, 170], [780, 440], [120, 440]]) + disc(270, 170, 270, 440) + recto(270, 440, 1, -1) + (e.baseMenor ? txtH(450, 140, e.baseMenor, 40) : '') + (e.baseMayor ? txtH(450, 500, e.baseMayor, 40) : '') + (e.altura ? txtH(284, 320, e.altura, 40, { a: 'start' }) : '') + (e.lado ? txtH(745, 300, e.lado, 38, { a: 'start' }) : ''); break;
      case 'paralelogramo': s = poly([[260, 170], [770, 170], [640, 440], [130, 440]]) + disc(260, 170, 260, 440) + recto(260, 440, 1, -1) + (e.base ? txtH(385, 500, e.base, 40) : '') + (e.altura ? txtH(274, 320, e.altura, 40, { a: 'start' }) : '') + (e.lado ? txtH(740, 300, e.lado, 38, { a: 'start' }) : ''); break;
      case 'poligono': { const n = p.lados || 5, V = Array.from({ length: n }, (_, i) => polar(450, 300, 195, 360 * i / n + 180 / n).map(v => +v.toFixed(1))); const m = [(V[0][0] + V[1][0]) / 2, (V[0][1] + V[1][1]) / 2]; s = poly(V) + (e.lado ? txtH(m[0] + (m[0] - 450) * .22, m[1] + (m[1] - 300) * .22 + 12, e.lado, 38) : '') + (e.nombre ? txt(450, 310, e.nombre, 36, { c: K }) : ''); break; }
      default: throw new Error('figura: tipo no válido "' + p.tipo + '" (triangulo_rectangulo, triangulo, rectangulo, cuadrado, circulo, trapecio, paralelogramo, poligono)');
    }
    return s;
  },
  angulo(p) {                                          // ángulos a partir de rayos (grados, sentido antihorario desde la derecha)
    const V = [300, 440], R = p.rayos, largo = (g) => { const c = Math.cos(g * Math.PI / 180), n = Math.sin(g * Math.PI / 180); return Math.min(c > 1e-6 ? (860 - V[0]) / c : c < -1e-6 ? (V[0] - 40) / -c : 1e9, n > 1e-6 ? (V[1] - 50) / n : n < -1e-6 ? (570 - V[1]) / -n : 1e9); }, et = p.etiquetas || [], pt = (g, r) => [V[0] + r * Math.cos(g * Math.PI / 180), V[1] - r * Math.sin(g * Math.PI / 180)];
    let s = '';
    for (let i = 0; i < R.length - 1; i++) {
      const a = R[i], b = R[i + 1], r = 110, A = pt(a, r), B = pt(b, r), recto = Math.abs(b - a) === 90;
      s += recto ? `<path d="M${V[0]},${V[1]} L${pt(a, 70)[0]},${pt(a, 70)[1]} L${pt(a, 70)[0] + pt(b, 70)[0] - V[0]},${pt(a, 70)[1] + pt(b, 70)[1] - V[1]} L${pt(b, 70)[0]},${pt(b, 70)[1]} Z" fill="#BFDBFE" fill-opacity=".6" stroke="#1E3A8A" stroke-width="4"/>`
        : `<path d="M${V[0]},${V[1]} L${A[0].toFixed(1)},${A[1].toFixed(1)} A${r},${r} 0 ${(b - a) > 180 ? 1 : 0} 0 ${B[0].toFixed(1)},${B[1].toFixed(1)} Z" fill="#BFDBFE" fill-opacity=".6" stroke="#1E3A8A" stroke-width="4"/>`;
      if (et[i]) { const m = pt((a + b) / 2, recto ? 120 : 180); s += txtH(m[0], m[1] + 12, et[i], 38); }
    }
    R.forEach(g => { const q = pt(g, largo(g)); s += `<line x1="${V[0]}" y1="${V[1]}" x2="${q[0].toFixed(1)}" y2="${q[1].toFixed(1)}" stroke="#111827" stroke-width="7" stroke-linecap="round"/>`; });
    return s + `<circle cx="${V[0]}" cy="${V[1]}" r="11" fill="#111827"/>` + (p.nota ? txt(450, 570, p.nota, 30, { w: 'normal', c: '#374151' }) : '');
  },
  lineaTiempo(p) {                                     // línea de tiempo con eventos alternados arriba y abajo
    const ev = p.eventos, n = ev.length, x0 = 150, x1 = 750, y = 320, slot = n > 1 ? (x1 - x0) / (n - 1) : 600, maxc = Math.max(12, Math.min(22, Math.floor(slot * 2 / 14)));
    let s = txt(450, 55, p.titulo || '', 36) + `<line x1="${x0 - 60}" y1="${y}" x2="${x1 + 60}" y2="${y}" stroke="#111827" stroke-width="8" stroke-linecap="round"/>`;
    ev.forEach((e, i) => {
      const x = n > 1 ? x0 + slot * i : 450, up = i % 2 === 0, col = PAL[i % PAL.length], L = envolver(e.texto || '', maxc);
      s += `<line x1="${x}" y1="${y + (up ? -18 : 18)}" x2="${x}" y2="${y + (up ? -62 : 62)}" stroke="#6B7280" stroke-width="4"/><circle cx="${x}" cy="${y}" r="18" fill="${col}" stroke="#111827" stroke-width="5"/>`;
      if (up) { s += txt(x, y - 78, String(e.anio), 34, { c: '#1E3A8A' }); L.forEach((l, k) => { s += txt(x, y - 78 - 38 * (L.length - k), l, 24, { w: 'normal' }); }); }
      else { s += txt(x, y + 98, String(e.anio), 34, { c: '#1E3A8A' }); L.forEach((l, k) => { s += txt(x, y + 98 + 36 * (k + 1), l, 24, { w: 'normal' }); }); }
    });
    return s;
  },
  flujo(p) {                                           // procesos, cadenas y ciclos (nodos con flechas); un nodo {oculto:true} muestra "?"
    const N = p.nodos.map(n => typeof n === 'string' ? { texto: n } : n), m = N.length, modo = p.modo || 'fila', cols0 = Math.min(m, p.porFila || 3), bw = modo === 'fila' ? Math.min(210, 800 / cols0 - 66) : 210, bh = 96; let pos = [];
    if (modo === 'ciclo') N.forEach((_, i) => { const a = (-90 + 360 * i / m) * Math.PI / 180; pos.push([450 + 300 * Math.cos(a), 325 + 190 * Math.sin(a)]); });
    else if (modo === 'columna') N.forEach((_, i) => pos.push([450, 120 + Math.min(130, 420 / Math.max(1, m - 1)) * i]));
    else { const cols = cols0, rows = Math.ceil(m / cols), dx = 800 / cols, dy = Math.min(190, 400 / Math.max(1, rows - 1 || 1)); N.forEach((_, i) => { const r = Math.floor(i / cols), c = i % cols, cc = r % 2 ? cols - 1 - c : c; pos.push([50 + dx / 2 + dx * cc, (rows === 1 ? 310 : 150 + dy * r)]); }); }
    let s = txt(450, 50, p.titulo || '', 34);
    const salida = (a, b) => { const dx = b[0] - a[0], dy = b[1] - a[1], t = Math.min(dx ? (bw / 2 + 6) / Math.abs(dx) : 9, dy ? (bh / 2 + 6) / Math.abs(dy) : 9); return [a[0] + dx * t, a[1] + dy * t]; };
    const F = p.flechas || [...N.slice(0, -1).map((_, i) => [i, i + 1]), ...((modo === 'ciclo' || p.cerrar) && m > 2 ? [[m - 1, 0]] : [])];
    F.forEach(([a, b, lab]) => { const A = salida(pos[a], pos[b]), B = salida(pos[b], pos[a]); s += flecha(A[0], A[1], B[0], B[1], '#374151', 5) + (lab ? txtH((A[0] + B[0]) / 2, (A[1] + B[1]) / 2 - 10, lab, 22, { w: 'normal' }) : ''); });
    N.forEach((n, i) => {
      const [x, y] = pos[i], col = PAL[i % PAL.length], L = envolver(n.texto || '', Math.max(8, Math.floor(bw / 15))), fs = (L.length > 3 || L.some(l => l.length > 11) || bw < 170) ? 22 : 26;
      s += `<rect x="${x - bw / 2}" y="${y - bh / 2}" width="${bw}" height="${bh}" rx="18" fill="${n.oculto ? '#F3F4F6' : col + '22'}" stroke="${n.oculto ? '#6B7280' : col}" stroke-width="5" ${n.oculto ? 'stroke-dasharray="10 7"' : ''}/>`;
      s += n.oculto ? txt(x, y + 18, '?', 56, { c: '#6B7280' }) : L.map((l, k) => txt(x, y + 9 + (k - (L.length - 1) / 2) * (fs + 6), l, fs)).join('');
    });
    return s;
  },
  aviso(p) {                                           // aviso, señal, cartel o etiqueta (texto discontinuo)
    const est = p.estilo || 'informativo', T = { prohibido: ['#FEF2F2', '#DC2626'], advertencia: ['#FEFCE8', '#A16207'], informativo: ['#EFF6FF', '#1D4ED8'], invitacion: ['#FDF4FF', '#A21CAF'], etiqueta: ['#F9FAFB', '#374151'], seguridad: ['#F0FDF4', '#15803D'] }, [bg, fg] = T[est] || T.informativo;
    const icono = ['prohibido', 'advertencia', 'informativo', 'seguridad'].includes(est) && p.icono !== false, cx = icono ? 560 : 450;
    let s = `<rect x="70" y="60" width="760" height="480" rx="36" fill="${bg}" stroke="${fg}" stroke-width="12"/>`;
    if (icono) {
      if (est === 'prohibido') s += `<circle cx="200" cy="200" r="78" fill="#fff" stroke="${fg}" stroke-width="20"/><line x1="145" y1="145" x2="255" y2="255" stroke="${fg}" stroke-width="20" stroke-linecap="round"/>`;
      else if (est === 'advertencia') s += `<polygon points="200,118 285,270 115,270" fill="#FACC15" stroke="${fg}" stroke-width="12" stroke-linejoin="round"/>` + txt(200, 252, '!', 84, { c: '#1F2937' });
      else if (est === 'seguridad') s += `<circle cx="200" cy="200" r="80" fill="${fg}"/>` + txt(200, 232, '✓', 90, { c: '#fff' });
      else s += `<circle cx="200" cy="200" r="80" fill="${fg}"/>` + txt(200, 232, 'i', 96, { c: '#fff' });
    }
    const T1 = envolver(p.titulo || '', icono ? 14 : 20), ts = T1.length > 2 ? 46 : 60, y0 = icono ? 190 : 170; T1.forEach((l, k) => { s += txt(cx, y0 + k * (ts + 10), l, ts, { c: fg }); });
    const yIni = Math.max(icono ? y0 + T1.length * (ts + 10) + 70 : y0 + T1.length * (ts + 10) + 40, 330);
    let fs = 36, cuerpo = []; const armar = () => { cuerpo = (p.lineas || []).map(ln => envolver(ln, Math.floor(700 / (fs * 0.56)))); return cuerpo.reduce((a, l) => a + l.length * (fs + 10) + 14, 0); };
    while (armar() > 520 - yIni && fs > 24) fs -= 2;
    let y = yIni; cuerpo.forEach(ls => { ls.forEach(l => { s += txt(450, y, l, fs, { w: 'normal', c: '#1F2937' }); y += fs + 10; }); y += 14; });
    return s;
  },
  dialogo(p) {                                         // conversación en globos de diálogo
    const M = p.mensajes; let y = 40, g = '';
    M.forEach((m, i) => {
      const izq = (m.lado || (i % 2 ? 'der' : 'izq')) === 'izq', L = envolver(m.texto, 34), w = Math.min(680, Math.max(260, Math.max(...L.map(l => l.length)) * 19 + 70)), h = 52 + L.length * 38 + 14, x = izq ? 50 : 850 - w;
      g += `<rect x="${x}" y="${y}" width="${w}" height="${h}" rx="30" fill="${izq ? '#E0F2FE' : '#DCFCE7'}" stroke="${izq ? '#0369A1' : '#15803D'}" stroke-width="4"/>` + txt(x + 26, y + 40, m.quien || '', 26, { a: 'start', c: izq ? '#0369A1' : '#15803D' });
      L.forEach((l, k) => { g += txt(x + 26, y + 82 + k * 38, l, 29, { a: 'start', w: 'normal' }); });
      y += h + 22;
    });
    const k = Math.min(1, 560 / (y - 22 + 20)); return `<g transform="translate(${450 * (1 - k)},${10}) scale(${k})">${g}</g>`;
  },
  circuloCromatico(p) {                                // círculo cromático de 12 colores (ocultos: índices que se muestran como "?")
    const C = [['Amarillo', '#FFE81A'], ['Amarillo anaranjado', '#FDB515'], ['Anaranjado', '#F7821B'], ['Rojo anaranjado', '#EF4F23'], ['Rojo', '#E21F26'], ['Rojo violáceo', '#B5195D'], ['Violeta', '#7B2D8E'], ['Azul violáceo', '#4B3F9E'], ['Azul', '#1F5FBF'], ['Azul verdoso', '#13919A'], ['Verde', '#2BA84A'], ['Amarillo verdoso', '#9BCB3B']];
    const oc = new Set(p.ocultos || []), cx = 290, cy = 310, R = 215, r = 98; let s = txt(450, 50, p.titulo || '', 34);
    C.forEach(([nom, col], i) => {
      const a0 = i * 30 - 15, a1 = a0 + 30, [x0, y0] = polar(cx, cy, R, a0), [x1, y1] = polar(cx, cy, R, a1), [x2, y2] = polar(cx, cy, r, a1), [x3, y3] = polar(cx, cy, r, a0);
      s += `<path d="M${x0.toFixed(1)},${y0.toFixed(1)} A${R},${R} 0 0 1 ${x1.toFixed(1)},${y1.toFixed(1)} L${x2.toFixed(1)},${y2.toFixed(1)} A${r},${r} 0 0 0 ${x3.toFixed(1)},${y3.toFixed(1)} Z" fill="${oc.has(i) ? '#F3F4F6' : col}" stroke="${oc.has(i) ? '#6B7280' : '#fff'}" stroke-width="5" stroke-linejoin="round"/>`;
      if (oc.has(i)) { const [qx, qy] = polar(cx, cy, (R + r) / 2, i * 30); s += txt(qx, qy + 16, '?', 46, { c: '#6B7280' }); }
    });
    if (p.centro) s += txt(cx, cy + 10, p.centro, 28, { w: 'normal' });
    if (p.nombres) C.forEach(([nom, col], i) => { if (oc.has(i)) return; const y = 100 + i * 38; s += `<rect x="560" y="${y - 22}" width="30" height="30" rx="6" fill="${col}" stroke="#1F2937" stroke-width="2"/>` + txt(602, y + 3, nom, 24, { a: 'start', w: 'normal' }); });
    return s;
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
      celda([para([new ImageRun({ type: 'png', data: escudo.buf, transformation: { width: wPx, height: hPx }, altText: { name: 'Escudo', title: 'Escudo institucional', description: 'Escudo de ' + (INST.nombre || 'la institución') } })], { alignment: AlignmentType.CENTER })],
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
    const L = (p.enunciado || '').length, base = p._base || (L <= 180 ? [216, 144] : L <= 420 ? [189, 126] : [162, 108]);
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
// Texto compartido por varias preguntas (continuo o discontinuo): caja gris con título, párrafos, fuente y apoyo visual opcional
function bloqueEstimulo(est, n0, n1, st, imgsEst, juntar) {
  const out = [], rango = n0 === n1 ? `la pregunta ${n0}` : `las preguntas ${n0} ${n1 === n0 + 1 ? 'y' : 'a'} ${n1}`;
  const sh = { type: ShadingType.CLEAR, fill: 'F5F5F5', color: 'auto' }, bd = { style: BorderStyle.SINGLE, size: 6, color: '7F7F7F', space: 4 };
  const border = { top: bd, bottom: bd, left: bd, right: bd }, ind = { left: 120, right: 120 };
  const rangoEn = n0 === n1 ? `question ${n0}` : `questions ${n0} ${n1 === n0 + 1 ? 'and' : 'to'} ${n1}`, ins = est.instruccion ? est.instruccion.replace(/\{preguntas\}/g, rango).replace(/\{questions\}/g, rangoEn) : `Responde ${rango} de acuerdo con ${est.tipo === 'discontinuo' ? 'la siguiente información' : 'el siguiente texto'}.`;
  out.push(para([run(ins, { bold: true, size: SZ_S })], { alignment: AlignmentType.LEFT, spacing: { before: 0, after: 50 }, keepNext: true, keepLines: true }));
  const partes = String(est.texto || '').split(/\n+/).filter(Boolean), num = (i) => est.numerarParrafos ? [run(`${i + 1}. `, { bold: true, size: SZ_S })] : [];
  if (partes.length || est.titulo) {                 // caja gris = celda de tabla (los bordes de párrafo fallan al partirse entre columnas)
    const izq = [];
    if (est.titulo) izq.push(para([run(est.titulo, { bold: true, size: SZ_S })], { alignment: AlignmentType.CENTER, spacing: { before: 0, after: 40 }, keepNext: true, keepLines: true }));
    partes.forEach((t, k) => izq.push(para([...num(k), run(t, { size: SZ_S })], { alignment: AlignmentType.JUSTIFIED, spacing: { before: 0, after: 40 }, widowControl: true })));
    if (est.fuente) izq.push(para([run(est.fuente, { italics: true, size: SZ_XS, color: '555555' })], { alignment: AlignmentType.RIGHT, spacing: { before: 0, after: 0 } }));
    const caja = { borders: bAll('7F7F7F', 6), shading: sh };
    if (est._sprite) out.push(tablaLateral(izq, est._sprite, (est.texto || '').length + (est.titulo || '').length + 10, 'Ilustración del texto', 0.32, caja));
    else if (juntar) out.push(tabla([new TableRow({ cantSplit: true, children: [celda(izq, COL_W, { ...caja, margins: { top: 60, bottom: 50, left: 110, right: 110 } })] })], [COL_W]));   // una sola celda: el texto no se parte entre columnas
    else {                                          // texto muy largo: una fila por párrafo, la caja puede partirse entre columnas
      const ln = line('7F7F7F', 6), ult = izq.length - 1;
      out.push(tabla(izq.map((pp, k) => new TableRow({ cantSplit: false, children: [celda([pp], COL_W, { shading: sh, borders: { top: k === 0 ? ln : none, bottom: k === ult ? ln : none, left: ln, right: ln }, margins: { top: k === 0 ? 60 : 12, bottom: k === ult ? 50 : 12, left: 110, right: 110 } })] })), [COL_W]));
    }
  }
  if (est.visual) out.push(...bloqueVisual({ num: est.id, enunciado: est.texto || '', visual: est.visual, _base: est.texto ? null : [270, 180] }, st, imgsEst));
  out.push(para([], { spacing: { before: 0, after: 60, line: 80, lineRule: 'exact' }, keepNext: true }));
  return out;
}
// Unidad de maquetación = una pregunta suelta, o un texto compartido con TODAS sus preguntas (siempre juntos y en orden)
function unidades(mc) {
  const out = [];
  for (const p of mc) { const u = out[out.length - 1]; if (p._est && u && u.est === p._est) u.qs.push(p); else out.push({ est: p._est || null, qs: [p] }); }
  return out;
}
const aplanar = (us) => us.flatMap(u => u.qs);
function reunirGrupos(mc) {                           // las preguntas de un mismo texto quedan contiguas (en su orden relativo)
  const vistos = new Set(), out = [];
  for (const p of mc) { if (!p.estimulo) { out.push(p); continue; } if (vistos.has(p.estimulo)) continue; vistos.add(p.estimulo); mc.filter(q => q.estimulo === p.estimulo).forEach(q => out.push(q)); }
  return out;
}
function estimarAltoUnidad(u) {                        // alto aproximado (pt) del texto compartido + sus preguntas
  const lin = (n, cpl) => Math.ceil(n / cpl), e = u.est; let h = 14;
  if (e) h += (e.titulo ? 14 : 0) + lin((e.texto || '').length, 40) * 11 + 12 + (e.fuente ? 12 : 0) + (e.visual ? 130 : 0) + (e._sprite ? 40 : 0);
  u.qs.forEach(p => { h += lin(((p.enunciado || '') + p.pregunta).length, 34) * 12 + (p._ops || []).reduce((a, o) => a + lin(o.texto.length, 30) * 12, 0) + 14 + (p.visual ? 110 : 0); });
  return h;
}
function bloqueUnidad(u, st, imgs, imgsEst, docente, juntarTodo) {
  const out = [];
  if (u.est) out.push(...bloqueEstimulo(u.est, u.qs[0].num, u.qs[u.qs.length - 1].num, st, imgsEst || {}, juntarTodo || estimarAltoUnidad(u) <= 520));
  u.qs.forEach(p => out.push(...bloqueMC(p, st, imgs, docente)));
  return out;
}
function bloqueMC(p, st, imgs, docente) {
  const el = [];
  if (p.enunciado) {
    el.push(para([run(`${p.num}. `, { bold: true }), run(p.enunciado)], { alignment: AlignmentType.JUSTIFIED, spacing: { before: 0, after: st.enun }, keepNext: true, keepLines: true }));
    el.push(...bloqueVisual(p, st, imgs));
    el.push(para([run(p.pregunta, { bold: true })], { alignment: AlignmentType.JUSTIFIED, spacing: { before: 30, after: st.opt + 40 }, keepNext: true, keepLines: true }));
  } else {                                          // pregunta de un texto compartido: sin enunciado propio
    el.push(para([run(`${p.num}. `, { bold: true }), run(p.pregunta, { bold: true })], { alignment: AlignmentType.JUSTIFIED, spacing: { before: 0, after: st.opt + 40 }, keepNext: true, keepLines: true }));
    el.push(...bloqueVisual(p, st, imgs));
  }
  const ops = p._ops, corto = !p._sprite && ops.every(o => o.texto.length <= 16);
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
  if (p._sprite) {                                  // pregunta completa (enunciado + pregunta + opciones) a la izquierda, ilustración a la derecha
    const cplL = Math.floor((COL_W * 0.66 - 200) / 20 / 5.3), lineas = Math.ceil(((p.enunciado || '').length + p.pregunta.length + 4) / cplL) + ops.reduce((a, o) => a + Math.ceil((o.texto.length + 4) / (cplL - 4)), 0);
    return [tablaLateral(el, p._sprite, 0, `Ilustración de la pregunta ${p.num}`, 0.34, null, lineas), para([], { spacing: { before: 0, after: 0, line: 20, lineRule: 'exact' } })];
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
async function analizarLayout(pdf, claveUlt, esperados = []) {
  const pgs = await paginasPDF(pdf), norm = (t) => t.replace(/\s+/g, ' ');
  const textos = pgs.map((_, i) => norm(cp.execFileSync('pdftotext', ['-f', String(i + 1), '-l', String(i + 1), pdf, '-']).toString()));
  let ult = -1; textos.forEach((t, i) => { if (t.includes(claveUlt)) ult = i; });
  const hoja = textos.findIndex(t => t.includes('Marca un solo'));
  const top = 92, bot = PAGE_H / 20 - MARG.bottom / 20, lx0 = MARG.left / 20, lx1 = lx0 + COL_W / 20, rx0 = lx1 + COL_GAP / 20, rx1 = rx0 + COL_W / 20 + 4;
  const gaps = [];
  for (let p = 0; p < Math.max(ult, 0); p++) { const [, l] = tinta(pgs[p], lx0, lx1, top, bot + 10), [, r] = tinta(pgs[p], rx0, rx1, top, bot + 10); gaps.push({ pagina: p + 1, izq: Math.round(bot - l), der: Math.round(bot - r) }); }
  const maxGap = gaps.length ? Math.max(...gaps.flatMap(g => [g.izq, g.der])) : 0;
  const todo = textos.join(' '), faltan = esperados.filter(t => !todo.includes(t));   // integridad: todo el texto del Word debe verse en el PDF
  return { paginas: pgs.length, paginaUltimaPregunta: ult + 1, paginaHoja: hoja + 1, hojaMismaPagina: hoja === ult, gaps, maxGap, faltan };
}

// ═══ 7b. IMÁGENES CON CHATGPT: prompt de hojas (sprites), recorte automático y colocación lateral ═══════
const IMG_POR_HOJA = 8, MAX_HOJAS = 3, MAX_IMG = IMG_POR_HOJA * MAX_HOJAS;

// Numera las imágenes (1..24) en el orden en que aparecen en la prueba. Respeta los números "imagen" ya asignados.
function asignarImagenes(ex) {
  const estMap = Object.fromEntries((ex.estimulos || []).map(e => [e.id, e])), cand = [], vistos = new Set(); let k = 0;
  (ex.preguntas || []).forEach(p => {
    if ((p.tipo || 'multiple') !== 'multiple') { if (p.imagenPrompt) cand.push({ ref: p, tipo: 'pregunta abierta', etiqueta: 'Pregunta abierta' }); return; }
    k++; const e = p.estimulo && estMap[p.estimulo];
    if (e && e.imagenPrompt && !vistos.has(e)) { vistos.add(e); cand.push({ ref: e, tipo: 'texto compartido', etiqueta: `Texto «${e.titulo || e.id}» (pregunta ${k})` }); }
    if (p.imagenPrompt) cand.push({ ref: p, tipo: 'pregunta', etiqueta: `Pregunta ${k}` });
  });
  const usados = new Set();
  cand.forEach(c => { const n = c.ref.imagen; if (Number.isInteger(n) && n >= 1 && n <= MAX_IMG && !usados.has(n)) usados.add(n); else delete c.ref.imagen; });
  let sig = 1; cand.forEach(c => { if (c.ref.imagen) return; while (usados.has(sig)) sig++; if (sig <= MAX_IMG) { c.ref.imagen = sig; usados.add(sig); } });
  const items = cand.filter(c => c.ref.imagen).map(c => ({ ...c, n: c.ref.imagen })).sort((a, b) => a.n - b.n);
  items.forEach(i => { i.hoja = Math.ceil(i.n / IMG_POR_HOJA); const pos = (i.n - 1) % IMG_POR_HOJA; i.fila = Math.floor(pos / 4) + 1; i.col = pos % 4 + 1; });
  const sobran = cand.length - items.length; if (sobran > 0) console.warn(`⚠  Hay ${sobran} imagen(es) de más: el máximo es ${MAX_IMG} (3 hojas de 8). Se ignoran las últimas.`);
  return items;
}
function estiloPorGrado(g) {
  g = Number(g);
  if (g && g <= 5) return 'ilustración digital limpia y amable, estilo de libro escolar para niños de primaria: formas redondeadas, colores vivos y suaves, contornos definidos de grosor uniforme';
  if (g && g <= 9) return 'ilustración digital clara y moderna, estilo de libro escolar para adolescentes: formas definidas, colores vivos pero equilibrados, contornos limpios y detalles reconocibles';
  return 'ilustración digital sobria y detallada, estilo de material educativo para jóvenes: proporciones realistas, colores naturales, contornos finos y buena iluminación';
}
function textoPromptImagenes(ex, items) {
  const m = ex.meta || {}, inst = INST.nombre || 'una institución educativa', hojas = Math.max(...items.map(i => i.hoja)), estilo = ex.config.estiloImagenes || estiloPorGrado(m.grado);
  const general = `Eres un ilustrador profesional de material escolar. Vas a crear HOJAS DE ILUSTRACIONES (sprite sheets) para una evaluación escolar de ${m.area || m.asignatura || 'la asignatura'}, tema «${m.tema || ''}», grado ${m.grado || ''}.° (${inst}, Colombia). Las ilustraciones servirán de apoyo visual para que la evaluación sea más gráfica y clara.

ESTILO (idéntico en todas las hojas): ${estilo}. Misma paleta, mismo grosor de línea y mismo nivel de detalle en las ${hojas} hojas. Usa contextos cercanos a estudiantes colombianos (colegio, familia, barrio, naturaleza del país) cuando la descripción lo permita. Personas diversas y respetuosas, sin estereotipos. Sin marcas, logotipos ni personajes con derechos de autor, y sin texto dentro de las ilustraciones (salvo que la descripción pida un texto exacto).

FORMATO DE CADA HOJA:
• UNA sola imagen horizontal por hoja, de 2048 × 1024 px (proporción 2:1), con 8 ilustraciones en una cuadrícula de 2 filas × 4 columnas (4 arriba y 4 abajo).
• Cada ilustración va centrada en su celda (de 512 × 512 px aprox.), COMPLETA (nada cortado) y ocupando cerca del 80 % de la celda.
• Deja un espacio vacío amplio entre ilustraciones: no deben tocarse, superponerse ni compartir sombras.
• Sin líneas de cuadrícula, sin marcos, sin números, sin títulos ni etiquetas.

FONDO — OBLIGATORIO:
• Fondo TRANSPARENTE REAL (canal alfa), entregado como archivo PNG.
• PROHIBIDO dibujar la cuadrícula gris y blanca que simula la transparencia; PROHIBIDO usar fondo blanco, de color o degradado.
• Sin sombras proyectadas sobre el fondo y sin halos ni bordes blancos alrededor de las figuras (si hay sombra, debe ser parte del objeto).
• Si no puedes entregar un PNG con transparencia real, avísame ANTES de generar la imagen.

CALIDAD: alta resolución, líneas limpias, bordes nítidos, colores vivos pero suaves, buena iluminación y proporciones correctas.
Las ilustraciones NO deben revelar la respuesta de ninguna pregunta: solo contextualizan la situación.

Te pediré las hojas una por una (${Array.from({ length: hojas }, (_, i) => 'Hoja ' + (i + 1)).join(', ')}). Responde solo «Entendido» y espera la primera hoja.`;
  const hojasTxt = [];
  for (let h = 1; h <= hojas; h++) {
    const its = items.filter(i => i.hoja === h), L = [];
    L.push(`HOJA ${h} de ${hojas} — ${IMG_POR_HOJA} ilustraciones (2 filas × 4 columnas)`, '', `Genera ahora la HOJA ${h}: UNA sola imagen PNG de 2048 × 1024 px con fondo TRANSPARENTE REAL (canal alfa), con el formato y las reglas del prompt general. Contenido, de izquierda a derecha y de arriba abajo:`, '');
    for (let pos = 0; pos < IMG_POR_HOJA; pos++) {
      const n = (h - 1) * IMG_POR_HOJA + pos + 1, it = its.find(i => i.n === n);
      if (pos === 0) L.push('FILA 1 (arriba)'); if (pos === 4) L.push('', 'FILA 2 (abajo)');
      L.push(it ? `${pos % 4 + 1}. [Imagen ${n} · ${it.etiqueta}] ${String(it.ref.imagenPrompt).trim()}` : `${pos % 4 + 1}. [Imagen ${n}] CELDA VACÍA: no dibujes nada en este espacio.`);
    }
    L.push('', 'Recuerda: 8 ilustraciones separadas entre sí, sin texto, sin cuadrícula, bordes limpios y fondo transparente real (PNG con canal alfa).');
    hojasTxt.push(L.join('\n'));
  }
  let md = `# Prompts para crear las imágenes de la evaluación con ChatGPT\n\n**Evaluación:** ${m.area || ''} · ${m.tema || ''} · grado ${m.grado || ''}.°  \n**Imágenes:** ${items.length} en ${hojas} hoja(s) de ${IMG_POR_HOJA}\n\n## Cómo se usa (4 pasos)\n1. Abre un chat **nuevo** en ChatGPT (con generación de imágenes) y pega el **PROMPT GENERAL**.\n2. Cuando responda «Entendido», pega el prompt de la **HOJA 1**. Descarga la imagen como **PNG**. Repite con la HOJA 2${hojas > 2 ? ' y la HOJA 3' : ''}.\n3. Si el fondo sale blanco o con cuadrícula gris, escribe: *«Quita el fondo y entrégame el PNG con transparencia real (canal alfa), sin cuadrícula.»*\n4. Sube las hojas a esta conversación **en orden** (Hoja 1, Hoja 2${hojas > 2 ? ', Hoja 3' : ''}). Yo recorto cada ilustración, la acomodo en la prueba al lado del texto y te entrego la evaluación terminada.\n\n---\n\n## 1) PROMPT GENERAL — pégalo primero\n\n\`\`\`text\n${general}\n\`\`\`\n`;
  hojasTxt.forEach((t, i) => { md += `\n---\n\n## ${i + 2}) HOJA ${i + 1} — pégalo después\n\n\`\`\`text\n${t}\n\`\`\`\n`; });
  return { md, general, hojas: hojasTxt };
}

// ── Lectura y limpieza de las hojas ────────────────────────────────────────
async function leerRGBA(file) { const { data, info } = await sharp(file, { limitInputPixels: false }).ensureAlpha().raw().toBuffer({ resolveWithObject: true }); return { d: data, w: info.width, h: info.height }; }
function tieneTransparencia(im) { let n = 0; for (let i = 3; i < im.d.length; i += 4) if (im.d[i] < 250) n++; return n > im.w * im.h * 0.005; }
function quitarFondoPorColor(im) {                    // rescate: hoja sin transparencia real (fondo blanco, de color o cuadrícula gris simulada)
  const { d, w, h } = im, bins = new Map(), q = (r, g, b) => (r >> 4) << 8 | (g >> 4) << 4 | (b >> 4);
  const mu = (x, y) => { const i = (y * w + x) * 4, k = q(d[i], d[i + 1], d[i + 2]); bins.set(k, (bins.get(k) || 0) + 1); };
  for (let x = 0; x < w; x++) { mu(x, 0); mu(x, h - 1); } for (let y = 0; y < h; y++) { mu(0, y); mu(w - 1, y); }
  const tot = 2 * (w + h), fondo = [...bins.entries()].filter(([, c]) => c > tot * 0.08).sort((a, b) => b[1] - a[1]).slice(0, 3).map(([k]) => [((k >> 8) & 15) * 16 + 8, ((k >> 4) & 15) * 16 + 8, (k & 15) * 16 + 8]);
  if (!fondo.length) return false;
  const tol = 34, par = (p) => { const i = p * 4; return fondo.some(c => Math.abs(d[i] - c[0]) <= tol && Math.abs(d[i + 1] - c[1]) <= tol && Math.abs(d[i + 2] - c[2]) <= tol); };
  const vis = new Uint8Array(w * h), pila = new Int32Array(w * h); let sp = 0;
  const push = (x, y) => { const p = y * w + x; if (!vis[p] && par(p)) { vis[p] = 1; pila[sp++] = p; } };
  for (let x = 0; x < w; x++) { push(x, 0); push(x, h - 1); } for (let y = 0; y < h; y++) { push(0, y); push(w - 1, y); }
  while (sp) { const p = pila[--sp], x = p % w, y = (p / w) | 0; if (x > 0) push(x - 1, y); if (x < w - 1) push(x + 1, y); if (y > 0) push(x, y - 1); if (y < h - 1) push(x, y + 1); }
  for (let p = 0; p < w * h; p++) if (vis[p]) d[p * 4 + 3] = 0;
  const borde = []; for (let p = 0; p < w * h; p++) { if (vis[p]) continue; const x = p % w, y = (p / w) | 0, i = p * 4; if (Math.min(d[i], d[i + 1], d[i + 2]) > 205 && ((x > 0 && vis[p - 1]) || (x < w - 1 && vis[p + 1]) || (y > 0 && vis[p - w]) || (y < h - 1 && vis[p + w]))) borde.push(p); }
  borde.forEach(p => { d[p * 4 + 3] = 0; });        // quita el halo claro de 1 px
  return true;
}
// Detecta las ilustraciones (componentes conexas separadas por transparencia) y las ordena por filas (arriba→abajo, izquierda→derecha)
function detectarObjetos(im, esperados) {
  const { d, w, h } = im, S = Math.max(1, Math.round(Math.max(w, h) / 640)), mw = Math.ceil(w / S), mh = Math.ceil(h / S), m = new Uint8Array(mw * mh);
  for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) if (d[(y * w + x) * 4 + 3] > 24) m[((y / S) | 0) * mw + ((x / S) | 0)] = 1;
  const et = new Int32Array(mw * mh).fill(-1), pila = new Int32Array(mw * mh); let cs = [];
  for (let s = 0; s < mw * mh; s++) {
    if (!m[s] || et[s] >= 0) continue; let sp = 0; const c = { x0: mw, y0: mh, x1: 0, y1: 0, a: 0 }; pila[sp++] = s; et[s] = cs.length;
    while (sp) { const p = pila[--sp], x = p % mw, y = (p / mw) | 0; c.a++; if (x < c.x0) c.x0 = x; if (x > c.x1) c.x1 = x; if (y < c.y0) c.y0 = y; if (y > c.y1) c.y1 = y;
      for (let dy = -1; dy <= 1; dy++) for (let dx = -1; dx <= 1; dx++) { const nx = x + dx, ny = y + dy; if (nx < 0 || ny < 0 || nx >= mw || ny >= mh) continue; const q = ny * mw + nx; if (m[q] && et[q] < 0) { et[q] = cs.length; pila[sp++] = q; } } }
    cs.push(c);
  }
  cs = cs.filter(c => c.a >= 4);
  for (let gap = 0; gap <= 0.07; gap += 0.005) {      // une las partes sueltas de una misma ilustración, solo si hay de más
    if (cs.length <= esperados) break; const G = gap * Math.max(mw, mh); let cambio = true;
    while (cambio && cs.length > esperados) { cambio = false;
      outer: for (let i = 0; i < cs.length; i++) for (let j = i + 1; j < cs.length; j++) { const a = cs[i], b = cs[j];
        if (a.x0 - G <= b.x1 && b.x0 - G <= a.x1 && a.y0 - G <= b.y1 && b.y0 - G <= a.y1) { cs[i] = { x0: Math.min(a.x0, b.x0), y0: Math.min(a.y0, b.y0), x1: Math.max(a.x1, b.x1), y1: Math.max(a.y1, b.y1), a: a.a + b.a }; cs.splice(j, 1); cambio = true; break outer; } } }
  }
  if (cs.length > esperados) cs = cs.sort((a, b) => b.a - a.a).slice(0, esperados);
  cs.forEach(c => { c.cy = (c.y0 + c.y1) / 2; c.cx = (c.x0 + c.x1) / 2; }); cs.sort((a, b) => a.cy - b.cy);
  const hs = cs.map(c => c.y1 - c.y0).sort((a, b) => a - b), med = hs[(hs.length / 2) | 0] || 1, filas = [];
  cs.forEach(c => { const f = filas[filas.length - 1]; if (f && Math.abs(c.cy - f.cy) < med * 0.6) { f.it.push(c); f.cy = f.it.reduce((s, x) => s + x.cy, 0) / f.it.length; } else filas.push({ cy: c.cy, it: [c] }); });
  return filas.flatMap(f => f.it.sort((a, b) => a.cx - b.cx)).map(c => ({ x0: c.x0 * S, y0: c.y0 * S, x1: Math.min(w - 1, (c.x1 + 1) * S), y1: Math.min(h - 1, (c.y1 + 1) * S), S }));
}
async function cortarObjeto(im, r, file) {           // recorte exacto al borde de la ilustración (sin márgenes) y PNG con transparencia
  const { d, w } = im; let x0 = r.x1, y0 = r.y1, x1 = r.x0, y1 = r.y0;
  for (let y = r.y0; y <= r.y1; y++) for (let x = r.x0; x <= r.x1; x++) if (d[(y * w + x) * 4 + 3] > 12) { if (x < x0) x0 = x; if (x > x1) x1 = x; if (y < y0) y0 = y; if (y > y1) y1 = y; }
  const cw = x1 - x0 + 1, ch = y1 - y0 + 1, buf = Buffer.alloc(cw * ch * 4);
  for (let y = 0; y < ch; y++) d.copy(buf, y * cw * 4, ((y0 + y) * w + x0) * 4, ((y0 + y) * w + x0 + cw) * 4);
  for (let i = 3; i < buf.length; i += 4) if (buf[i] < 10) buf[i] = 0;
  let s = sharp(buf, { raw: { width: cw, height: ch, channels: 4 } }); if (Math.max(cw, ch) > 900) s = s.resize({ width: cw >= ch ? 900 : undefined, height: ch > cw ? 900 : undefined, fit: 'inside' });
  await s.png({ compressionLevel: 9 }).toFile(file); const meta = await sharp(file).metadata(); return { file, w: meta.width, h: meta.height };
}
async function cortarHojas(archivos, items, outDir) {   // → { [n]: {file,w,h} }
  const dir = path.join(outDir, 'img', 'sprites'); fs.mkdirSync(dir, { recursive: true }); const res = {}, hojas = Math.max(...items.map(i => i.hoja));
  if (archivos.length !== hojas) console.warn(`⚠  La prueba necesita ${hojas} hoja(s) de imágenes y recibí ${archivos.length}.`);
  for (let h = 1; h <= Math.min(hojas, archivos.length); h++) {
    const its = items.filter(i => i.hoja === h).sort((a, b) => a.n - b.n), im = await leerRGBA(archivos[h - 1]);
    console.log(`🖼️  Hoja ${h}: ${path.basename(archivos[h - 1])} (${im.w}×${im.h} px)`);
    if (!tieneTransparencia(im)) { const ok = quitarFondoPorColor(im); console.warn(`⚠  La hoja ${h} NO tiene transparencia real (${ok ? 'se quitó el fondo por color; revisa los bordes' : 'no pude quitar el fondo'}). Pide a ChatGPT un PNG con canal alfa para mejor calidad.`); }
    const cajas = detectarObjetos(im, its.length);
    if (cajas.length !== its.length) console.warn(`⚠  Hoja ${h}: encontré ${cajas.length} ilustraciones y esperaba ${its.length}. Revisa la previsualización.`);
    for (let k = 0; k < Math.min(cajas.length, its.length); k++) res[its[k].n] = await cortarObjeto(im, cajas[k], path.join(dir, `imagen_${String(its[k].n).padStart(2, '0')}.png`));
  }
  return res;
}
async function previsualizarSprites(res, outFile) {   // hoja de contacto numerada para comprobar los recortes de un vistazo
  const ns = Object.keys(res).map(Number).sort((a, b) => a - b), C = 220, cols = Math.min(8, Math.max(1, ns.length)), filas = Math.ceil(ns.length / cols), comp = [];
  for (let i = 0; i < ns.length; i++) {
    const x = (i % cols) * C, y = Math.floor(i / cols) * C, t = await sharp(res[ns[i]].file).resize({ width: C - 24, height: C - 24, fit: 'inside' }).png().toBuffer(), mt = await sharp(t).metadata();
    comp.push({ input: t, left: x + Math.round((C - mt.width) / 2), top: y + Math.round((C - mt.height) / 2) });
    comp.push({ input: Buffer.from(`<svg xmlns="http://www.w3.org/2000/svg" width="64" height="36"><rect width="64" height="36" rx="8" fill="#1F2937"/><text x="32" y="26" font-family="DejaVu Sans, Arial" font-size="22" font-weight="bold" fill="#fff" text-anchor="middle">${ns[i]}</text></svg>`), left: x + 6, top: y + 6 });
  }
  await sharp({ create: { width: cols * C, height: filas * C, channels: 4, background: '#E5E7EB' } }).composite(comp).png().toFile(outFile);
}
// (estímulos con ilustración) tabla: caja gris a la izquierda e ilustración a la derecha
function tablaLateral(izq, spr, nchars, alt, frac = 0.36, caja = null, lineasExactas = 0) {
  const wR = Math.round(COL_W * frac), wL = COL_W - wR, cpl = Math.max(18, Math.floor((wL - 200) / 20 / 5.3)), hTxt = Math.max(62, (lineasExactas || (Math.ceil(nchars / cpl) + 1)) * 12.6 + 8);
  const maxW = (wR - 150) / 20 * 96 / 72, asp = spr.w / spr.h; let w = Math.min(maxW, hTxt * 96 / 72 * asp); w = Math.max(w, Math.min(72, maxW)); const h = w / asp;
  const nb = { top: none, bottom: none, left: none, right: none };
  return tabla([new TableRow({ cantSplit: true, children: [
    celda(izq, wL, caja ? { borders: caja.borders, shading: caja.shading, margins: { top: 60, bottom: 40, left: 110, right: 110 } } : { borders: nb, margins: { top: 0, bottom: 0, left: 0, right: 100 } }),
    celda([para([new ImageRun({ type: 'png', data: fs.readFileSync(spr.file), transformation: { width: Math.round(w), height: Math.round(h) }, altText: { name: 'ilustracion', title: alt, description: alt } })], { alignment: AlignmentType.CENTER, spacing: { before: 0, after: 0 }, keepNext: true })], wR, { borders: nb, verticalAlign: VerticalAlign.CENTER, margins: { top: 0, bottom: 0, left: 40, right: 0 } })] })],
    [wL, wR], { borders: { top: none, bottom: none, left: none, right: none, insideHorizontal: none, insideVertical: none } });
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
    para([run((AREA && AREA.niveles_cortos ? 'Niveles de desempeño: ' + AREA.niveles_cortos.map((n, i) => `N${i + 1} = ${n}`).join(' · ') + '. ' : 'N1–N4 = nivel de desempeño SABER de cada pregunta de selección múltiple. ') + 'Úsala para comprobar que la prueba cubre todas las competencias y no concentra la dificultad en un solo nivel.', { size: SZ_XS, italics: true })], { alignment: AlignmentType.JUSTIFIED, spacing: { after: 80 }, keepNext: true }),
    tabla(filas, [wC, wComp, wP, wT, wN, wN, wN, wN])
  ];
}
async function armar(ctx, nivel, hojaPropia) {
  const { ex, imgs, bubs, escudo, outDir } = ctx, st = NIVELES[nivel], m = ex.meta;
  let mc = ctx.mc.slice(); const ab = ctx.ab.slice();
  mc.forEach(p => p._ops = [p.correcta, ...p.distractores].map((t, i) => ({ letra: 'ABCD'[i], texto: String(t), ok: i === 0, error: '' })));
  const hdrCfg = Object.assign({ tituloMembrete: ['GESTION ACADEMICA', 'SEGUIMIENTO ACADEMICO', 'EVALUACION DE COMPRENSION'], codFormato: 'GA-F16', codVersion: 'Versión 1.0', codFecha: '04-04-2016' }, INST.membrete || {}, ex.config.membrete || {});
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
    const uni0 = unidades(mc);
    const bl = [...uni0.map(u => () => bloqueUnidad(u, st, imgs, ctx.imgsEst, false, ctx.juntarTodo)), () => [bloqueParte('PARTE I')], ...(ab.length ? [() => [bloqueParte('PARTE II')], ...ab.map(p => () => bloqueAbierta(p, st, false))] : [])];
    const hs = await medirBloques(bl, outDir, 'b' + nivel, st.fin);
    const hdoc = new Document({ styles: estilosDoc, compatibility: { version: 15 }, sections: [{ properties: sizeOf({}), headers: { default: makeHeader(escudo, hdrCfg) }, children: sec1(false) }] });
    const fS = path.join(outDir, `_sec1_${nivel}.docx`); fs.writeFileSync(fS, await Packer.toBuffer(hdoc)); const pS = aPDF(fS, outDir);
    if (hs && pS) {
      const pg = (await paginasPDF(pS))[0], [, l] = tinta(pg, 0, pg.w, 92, pg.h), capFirst = (PAGE_H / 20 - MARG.bottom / 20) - (l + 4 + 5), capFull = (PAGE_H - MARG.top - MARG.bottom) / 20;
      const itM = uni0.map((u, i) => ({ p: u, h: hs[i] })), hP1 = hs[uni0.length], hP2 = ab.length ? hs[uni0.length + 1] : 0, itA = ab.map((p, i) => ({ p, h: hs[uni0.length + 2 + i] }));
      const r = empaquetar(itM, itA, hP1, hP2, capFirst, capFull);
      const nuevos = aplanar(r.mc.map(x => x.p)); nuevos.forEach((x, i) => { if (x !== mc[i]) movidas++; }); mc = nuevos; r.ab.forEach((x, i) => { ab[i] = x.p; });
      medido = true; try { fs.unlinkSync(fS); fs.unlinkSync(pS); } catch (e) {}
    } else console.warn('⚠  No se pudo medir con LibreOffice; se mantiene el orden original.');
  }
  mc.forEach((p, i) => p.num = i + 1); ab.forEach((p, i) => p.num = mc.length + i + 1);

  const seed = ([...(m.tema || 'examen')].reduce((a, c) => a * 31 + c.charCodeAt(0) >>> 0, 7) + (ctx.seedExtra || 0)) >>> 0;
  const dist = asignarLetras(mc, seed), claves = mc.map(p => p._clave);
  const pc = (l) => Math.round(100 * dist[l] / mc.length);
  const FRASES = { matematicas: '¡Cada problema tiene solución! Confía en ti.', ciencias: '¡El conocimiento es tu superpoder! Da lo mejor.', lenguaje: '¡Las palabras son tu fuerza! Lee, piensa y triunfa.', sociales: '¡Conocer el mundo comienza contigo! Éxito.', tecnologia: '¡La tecnología está en tus manos! Demuéstralo.', general: '¡Tú puedes! Respira, enfócate y demuestra todo lo que sabes.' };
  const frase = (AREA && AREA.frase_final) || FRASES[(m.areaClave || 'general')] || FRASES.general;
  const cuerpo = (docente) => { const els = [bloqueParte(`PARTE I — SELECCIÓN MÚLTIPLE CON ÚNICA RESPUESTA (${mc.length})`)]; unidades(mc).forEach(u => els.push(...bloqueUnidad(u, st, imgs, ctx.imgsEst, docente, ctx.juntarTodo))); if (ab.length) { els.push(bloqueParte(`PARTE II — PREGUNTAS ABIERTAS (${ab.length})`)); ab.forEach(p => els.push(...bloqueAbierta(p, st, docente))); } return els; };
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
  const ult = mc.length && !ab.length ? (mc[mc.length - 1].enunciado || mc[mc.length - 1].pregunta) : (ab.length ? ab[ab.length - 1].enunciado : '');
  const ult2 = (t) => String(t || '').replace(/\s+/g, ' ').trim().split(' ').slice(-2).join(' ');
  const esperados = [...mc.flatMap(p => [ult2(p.enunciado), ult2(p.pregunta), ...p._ops.map(o => ult2(o.texto))]), ...ab.map(p => ult2(p.pregunta)), ...(ctx.ex.estimulos || []).filter(e => mc.some(p => p._est === e)).map(e => ult2(e.texto))].filter(t => t.length > 3);
  return { fE, fD, claves, dist, pc, movidas, medido, esperados, claveUlt: ult.replace(/\s+/g, ' ').split(' ').slice(0, 4).join(' ') };
}

(async () => {
  const args = process.argv.slice(2), jsonPath = args.find(a => a.endsWith('.json'));
  if (!jsonPath) { console.error('Uso: node build_exam.js exam.json [--nivel 0-3] [--out carpeta] [--sin-ajuste] [--rapido] [--versiones 2] [--institucion archivo.json]'); process.exit(2); }
  cargarInstitucion(args);
  const nivelFijo = args.includes('--nivel') ? Number(args[args.indexOf('--nivel') + 1]) : null;
  const outDir = args.includes('--out') ? args[args.indexOf('--out') + 1] : './salida';
  fs.mkdirSync(path.join(outDir, 'img'), { recursive: true });
  const ex = JSON.parse(fs.readFileSync(jsonPath, 'utf8')); ex.config = ex.config || {}; ex.meta = ex.meta || {};

  cargarArea(ex.meta.areaClave, ex.meta.grado);
  const { E, W } = validar(ex); W.forEach(w => console.warn('⚠  ' + w));
  if (E.length) { console.error('\n❌ LA PRUEBA NO PASA LA VALIDACIÓN. Corrige y vuelve a ejecutar:\n' + E.map(e => '   • ' + e).join('\n')); process.exit(1); }
  console.log('✅ Validación OK: opciones únicas, respuesta presente, longitudes en rango.');

  const estMap = Object.fromEntries((ex.estimulos || []).map(e => [e.id, e]));
  // ── Imágenes con ChatGPT: (1) --prompt-imagenes escribe el prompt de las hojas; (2) --sprites hoja1.png hoja2.png… las recorta y las coloca ──
  const itemsImg = asignarImagenes(ex), sp = args.indexOf('--sprites');
  if (args.includes('--prompt-imagenes')) {
    if (!itemsImg.length) { console.error('❌ Ninguna pregunta ni texto tiene "imagenPrompt". Agrega una descripción detallada de la ilustración en cada pregunta que la necesite.'); process.exit(1); }
    const t = textoPromptImagenes(ex, itemsImg);
    fs.writeFileSync(path.join(outDir, 'prompt_imagenes_chatgpt.md'), t.md);
    fs.writeFileSync(path.join(outDir, 'imagenes_mapa.json'), JSON.stringify(itemsImg.map(i => ({ imagen: i.n, hoja: i.hoja, fila: i.fila, columna: i.col, para: i.etiqueta, descripcion: i.ref.imagenPrompt })), null, 2));
    fs.writeFileSync(jsonPath, JSON.stringify(ex, null, 2));       // guarda el número de cada imagen en el exam.json (así no cambia después)
    console.log(`🖼️  ${itemsImg.length} imagen(es) en ${Math.max(...itemsImg.map(i => i.hoja))} hoja(s). Prompts → ${path.join(outDir, 'prompt_imagenes_chatgpt.md')}`);
    process.exit(0);
  }
  let spritesCortados = null;
  if (sp >= 0) {
    let arch = []; for (let k = sp + 1; k < args.length && !args[k].startsWith('--'); k++) arch.push(args[k]);
    if (arch.length === 1 && fs.existsSync(arch[0]) && fs.statSync(arch[0]).isDirectory()) { const dd = arch[0]; arch = fs.readdirSync(dd).filter(f => /\.(png|webp|jpe?g)$/i.test(f)).sort((a, b) => a.localeCompare(b, undefined, { numeric: true })).map(f => path.join(dd, f)); }
    if (!arch.length) { console.error('❌ --sprites necesita las hojas (PNG) en orden: --sprites hoja1.png hoja2.png hoja3.png  (o una carpeta).'); process.exit(1); }
    if (!itemsImg.length) { console.error('❌ La prueba no tiene "imagenPrompt": no hay a dónde colocar las imágenes.'); process.exit(1); }
    spritesCortados = await cortarHojas(arch, itemsImg, outDir);
    await previsualizarSprites(spritesCortados, path.join(outDir, 'sprites_previsualizacion.png'));
    itemsImg.forEach(i => { if (spritesCortados[i.n]) i.ref._sprite = spritesCortados[i.n]; });
    console.log(`🖼️  ${Object.keys(spritesCortados).length} ilustración(es) recortadas y colocadas · previsualización: ${path.join(outDir, 'sprites_previsualizacion.png')}`);
  } else if (itemsImg.length) console.warn(`⚠  ${itemsImg.length} imagen(es) pendientes (preguntas con "imagenPrompt"): genera el prompt con --prompt-imagenes y luego usa --sprites con las hojas de ChatGPT.`);
  const mc = reunirGrupos(ex.preguntas.filter(p => (p.tipo || 'multiple') === 'multiple')), ab = ex.preguntas.filter(p => p.tipo === 'abierta');
  mc.forEach(p => { if (p.estimulo) p._est = estMap[p.estimulo]; });
  mc.forEach((p, i) => p.num = i + 1); ab.forEach((p, i) => p.num = mc.length + i + 1);
  const imgs = {}; let nImg = 0;
  for (const p of ex.preguntas) {                      // PNG de alta resolución, una vez por pregunta
    const v = p.visual; if (!v || !(v.tipo === 1 || v.tipo === 3)) continue;
    const plantilla = PLANTILLAS[v.plantilla]; if (!plantilla) { console.error(`❌ P${p.num}: plantilla "${v.plantilla}" no existe. Disponibles: ${Object.keys(PLANTILLAS).join(', ')}`); process.exit(1); }
    const f = path.join(outDir, 'img', `pregunta_${String(p.num).padStart(2, '0')}.png`); await renderPNG(plantilla(v.params || {}), f); imgs[p.num] = f; nImg++;
  }
  const imgsEst = {};
  for (const e of (ex.estimulos || [])) {              // apoyos visuales de los textos compartidos
    const v = e.visual; if (!v || !(v.tipo === 1 || v.tipo === 3)) continue;
    const plantilla = PLANTILLAS[v.plantilla]; if (!plantilla) { console.error(`❌ Estímulo ${e.id}: plantilla "${v.plantilla}" no existe. Disponibles: ${Object.keys(PLANTILLAS).join(', ')}`); process.exit(1); }
    const f = path.join(outDir, 'img', `estimulo_${e.id}.png`); await renderPNG(plantilla(v.params || {}), f); imgsEst[e.id] = f; nImg++;
  }
  // La imagen se asocia al OBJETO pregunta (no al número), así sobrevive al reordenamiento
  const imgPorPregunta = new Map(); for (const p of ex.preguntas) if (imgs[p.num]) imgPorPregunta.set(p, imgs[p.num]);
  const imgsProxy = new Proxy({}, { get: (_, n) => { for (const [p, f] of imgPorPregunta) if (String(p.num) === String(n)) return f; } });
  const bubs = {}; for (const L of 'ABCD') { bubs[L] = path.join(outDir, 'img', `ovalo_${L}.png`); bubs['F' + L] = path.join(outDir, 'img', `ovalo_${L}_lleno.png`); await bubble(L, false, bubs[L]); await bubble(L, true, bubs['F' + L]); }
  const escudo = await prepararEscudo();

  const rapido = args.includes('--rapido');            // prueba rápida: sin medir maquetación (solo para comprobar contenido)
  const ctx = { ex, mc, ab, imgs: imgsProxy, imgsEst, bubs, escudo, outDir, empaquetar: !args.includes('--sin-ajuste') && !rapido };
  const niveles = nivelFijo !== null ? [nivelFijo] : [0, 1, 2];
  const optimizar = async (cx) => {
    if (rapido) { const r = await armar(cx, 1, false); r.nivel = 1; r.hojaPropia = false; r.an = null; console.log('⚡ Modo rápido: documentos generados sin medir la maquetación'); return r; }
    let mejor = null; const hayEst = cx.mc.some(p => p._est), modos = hayEst ? [false, true] : [false];
    const bueno = (r) => r && r.an && r.an.maxGap <= 55 && !r.an.faltan.length;
    for (const nv of niveles) for (const juntar of modos) for (const hojaPropia of [false, true]) {
      if (bueno(mejor)) break;
      if (hojaPropia && mejor && mejor.an && mejor.an.hojaMismaPagina && !mejor.an.faltan.length) continue;
      cx.juntarTodo = juntar;
      console.log(`🧩 Maquetando (compactación ${nv}${juntar ? ', textos compartidos sin partir' : ''}${hojaPropia ? ', hoja de respuestas en página propia' : ''})…`);
      const r = await armar(cx, nv, hojaPropia); const pdf = aPDF(r.fE, outDir);
      r.nivel = nv; r.hojaPropia = hojaPropia; r.juntar = juntar;
      if (pdf) {
        r.an = await analizarLayout(pdf, r.claveUlt, r.esperados); try { fs.unlinkSync(pdf); } catch (e) {}
        console.log(`   páginas: ${r.an.paginas} · hueco máx. al pie de columna: ${r.an.maxGap} pt · hoja en la misma página que las preguntas: ${r.an.hojaMismaPagina ? 'sí' : 'no'}${r.an.faltan.length ? ` · ⚠ faltan ${r.an.faltan.length} fragmento(s) de texto en el PDF` : ''}`);
      }
      const score = r.an ? r.an.maxGap + (r.an.hojaMismaPagina || hojaPropia ? 0 : 200) + 1000 * r.an.faltan.length : 999;
      r.score = score; if (!mejor || score < mejor.score) mejor = r;
    }
    cx.juntarTodo = mejor.juntar;
    if (mejor.an && mejor.an.faltan.length) console.warn(`⚠  Con ninguna maquetación se vio TODO el texto en la medición (faltan: ${mejor.an.faltan.slice(0, 3).join(' | ')}). Abre el Word y revisa esas preguntas.`);
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
    mejorB = await optimizar({ ...ctx, mc: aplanar(barajar(unidades(ctx.mc))), ab: barajar(ctx.ab), sufijo: '_VersionB', etiquetaVersion: 'B', seedExtra: 7919 });
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
