#!/usr/bin/env node
/* probar-banco.js — Prueba el ciclo completo del banco de preguntas en una carpeta temporal:
 * guardar (sin duplicar, aunque la pregunta traiga su propio id) → listar → armar → excluir → exigir «aprobada» → construir el Word.
 * Uso: node herramientas/mantenimiento/probar-banco.js   (requiere haber instalado las dependencias del prompt de pruebas) */
const fs = require('fs'), os = require('os'), path = require('path'), { spawnSync } = require('child_process');
const raiz = path.join(__dirname, '..', '..'), H = (f) => path.join(raiz, 'herramientas', f);
const PROMPT = path.join(raiz, 'prompts/evaluacion/creacion-pruebas-saber');
const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'banco-')), banco = path.join(tmp, 'banco');
const run = (script, a, opts = {}) => spawnSync('node', [script, ...a], { encoding: 'utf8', ...opts });
const E = []; const chequea = (ok, msg) => { if (!ok) E.push(msg); };

const ejemplo = JSON.parse(fs.readFileSync(path.join(PROMPT, 'examples/exam_ejemplo.json'), 'utf8'));
ejemplo.preguntas.forEach((p, i) => { p.id = i + 1; });                    // caso de riesgo: preguntas que traen su propio id numérico
const fuente = path.join(tmp, 'exam.json'); fs.writeFileSync(fuente, JSON.stringify(ejemplo));

// 1. guardar dos veces: la segunda no debe duplicar
let r = run(H('guardar-en-banco.js'), [fuente, '--banco', banco, '--estado', 'aprobada']); chequea(r.status === 0 && /16 nuevas/.test(r.stdout), 'guardar: debía guardar 16 preguntas nuevas\n' + r.stdout + r.stderr);
r = run(H('guardar-en-banco.js'), [fuente, '--banco', banco, '--estado', 'aprobada']); chequea(/0 nuevas, 16 repetidas/.test(r.stdout), 'guardar: la segunda vez no debe duplicar (¿el id de la prueba pisó la huella?)\n' + r.stdout);
const guardado = JSON.parse(fs.readFileSync(path.join(banco, 'matematicas/5/fracciones.json'), 'utf8'));
chequea(guardado.preguntas.every(p => /^[0-9a-f]{10}$/.test(p.id) && p.idEnLaPrueba !== undefined), 'guardar: cada pregunta debe tener id de 10 caracteres hex y conservar idEnLaPrueba');

// 2. listar y armar
r = run(H('armar-desde-banco.js'), ['--listar', '--banco', banco]); chequea(/matematicas\/5\/fracciones: 0 \/ 0 \/ 16/.test(r.stdout), 'listar: debía mostrar 0 / 0 / 16\n' + r.stdout);
const out1 = path.join(tmp, 'e1.json'), out2 = path.join(tmp, 'e2.json'), out3 = path.join(tmp, 'e3.json');
const base = ['--area', 'matematicas', '--grado', '5', '--n', '8', '--abiertas', '1', '--banco', banco];
r = run(H('armar-desde-banco.js'), [...base, '--semilla', '3', '--out', out1]); chequea(r.status === 0, 'armar: falló\n' + r.stderr);
run(H('armar-desde-banco.js'), [...base, '--semilla', '3', '--out', out2]);
const e1 = JSON.parse(fs.readFileSync(out1, 'utf8')), e2 = JSON.parse(fs.readFileSync(out2, 'utf8'));
chequea(e1.preguntas.filter(p => p.tipo !== 'abierta').length === 8 && e1.preguntas.filter(p => p.tipo === 'abierta').length === 1, 'armar: debía haber 8 de selección múltiple y 1 abierta');
chequea(JSON.stringify(e1.origenBanco.ids) === JSON.stringify(e2.origenBanco.ids), 'armar: con la misma semilla el resultado debe ser igual');
chequea(new Set(e1.origenBanco.ids).size === e1.origenBanco.ids.length, 'armar: hay preguntas repetidas');
chequea(new Set(e1.preguntas.filter(p => p.tipo !== 'abierta').map(p => p.competencia)).size >= 3, 'armar: debía repartir entre las 3 competencias');
chequea(e1.preguntas.every(p => !('estado' in p) && !('fechaRegistro' in p) && !('idEnLaPrueba' in p)), 'armar: el exam.json no debe llevar campos internos del banco');

// 3. excluir lo ya usado: quedan 6 de 14 → pedir 6 funciona y no repite; pedir 8 falla
r = run(H('armar-desde-banco.js'), ['--area', 'matematicas', '--grado', '5', '--n', '6', '--banco', banco, '--excluir', out1, '--out', out3]); chequea(r.status === 0, 'excluir: pedir 6 debía funcionar\n' + r.stderr);
if (r.status === 0) { const usados = new Set(e1.origenBanco.ids); chequea(JSON.parse(fs.readFileSync(out3, 'utf8')).origenBanco.ids.every(i => !usados.has(i)), 'excluir: repitió preguntas de la prueba anterior'); }
r = run(H('armar-desde-banco.js'), ['--area', 'matematicas', '--grado', '5', '--n', '8', '--banco', banco, '--excluir', out1, '--out', out3]); chequea(r.status === 1 && /Solo hay 6/.test(r.stderr), 'excluir: pedir 8 debía fallar con un mensaje claro');

// 4. solo preguntas aprobadas por defecto
const bancoBorrador = path.join(tmp, 'banco-borrador'); run(H('guardar-en-banco.js'), [fuente, '--banco', bancoBorrador]);
r = run(H('armar-desde-banco.js'), ['--area', 'matematicas', '--grado', '5', '--banco', bancoBorrador, '--out', out3]); chequea(r.status === 1 && /aprobada/.test(r.stderr), 'estado: con preguntas en borrador debía negarse');
r = run(H('armar-desde-banco.js'), ['--area', 'matematicas', '--grado', '5', '--n', '5', '--estado', 'borrador', '--banco', bancoBorrador, '--out', out3]); chequea(r.status === 0, 'estado: --estado borrador debía permitirlo');

// 5. la prueba armada debe construirse en Word
r = run(path.join(PROMPT, 'scripts/build_exam.js'), [out1, '--out', path.join(tmp, 'salida'), '--rapido'], { cwd: PROMPT });
const docx = fs.existsSync(path.join(tmp, 'salida')) ? fs.readdirSync(path.join(tmp, 'salida')).filter(f => f.endsWith('.docx')) : [];
chequea(r.status === 0 && docx.length === 2, 'build_exam: la prueba armada desde el banco no generó los dos Word\n' + (r.stderr || '').slice(0, 500));

fs.rmSync(tmp, { recursive: true, force: true });
if (E.length) { console.error(E.map(e => '❌ ' + e).join('\n')); process.exit(1); }
console.log('✅ Banco de preguntas: guarda sin duplicar, arma pruebas equilibradas y repetibles, excluye lo usado, exige «aprobada» y genera el Word.');
