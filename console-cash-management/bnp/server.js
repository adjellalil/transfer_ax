// server.js — Cash Management Console
// Serveur Node.js minimaliste, modules natifs UNIQUEMENT (zéro dépendance npm).
'use strict';
const http = require('http');
const fs = require('fs');
const path = require('path');
const { spawn } = require('child_process');
const { URL } = require('url');

const APP_ROOT = __dirname;

// --- Configuration ---------------------------------------------------------
function loadConfig() {
  const defaults = {
    port: 3000, livrables_path: '../local/livrables', python_path: 'python',
    sources_originals_path: '../local/sources/originals', sources_work_path: '../local/sources/work',
    configuration_path: '../configuration', aggregators_path: '../local/catalog/aggregators.json',
    default_output_path: '../local/outputs',
    default_filename_template: '{YYYY}{MM}{DD}_{HH}{MI}_{NAME}'
  };
  try {
    return Object.assign(defaults, JSON.parse(fs.readFileSync(path.join(APP_ROOT, 'config.json'), 'utf8')));
  } catch (e) {
    console.warn('[config] config.json illisible — valeurs par défaut. (%s)', e.message);
    return defaults;
  }
}
let config = loadConfig();
const CONFIG_PATH = path.join(APP_ROOT, 'config.json');
const resolveCfg = (p) => path.resolve(APP_ROOT, p);

// --- Catalogue : les 3 JSON maitres de configuration/ (lus directement) -----
const configurationDir = () => resolveCfg(config.configuration_path || '../configuration');
const aggregatorsFile = () => resolveCfg(config.aggregators_path || '../local/catalog/aggregators.json');
function readJsonFile(fp) {
  try { return JSON.parse(fs.readFileSync(fp, 'utf8')); } catch (e) { return null; }
}
function loadConfiguration() {
  const dir = configurationDir();
  return {
    data: readJsonFile(path.join(dir, 'data.json')),
    sources: readJsonFile(path.join(dir, 'sources.json')),
    livrables: readJsonFile(path.join(dir, 'livrables.json')),
  };
}

function resolveLivrablesPath() {
  const primary = resolveCfg(config.livrables_path);
  if (fs.existsSync(primary) && fs.statSync(primary).isDirectory()) return primary;
  const fallback = path.resolve(APP_ROOT, '..', 'local', 'livrables');
  if (fs.existsSync(fallback) && fs.statSync(fallback).isDirectory()) {
    console.log('[livrables] "%s" introuvable → fallback : %s', config.livrables_path, fallback);
    return fallback;
  }
  return primary;
}
let LIVRABLES_PATH = resolveLivrablesPath();
let SOURCES_WORK = resolveCfg(config.sources_work_path);
let SOURCES_ORIG = resolveCfg(config.sources_originals_path);

function reloadPaths() {
  config = loadConfig();
  LIVRABLES_PATH = resolveLivrablesPath();
  SOURCES_WORK = resolveCfg(config.sources_work_path);
  SOURCES_ORIG = resolveCfg(config.sources_originals_path);
}

function saveConfig(next) {
  const merged = Object.assign({}, config, next);
  try {
    fs.writeFileSync(CONFIG_PATH, JSON.stringify(merged, null, 2), 'utf8');
    reloadPaths();
    return true;
  } catch (e) {
    return e.message;
  }
}

// --- Helpers ---------------------------------------------------------------
function sendJson(res, status, obj) {
  res.writeHead(status, { 'Content-Type': 'application/json; charset=utf-8' });
  res.end(JSON.stringify(obj));
}
const safeName = (s) =>
  typeof s === 'string' && s.length > 0 && !s.includes('/') && !s.includes('\\') && !s.includes('..');

function readBody(req) {
  return new Promise((resolve) => {
    let b = '';
    req.on('data', (c) => { b += c; if (b.length > 5e6) req.destroy(); });
    req.on('end', () => { try { resolve(JSON.parse(b || '{}')); } catch (e) { resolve(null); } });
  });
}

// Découpe une chaîne d'arguments CLI en argv (gère les guillemets).
function parseArgs(str) {
  if (!str) return [];
  const out = []; let cur = '', q = null, has = false;
  for (const c of str) {
    if (q) { if (c === q) q = null; else cur += c; has = true; }
    else if (c === '"' || c === "'") { q = c; has = true; }
    else if (/\s/.test(c)) { if (has) { out.push(cur); cur = ''; has = false; } }
    else { cur += c; has = true; }
  }
  if (has) out.push(cur);
  return out;
}

// Remplace les tokens de date/nom dans une valeur ({YYYY}{MM}{DD}{HH}{MI}{NAME}).
function resolveTokens(str, name) {
  const d = new Date();
  const p2 = (n) => String(n).padStart(2, '0');
  return String(str)
    .replace(/\{YYYY\}/g, d.getFullYear())
    .replace(/\{MM\}/g, p2(d.getMonth() + 1))
    .replace(/\{DD\}/g, p2(d.getDate()))
    .replace(/\{HH\}/g, p2(d.getHours()))
    .replace(/\{MI\}/g, p2(d.getMinutes()))
    .replace(/\{NAME\}/g, name || 'output');
}

// Construit argv depuis args (objet {flag:valeur} -> --flag valeur, ou chaîne libre).
function buildArgv(args) {
  if (typeof args === 'string') return parseArgs(args);
  if (!args || typeof args !== 'object') return [];
  const argv = [];
  const name = args.name || args.NAME || '';
  for (const [k, v] of Object.entries(args)) {
    if (k === 'name' || v === undefined || v === null || v === '') continue;
    const flag = k.startsWith('--') ? k : '--' + k.replace(/_/g, '-');
    argv.push(flag, resolveTokens(String(v), name));
  }
  return argv;
}

// Affichage seul : les chips "Sources requises" parsees des docstrings sont indicatives.
// La cartographie canonique faisant foi vit dans configuration/sources.json (champ "alias").
function normalizeSourceName(name) {
  if (!name || typeof name !== 'string') return name;
  return name.trim().toUpperCase();
}

// --- Parsing du docstring d'en-tête (titre/code/sections/arguments/DECOMPOSITION) ----
function parseDocstring(src) {
  const m = src.match(/"""([\s\S]*?)"""/);
  const lines = (m ? m[1] : '').split(/\r?\n/);
  let title = '', code = '';
  for (const ln of lines) {
    const t = ln.trim();
    if (!t || /^[-=*#]{3,}$/.test(t)) continue; // ignore lignes vides et séparateurs
    const cm = t.match(/\[([A-Z0-9]{5})\]/);
    if (cm) { code = cm[1]; title = t.slice(0, cm.index).trim() || t; } else { title = t; }
    break;
  }
  // Sections repérées par titre connu (tolère ':' final, style souligné OU indenté).
  const KNOWN = ['DESCRIPTION', 'SOURCES REQUISES', 'OUTPUTS PRODUITS', 'ARGUMENTS CLI', 'DECOMPOSITION'];
  const norm = (s) => s.trim().replace(/\s*:\s*$/, '').toUpperCase();
  const marks = [];
  lines.forEach((ln, i) => { if (KNOWN.includes(norm(ln))) marks.push([i, norm(ln)]); });
  const sections = {};
  marks.forEach(([i, name], k) => {
    const end = k + 1 < marks.length ? marks[k + 1][0] : lines.length;
    sections[name] = lines.slice(i + 1, end).filter((l) => !/^\s*[-=]{3,}\s*$/.test(l));
  });
  const args = [];
  for (const ln of (sections['ARGUMENTS CLI'] || [])) {
    const am = ln.trim().match(/^(?:[-*]\s+)?(--[\w-]+)\s+(.*?)\s*\((oblig\w*|opt\w*|requis|optionnel)\)\s*:?\s*(.*)$/i);
    if (am) args.push({ flag: am[1], metavar: am[2].trim(), required: /^(oblig|requis)/i.test(am[3]), description: (am[4] || '').trim() });
  }
  // Repli : si aucun argument structuré, collecte les --flags en début de ligne du docstring.
  if (!args.length) {
    const seen = new Set();
    for (const ln of lines) {
      const fm = ln.trim().match(/^(?:[-*]\s+)?(--[\w-]+)\b(.*)$/);
      if (fm && !seen.has(fm[1])) { seen.add(fm[1]); args.push({ flag: fm[1], metavar: '', required: false, description: fm[2].trim().replace(/^[:.\s]+/, '') }); }
    }
  }
  const decompositionText = (sections['DECOMPOSITION'] || []).map((l) => l.replace(/\s+$/, '')).join('\n').replace(/\n{3,}/g, '\n\n').trim();
  const tree = buildTree(sections['DECOMPOSITION'] || []);
  const description = (sections['DESCRIPTION'] || []).map((l) => l.trim()).filter(Boolean).join(' ');
  const sources = (sections['SOURCES REQUISES'] || []).map((l) => normalizeSourceName(extractSourceRef(l))).filter(Boolean);
  const outputs = (sections['OUTPUTS PRODUITS'] || []).map((l) => l.trim()).filter(Boolean);
  return { title, code, description, sources, outputs, arguments: args, tree, decompositionText };
}

function extractSourceRef(line) {
  const trimmed = line.trim();
  if (!trimmed) return '';
  const match = trimmed.match(/^([A-Z][A-Z0-9_\- ]*?)(?:\s{2,}|\s*[:(]|$)/);
  return match ? match[1].trim() : trimmed.split(/\s+/)[0];
}

function buildTree(lines) {
  const roots = []; const stack = [];
  for (const raw of lines) {
    const m = raw.replace(/\s+$/, '').match(/^(\s*)(\d+(?:\.\d+)*)[.)]?\s+(.*)$/);
    if (!m) continue;
    const depth = m[2].split('.').length;
    const node = { label: m[3].trim(), children: [] };
    while (stack.length && stack[stack.length - 1].depth >= depth) stack.pop();
    if (stack.length) stack[stack.length - 1].node.children.push(node); else roots.push(node);
    stack.push({ depth, node });
  }
  return roots;
}

// --- subprocess _explore.py (preview/filter/stats) -> JSON ------------------
function runExplore(exArgs) {
  return new Promise((resolve) => {
    const child = spawn(config.python_path, [path.join(APP_ROOT, '_explore.py'), ...exArgs], {
      cwd: APP_ROOT, env: { ...process.env, PYTHONIOENCODING: 'utf-8' }
    });
    let out = '', err = '';
    child.stdout.on('data', (d) => (out += d));
    child.stderr.on('data', (d) => (err += d));
    child.on('error', (e) => resolve({ error: `Python introuvable : ${e.message}` }));
    child.on('close', (code) => {
      if (out.trim()) { try { return resolve(JSON.parse(out)); } catch (e) { return resolve({ error: 'Sortie illisible' }); } }
      resolve({ error: err.trim() || `_explore.py code ${code}` });
    });
  });
}

// --- Endpoints lecture seule -----------------------------------------------
function handleLivrables(res) {
  let dirs;
  try { dirs = fs.readdirSync(LIVRABLES_PATH, { withFileTypes: true }); } catch (e) { return sendJson(res, 200, []); }
  const out = [];
  for (const d of dirs) {
    if (!d.isDirectory()) continue;
    const dir = path.join(LIVRABLES_PATH, d.name);
    let files = []; try { files = fs.readdirSync(dir); } catch (e) { /* ignore */ }
    const scripts = files.filter((f) => f.toLowerCase().endsWith('.py') && f !== '_shared.py' && f !== 'a_console.py').sort();
    let readme = null;
    const rp = path.join(dir, 'README.md');
    if (fs.existsSync(rp)) { try { readme = fs.readFileSync(rp, 'utf8'); } catch (e) { /* ignore */ } }
    out.push({ ref: d.name, scripts, readme });
  }
  out.sort((a, b) => a.ref.localeCompare(b.ref));
  sendJson(res, 200, out);
}

function handleCatalog(res) {
  const c = loadConfiguration();
  if (!c.sources) {
    return sendJson(res, 200, { error: 'configuration/ illisible (data/sources/livrables.json)',
      data: { datas: {} }, sources: { sources: {} }, livrables: { livrables: {} } });
  }
  sendJson(res, 200, c);
}

function handleAggregators(res) {
  const agg = readJsonFile(aggregatorsFile());
  if (!agg || !agg.aggregators) return sendJson(res, 200, []);
  const list = Object.entries(agg.aggregators).map(([ref, v]) => ({ ref, ...v }));
  sendJson(res, 200, list);
}

// Enregistre le chemin_local d'une source dans configuration/sources.json.
async function handleSourcePath(req, res) {
  const body = await readBody(req);
  if (!body || !body.ref) return sendJson(res, 400, { error: 'Paramètre "ref" manquant' });
  const fp = path.join(configurationDir(), 'sources.json');
  const doc = readJsonFile(fp);
  if (!doc || !doc.sources) return sendJson(res, 500, { error: 'sources.json illisible' });
  if (!doc.sources[body.ref]) return sendJson(res, 404, { error: `Source inconnue : ${body.ref}` });
  doc.sources[body.ref].chemin_local = String(body.chemin_local || '');
  try { fs.writeFileSync(fp, JSON.stringify(doc, null, 2), 'utf8'); }
  catch (e) { return sendJson(res, 500, { error: e.message }); }
  sendJson(res, 200, { ref: body.ref, chemin_local: doc.sources[body.ref].chemin_local });
}

// Valide le fichier reel d'une source : compare ses colonnes a celles attendues au catalogue.
async function handleSourceValidate(params, res) {
  const ref = params.get('ref');
  const c = loadConfiguration();
  const src = c.sources && c.sources.sources && c.sources.sources[ref];
  if (!src) return sendJson(res, 404, { error: `Source inconnue : ${ref}` });
  const p = (src.chemin_local || '').trim();
  const expected = (src.colonnes || []).map((col) => col.nom_dans_fichier);
  if (!p) return sendJson(res, 200, { ok: null, reason: 'Aucun chemin_local renseigné', expected });
  const prev = await runExplore(['--action', 'preview', '--path', p, '--limit', '1']);
  if (prev.error) return sendJson(res, 200, { ok: false, error: prev.error, expected });
  const actual = prev.columns || [];
  const norm = (s) => String(s).trim().toLowerCase();
  const aset = new Set(actual.map(norm)), eset = new Set(expected.map(norm));
  const missing = expected.filter((e) => !aset.has(norm(e)));
  const extra = actual.filter((a) => !eset.has(norm(a)));
  const namesMatch = expected.length > 0 && missing.length === 0;
  const countMatch = expected.length > 0 && actual.length === expected.length;
  sendJson(res, 200, {
    ok: expected.length === 0 ? null : (namesMatch && countMatch),
    reason: expected.length === 0 ? 'Colonnes non documentées au catalogue' : undefined,
    expected, actual, missing, extra,
    count_expected: expected.length, count_actual: actual.length,
    names_match: namesMatch, count_match: countMatch, total_rows: prev.total_rows,
  });
}

function handleConfig(res) {
  sendJson(res, 200, config);
}

async function handleConfigSave(req, res) {
  const body = await readBody(req);
  if (!body || typeof body !== 'object') return sendJson(res, 400, { error: 'JSON invalide' });
  const error = saveConfig(body);
  if (error !== true) return sendJson(res, 500, { error: error });
  sendJson(res, 200, config);
}

function handleScriptTree(params, res) {
  const livrable = params.get('livrable'), script = params.get('script');
  if (!safeName(livrable) || !safeName(script)) return sendJson(res, 400, { error: 'Paramètres invalides' });
  const sp = path.resolve(LIVRABLES_PATH, livrable, script);
  if (!sp.startsWith(path.resolve(LIVRABLES_PATH) + path.sep) || !fs.existsSync(sp))
    return sendJson(res, 404, { error: 'Script introuvable' });
  try { sendJson(res, 200, { script, ...parseDocstring(fs.readFileSync(sp, 'utf8')) }); }
  catch (e) { sendJson(res, 500, { error: e.message }); }
}

function handleBrowse(params, res) {
  const pathParam = params.get('path');
  let target = (pathParam && pathParam.trim()) ? pathParam : (fs.existsSync(SOURCES_WORK) ? SOURCES_WORK : APP_ROOT);
  let abs = path.resolve(target), stat;
  try { stat = fs.statSync(abs); } catch (e) { return sendJson(res, 404, { error: `Chemin introuvable : ${abs}` }); }
  if (!stat.isDirectory()) abs = path.dirname(abs);
  let entries;
  try { entries = fs.readdirSync(abs, { withFileTypes: true }); } catch (e) { return sendJson(res, 500, { error: e.message }); }
  const list = entries.map((d) => ({ name: d.name, type: d.isDirectory() ? 'dir' : 'file', full: path.join(abs, d.name) }))
    .sort((a, b) => (a.type !== b.type ? (a.type === 'dir' ? -1 : 1) : a.name.localeCompare(b.name)));
  const parent = path.dirname(abs);
  sendJson(res, 200, { path: abs, parent: parent === abs ? null : parent, entries: list });
}

// --- Exécution streamée (NDJSON) -------------------------------------------
function streamPython(res, scriptPath, argv, label) {
  res.writeHead(200, { 'Content-Type': 'application/x-ndjson; charset=utf-8', 'Cache-Control': 'no-cache', 'X-Accel-Buffering': 'no' });
  const send = (o) => { try { res.write(JSON.stringify(o) + '\n'); } catch (e) { /* socket fermé */ } };
  send({ t: 'meta', d: `▸ ${label} : ${config.python_path} ${path.basename(scriptPath)} ${argv.join(' ')}`.trim() });
  console.log('[run] %s %j', scriptPath, argv);
  let child;
  try {
    child = spawn(config.python_path, [scriptPath, ...argv], {
      cwd: APP_ROOT, env: { ...process.env, PYTHONIOENCODING: 'utf-8', PYTHONUNBUFFERED: '1' }
    });
  } catch (e) { send({ t: 'err', d: `Échec lancement Python : ${e.message}` }); send({ t: 'end', code: 2 }); return res.end(); }
  child.stdout.setEncoding('utf8'); child.stderr.setEncoding('utf8');
  child.stdout.on('data', (d) => send({ t: 'out', d }));
  child.stderr.on('data', (d) => send({ t: 'err', d }));
  child.on('error', (e) => send({ t: 'err', d: `Erreur process : ${e.message}` }));
  child.on('close', (code) => { send({ t: 'end', code }); res.end(); });
  res.on('close', () => { if (child && !child.killed) child.kill(); });
}

async function handleScriptRun(req, res) {
  const body = await readBody(req);
  if (!body) return sendJson(res, 400, { error: 'JSON invalide' });
  const { livrable, script, args } = body;
  if (!safeName(livrable) || !safeName(script) || !script.toLowerCase().endsWith('.py'))
    return sendJson(res, 400, { error: 'Paramètres livrable/script invalides' });
  const sp = path.resolve(LIVRABLES_PATH, livrable, script);
  if (!sp.startsWith(path.resolve(LIVRABLES_PATH) + path.sep)) return sendJson(res, 400, { error: 'Chemin hors périmètre' });
  if (!fs.existsSync(sp)) return sendJson(res, 404, { error: `Script introuvable : ${script}` });
  streamPython(res, sp, buildArgv(args), 'Lancement');
}

function handleScriptRead(params, res) {
  const livrable = params.get('livrable'), script = params.get('script');
  if (!safeName(livrable) || !safeName(script) || !script.toLowerCase().endsWith('.py'))
    return sendJson(res, 400, { error: 'Paramètres invalides' });
  const sp = path.resolve(LIVRABLES_PATH, livrable, script);
  if (!sp.startsWith(path.resolve(LIVRABLES_PATH) + path.sep) || !fs.existsSync(sp))
    return sendJson(res, 404, { error: 'Script introuvable' });
  try { sendJson(res, 200, { script, content: fs.readFileSync(sp, 'utf8') }); }
  catch (e) { sendJson(res, 500, { error: e.message }); }
}

async function handleScriptSave(req, res) {
  const body = await readBody(req);
  if (!body) return sendJson(res, 400, { error: 'JSON invalide' });
  const { livrable, script, content } = body;
  if (!safeName(livrable) || !safeName(script) || !script.toLowerCase().endsWith('.py') || typeof content !== 'string')
    return sendJson(res, 400, { error: 'Paramètres invalides' });
  const sp = path.resolve(LIVRABLES_PATH, livrable, script);
  if (!sp.startsWith(path.resolve(LIVRABLES_PATH) + path.sep)) return sendJson(res, 400, { error: 'Chemin hors périmètre' });
  try { fs.writeFileSync(sp, content, 'utf8'); sendJson(res, 200, { script, content }); }
  catch (e) { sendJson(res, 500, { error: e.message }); }
}

async function handleScriptCreate(req, res) {
  const body = await readBody(req);
  if (!body) return sendJson(res, 400, { error: 'JSON invalide' });
  const { livrable, script, content } = body;
  if (!safeName(livrable) || !safeName(script) || !script.toLowerCase().endsWith('.py') || typeof content !== 'string')
    return sendJson(res, 400, { error: 'Paramètres invalides' });
  const dir = path.resolve(LIVRABLES_PATH, livrable);
  if (!dir.startsWith(path.resolve(LIVRABLES_PATH) + path.sep) || !fs.existsSync(dir))
    return sendJson(res, 404, { error: 'Livrable introuvable' });
  const sp = path.resolve(dir, script);
  if (!sp.startsWith(dir + path.sep)) return sendJson(res, 400, { error: 'Nom de script invalide' });
  if (fs.existsSync(sp)) return sendJson(res, 409, { error: 'Script existe déjà' });
  try { fs.writeFileSync(sp, content, 'utf8'); sendJson(res, 200, { script, content }); }
  catch (e) { sendJson(res, 500, { error: e.message }); }
}

async function handleAggregatorRun(req, res) {
  const body = await readBody(req);
  if (!body) return sendJson(res, 400, { error: 'JSON invalide' });
  const ref = body.aggregator;
  const cat = readJsonFile(aggregatorsFile());
  if (!cat || !cat.aggregators) return sendJson(res, 500, { error: 'aggregators.json illisible' });
  const agg = cat.aggregators[ref];
  if (!agg) return sendJson(res, 404, { error: 'Agrégateur inconnu' });
  if (agg.status !== 'actif') return sendJson(res, 400, { error: `Agrégateur "${ref}" non disponible (statut: ${agg.status})` });
  const sp = path.resolve(APP_ROOT, '..', 'local', agg.script);
  if (!fs.existsSync(sp)) return sendJson(res, 404, { error: `Script agrégateur introuvable : ${agg.script}` });
  const argv = ['--input-folder', resolveCfg(path.join('..', 'local', agg.input_folder)),
    '--output-file', resolveCfg(path.join('..', 'local', agg.output_file))];
  streamPython(res, sp, argv, 'Agrégation');
}

// --- Serveur ---------------------------------------------------------------
const server = http.createServer(async (req, res) => {
  try {
    const u = new URL(req.url, 'http://localhost');
    const p = u.pathname, q = u.searchParams;
    if (req.method === 'GET' && (p === '/' || p === '/index.html')) {
      res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
      return res.end(fs.readFileSync(path.join(APP_ROOT, 'index.html')));
    }
    if (req.method === 'GET' && p === '/api/livrables') return handleLivrables(res);
    if (req.method === 'GET' && p === '/api/catalog') return handleCatalog(res);
    if (req.method === 'GET' && p === '/api/aggregators') return handleAggregators(res);
    if (req.method === 'GET' && p === '/api/script/tree') return handleScriptTree(q, res);
    if (req.method === 'GET' && p === '/api/script/read') return handleScriptRead(q, res);
    if (req.method === 'GET' && p === '/api/browse') return handleBrowse(q, res);
    if (req.method === 'GET' && p === '/api/config') return handleConfig(res);
    if (req.method === 'POST' && p === '/api/config') return handleConfigSave(req, res);
    if (req.method === 'POST' && p === '/api/source/path') return handleSourcePath(req, res);
    if (req.method === 'GET' && p === '/api/source/validate') return handleSourceValidate(q, res);
    if (req.method === 'POST' && p === '/api/script/save') return handleScriptSave(req, res);
    if (req.method === 'POST' && p === '/api/script/create') return handleScriptCreate(req, res);
    if (req.method === 'GET' && p === '/api/file/preview')
      return sendJson(res, 200, await runExplore(['--action', 'preview', '--path', q.get('path') || '', '--limit', q.get('limit') || '100']));
    if (req.method === 'POST' && p === '/api/file/filter') {
      const b = await readBody(req); if (!b) return sendJson(res, 400, { error: 'JSON invalide' });
      return sendJson(res, 200, await runExplore(['--action', 'filter', '--path', b.path || '', '--limit', String(b.limit || 100), '--filters', JSON.stringify(b.filters || [])]));
    }
    if (req.method === 'POST' && p === '/api/file/stats') {
      const b = await readBody(req); if (!b) return sendJson(res, 400, { error: 'JSON invalide' });
      return sendJson(res, 200, await runExplore(['--action', 'stats', '--path', b.path || '', '--column', b.column || '']));
    }
    if (req.method === 'POST' && p === '/api/script/run') return handleScriptRun(req, res);
    if (req.method === 'POST' && p === '/api/aggregator/run') return handleAggregatorRun(req, res);
    sendJson(res, 404, { error: 'Route inconnue' });
  } catch (e) {
    console.error('[erreur]', e);
    if (!res.headersSent) sendJson(res, 500, { error: e.message }); else { try { res.end(); } catch (_) { /* ignore */ } }
  }
});

const PORT = process.env.PORT ? Number(process.env.PORT) : config.port;
server.on('error', (e) => {
  if (e.code === 'EADDRINUSE') console.error('[erreur] Port %s déjà utilisé. Fermez l\'app qui l\'occupe ou changez "port" dans config.json.', PORT);
  else console.error('[erreur serveur]', e.message);
  process.exit(1);
});
server.listen(PORT, () => {
  console.log('=== Cash Management Console ===');
  console.log('Livrables : %s', LIVRABLES_PATH);
  console.log('Sources   : %s', SOURCES_WORK);
  console.log('Python    : %s', config.python_path);
  console.log('Serveur prêt → http://localhost:%s', PORT);
});
