/**
 * Nico Roadmap API (v3: status, blocked-by, safe merge)
 * Reads and saves the whole plan for the roadmap app.
 *
 * Setup (once):
 * 1. In the Nico Roadmap sheet: Extensions > Apps Script. Delete what's there, paste this file, save.
 * 2. Project Settings (gear icon) > Script properties > Add property:
 *      Name: TOKEN    Value: any long password you make up
 * 3. Deploy > New deployment > type "Web app".
 *      Execute as: Me.  Who has access: Anyone.
 *    Authorise when asked, then copy the Web app URL.
 *
 * GET  <url>?token=XXX          returns the plan as JSON
 * POST <url>  body {token, plan} saves the whole plan
 */

const TABS = { projects: 'Projects', tasks: 'Tasks', blocks: 'Blocked days', settings: 'Settings', backlog: 'Backlog' };
const DAY_NUMS = ['1', '2', '3', '4', '5', '6', '0'];

function checkToken_(t) {
  const real = PropertiesService.getScriptProperties().getProperty('TOKEN');
  if (!real || t !== real) throw new Error('unauthorised');
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    checkToken_(e.parameter.token);
    return json_({ ok: true, plan: readPlan_() });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const body = JSON.parse(e.postData.contents);
    checkToken_(body.token);
    lock.waitLock(10000);
    const merged = body.base ? mergePlan_(body.base, body.plan, readPlan_()) : body.plan;
    writePlan_(merged);
    return json_({ ok: true, plan: merged, savedAt: new Date().toISOString() });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function rows_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, Math.max(sh.getLastColumn(), 9)).getDisplayValues();
}

function num_(v) { const n = parseFloat(String(v).replace('%', '')); return isNaN(n) ? 0 : n; }

function readPlan_() {
  const plan = { projects: {}, order: [], done: {}, blocks: {}, backlog: [], settings: { day: {}, eve: {}, realism: 80, stock: '' } };

  rows_(TABS.projects)
    .filter(r => r[0])
    .sort((a, b) => num_(a[4]) - num_(b[4]))
    .forEach(r => {
      plan.order.push(r[0]);
      const p = { name: r[1], tag: r[2], color: r[3], tasks: [] };
      if (r[5]) p.deadline = r[5];
      plan.projects[r[0]] = p;
    });

  rows_(TABS.tasks)
    .filter(r => r[0] && plan.projects[r[1]])
    .sort((a, b) => num_(a[2]) - num_(b[2]))
    .forEach(r => {
      const t = { id: r[0], name: r[3], h: num_(r[4]) };
      if (String(r[5]).toUpperCase() === 'TRUE') t.afterStock = true;
      const st = String(r[7] || '').toLowerCase();
      if (st === 'in progress') t.status = 'doing';
      if (st === 'blocked') t.status = 'blocked';
      if (r[8]) t.wait = r[8];
      plan.projects[r[1]].tasks.push(t);
      if (r[6]) plan.done[r[0]] = r[6];
    });

  rows_(TABS.blocks).filter(r => r[0] && r[1]).forEach(r => { plan.blocks[r[0]] = r[1]; });

  const set = rows_(TABS.settings);
  set.slice(0, 7).forEach(r => { plan.settings.day[r[1]] = num_(r[2]); plan.settings.eve[r[1]] = num_(r[3]); });
  set.forEach(r => {
    if (r[0] === 'Realism') plan.settings.realism = num_(r[1]);
    if (r[0] === 'Shirts arrive') plan.settings.stock = r[1];
  });

  rows_(TABS.backlog).forEach(r => { if (r[0]) plan.backlog.push(r[0]); });
  return plan;
}

function writeTable_(name, rows, width) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  const last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, width).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, width).setValues(rows);
}

function writePlan_(plan) {
  const txt = v => (v ? "'" + v : '');
  const order = (plan.order || []).filter(id => plan.projects[id]);
  Object.keys(plan.projects).forEach(id => { if (order.indexOf(id) < 0) order.push(id); });

  writeTable_(TABS.projects, order.map((id, i) => {
    const p = plan.projects[id];
    return [id, p.name || '', p.tag || '', p.color || '', i + 1, txt(p.deadline)];
  }), 6);

  const tasks = [];
  order.forEach(id => (plan.projects[id].tasks || []).forEach((t, j) => {
    const status = t.status === 'doing' ? 'In progress' : t.status === 'blocked' ? 'Blocked' : '';
    tasks.push([t.id, id, j + 1, t.name || '', +t.h || 0, !!t.afterStock, txt((plan.done || {})[t.id]), status, t.wait || '']);
  }));
  writeTable_(TABS.tasks, tasks, 9);
  const tsh = SpreadsheetApp.getActive().getSheetByName(TABS.tasks);
  if (tasks.length) tsh.getRange(2, 6, tasks.length, 1).insertCheckboxes();

  writeTable_(TABS.blocks, Object.keys(plan.blocks || {}).sort()
    .map(k => [txt(k), plan.blocks[k] === 'hilton' ? 'shift' : plan.blocks[k]]), 2);

  const s = plan.settings || {};
  const ssh = SpreadsheetApp.getActive().getSheetByName(TABS.settings);
  ssh.getRange(2, 3, 7, 2).setValues(DAY_NUMS.map(d => [+((s.day || {})[d]) || 0, +((s.eve || {})[d]) || 0]));
  ssh.getRange(11, 2, 2, 1).setValues([[(+s.realism || 80) / 100], [txt(s.stock)]]);

  writeTable_(TABS.backlog, (plan.backlog || []).map(b => [b]), 1);
}

// ---- 3-way merge: apply only what this device changed (base -> plan) onto the current Sheet ----
function same_(a, b) { return JSON.stringify(a) === JSON.stringify(b); }
function isObj_(v) { return v && typeof v === 'object' && !Array.isArray(v); }
function norm_(p) {
  if (!p) return p;
  const q = JSON.parse(JSON.stringify(p));
  Object.keys(q.projects || {}).forEach(id => {
    const ts = q.projects[id].tasks || [], m = {};
    ts.forEach(t => { m[t.id] = t; });
    q.projects[id].tasks = { o: ts.map(t => t.id), m: m };
  });
  return q;
}
function denorm_(q) {
  Object.keys(q.projects || {}).forEach(id => {
    const tk = q.projects[id].tasks || { o: [], m: {} }, m = tk.m || {}, o = (tk.o || []).filter(x => m[x]);
    Object.keys(m).forEach(x => { if (o.indexOf(x) < 0) o.push(x); });
    q.projects[id].tasks = o.map(x => m[x]);
  });
  return q;
}
function m3_(b, n, c) {
  if (same_(n, b)) return c;
  if (same_(c, b)) return n;
  if (isObj_(b) && isObj_(n) && isObj_(c)) {
    const out = {}, keys = {};
    [b, n, c].forEach(o => Object.keys(o).forEach(k => { keys[k] = 1; }));
    Object.keys(keys).forEach(k => { const v = m3_(b[k], n[k], c[k]); if (v !== undefined) out[k] = v; });
    return out;
  }
  return n;
}
function mergePlan_(base, plan, cur) {
  if (!base) return plan;
  return denorm_(m3_(norm_(base), norm_(plan), norm_(cur)));
}
