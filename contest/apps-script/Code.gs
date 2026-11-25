/**
 * منصة مسابقات المحقق العلمي — الخادم (Google Apps Script) — الإصدار 2
 * مدرسة أبي الدرداء المتوسطة
 *
 * طريقة التركيب (مرة واحدة فقط):
 * 1) أنشئ جدول Google Sheets جديدًا باسم «منصة مسابقات المحقق العلمي».
 * 2) من القائمة: الإضافات ← Apps Script، واحذف المحتوى وألصق هذا الملف كاملًا.
 * 3) غيّر قيمة ADMIN_KEY بالأسفل إلى رمز سري خاص بك (لا تشاركه مع الطلاب).
 * 4) نشر ← نشر جديد ← النوع: تطبيق ويب ← التنفيذ باسم: أنا ← من يمكنه الوصول: أي شخص.
 * 5) انسخ رابط تطبيق الويب (ينتهي بـ /exec) وضعه في لوحة المعلم ← الإعدادات.
 *
 * عند تعديل هذا الملف لاحقًا: نشر ← إدارة عمليات النشر ← تعديل ← إصدار جديد (حتى يبقى الرابط نفسه).
 */

const ADMIN_KEY = 'ضع-رمزك-السري-هنا';
const SPREADSHEET_ID = ''; // اتركه فارغًا إذا كان السكربت مربوطًا بالجدول نفسه

const SH = { settings: 'الإعدادات', comps: 'المسابقات', results: 'النتائج', assets: 'الصور' };
const CHUNK = 45000;
const RESULT_H = ['submissionId', 'year', 'compId', 'compTitle', 'grade', 'className', 'studentName', 'normName',
  'caseNos', 'caseTitles', 'score', 'maxScore', 'percentage', 'level', 'durationSec', 'startedAt', 'submittedAt',
  'answers', 'reflection', 'source'];
const COMP_H = ['id', 'year', 'title', 'status', 'updatedAt', 'createdAt', 'json'];
const ASSET_H = ['id', 'createdAt', 'data'];
const DEFAULT_SETTINGS = {
  year: '1448',
  schoolName: 'مدرسة أبي الدرداء المتوسطة',
  eduAdmin: 'الإدارة العامة للتعليم بمحافظة الأحساء',
  teacherName: 'علي طاهر البقاشي',
  principalName: 'فوزي محمد الهلال',
  slogan: 'معًا نصنع جيلًا مبدعًا في العلوم',
  logo: 'assets/logo.jpg'
};

/* ============ نقاط الدخول ============ */
function doGet(e) {
  const p = (e && e.parameter) || {};
  return respond_(route_(p), p.callback);
}
function doPost(e) {
  let p = {};
  try { p = JSON.parse(e.postData.contents); } catch (err) { p = (e && e.parameter) || {}; }
  return respond_(route_(p), null);
}
function respond_(obj, cb) {
  const txt = JSON.stringify(obj);
  if (cb && /^[\w$.]+$/.test(cb)) {
    return ContentService.createTextOutput(cb + '(' + txt + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(txt).setMimeType(ContentService.MimeType.JSON);
}

const PUBLIC_ACTIONS = {
  ping: () => ({ ok: true, version: 2, time: new Date().toISOString() }),
  publicSettings: () => ({ ok: true, settings: publicSettings_() }),
  listOpen: api_listOpen_,
  getComp: api_getComp_,
  check: api_check_,
  start: api_start_,
  submit: api_submit_,
  asset: api_asset_
};
const ADMIN_ACTIONS = {
  adminInit: api_adminInit_,
  getCompFull: api_getCompFull_,
  saveComp: api_saveComp_,
  setCompStatus: api_setCompStatus_,
  deleteComp: api_deleteComp_,
  uploadAsset: api_uploadAsset_,
  cleanAssets: api_cleanAssets_,
  results: api_results_,
  deleteResult: api_deleteResult_,
  deleteResults: api_deleteResults_,
  importResults: api_importResults_,
  saveSettings: api_saveSettings_,
  newYear: api_newYear_,
  backup: api_backup_
};

function route_(p) {
  try {
    const a = String(p.action || 'ping');
    if (PUBLIC_ACTIONS[a]) return PUBLIC_ACTIONS[a](p);
    if (ADMIN_ACTIONS[a]) {
      if (ADMIN_KEY === 'ضع-رمزك-السري-هنا') return { ok: false, error: 'غيّر قيمة ADMIN_KEY في ملف Code.gs ثم انشر إصدارًا جديدًا.' };
      if (String(p.key || '') !== ADMIN_KEY) return { ok: false, error: 'رمز المعلم غير صحيح.' };
      return ADMIN_ACTIONS[a](p);
    }
    return { ok: false, error: 'إجراء غير معروف: ' + a };
  } catch (err) {
    return { ok: false, error: String((err && err.message) || err) };
  }
}

/* ============ أدوات الجداول ============ */
function book_() {
  return SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}
function sheet_(name, headers) {
  const b = book_();
  let s = b.getSheetByName(name);
  if (!s) {
    s = b.insertSheet(name);
    s.getRange(1, 1, s.getMaxRows(), Math.max(headers.length, 26)).setNumberFormat('@');
    s.getRange(1, 1, 1, headers.length).setValues([headers]);
    s.setFrozenRows(1);
  }
  return s;
}
function values_(s) {
  const lr = s.getLastRow(), lc = s.getLastColumn();
  if (lr < 1 || lc < 1) return [];
  return s.getRange(1, 1, lr, lc).getValues();
}
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try { return fn(); } finally { lock.releaseLock(); }
}
function chunks_(str) {
  const out = [];
  for (let i = 0; i < str.length; i += CHUNK) out.push('~' + str.slice(i, i + CHUNK));
  return out.length ? out : ['~'];
}
function unchunk_(row, start) {
  let str = '';
  for (let i = start; i < row.length; i++) {
    const v = row[i];
    if (v === '' || v == null) break;
    str += String(v).replace(/^~/, '');
  }
  return str;
}
function writeRow_(s, r, row) {
  const lc = s.getLastColumn();
  s.getRange(r, 1, 1, row.length).setNumberFormat('@').setValues([row]);
  if (lc > row.length) s.getRange(r, row.length + 1, 1, lc - row.length).clearContent();
}
function safe_(v) { return String(v == null ? '' : v).replace(/^[=+\-@]+/, '').trim(); }
function uid_(prefix) { return prefix + Utilities.getUuid().replace(/-/g, '').slice(0, 12); }
function normName_(s) {
  return String(s || '').trim().toLowerCase()
    .replace(/[أإآ]/g, 'ا').replace(/ى/g, 'ي').replace(/ؤ/g, 'و').replace(/ئ/g, 'ي').replace(/ة/g, 'ه')
    .replace(/[ً-ْـ]/g, '').replace(/\s+/g, ' ');
}
function hash_(s) {
  let h = 2166136261;
  for (let i = 0; i < s.length; i++) { h ^= s.charCodeAt(i); h = Math.imul(h, 16777619); }
  return h >>> 0;
}
function level_(p) {
  if (p >= 90) return 'ممتاز';
  if (p >= 80) return 'جيد جدًا';
  if (p >= 70) return 'جيد';
  if (p >= 60) return 'مقبول';
  return 'يحتاج دعمًا';
}

/* ============ الإعدادات ============ */
function getSettings_() {
  const v = values_(sheet_(SH.settings, ['key', 'value']));
  const o = Object.assign({}, DEFAULT_SETTINGS);
  for (let i = 1; i < v.length; i++) if (v[i][0]) o[String(v[i][0])] = String(v[i][1]);
  return o;
}
function setSettings_(obj) {
  const s = sheet_(SH.settings, ['key', 'value']);
  const cur = Object.assign(getSettings_(), obj || {});
  const rows = Object.keys(cur).map(k => [k, safe_(cur[k])]);
  const lr = s.getLastRow();
  if (lr > 1) s.getRange(2, 1, lr - 1, 2).clearContent();
  if (rows.length) s.getRange(2, 1, rows.length, 2).setNumberFormat('@').setValues(rows);
  return cur;
}
function publicSettings_() {
  const s = getSettings_();
  return { year: s.year, schoolName: s.schoolName, eduAdmin: s.eduAdmin, teacherName: s.teacherName, slogan: s.slogan, logo: s.logo };
}

/* ============ المسابقات ============ */
function compsTable_() {
  const s = sheet_(SH.comps, COMP_H);
  return { s: s, v: values_(s) };
}
function compFromRow_(row, withConfig) {
  const c = { id: String(row[0]), year: String(row[1]), title: String(row[2]), status: String(row[3] || 'draft'),
    updatedAt: String(row[4]), createdAt: String(row[5]) };
  if (withConfig) {
    try { c.config = JSON.parse(unchunk_(row, 6) || '{}'); } catch (e) { c.config = {}; }
  }
  return c;
}
function allComps_(withConfig) {
  const t = compsTable_();
  const out = [];
  for (let i = 1; i < t.v.length; i++) if (t.v[i][0]) out.push(compFromRow_(t.v[i], withConfig));
  return out;
}
function findComp_(id) {
  const t = compsTable_();
  for (let i = 1; i < t.v.length; i++) if (String(t.v[i][0]) === String(id)) return Object.assign(compFromRow_(t.v[i], true), { _row: i + 1 });
  return null;
}
function maxScoreOf_(cfg, idxs) {
  return idxs.reduce((t, i) => t + (cfg.cases[i].questions || []).reduce((s, q) => s + (Number(q.points) || 1), 0), 0);
}
function publicCase_(k, i) {
  return {
    no: i + 1, title: k.title || '', shortName: k.shortName || '', image: k.image || '', intro: k.intro || '',
    questions: (k.questions || []).map(q => ({ text: q.text || '', options: q.options || [], points: Number(q.points) || 1,
      section: q.section || '', sub: !!q.sub, layout: q.layout || '' }))
  };
}
function publicComp_(c) {
  const cfg = c.config || {};
  const cases = cfg.cases || [];
  const allIdx = cases.map((_, i) => i);
  const meta = {
    id: c.id, year: c.year, status: c.status, title: cfg.title || c.title, subtitle: cfg.subtitle || '',
    description: cfg.description || '', grades: cfg.grades || [], classes: cfg.classes || [], mode: cfg.mode || 'single',
    attempts: Number(cfg.attempts) || 1, timeLimitMin: Number(cfg.timeLimitMin) || 0, shuffleOptions: !!cfg.shuffleOptions,
    allowReflection: cfg.allowReflection !== false, caseCount: cases.length, cover: cfg.cover || '',
    maxScore: cfg.mode === 'all' ? maxScoreOf_(cfg, allIdx) : (cases.length ? maxScoreOf_(cfg, [0]) : 0)
  };
  if (meta.mode === 'choose') meta.caseList = cases.map((k, i) => ({ no: i + 1, title: k.title, shortName: k.shortName, image: k.image }));
  if (cfg.rosterAutocomplete && cfg.rosters) meta.rosters = cfg.rosters;
  return meta;
}
function validateConfig_(cfg) {
  if (!cfg || typeof cfg !== 'object') throw new Error('بيانات المسابقة غير صالحة.');
  if (!String(cfg.title || '').trim()) throw new Error('اكتب عنوان المسابقة.');
  if (!Array.isArray(cfg.cases) || !cfg.cases.length) throw new Error('أضف نشاطًا واحدًا على الأقل.');
  if (!Array.isArray(cfg.classes) || !cfg.classes.length) throw new Error('حدّد الفصول المشاركة.');
  cfg.cases.forEach((k, i) => {
    if (!Array.isArray(k.questions) || !k.questions.length) throw new Error('النشاط ' + (i + 1) + ' لا يحتوي أسئلة.');
    k.questions.forEach((q, j) => {
      if (!Array.isArray(q.options) || q.options.length < 2) throw new Error('السؤال ' + (j + 1) + ' في النشاط ' + (i + 1) + ' يحتاج خيارين على الأقل.');
      const c = Number(q.correct);
      if (!(c >= 0 && c < q.options.length)) throw new Error('حدّد الإجابة الصحيحة للسؤال ' + (j + 1) + ' في النشاط ' + (i + 1) + '.');
    });
  });
}

function api_listOpen_() {
  const list = allComps_(true).filter(c => c.status === 'open').map(publicComp_);
  return { ok: true, settings: publicSettings_(), comps: list };
}
function api_getComp_(p) {
  const c = findComp_(p.id);
  if (!c || c.status === 'draft') return { ok: false, error: 'المسابقة غير موجودة أو لم تُنشر بعد.' };
  return { ok: true, settings: publicSettings_(), comp: publicComp_(c) };
}

/* ============ مشاركة الطالب ============ */
function resultsTable_() {
  const s = sheet_(SH.results, RESULT_H);
  return { s: s, v: values_(s) };
}
function rowToResult_(r) {
  const o = {};
  RESULT_H.forEach((h, i) => { o[h] = r[i] === undefined ? '' : r[i]; });
  ['score', 'maxScore', 'percentage', 'durationSec'].forEach(k => { o[k] = Number(o[k]) || 0; });
  o.year = String(o.year); o.className = String(o.className);
  return o;
}
function priorAttempts_(v, compId, className, norm) {
  let n = 0;
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][2]) === String(compId) && String(v[i][5]) === String(className) && String(v[i][7]) === norm) n++;
  }
  return n;
}
function classOf_(cfg, className) {
  return (cfg.classes || []).find(x => String(x.n) === String(className));
}
function assignment_(c, className, name, caseNo) {
  const cfg = c.config;
  const n = cfg.cases.length;
  const mode = cfg.mode || 'single';
  if (mode === 'all') return cfg.cases.map((_, i) => i);
  if (mode === 'choose') {
    const k = Number(caseNo) - 1;
    if (!(k >= 0 && k < n)) throw new Error('اختر النشاط أولًا.');
    return [k];
  }
  const norm = normName_(name);
  const roster = (cfg.rosters || {})[className];
  if (Array.isArray(roster)) {
    const idx = roster.map(normName_).indexOf(norm);
    if (idx >= 0) return [idx % n];
  }
  return [hash_(c.id + '|' + className + '|' + norm) % n];
}
function validateStudent_(c, p) {
  if (!c) throw new Error('المسابقة غير موجودة.');
  const name = safe_(p.studentName || p.name).replace(/\s+/g, ' ');
  if (name.length < 3) throw new Error('اكتب اسمك كاملًا.');
  const className = safe_(p.className);
  const cls = classOf_(c.config, className);
  if (!cls) throw new Error('اختر فصلك من القائمة.');
  return { name: name, className: className, grade: cls.g, norm: normName_(name) };
}
function api_check_(p) {
  const c = findComp_(p.compId);
  const st = validateStudent_(c, p);
  const t = resultsTable_();
  const used = priorAttempts_(t.v, c.id, st.className, st.norm);
  const allowed = Number(c.config.attempts) || 1;
  return { ok: true, submitted: used >= allowed, attemptsUsed: used, attemptsLeft: Math.max(0, allowed - used) };
}
function api_start_(p) {
  const c = findComp_(p.compId);
  if (!c || c.status === 'draft') return { ok: false, error: 'المسابقة غير متاحة.' };
  if (c.status !== 'open') return { ok: false, error: 'هذه المسابقة مغلقة حاليًا. راجع معلمك.' };
  const st = validateStudent_(c, p);
  const t = resultsTable_();
  const used = priorAttempts_(t.v, c.id, st.className, st.norm);
  const allowed = Number(c.config.attempts) || 1;
  if (used >= allowed) return { ok: true, submitted: true, attemptsUsed: used };
  const idxs = assignment_(c, st.className, st.name, p.caseNo);
  return {
    ok: true, submitted: false, attemptsUsed: used, attemptsLeft: allowed - used,
    cases: idxs.map(i => publicCase_(c.config.cases[i], i)),
    maxScore: maxScoreOf_(c.config, idxs), timeLimitMin: Number(c.config.timeLimitMin) || 0
  };
}
function api_submit_(p) {
  return withLock_(() => {
    const c = findComp_(p.compId);
    if (!c || c.status === 'draft') return { ok: false, error: 'المسابقة غير متاحة.' };
    const st = validateStudent_(c, p);
    const cfg = c.config;
    const t = resultsTable_();
    const sid = safe_(p.submissionId) || uid_('s');
    for (let i = 1; i < t.v.length; i++) {
      if (String(t.v[i][0]) === sid) return { ok: true, duplicate: true, result: buildFeedback_(c, rowToResult_(t.v[i])) };
    }
    const used = priorAttempts_(t.v, c.id, st.className, st.norm);
    const allowed = Number(cfg.attempts) || 1;
    if (used >= allowed) return { ok: false, alreadySubmitted: true, error: 'سبق تسليم مشاركتك، ولا توجد محاولات متبقية.' };

    let idxs = assignment_(c, st.className, st.name, p.caseNo);
    const given = Array.isArray(p.answers) ? p.answers : [];
    const answers = idxs.map(i => {
      const k = cfg.cases[i];
      const g = given.find(x => Number(x.caseNo) === i + 1) || { a: [] };
      const qs = k.questions.map((q, j) => {
        const raw = g.a ? g.a[j] : null;
        const a = raw === null || raw === undefined || raw === '' ? null : Number(raw);
        const ok = a !== null && a === Number(q.correct);
        return { a: a, c: ok, p: ok ? (Number(q.points) || 1) : 0 };
      });
      return { caseNo: i + 1, qs: qs };
    });
    const caseScores = answers.map(x => x.qs.reduce((s, q) => s + q.p, 0));
    const score = caseScores.reduce((s, v) => s + v, 0);
    const maxScore = maxScoreOf_(cfg, idxs);
    const pct = maxScore ? Math.round(score / maxScore * 100) : 0;
    const now = new Date();
    let started = new Date(p.startedAt || now);
    if (isNaN(started.getTime()) || started > now) started = now;
    const dur = Math.max(1, Math.round((now - started) / 1000));
    const row = {
      submissionId: sid, year: c.year, compId: c.id, compTitle: cfg.title || c.title, grade: st.grade,
      className: st.className, studentName: st.name, normName: st.norm,
      caseNos: idxs.map(i => i + 1).join('،'), caseTitles: idxs.map(i => cfg.cases[i].title).join('، '),
      score: score, maxScore: maxScore, percentage: pct, level: level_(pct), durationSec: dur,
      startedAt: started.toISOString(), submittedAt: now.toISOString(),
      answers: JSON.stringify(answers), reflection: safe_(p.reflection).slice(0, 2000), source: 'v2'
    };
    t.s.appendRow(RESULT_H.map(h => row[h]));
    return { ok: true, result: buildFeedback_(c, row) };
  });
}
function buildFeedback_(c, row) {
  const cfg = c.config;
  let answers = [];
  try { answers = typeof row.answers === 'string' ? JSON.parse(row.answers) : (row.answers || []); } catch (e) { answers = []; }
  const fb = {
    score: Number(row.score), maxScore: Number(row.maxScore), percentage: Number(row.percentage), level: row.level,
    studentName: row.studentName, className: row.className, durationSec: Number(row.durationSec),
    caseScores: answers.map(x => {
      const k = cfg.cases[x.caseNo - 1] || { title: '', questions: [] };
      return { no: x.caseNo, title: k.title, score: (x.qs || []).reduce((s, q) => s + (q.p || 0), 0), max: maxScoreOf_(cfg, [x.caseNo - 1]) };
    })
  };
  if (cfg.showReview) {
    fb.review = answers.map(x => {
      const k = cfg.cases[x.caseNo - 1];
      return { no: x.caseNo, title: k.title, items: k.questions.map((q, j) => ({
        text: q.text, chosen: x.qs[j] && x.qs[j].a !== null ? q.options[x.qs[j].a] : '', correct: q.options[q.correct], ok: !!(x.qs[j] && x.qs[j].c)
      })) };
    });
  }
  return fb;
}

/* ============ الصور ============ */
function api_asset_(p) {
  const v = values_(sheet_(SH.assets, ASSET_H));
  for (let i = 1; i < v.length; i++) if (String(v[i][0]) === String(p.id)) return { ok: true, id: p.id, dataUrl: unchunk_(v[i], 2) };
  return { ok: false, error: 'الصورة غير موجودة.' };
}
function api_uploadAsset_(p) {
  const d = String(p.dataUrl || '');
  if (!/^data:image\/(png|jpe?g|webp|gif);base64,/.test(d)) return { ok: false, error: 'صيغة الصورة غير مدعومة.' };
  if (d.length > 3000000) return { ok: false, error: 'الصورة كبيرة جدًا.' };
  return withLock_(() => {
    const s = sheet_(SH.assets, ASSET_H);
    const want = safe_(p.id);
    if (want) {
      const v = values_(s);
      for (let i = 1; i < v.length; i++) if (String(v[i][0]) === want) return { ok: true, id: want, ref: 'asset:' + want, existed: true };
    }
    const id = want || uid_('a');
    const row = [id, new Date().toISOString()].concat(chunks_(d));
    writeRow_(s, s.getLastRow() + 1, row);
    return { ok: true, id: id, ref: 'asset:' + id };
  });
}
function api_cleanAssets_() {
  return withLock_(() => {
    const used = JSON.stringify(allComps_(true).map(c => c.config)) + JSON.stringify(getSettings_());
    const s = sheet_(SH.assets, ASSET_H);
    const v = values_(s);
    let removed = 0;
    for (let i = v.length - 1; i >= 1; i--) {
      if (v[i][0] && used.indexOf('asset:' + v[i][0]) < 0) { s.deleteRow(i + 1); removed++; }
    }
    return { ok: true, removed: removed };
  });
}

/* ============ إدارة المعلم ============ */
function resultCounts_() {
  const v = resultsTable_().v;
  const byComp = {}, byYear = {};
  for (let i = 1; i < v.length; i++) {
    if (!v[i][0]) continue;
    byComp[v[i][2]] = (byComp[v[i][2]] || 0) + 1;
    byYear[v[i][1]] = (byYear[v[i][1]] || 0) + 1;
  }
  return { byComp: byComp, byYear: byYear, total: v.length > 1 ? v.length - 1 : 0 };
}
function api_adminInit_() {
  const counts = resultCounts_();
  const comps = allComps_(true).map(c => {
    const cfg = c.config || {};
    return { id: c.id, year: c.year, title: cfg.title || c.title, status: c.status, updatedAt: c.updatedAt, createdAt: c.createdAt,
      grades: cfg.grades || [], classes: cfg.classes || [], mode: cfg.mode || 'single', caseCount: (cfg.cases || []).length,
      results: counts.byComp[c.id] || 0 };
  });
  return { ok: true, settings: getSettings_(), comps: comps, years: counts.byYear, totalResults: counts.total };
}
function api_getCompFull_(p) {
  const c = findComp_(p.id);
  if (!c) return { ok: false, error: 'المسابقة غير موجودة.' };
  delete c._row;
  return { ok: true, comp: c };
}
function api_saveComp_(p) {
  const cfg = p.config;
  validateConfig_(cfg);
  return withLock_(() => {
    const t = compsTable_();
    const now = new Date().toISOString();
    let id = safe_(p.id), r = -1, createdAt = now;
    if (id) {
      for (let i = 1; i < t.v.length; i++) if (String(t.v[i][0]) === id) { r = i + 1; createdAt = String(t.v[i][5] || now); }
    } else id = uid_('c');
    const year = safe_(p.year) || getSettings_().year;
    const status = ['draft', 'open', 'closed'].indexOf(p.status) >= 0 ? p.status : (r > 0 ? String(t.v[r - 1][3]) : 'draft');
    const row = [id, year, safe_(cfg.title), status, now, createdAt].concat(chunks_(JSON.stringify(cfg)));
    writeRow_(t.s, r > 0 ? r : t.s.getLastRow() + 1, row);
    return { ok: true, id: id, year: year, status: status };
  });
}
function api_setCompStatus_(p) {
  if (['draft', 'open', 'closed'].indexOf(p.status) < 0) return { ok: false, error: 'حالة غير صالحة.' };
  return withLock_(() => {
    const c = findComp_(p.id);
    if (!c) return { ok: false, error: 'المسابقة غير موجودة.' };
    const s = compsTable_().s;
    s.getRange(c._row, 4, 1, 2).setValues([[p.status, new Date().toISOString()]]);
    return { ok: true };
  });
}
function api_deleteComp_(p) {
  return withLock_(() => {
    const c = findComp_(p.id);
    if (!c) return { ok: false, error: 'المسابقة غير موجودة.' };
    compsTable_().s.deleteRow(c._row);
    let removed = 0;
    if (p.withResults) removed = deleteResultsWhere_(r => String(r[2]) === String(p.id));
    return { ok: true, removedResults: removed };
  });
}
function api_results_(p) {
  const v = resultsTable_().v;
  const rows = [];
  for (let i = 1; i < v.length; i++) {
    if (!v[i][0]) continue;
    if (p.year && String(v[i][1]) !== String(p.year)) continue;
    if (p.compId && String(v[i][2]) !== String(p.compId)) continue;
    rows.push(rowToResult_(v[i]));
  }
  return { ok: true, rows: rows };
}
function api_deleteResult_(p) {
  return withLock_(() => {
    const n = deleteResultsWhere_(r => String(r[0]) === String(p.submissionId));
    return n ? { ok: true, removed: n } : { ok: false, error: 'لم يتم العثور على المشاركة.' };
  });
}
function deleteResultsWhere_(pred) {
  const t = resultsTable_();
  if (t.v.length < 2) return 0;
  const keep = [], cols = RESULT_H.length;
  let removed = 0;
  for (let i = 1; i < t.v.length; i++) {
    if (!t.v[i][0]) continue;
    if (pred(t.v[i])) removed++;
    else keep.push(RESULT_H.map((_, j) => t.v[i][j] === undefined ? '' : t.v[i][j]));
  }
  if (!removed) return 0;
  const lr = t.s.getLastRow();
  if (lr > 1) t.s.getRange(2, 1, lr - 1, Math.max(cols, t.s.getLastColumn())).clearContent();
  if (keep.length) t.s.getRange(2, 1, keep.length, cols).setNumberFormat('@').setValues(keep);
  return removed;
}
function api_deleteResults_(p) {
  return withLock_(() => {
    let pred;
    if (p.scope === 'all') pred = () => true;
    else if (p.scope === 'year') pred = r => String(r[1]) === String(p.year);
    else if (p.scope === 'beforeYear') pred = r => String(r[1]) !== String(p.year);
    else if (p.scope === 'comp') pred = r => String(r[2]) === String(p.compId);
    else return { ok: false, error: 'حدّد نطاق الحذف.' };
    return { ok: true, removed: deleteResultsWhere_(pred) };
  });
}
function api_importResults_(p) {
  const rows = Array.isArray(p.rows) ? p.rows : [];
  return withLock_(() => {
    const t = resultsTable_();
    const have = {};
    for (let i = 1; i < t.v.length; i++) have[String(t.v[i][0])] = 1;
    const add = [];
    rows.forEach(r => {
      const sid = safe_(r.submissionId) || uid_('s');
      if (have[sid]) return;
      have[sid] = 1;
      const o = Object.assign({}, r, {
        submissionId: sid, studentName: safe_(r.studentName), className: safe_(r.className),
        normName: normName_(r.studentName), answers: typeof r.answers === 'string' ? r.answers : JSON.stringify(r.answers || []),
        source: r.source || 'import'
      });
      add.push(RESULT_H.map(h => (o[h] === undefined || o[h] === null) ? '' : o[h]));
    });
    if (add.length) t.s.getRange(t.s.getLastRow() + 1, 1, add.length, RESULT_H.length).setNumberFormat('@').setValues(add);
    return { ok: true, added: add.length, skipped: rows.length - add.length };
  });
}
function api_saveSettings_(p) {
  return withLock_(() => ({ ok: true, settings: setSettings_(p.settings || {}) }));
}
function api_newYear_(p) {
  const year = safe_(p.year);
  if (!/^\d{4}$/.test(year)) return { ok: false, error: 'اكتب العام الدراسي الجديد بأربعة أرقام، مثل 1449.' };
  return withLock_(() => {
    const old = getSettings_().year;
    setSettings_({ year: year });
    let removed = 0, cloned = 0, closed = 0;
    if (p.deleteOldResults) removed = deleteResultsWhere_(r => String(r[1]) !== year);
    const t = compsTable_();
    const now = new Date().toISOString();
    for (let i = 1; i < t.v.length; i++) {
      const row = t.v[i];
      if (!row[0] || String(row[1]) === year) continue;
      if (p.closeOld && String(row[3]) === 'open') { t.s.getRange(i + 1, 4).setValue('closed'); closed++; }
    }
    (Array.isArray(p.cloneIds) ? p.cloneIds : []).forEach(id => {
      const c = findComp_(id);
      if (!c) return;
      const cfg = c.config;
      cfg.title = String(cfg.title || '').replace(new RegExp(old, 'g'), year);
      const row = [uid_('c'), year, safe_(cfg.title), 'draft', now, now].concat(chunks_(JSON.stringify(cfg)));
      writeRow_(t.s, t.s.getLastRow() + 1, row);
      cloned++;
    });
    return { ok: true, year: year, removedResults: removed, cloned: cloned, closed: closed };
  });
}
function api_backup_() {
  const comps = allComps_(true);
  const results = api_results_({}).rows;
  const av = values_(sheet_(SH.assets, ASSET_H));
  const assets = [];
  for (let i = 1; i < av.length; i++) if (av[i][0]) assets.push({ id: String(av[i][0]), dataUrl: unchunk_(av[i], 2) });
  return { ok: true, version: 2, createdAt: new Date().toISOString(), settings: getSettings_(), comps: comps, results: results, assets: assets };
}
