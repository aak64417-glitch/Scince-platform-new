/* أدوات مشتركة — منصة مسابقات المحقق العلمي */
(function () {
  const qs = new URLSearchParams(location.search);
  const cfg = window.CONTEST_CONFIG || {};

  function getApi() {
    return (qs.get('api') || cfg.api || localStorage.getItem('contest_api') || '').trim();
  }

  async function api(action, data, opts) {
    const url = (opts && opts.url) || getApi();
    if (!url) throw new Error('لم يتم ضبط رابط الخادم بعد.');
    const ctrl = new AbortController();
    const timer = setTimeout(() => ctrl.abort(), (opts && opts.timeout) || 45000);
    try {
      const r = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify(Object.assign({ action: action }, data || {})),
        signal: ctrl.signal,
        redirect: 'follow'
      });
      const txt = await r.text();
      let j;
      try { j = JSON.parse(txt); } catch (e) { throw new Error('استجابة غير متوقعة من الخادم. تأكد من نشر السكربت بصلاحية «أي شخص».'); }
      return j;
    } catch (e) {
      if (e.name === 'AbortError') throw new Error('انتهت مهلة الاتصال بالخادم.');
      if (e instanceof TypeError) throw new Error('تعذر الاتصال بالخادم. تحقق من الإنترنت.');
      throw e;
    } finally { clearTimeout(timer); }
  }

  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, m => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[m]));
  }
  function normName(s) {
    return String(s || '').trim().toLowerCase()
      .replace(/[أإآ]/g, 'ا').replace(/ى/g, 'ي').replace(/ؤ/g, 'و').replace(/ئ/g, 'ي').replace(/ة/g, 'ه')
      .replace(/[ً-ْـ]/g, '').replace(/\s+/g, ' ');
  }
  function hash(s) {
    let h = 2166136261;
    for (let i = 0; i < s.length; i++) { h ^= s.charCodeAt(i); h = Math.imul(h, 16777619); }
    return h >>> 0;
  }
  const LEVELS = [
    { min: 90, name: 'ممتاز', cls: 'excellent', badge: 'المحقق الذهبي', icon: '🥇', color: '#1f9d63' },
    { min: 80, name: 'جيد جدًا', cls: 'verygood', badge: 'المحقق الفضي', icon: '🥈', color: '#2f7fb8' },
    { min: 70, name: 'جيد', cls: 'good', badge: 'المحقق البرونزي', icon: '🥉', color: '#b8860b' },
    { min: 60, name: 'مقبول', cls: 'pass', badge: 'محقق مبتدئ', icon: '🔎', color: '#c96a12' },
    { min: 0, name: 'يحتاج دعمًا', cls: 'support', badge: 'محقق تحت التدريب', icon: '🧭', color: '#c0392b' }
  ];
  function level(p) { p = Number(p) || 0; return LEVELS.find(l => p >= l.min); }
  const GRADES = { 1: 'الأول المتوسط', 2: 'الثاني المتوسط', 3: 'الثالث المتوسط' };
  function gradeName(g) { return GRADES[g] ? 'الصف ' + GRADES[g] : 'الصف ' + g; }
  function gradesLabel(arr) {
    arr = (arr || []).map(Number).sort();
    if (arr.length === 3) return 'جميع صفوف المتوسط';
    return arr.map(gradeName).join(' و');
  }
  const MODES = { single: 'نشاط واحد موزّع لكل طالب', all: 'جميع الأنشطة لكل طالب', choose: 'الطالب يختار نشاطه' };

  const assetCache = {};
  async function resolveImage(ref) {
    if (!ref) return '';
    if (!String(ref).startsWith('asset:')) return ref;
    const id = String(ref).slice(6);
    if (assetCache[id]) return assetCache[id];
    try { const c = sessionStorage.getItem('asset_' + id); if (c) return (assetCache[id] = c); } catch (e) { }
    const j = await api('asset', { id: id });
    if (!j.ok) return '';
    assetCache[id] = j.dataUrl;
    try { sessionStorage.setItem('asset_' + id, j.dataUrl); } catch (e) { }
    return j.dataUrl;
  }
  async function hydrateImages(root) {
    const imgs = (root || document).querySelectorAll('img[data-ref]');
    await Promise.all([...imgs].map(async img => {
      const src = await resolveImage(img.dataset.ref).catch(() => '');
      if (src) img.src = src; else img.closest('.img-holder') ? img.closest('.img-holder').classList.add('noimg') : img.remove();
      img.removeAttribute('data-ref');
    }));
  }
  function fmtDate(iso, withTime) {
    if (!iso) return '';
    const d = new Date(iso);
    if (isNaN(d)) return String(iso);
    const o = { timeZone: 'Asia/Riyadh', year: 'numeric', month: 'long', day: 'numeric' };
    if (withTime) Object.assign(o, { hour: 'numeric', minute: '2-digit' });
    try { return d.toLocaleString('ar-SA-u-ca-islamic-umalqura-nu-latn', o); } catch (e) { return d.toLocaleString(); }
  }
  function secs(s) { s = Math.round(+s || 0); const m = Math.floor(s / 60), r = s % 60; return m ? (r ? `${m} دقيقة و${r} ثانية` : `${m} دقيقة`) : `${r} ثانية`; }

  window.C = { qs, getApi, api, esc, normName, hash, LEVELS, level, gradeName, gradesLabel, GRADES, MODES, resolveImage, hydrateImages, fmtDate, secs };
})();
