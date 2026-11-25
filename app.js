'use strict';
// Files from outside (Excel, zip) are read by a library; freezing the base object prototypes blocks the known
// "prototype pollution" attacks through crafted files (SheetJS 0.18.5, CVE-2023-30533).
Object.freeze(Object.prototype); Object.freeze(Array.prototype);
// the layout of the files exchanged with the Excel workbook; raised when the layout changes in a way the other side must know
const FILE_FORMAT = 1;
/* =============== defaults (from the budget workbook) =============== */
const DEFAULTS = {"cfg": {"groups": [], "cats": [], "pays": ["מזומן", "ביט / פייבוקס", "עו\"ש (הו\"ק / העברות)"], "people": [], "me": "", "senders": {}, "defPay": {}, "start": "", "dateOrder": "auto"}, "kw": []};
const KEY = 'kesef-shelanu-v1';

/* =============== helpers =============== */
const $ = (s, r = document) => r.querySelector(s);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const num = v => { const n = parseFloat(String(v ?? '').replace(/,/g, '').replace(/[^\d.\-]/g, '')); return isFinite(n) ? n : 0; };
const r2 = n => Math.round((+n || 0) * 100) / 100;
const fmt = (n, dec) => { n = r2(n); const d = dec ?? (Number.isInteger(n) ? 0 : 2); return n.toLocaleString('he-IL', {minimumFractionDigits: d, maximumFractionDigits: d}); };
const ils = (n, dec) => `${fmt(n, dec)} ₪`;
const pad = n => String(n).padStart(2, '0');
const todayStr = () => { const d = new Date(); return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`; };
const curMonth = () => todayStr().slice(0, 7);
const addM = (m, k) => { let [y, mo] = m.split('-').map(Number); mo += k; while (mo > 12) { mo -= 12; y++; } while (mo < 1) { mo += 12; y--; } return `${y}-${pad(mo)}`; };
const mDiff = (a, b) => { const [y1, m1] = a.split('-').map(Number), [y2, m2] = b.split('-').map(Number); return (y2 - y1) * 12 + (m2 - m1); };
const HEB_M = ['ינואר','פברואר','מרץ','אפריל','מאי','יוני','יולי','אוגוסט','ספטמבר','אוקטובר','נובמבר','דצמבר'];
const mLabel = m => { const [y, mo] = m.split('-').map(Number); return `${HEB_M[mo-1]} ${String(y).slice(2)}`; };
const mName = m => HEB_M[+m.split('-')[1] - 1];
const sheetName = m => { const [y, mo] = m.split('-').map(Number); return `${mo}.${String(y).slice(2)}`; };
const dShort = ds => { const [y, m, d] = ds.split('-'); return `${d}.${m}.${y.slice(2)}`; };
const ddmm = ds => { const [, m, d] = ds.split('-'); return `${d}/${m}`; };
const daysIn = m => { const [y, mo] = m.split('-').map(Number); return new Date(y, mo, 0).getDate(); };
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
const clone = o => JSON.parse(JSON.stringify(o));
// the Excel library is big, so it is loaded only when a file is read or written (from the phone's copy, never from the internet)
let xlsxP = null;
function loadXLSX() {
  if (window.XLSX) return Promise.resolve(window.XLSX);
  return xlsxP = xlsxP || new Promise((res, rej) => { const s = document.createElement('script'); s.src = 'lib/xlsx.full.min.js'; s.onload = () => res(window.XLSX); s.onerror = () => { xlsxP = null; rej(new Error('לא הצלחתי לטעון את רכיב האקסל')); }; document.head.appendChild(s); });
}
function toast(t, ms, undo) {
  const el = $('#toast'); el.dataset.t = t; toast.at = Date.now();
  el.innerHTML = '<span>' + esc(t) + '</span>' + (undo ? '<button class="undo" data-act="undo">ביטול</button>' : '');
  el.classList.toggle('act', !!undo); el.classList.add('on');
  clearTimeout(toast.t); toast.t = setTimeout(() => el.classList.remove('on', 'act'), ms || 2400);
}

const I = {
  menu: '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"><path d="M4 7h16M4 12h16M4 17h16"/></svg>',
  r: '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><path d="M9 6l6 6-6 6"/></svg>',
  l: '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><path d="M15 6l-6 6 6 6"/></svg>',
  arrowL: '<svg class="arrow" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M19 12H5M11 6l-6 6 6 6"/></svg>',
  plus: '<svg width="26" height="26" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.6" stroke-linecap="round"><path d="M12 5v14M5 12h14"/></svg>',
  home: '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linejoin="round"><path d="M4 10.5L12 4l8 6.5V19a1 1 0 01-1 1h-4v-5h-6v5H5a1 1 0 01-1-1z"/></svg>',
  flow: '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="6" width="18" height="12" rx="3"/><circle cx="12" cy="12" r="2.5"/></svg>',
  check: '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="4" y="4" width="16" height="16" rx="4"/><path d="M8.5 12.5l2.5 2.5 5-6"/></svg>',
  star: '<svg width="34" height="34" viewBox="0 0 24 24"><path fill="#E98E77" d="M12 1c.9 6.4 3.6 9.1 10 11-6.4 1.9-9.1 4.6-10 11-.9-6.4-3.6-9.1-10-11 6.4-1.9 9.1-4.6 10-11z"/></svg>',
  hour: '<svg width="34" height="34" viewBox="0 0 24 24"><path fill="#C9B3EA" d="M3 3h18c0 4-4 6.5-6.5 9C17 14.5 21 17 21 21H3c0-4 4-6.5 6.5-9C7 9.5 3 7 3 3z"/></svg>',
  ast: '<svg width="34" height="34" viewBox="0 0 24 24" fill="#8FB6E3">' + [0,30,60,90,120,150].map(a => `<rect x="10.4" y="1" width="3.2" height="22" rx="1.6" transform="rotate(${a} 12 12)"/>`).join('') + '</svg>',
  moon: '<svg width="13" height="13" viewBox="0 0 24 24"><path fill="#F5C518" d="M20 15A9 9 0 019 4a9 9 0 1011 11z"/></svg>',
  flower: '<svg width="180" height="180" viewBox="0 0 100 100">' + [0,40,80,120,160,200,240,280,320].map(a => `<ellipse cx="50" cy="22" rx="11" ry="22" fill="#FBE9E1" transform="rotate(${a} 50 50)"/>`).join('') + '<circle cx="50" cy="50" r="12" fill="#FBE9E1"/><path d="M44 50l4 4 8-9" stroke="#1B1E3A" stroke-width="2.5" fill="none" stroke-linecap="round"/></svg>',
};
const SRC_ICON = {chat: '💬', fixed: '🔁', manual: '✍️', income: '💰'};
const G_BG = ['#E3F5EF', '#EEEAFD', '#FFEDE3', '#FDEAF1', '#E4F0FB', '#FFF3D6', '#EAF4E0', '#F1EFF4'];
const G_INK = ['#1E6B5C', '#4A3696', '#8A4320', '#8E2453', '#1F5A99', '#7A5A00', '#3D6B1E', '#22243F'];
const G_BAR = ['#8FD5C6', '#B3A7F0', '#FFBE9C', '#F4A9C6', '#9FC5EC', '#F2D27A', '#B5DB94', '#C3C2D4'];
const G_ICON = [
  [/שוטפ|מזון|סופר|קניות|מכולת/, '<path d="M3 4h2l2.4 11h10.2L20 8H7"></path><path d="M10 20h.01M17 20h.01"></path>'],
  [/דיור|בית|משכנתא|שכירות/, '<path d="M4 11l8-6 8 6v8a1 1 0 01-1 1h-5v-5h-4v5H5a1 1 0 01-1-1z"></path>'],
  [/רכב|תחבורה|דלק/, '<path d="M5 16V11l2-5h10l2 5v5"></path><path d="M4 16h16M7.5 16v2M16.5 16v2"></path>'],
  [/ביטוח/, '<path d="M12 3l7 3v5c0 4.5-3 8-7 10-4-2-7-5.5-7-10V6z"></path>'],
  [/מתנ|תרומ|צדק/, '<path d="M4 10h16v10H4zM3 7h18v3H3zM12 7v13"></path><path d="M12 7c-2-3-5-3-5-1s3 1 5 1c2 0 5 1 5-1s-3-2-5 1"></path>'],
  [/מס|עמל|הלווא|ריבית/, '<path d="M6 3h12v18l-3-2-3 2-3-2-3 2z"></path><path d="M9 8h6M9 12h6"></path>'],
  [/חסכו|השקע/, '<path d="M4 15c0-4 3.6-7 8-7s8 3 8 7-3.6 5-8 5-8-1-8-5z"></path><path d="M12 8V4M9 13h6"></path>'],
  [/הכנס|משכור|עבודה/, '<path d="M4 7h16v12H4z"></path><path d="M16 13h.01M4 10h16"></path>'],
  [/ילד|חינוך|לימוד|גן|חוג/, '<path d="M3 8l9-4 9 4-9 4z"></path><path d="M7 10v5c3 2 7 2 10 0v-5"></path>'],
  [/בריאות|רפוא|טיפול/, '<path d="M12 20s-7-4.5-7-10a4 4 0 017-2.5A4 4 0 0119 10c0 5.5-7 10-7 10z"></path>'],
  [/פנאי|בילוי|חופש|תרבות/, '<circle cx="12" cy="12" r="4"></circle><path d="M12 2v2M12 20v2M2 12h2M20 12h2M5 5l1.5 1.5M17.5 17.5L19 19M5 19l1.5-1.5M17.5 6.5L19 5"></path>'],
];
function groupStyle(g) {
  const i = Math.max(0, D.cfg.groups.findIndex(x => x.id === g.id)) % G_BG.length;
  const m = G_ICON.find(([re]) => re.test(g.name));
  const path = m ? m[1] : '<circle cx="12" cy="12" r="3.5"></circle>';
  return {bg: G_BG[i], bar: G_BAR[i], icon: `<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="${G_INK[i]}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${path}</svg>`};
}

/* =============== storage =============== */
let D = null;
function blank() { const d = clone(DEFAULTS); if (!d.cfg.start) d.cfg.start = addM(curMonth(), -1) + '-01'; return {v: 6, cfg: d.cfg, kw: d.kw, chat: [], ov: {}, fixed: [], fixedM: {}, manual: [], meta: {dirty: false}}; }
// The data lives in the browser's database on this phone (IndexedDB): much more room than the old
// localStorage (about 5MB), and still only on this phone. Data from older versions is moved over once.
const IDB = {
  db: null,
  open() { return this.db ? Promise.resolve(this.db) : new Promise((res, rej) => { const r = indexedDB.open('kesef-data', 1); r.onupgradeneeded = () => r.result.createObjectStore('kv'); r.onsuccess = () => res(this.db = r.result); r.onerror = () => rej(r.error); }); },
  async get(k) { const db = await this.open(); return new Promise((res, rej) => { const q = db.transaction('kv').objectStore('kv').get(k); q.onsuccess = () => res(q.result); q.onerror = () => rej(q.error); }); },
  async put(k, v) { const db = await this.open(); return new Promise((res, rej) => { const t = db.transaction('kv', 'readwrite'); t.objectStore('kv').put(v, k); t.oncomplete = () => res(); t.onerror = () => rej(t.error); t.onabort = () => rej(t.error); }); },
  async del(k) { const db = await this.open(); return new Promise(res => { const t = db.transaction('kv', 'readwrite'); t.objectStore('kv').delete(k); t.oncomplete = res; t.onerror = res; }); },
};
let useLS = false;                                  // fallback when the browser database can't be opened
async function loadStore() {
  try {
    const s = await IDB.get('D');
    if (s) return JSON.parse(s);
    const old = localStorage.getItem(KEY);
    if (!old) return null;
    await IDB.put('D', old);
    if ((await IDB.get('D')) === old) localStorage.removeItem(KEY);       // moved and checked – only then removed
    return JSON.parse(old);
  } catch { useLS = true; try { const s = localStorage.getItem(KEY); return s ? JSON.parse(s) : null; } catch { return null; } }
}
let writing = null, pendingJSON = null, lastJSON = null;
function wipeStore() {
  try { localStorage.removeItem(KEY); localStorage.removeItem(LOCK_KEY); } catch {}
  pendingJSON = null; lastJSON = null; undoBase = null;
  IDB.del('D').catch(() => {}); IDB.del('snap').catch(() => {});
}
function persist(json) {
  if (useLS) { try { localStorage.setItem(KEY, json); } catch { toast('השמירה בטלפון נכשלה – כדאי ליצור גיבוי עכשיו', 5000); } return; }
  pendingJSON = json;
  if (writing) return;
  writing = (async () => {
    while (pendingJSON != null) { const j = pendingJSON; pendingJSON = null; try { await IDB.put('D', j); } catch { toast('השמירה בטלפון נכשלה – כדאי ליצור גיבוי עכשיו', 5000); } }
    writing = null;
  })();
}
// ---- undo: a change made by a tap can be undone from the message that follows it ----
let lastTapAt = 0, undoBase = null, undoGroupTill = 0, undoTimer = 0;
document.addEventListener('click', ev => { if (!ev.target.closest('input, textarea, select')) lastTapAt = Date.now(); }, true);
function save(dirty = true, opts = {}) {
  if (dirty) { if (!D.meta.dirty) D.meta.dirtySince = Date.now(); D.meta.dirty = true; }
  cacheClear();
  const json = JSON.stringify(D);
  if (lastJSON && json !== lastJSON && !opts.noUndo && (opts.undo || (dirty && Date.now() - lastTapAt < 4000))) {
    if (Date.now() > undoGroupTill) undoBase = lastJSON;      // several saves of one action = one undo
    undoGroupTill = Date.now() + 900;
    clearTimeout(undoTimer); undoTimer = setTimeout(offerUndo, 0);
  }
  lastJSON = json;
  persist(json);
}
function offerUndo() {
  if (!undoBase) return;
  const el = $('#toast'), recent = el.classList.contains('on') && Date.now() - (toast.at || 0) < 400;
  toast(recent ? el.dataset.t : 'נשמר', 5000, true);
}
function doUndo() {
  if (!undoBase) return;
  D = JSON.parse(undoBase); undoBase = null; undoGroupTill = 0;
  cacheClear(); lastJSON = JSON.stringify(D); persist(lastJSON);
  render(); refreshSheets(); toast('בוטל ✓');
}
// ---- one calculation per change: month lists are kept until the data changes ----
const CACHE = new Map();
const cacheClear = () => CACHE.clear();
const cached = (k, fn) => { if (!CACHE.has(k)) CACHE.set(k, fn()); return CACHE.get(k); };
function migrate() {
  if (!D || (D.v || 1) >= 2) return;
  D.v = 2; save(false);
}
function migrate3() {
  if (!D || (D.v || 1) >= 3) return;
  const g = D.cfg.groups.find(x => x.name === 'הכנסות');
  if (g) for (const c of D.cfg.cats) if (c.group === g.id) c.kind = 'income';
  D.v = 3; save(false);
}
function migrate4() {
  if (!D || (D.v || 1) >= 4) return;
  for (const it of D.chat || []) {
    if (it.incoming || REFUND_EXP.test(it.text || '') || !CREDIT_W.test(it.text || '')) continue;
    it.incoming = true; it.refundExp = false; it.refundAmt = 0;
    it.cat = suggestCat(refundCore(it.desc), false) || ''; it.autoSkip = it.cat === SKIP; if (it.autoSkip) it.cat = '';
  }
  D.v = 4; save(false);
}
// v5: incoming lines ("התקבל ...") are matched to income sections by keywords too, not only to "הכנסות נוספות"
function migrate5() {
  if (!D || (D.v || 1) >= 5) return;
  recategorize();
  D.v = 5; save(false, {noUndo: true});
}
// v6: "התקבל ..." without an income keyword is a refund, not "הכנסות נוספות" – applied to every line already in the app
function migrate6() {
  if (!D || (D.v || 1) >= 6) return;
  recategorize();
  D.v = 6; save(false, {noUndo: true});
}
try { navigator.storage && navigator.storage.persist && navigator.storage.persist(); } catch {}

/* =============== model =============== */
const cats = () => D.cfg.cats;
const catById = id => cats().find(c => c.id === id);
const catByName = n => cats().find(c => c.name === n);
const groupById = id => D.cfg.groups.find(g => g.id === id);
const isIncome = c => c && c.kind === 'income';
const incomeGroup = () => D.cfg.groups.find(g => g.name === 'הכנסות');
const inIncomeGroup = c => c && incomeGroup() && c.group === incomeGroup().id;
const kindSelect = (c, attr) => inIncomeGroup(c) ? `<span class="note" ${attr.includes('class') ? '' : ''}>הכנסה</span>` : `<select ${attr}>${Object.entries(KIND_N).filter(([k]) => k !== 'income').map(([k, n]) => `<option value="${k}" ${c.kind === k ? 'selected' : ''}>${n}</option>`).join('')}</select>`;
const KIND_N = {income: 'הכנסה', fixed: 'קבועה', variable: 'משתנה', saving: 'חיסכון'};

/* ---- refunds: money back is tied to the expense that waited for it, and is counted there as a minus ----
   "450 עבור השתלמות, צריך להתקבל החזר" … "התקבל 445 החזר" → the 445 goes into the section of the 450, as −445.
   One expense can get its money back in several messages; once about 90% came back it is closed by itself.
   When it is not clear which expense a refund belongs to, the "לבדיקה" screen asks once. */
const RF_DONE = 0.9;
const RF_STOP = /^(התקבל|התקבלו|קיבלנו|קיבלתי|קיבלה|קיבל|נכנס|נכנסו|הגיע|הגיעו|החזר|ההחזר|החזרים|זיכוי|יתקבל|להתקבל|לקבל|נקבל|צריך|צפוי|מגיע|אמור|יוחזר|יוחזרו|יחזירו|עבור|של|על|את|עם|שח|דרך|קופה|קטנה|אשראי|מזומן|ביט|בביט|פייבוקס|בפייבוקס|העברה|בהעברה|כרטיס)$/;
function rfWords(t) {
  const s = new Set();
  for (const w0 of String(t || '').replace(/[()"'.,:;!?\-–״׳]/g, ' ').split(/\s+/)) {
    if (!w0 || /\d/.test(w0) || RF_STOP.test(w0)) continue;
    const w = w0.toLowerCase(), b = w.replace(/^(מה|וה|שה|כש|ה|מ|ל|ב|ו|ש)/, '');
    if (w.length >= 3) s.add(w);
    if (b.length >= 3 && !RF_STOP.test(b)) s.add(b);
  }
  return s;
}
// the expense's own words, without "צריך להתקבל החזר" and the like
const rfShort = d => String(d || '').replace(new RegExp(REFUND_EXP.source, 'g'), ' ').replace(/(?<![֐-׿])(צריך|עבור|אמור)(?![֐-׿])/g, ' ').replace(/[,]+/g, ' ').replace(/ {2,}/g, ' ').trim();
const chatItem = id => D.chat.find(x => x.id === id) || (D.rfPre || []).find(x => x.id === id);
const rfSkipped = it => { const o = D.ov[it.id] || {}; return o.skip === true || (it.autoSkip && o.skip !== false); };
// money that came in is a refund (it comes off an expense) – unless a keyword ("משכורת") or a hand choice puts it in an income section
const rfIsIn = it => { if (!it || !it.incoming) return false; if (isRefundText(it.text || '')) return true; const o = D.ov[it.id] || {}; return !isIncome(catById(o.cat || it.cat)); };
// the section of a line of money in: an income section only by an income keyword, otherwise an expense section (refund)
const inCat = (text, desc) => (!isRefundText(text) && suggestCat(desc, true)) || suggestCat(refundCore(desc), false) || '';
const rfAmtOf = it => { const o = D.ov[it.id] || {}; return o.amount != null ? o.amount : it.amount; };
// how much should come back: "החזר של 300" if written, otherwise the whole expense (after a corrected amount)
const rfWant = it => { const o = D.ov[it.id] || {}; return it.refundAmt && it.refundAmt !== it.amount ? it.refundAmt : rfAmtOf(it); };
function refundLinks() { return cached('rfl', refundLinks_); }
function refundLinks_() {
  const of = {}, ask = {}, got = {}, list = {}, want = {}, byId = {}, pre = {};
  const open = [], ids = new Set(D.chat.map(i => i.id));
  const preItems = (D.rfPre || []).filter(p => p.date < D.cfg.start && !ids.has(p.id));
  for (const p of preItems) pre[p.id] = true;
  const chron = [...preItems, ...D.chat].filter(it => !rfSkipped(it)).sort((a, b) => tKey(a).localeCompare(tKey(b)));
  const rem = x => want[x.it.id] - got[x.it.id];
  const closed = x => { const o = D.ov[x.it.id] || {}; return got[x.it.id] >= want[x.it.id] * RF_DONE || (o.refundDone === true && got[x.it.id] > 0); };
  for (const it of chron) {
    if (it.refundExp && !it.incoming) {
      byId[it.id] = it; want[it.id] = rfWant(it); got[it.id] = 0; list[it.id] = [];
      open.push({it, words: rfWords(rfShort(it.desc))});
      continue;
    }
    if (!rfIsIn(it)) continue;
    const o = D.ov[it.id] || {}, a = rfAmtOf(it);
    if (o.rfOf === 'none') continue;
    let e = null;
    if (o.rfOf) e = open.find(x => x.it.id === o.rfOf) || null;
    else {
      const age = x => (Date.parse(it.date) - Date.parse(x.it.date)) / 864e5;
      const cands = open.filter(x => !closed(x) && x.it.date <= it.date && age(x) <= 365 && a <= rem(x) * 1.15 + 10);
      if (!cands.length) continue;
      const rw = rfWords(it.desc), shared = x => [...x.words].some(w => rw.has(w));
      const close = x => Math.abs(rem(x) - a) <= Math.max(10, rem(x) * 0.05);
      const W = cands.filter(shared), pool = W.length ? W : cands, C = pool.filter(close);
      if (W.length === 1) e = W[0];
      else if (C.length === 1) e = C[0];
      else {
        ask[it.id] = pool.map(x => ({x, s: (shared(x) ? 2 : 0) + (close(x) ? 1 : 0)}))
          .sort((p, q) => q.s - p.s || Math.abs(rem(p.x) - a) - Math.abs(rem(q.x) - a) || tKey(q.x.it).localeCompare(tKey(p.x.it))).slice(0, 4).map(p => p.x.it.id);
        continue;
      }
    }
    if (e) { of[it.id] = e.it.id; got[e.it.id] += a; list[e.it.id].push(it.id); }
  }
  return {of, ask, got, list, want, byId, pre};
}
// state of an expense that waits for a refund
function rfState(id) {
  const L = refundLinks(), o = D.ov[id] || {}, want = L.want[id] || 0, got = L.got[id] || 0;
  const auto = want > 0 && got >= want * RF_DONE;
  return {want, got, auto, done: o.refundDone === true || (o.refundDone !== false && auto), refunds: L.list[id] || []};
}
// expenses a refund may be tied to by hand: waiting for money back, written before it, in the last year
function rfChoices(refundId) {
  const it = chatItem(refundId); if (!it) return [];
  const L = refundLinks();
  return Object.values(L.byId).filter(e => e.date <= it.date && (Date.parse(it.date) - Date.parse(e.date)) / 864e5 <= 365)
    .sort((a, b) => tKey(b).localeCompare(tKey(a))).slice(0, 8).map(e => e.id);
}
const rfLabel = id => { const e = refundLinks().byId[id]; if (!e) return ''; const st = rfState(id);
  return `${rfShort(e.desc) || 'הוצאה'} · ${ils(st.want)}${st.got ? ` (חזרו ${ils(st.got)})` : ''} · ${dShort(e.date)}`; };

function chatTx(it, withSkipped) {
  const o = D.ov[it.id] || {};
  const skip = o.skip === true || (it.autoSkip && o.skip !== false);
  if (skip && !withSkipped) return null;
  let cat = o.cat || it.cat || '', desc = it.desc || it.text, rfOf = '', rfAsk = null, rfPre = false;
  if (it.incoming) {
    const L = refundLinks();
    rfOf = L.of[it.id] || ''; rfAsk = L.ask[it.id] || null; rfPre = !!(rfOf && L.pre[rfOf]);
    if (rfOf) {
      const e = L.byId[rfOf], eo = D.ov[rfOf] || {};
      // an expense from before the budget start may have no section yet: then the refund's own choice counts
      cat = eo.cat || e.cat || (rfPre ? cat : '');
      const sd = rfShort(e.desc); if (sd) desc = `${desc} · על ${sd}${rfPre ? ` (${dShort(e.date)})` : ''}`;
    }
  }
  const c = catById(cat);
  const raw = o.amount != null ? o.amount : it.amount;
  const amt = it.incoming && !isIncome(c) ? -raw : raw;
  const rs = it.refundExp && !it.incoming ? rfState(it.id) : null;
  return {id: it.id, src: 'chat', date: it.date, time: it.time, month: it.date.slice(0, 7), amount: amt, cat: c ? cat : '', desc, text: it.text,
    pay: o.pay || it.pay, person: it.person, note: o.note || '', refundExp: !!rs, refundAmt: rs ? rs.want : 0, refundDone: rs ? rs.done : false, refundGot: rs ? rs.got : 0,
    rfOf, rfAsk, rfPre, preBudget: it.date < D.cfg.start, pending: !c, incoming: it.incoming, skipped: skip, autoSkip: !!it.autoSkip};
}
function fixedActive(f, m) {
  if (f.from && m < f.from) return false;
  if (f.to && m > f.to) return false;
  if (f.freq === '2m') return mDiff(f.from || m, m) % 2 === 0;
  if (f.freq === 'list') return (f.months || []).includes(+m.slice(5));
  return true;
}
function fixedFor(m) { return cached('fx|' + m, () => fixedFor_(m)); }
function fixedFor_(m) {
  return D.fixed.filter(f => fixedActive(f, m)).map(f => {
    const k = f.id + '|' + m, o = D.fixedM[k] || {};
    return {id: 'f:' + k, key: k, fid: f.id, src: 'fixed', date: `${m}-${pad(Math.min(f.day || 1, daysIn(m)))}`, month: m, amount: o.amount != null ? o.amount : num(f.amount),
      cat: catById(f.cat) ? f.cat : '', desc: f.name, pay: f.pay || '', person: '', ok: !!o.ok, note: o.note || '', pending: !catById(f.cat)};
  });
}
function manualTx(x) { return {...x, src: 'manual', month: x.date.slice(0, 7), pending: !catById(x.cat)}; }
// The monthly amount set for an income category counts as income from the 1st of the month.
// Until it is confirmed it covers only what has not come in yet from other entries in that category.
function expectedIncome(m, real) {
  const out = [];
  for (const c of cats()) {
    if (!isIncome(c) || c.archived || !(num(c.budget) > 0)) continue;
    if (m < minMonth()) continue;
    if (D.fixed.some(f => f.cat === c.id && fixedActive(f, m))) continue;      // a fixed item already covers it
    const k = 'inc:' + c.id + '|' + m, o = D.fixedM[k] || {};
    const got = real.filter(t => t.cat === c.id).reduce((s, t) => s + t.amount, 0);
    const amount = o.ok ? (o.amount != null ? o.amount : num(c.budget)) : Math.max(0, num(c.budget) - got);
    if (!o.ok && amount <= 0) continue;
    out.push({id: 'i:' + c.id + '|' + m, key: k, src: 'income', date: `${m}-01`, month: m, amount, cat: c.id, desc: c.name + (o.ok ? '' : ' (צפוי)'), pay: 'עו"ש (הו"ק / העברות)', person: '', ok: !!o.ok, note: o.note || '', expected: num(c.budget)});
  }
  return out;
}
function monthTx(m) { return cached('mt|' + m, () => monthTx_(m)); }
function monthTx_(m) {
  const out = [];
  for (const it of D.chat) if (it.date.slice(0, 7) === m) { const t = chatTx(it); if (t) out.push(t); }
  out.push(...fixedFor(m));
  for (const x of D.manual) if (x.date.slice(0, 7) === m) out.push(manualTx(x));
  out.push(...expectedIncome(m, out.filter(t => !t.pending)));
  return out.sort((a, b) => b.date.localeCompare(a.date) || String(b.time || '').localeCompare(String(a.time || '')));
}
const isClosed = m => !!(D.closed && D.closed[m]);
function monthData(m) { return cached('md|' + m, () => monthData_(m)); }
function monthData_(m) {
  const list = monthTx(m);
  let byCat = {}, pend = 0, pendSum = 0;
  for (const t of list) { if (t.pending) { pend++; pendSum += t.amount; continue; } byCat[t.cat] = (byCat[t.cat] || 0) + t.amount; }
  const closed = isClosed(m);
  if (closed) { byCat = {...D.closed[m]}; pend = 0; }
  const K = {income: {b: 0, a: 0, e: 0}, fixed: {b: 0, a: 0, e: 0}, variable: {b: 0, a: 0, e: 0}, saving: {b: 0, a: 0, e: 0}};
  for (const c of cats()) {
    const a = byCat[c.id] || 0, b = num(c.budget);
    if (c.archived && !a) continue;
    const k = K[c.kind] || K.variable; k.b += b; k.a += a; k.e += Math.max(b, a);
  }
  const incOpen = list.filter(t => t.src === 'income' && !t.ok).reduce((s, t) => s + t.amount, 0);
  return {list, byCat, K, pend, pendSum, incOpen, closed};
}
function pendingChat() { return cached('pend', () => pendingChat_()); }
// a refund tied to an expense follows that expense's section, and one that waits for "which expense?" is asked separately
// money that came in in a month already closed in Excel is asked about too: the refund rule changed those lines after the month was closed,
// and a line without a section is not exported – so Excel would never hear about it
const rvMonthOk = t => !isClosed(t.month) || t.incoming;
function pendingChat_() { return D.chat.map(it => chatTx(it)).filter(t => t && t.pending && (!t.rfOf || t.rfPre) && !t.rfAsk && t.date >= D.cfg.start && rvMonthOk(t)).sort((a, b) => a.date.localeCompare(b.date)); }
function rfAsks() { return cached('rfask', () => D.chat.map(it => chatTx(it)).filter(t => t && t.rfAsk && t.date >= D.cfg.start && rvMonthOk(t)).sort((a, b) => a.date.localeCompare(b.date))); }
const closedTag = t => isClosed(t.month) ? ` · <b style="color:var(--pink)">חודש סגור ${sheetName(t.month)}</b>` : '';
const reviewMonths = () => [addM(curMonth(), -1), curMonth()].filter(m => m >= minMonth() && !isClosed(m));
function dupPairs() { return cached('dup', () => dupPairs_()); }
function dupPairs_() {
  const by = {}, out = [];
  for (const it of D.chat) { const t = chatTx(it); if (!t || isClosed(t.month) || t.date < D.cfg.start) continue; const k = t.date + '|' + Math.abs(t.amount); (by[k] = by[k] || []).push(t); }
  for (const k in by) { const g = by[k]; for (let i = 0; i < g.length; i++) for (let j = i + 1; j < g.length; j++) {
    const a = g[i], b = g[j]; if ((D.ov[a.id] || {}).notDup || (D.ov[b.id] || {}).notDup) continue;
    if (!!a.incoming !== !!b.incoming) continue;      // an expense and its refund on the same day are not a double entry
    const words = t => new Set(String(t.desc).replace(/[()"',.\-]/g, ' ').split(/\s+/).filter(w => w.length >= 3 && !/^(העברה|בביט|ביט|פייבוקס|בפייבוקס|מזומן|דרך|עבור)$/.test(w)));
    const wa = words(a), shared = [...words(b)].some(w => wa.has(w));
    if (a.desc === b.desc || (a.person !== b.person && shared)) out.push([a, b]); } }
  return [...out, ...fixedDupPairs()];
}
// a fixed charge from the workbook that was also written in the group in the same month
function fixedDupPairs() {
  const out = [], fx = {};
  for (const it of D.chat) {
    const t = chatTx(it); if (!t || t.pending || t.incoming || isClosed(t.month) || t.date < D.cfg.start || (D.ov[t.id] || {}).notDup) continue;
    const list = fx[t.month] = fx[t.month] || fixedFor(t.month).filter(f => (D.fixed.find(x => x.id === f.fid) || {}).xkey);
    const f = list.find(f => Math.abs(f.amount - t.amount) <= Math.max(1, f.amount * 0.02) && (fxShare(f.desc, t.desc) || (f.cat === t.cat && Math.abs(f.amount - t.amount) < 0.01)));
    if (f) out.push([{...f, person: 'קבועה מהאקסל', text: f.desc}, t]);
  }
  return out;
}
function openFixed(m) { return fixedFor(m).filter(f => !f.ok && !(D.fixed.find(x => x.id === f.fid) || {}).xkey); }
function openRefunds() { return D.chat.map(it => chatTx(it)).filter(t => t && t.refundExp && !t.refundDone); }
function openIncome(m) { const l = monthTx(m); return l.filter(t => t.src === 'income' && !t.ok); }
function dueFixed() { const out = []; for (const m of reviewMonths()) out.push(...openFixed(m).filter(f => f.date <= todayStr()), ...openIncome(m)); return out; }
function entryByKey(k) {
  const m = k.split('|')[1];
  if (k.startsWith('inc:')) return monthTx(m).find(t => t.key === k) || (() => { const c = catById(k.slice(4).split('|')[0]); return {key: k, amount: num(c && c.budget), cat: c && c.id}; })();
  const fid = k.split('|')[0]; return fixedFor(m).find(x => x.fid === fid);
}
function reviewCount() { return pendingChat().length + rfAsks().length + multiNums().length + dueFixed().length + dupPairs().length + ((D.meta.fxAsk || []).length) + ((D.meta.fxChanges || []).length ? 1 : 0); }

const SKIP = '__skip';
const kwCatName = c => c === SKIP ? 'לא לייבא' : (catById(c) ? catById(c).name : '?');
// The longest keyword found in the text decides ("סופר פארם" beats "סופר"); on a tie, the one added last.
function suggestCat(desc, wantIncome) {
  const d = ' ' + String(desc || '').toLowerCase() + ' ';
  let hit = null, len = -1;
  for (const [k, c] of D.kw) {
    const kk = String(k).toLowerCase(); if (!kk.trim() || !d.includes(kk)) continue;
    if (c !== SKIP) { const cat = catById(c); if (!cat || cat.archived || !(wantIncome === null || isIncome(cat) === wantIncome)) continue; }
    const l = kk.trim().length; if (l >= len) { hit = c; len = l; }
  }
  return hit;
}

/* =============== WhatsApp parser (same rules as the workbook import) =============== */
// a credit word anywhere in the message ("זיכוי מאושר עד 120", "שופרסל זיכוי 45", "הוחזרו 60") = money came back, not an expense
const CREDIT_W = /(?<![\u0590-\u05FF])[וה]?(זיכוי|זיכויים|זיכו|הוחזר|הוחזרה|הוחזרו)(?![\u0590-\u05FF])/;
// an expense that waits for money back: "צריך להתקבל החזר", "יתקבל החזר", "מגיע החזר", "יוחזר", "קופה קטנה"…
const REFUND_EXP = /(יתקבל|להתקבל|לקבל|נקבל|צפוי|מגיע)\s*(ה)?(החזר|זיכוי)|(?<![֐-׿])(יוחזר|יוחזרו|יחזירו|יזוכה)(?![֐-׿])|החזר\s*(דרך\s*)?קופה קטנה|קופה קטנה/;
// money back that arrived: "התקבל 445 החזר", "קיבלנו החזר 445", "445 התקבל החזר"
const IN_VERB = /^(התקבל|התקבלו|קיבלנו|קיבלתי|קיבלה|קיבל|נכנס|נכנסו|הגיע|הגיעו)(?![֐-׿])/;
const isRefundText = t => /החזר/.test(t) || CREDIT_W.test(t);
const refundCore = d => String(d || '').replace(/^החזר\s*/, '').replace(new RegExp(CREDIT_W.source, 'g'), ' ').replace(/ {2,}/g, ' ').trim();
const BIDI = /[‎‏‪-‮⁦-⁩﻿]/g;
const SPACES = /[     \t]/g;
const HEAD = /^\[?(\d{1,2})[./-](\d{1,2})[./-](\d{2,4}),?\s+(\d{1,2}):(\d{2})(?::\d{2})?\s*([AaPp]\.?[Mm]\.?|לפנה"צ|אחה"צ|לפנה״צ|אחה״צ)?\]?\s*(?:-\s*)?(.*)$/;
const NOAMOUNT_SYS = /^(מחקת את ההודעה|הודעה זו נמחקה|ההודעה נמחקה|<מדיה הושמטה>|<Media omitted>|This message was deleted|You deleted this message)/;
const EDITED = /\s*<(ההודעה נערכה|This message was edited)>\s*$/;
function cleanLine(t) { return t.replace(BIDI, '').replace(SPACES, ' ').replace(/ {2,}/g, ' ').trim(); }
function firstNumber(msg) {
  const m = msg.match(/\d+(?:,\d{3})*(?:[.,]\d+)?/);
  if (!m) return null;
  let s = m[0];
  if (/,\d{3}(?!\d)/.test(s) && !/[.]/.test(s)) s = s.replace(/,/g, '');           // 1,280 → 1280
  else if (/,\d{1,2}$/.test(s)) s = s.replace(/,(\d{1,2})$/, '.$1').replace(/,/g, ''); // 17,5 → 17.5
  else s = s.replace(/,/g, '');
  return {value: parseFloat(s), index: m.index, len: m[0].length};
}
// Cards are told apart by company: "אשראי דני כאל", "אשראי דני ישראכרט"… A new card written in the group is added to the list by itself.
const CARD_CO = ['כאל', 'ישראכרט', 'מקס', 'max', 'לאומי קארד', 'לאומיקארד', 'אמקס', 'אמריקן אקספרס', 'דיינרס', 'ויזה', 'visa', 'מאסטרקארד', 'פועלים', 'דיסקונט', 'מזרחי', 'בינלאומי', 'וואן זירו', 'one zero', 'פלוס', 'כאל פלוס'];
const reEsc = s => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
const END = '(?=$|[\\s(),.:\\-–])';
const isCardName = p => /^אשראי\s/.test(p);
function detectPay(msg, person, added) {
  const cand = [];
  for (const p of D.cfg.pays) if (isCardName(p)) { const m = msg.match(new RegExp('(^|[\\s(])([לב]?' + reEsc(p) + ')' + END, 'i')); if (m) cand.push({pay: p, phrase: m[2]}); }
  const ppl = [...people()].sort((a, b) => b.length - a.length).map(reEsc).join('|');
  const co = [...CARD_CO].sort((a, b) => b.length - a.length).map(reEsc).join('|');
  let m = !ppl ? null : msg.match(new RegExp('[לב]?אשראי\\s+(' + ppl + ')\\s+(' + co + ')' + END, 'i')) || msg.match(new RegExp('[לב]?אשראי\\s+(' + ppl + ')' + END));
  if (m) {
    let pay = m[2] ? `אשראי ${m[1]} ${m[2]}` : `אשראי ${m[1]}`;
    const dp = (D.cfg.defPay || {})[m[1]] || '';
    if (!m[2] && dp.startsWith(pay + ' ')) pay = dp;           // "אשראי דני" without a company → that person's default card
    cand.push({pay, phrase: m[0], auto: true});
  }
  if (cand.length) {
    const best = cand.sort((a, b) => b.phrase.length - a.phrase.length)[0];
    const known = D.cfg.pays.find(p => p.toLowerCase() === best.pay.toLowerCase());
    if (known) best.pay = known; else if (added) { D.cfg.pays.push(best.pay); added.push(best.pay); }
    return best;
  }
  if (/מזומן/.test(msg)) return {pay: 'מזומן'};
  if (/העברה בנקאית|הוראת קבע|הו"ק/.test(msg)) return {pay: 'עו"ש (הו"ק / העברות)'};
  if (/(^|[\s(])[בל]?(ביט|פייבוקס)(?=$|[\s).,])/.test(msg)) return {pay: 'ביט / פייבוקס'};
  return {pay: (D.cfg.defPay || {})[person] || ''};
}
function senderToPerson(s) {
  const map = D.cfg.senders || {};
  if (map[s]) return map[s];
  if (/^(את\/ה|אתה|את|You|you)$/.test(s)) return D.cfg.me || people()[0];
  const p = [...people()].sort((a, b) => b.length - a.length).find(p => s.includes(p));
  return p || s;
}
const people = () => D.cfg.people || [];
function parseChat(text) {
  const lines = text.split(/\r?\n/);
  const heads = [];
  for (const raw of lines) { const t = cleanLine(raw); const h = t.match(HEAD); if (h) heads.push([+h[1], +h[2]]); }
  let order = D.cfg.dateOrder;
  if (order === 'auto' || !order) {
    const a = heads.some(h => h[0] > 12), b = heads.some(h => h[1] > 12);
    order = b && !a ? 'md' : 'dm';
  }
  const items = [], senders = new Set(), seen = {}, addedPays = [];
  let cur = null;
  for (const raw of lines) {
    let t = cleanLine(raw); if (!t) continue;
    let msg = t;
    const h = t.match(HEAD);
    if (h) {
      let [, a, b, y, hh, mi, ap, rest] = h;
      let d = +a, mo = +b; if (order === 'md') { d = +b; mo = +a; }
      y = +y; if (y < 100) y += 2000;
      if (d < 1 || d > 31 || mo < 1 || mo > 12) { cur = null; continue; }
      let H = +hh; if (ap) { const pm = /p|אחה/i.test(ap); H = H % 12 + (pm ? 12 : 0); }
      const ci = rest.indexOf(': ');
      if (ci < 0 && !rest.endsWith(':')) { cur = null; continue; }           // system line (no sender)
      const sender = rest.slice(0, ci < 0 ? rest.length - 1 : ci).trim();
      msg = ci < 0 ? '' : rest.slice(ci + 2);
      cur = {date: `${y}-${pad(mo)}-${pad(d)}`, time: `${pad(H)}:${pad(+mi)}`, sender};
      senders.add(sender);
    }
    if (!cur) continue;
    msg = msg.replace(EDITED, '').trim();
    if (!msg || NOAMOUNT_SYS.test(msg)) continue;
    const n = firstNumber(msg);
    if (!n || !(n.value > 0)) continue;
    const multi = (msg.match(/\d+(?:[.,]\d+)*/g) || []).length > 1;      // "2 פיצות 90": the first number may not be the amount
    const person = senderToPerson(cur.sender), pd = detectPay(msg, person, addedPays);
    let desc = ' ' + msg.slice(0, n.index) + msg.slice(n.index + n.len) + ' ';
    if (pd.phrase) desc = desc.split(pd.phrase).join(' ');
    for (const ph of ['כרטיס']) desc = desc.split(ph).join('');
    for (const ph of [' שח ', ' ש"ח ', ' ש״ח ', ' ₪ ', '₪']) desc = desc.split(ph).join(' ');
    desc = desc.replace(/ {2,}/g, ' ').replace(/^[\s\-–,.:]+|[\s\-–,]+$/g, '').trim();
    const expRe = REFUND_EXP.test(msg);
    const incoming = msg.startsWith('התקבל') || desc.startsWith('התקבל') || ((IN_VERB.test(msg) || IN_VERB.test(desc)) && isRefundText(msg)) || desc.startsWith('החזר') || (!expRe && CREDIT_W.test(msg));
    let cat = '';
    if (incoming) {
      cat = inCat(msg, desc);     // "התקבל משכורת" → the salary section; "התקבל 300" → a refund, off an expense
    } else cat = suggestCat(desc, false) || '';
    const autoSkip = cat === SKIP; if (autoSkip) cat = '';
    const refundExp = !incoming && expRe;
    const ra = refundExp ? (msg.match(/החזר\s*של\s*(\d+(?:\.\d+)?)/) || [])[1] : null;
    // the id depends only on what was written and when – not on the contact's name in the phone or on settings
    const base = `${cur.date}T${cur.time}|${n.value}|${hash8(normMsg(msg))}`;
    seen[base] = (seen[base] || 0) + 1;
    items.push({id: base + '#' + seen[base], date: cur.date, time: cur.time, sender: cur.sender, person, text: msg, desc, amount: n.value, incoming, cat,
      pay: pd.pay, autoSkip, refundExp, multi, refundAmt: ra ? +ra : (refundExp ? n.value : 0)});
  }
  // expenses from the year before the budget start that wait for a refund are kept aside (not counted anywhere),
  // only so that a refund arriving after the start still comes off that expense's section
  const preFrom = new Date(Date.parse(D.cfg.start) - 365 * 864e5).toISOString().slice(0, 10);
  const pre = items.filter(i => i.date < D.cfg.start && i.date >= preFrom && i.refundExp && !i.incoming);
  return {items: items.filter(i => i.date >= D.cfg.start), all: items.length, order, senders: [...senders], addedPays, pre, preSeen: items.some(i => i.date < D.cfg.start)};
}
async function readChatFile(file) {
  const u8 = new Uint8Array(await file.arrayBuffer());
  if (u8[0] === 0x50 && u8[1] === 0x4b) {
    const XLSX = await loadXLSX();
    const z = XLSX.CFB.read(u8, {type: 'array'});
    const e = z.FileIndex.filter(f => /\.txt$/i.test(f.name) && f.content && f.content.length).sort((a, b) => b.content.length - a.content.length)[0];
    if (!e) throw new Error('בקובץ ה-zip לא נמצא קובץ txt');
    return new TextDecoder('utf-8').decode(e.content);
  }
  return new TextDecoder('utf-8').decode(u8);
}
const tKey = i => `${i.date}T${i.time}`;
const normMsg = t => String(t || '').replace(EDITED, '').replace(/\s+/g, ' ').trim().toLowerCase();
function hash8(s) { let h = 0x811c9dc5; for (let i = 0; i < s.length; i++) { h ^= s.charCodeAt(i); h = Math.imul(h, 0x01000193); } return (h >>> 0).toString(16).padStart(8, '0'); }
const noNums = t => new Set(normMsg(t).replace(/[\d.,₪]+/g, ' ').split(' ').filter(w => w.length >= 2));
// A line that is already in the app keeps its id (and with it the chosen section, corrections, and its place in Excel) when:
// 1. its contact was renamed in the phone (same time, amount and text), or 2. the message was edited in WhatsApp
// (same minute, same writer or mostly the same words).
function keepIds(items) {
  const byId = new Map(D.chat.map(o => [o.id, o])), used = new Set(), out = {renamed: 0, edited: 0};
  const sameTxt = new Map();
  for (const o of D.chat) { const k = `${tKey(o)}|${o.amount}|${normMsg(o.text)}`; (sameTxt.get(k) || sameTxt.set(k, []).get(k)).push(o); }
  const rest = [];
  for (const i of items) {
    if (byId.has(i.id) && !used.has(i.id)) { used.add(i.id); continue; }
    const c = (sameTxt.get(`${tKey(i)}|${i.amount}|${normMsg(i.text)}`) || []).find(o => !used.has(o.id));
    if (c) { used.add(c.id); i.id = c.id; out.renamed++; continue; }
    rest.push(i);
  }
  const newIds = new Set(items.map(i => i.id));
  for (const i of rest) {
    const wi = noNums(i.text);
    let best = null, bestS = 0;
    for (const o of D.chat) {
      if (used.has(o.id) || newIds.has(o.id) || tKey(o) !== tKey(i)) continue;
      const wo = noNums(o.text), common = [...wi].filter(w => wo.has(w)).length, sim = common / Math.max(1, Math.max(wi.size, wo.size));
      const s = (o.sender === i.sender ? 1 : 0) + sim;
      if (s > bestS && (o.sender === i.sender || sim >= 0.6)) { best = o; bestS = s; }
    }
    if (best) { used.add(best.id); i.id = best.id; i.edited = true; out.edited++; }
  }
  return out;
}
function mergeChat(items) {
  items.sort((a, b) => a.date.localeCompare(b.date) || a.time.localeCompare(b.time));
  if (!items.length) return {add: 0, del: 0, keep: D.chat.length};
  const kept0 = keepIds(items);
  const min = items[0].date, upTo = tKey(items[items.length - 1]), newIds = new Set(items.map(i => i.id)), oldIds = new Set(D.chat.map(i => i.id));
  const kept = D.chat.filter(o => o.date < min);
  const after = D.chat.filter(o => tKey(o) > upTo && !newIds.has(o.id));
  const del = D.chat.filter(o => o.date >= min && tKey(o) <= upTo && !newIds.has(o.id)).length;
  const add = items.filter(i => !oldIds.has(i.id)).length;
  D.chat = [...kept, ...items, ...after];
  return {add, del, kept: kept.length, renamed: kept0.renamed, edited: kept0.edited};
}
async function importFile(file, fromShare) {
  if (!file) return;
  if (!D) D = blank();
  let text;
  try { text = await readChatFile(file); } catch (e) { toast(e.message || 'לא הצלחתי לקרוא את הקובץ', 4000); return; }
  const r = parseChat(text);
  if (!r.all) { toast('לא נמצאו בקובץ הודעות עם סכום. האם זה ייצוא צ\'אט של ווטסאפ?', 5000); return; }
  const before = pendingChat().length;
  const m = mergeChat(r.items);
  if (r.preSeen) D.rfPre = r.pre;
  const last = r.items.length ? r.items[r.items.length - 1] : null;
  const prevUntil = (D.meta.lastImport || {}).until || '', until = last ? `${last.date} ${last.time}` : '';
  D.meta.lastImport = {at: Date.now(), file: fromShare ? 'שיתוף מווטסאפ' : file.name, until: until && until > prevUntil ? until : (prevUntil || until), order: r.order};
  D.meta.senders = r.senders;
  save(false, {undo: true});
  const pend = pendingChat().length, rAsk = rfAsks().length;
  const L = refundLinks(), rTied = r.items.filter(i => L.of[i.id]).length;
  const lastM = last ? last.date.slice(0, 7) : S.month;
  if (lastM > S.month || !monthTx(S.month).length) S.month = lastM <= curMonth() ? lastM : curMonth();
  render();
  openSheet(() => `${head('הייבוא הושלם')}
    <div class="stat"><div><b>${r.items.length}</b><span>שורות עם סכום</span></div><div><b>${m.add}</b><span>חדשות</span></div><div><b class="${pend ? 'red' : ''}">${pend}</b><span>לבחור סעיף</span></div></div>
    <p class="note">${last ? `ההודעה האחרונה: ${dShort(last.date)} ${last.time}.` : ''} ${m.del ? `${m.del} שורות שנמחקו מהצ'אט הוסרו.` : ''} ${m.edited ? `${m.edited} הודעות שנערכו בווטסאפ עודכנו – הסעיף והתיקונים נשמרו.` : ''} סדר התאריך בקובץ: ${r.order === 'md' ? 'חודש/יום' : 'יום/חודש'}.</p>
    ${rTied || rAsk ? `<p class="note">↩︎ ${rTied ? `${rTied === 1 ? 'החזר אחד חובר' : rTied + ' החזרים חוברו'} להוצאות שחיכו להם וירדו מהן כמינוס.` : ''}${rAsk ? ` ${rAsk === 1 ? 'החזר אחד מחכה' : rAsk + ' החזרים מחכים'} לבחירה – על איזו הוצאה.` : ''}</p>` : ''}
    ${r.addedPays.length ? `<p class="note">אמצעי תשלום חדשים שזוהו: <b>${r.addedPays.map(esc).join(', ')}</b></p>` : ''}
    ${r.senders.length ? `<p class="note">שולחים: ${r.senders.map(s => `${esc(s)} ← <b>${esc(senderToPerson(s))}</b>`).join(' · ')}</p>` : ''}
    <div class="banner" style="background:var(--blue-soft)"><span>${fromShare ? '🔒 הצ\'אט נקלט ישר מווטסאפ – שום קובץ לא נשמר בטלפון, ואין מה למחוק.' : '🔒 קובץ הייצוא נשאר בתיקיית ההורדות ומכיל את כל השיחה. אפשר למחוק אותו עכשיו – הנתונים כבר באפליקציה.'}</span></div>
    ${pend || rAsk ? `<button class="primary" data-act="go-review">${pend ? `לבחירת סעיפים (${pend + rAsk})` : `לבחירת הוצאה להחזר (${rAsk})`}</button>` : '<button class="primary" data-act="close">מעולה</button>'}`);
}

/* =============== UI state & render =============== */
const S = {view: (location.hash || '').replace('#', '') || 'home', month: curMonth(), dismissed: new Set(), f: {q: '', src: '', g: '', pay: '', who: '', sort: ''}, rv: 'todo'};
const VIEWS = ['home', 'data', 'review'];
if (!VIEWS.includes(S.view)) S.view = 'home';

function render() {
  cacheClear();
  const app = $('#app'), nav = $('#nav');
  if (!D) {
    nav.hidden = true;
    app.innerHTML = `<div class="empty" style="padding-top:70px">${I.flower}<h2>איפה <span class="script">הכסף</span></h2>
      ${S.pendingShare ? `<div class="banner" style="max-width:340px;margin:12px auto"><span>💬 יש צ'אט ששותף מווטסאפ ומחכה. קודם משחזרים מגיבוי – והוא ייובא מיד אחר כך.</span></div>` : ''}
      <p style="line-height:1.6">פעם ראשונה כאן? קודם <b>שחזור מגיבוי</b> של האפליקציה הקודמת – כך חוזרים הסעיפים, מילות המפתח וכל הנתונים.<br>אחר כך: בקבוצה ← ⋮ ← עוד ← <b>ייצוא צ'אט</b> ← ללא מדיה ← <b>איפה הכסף</b>.</p>
      <button class="primary" data-act="restore" style="max-width:300px;margin:20px auto 0">שחזור מגיבוי</button>
      <button class="ghost" data-act="${S.pendingShare ? 'share-now' : 'import'}" style="max-width:300px;margin:10px auto 0">${S.pendingShare ? 'לייבא בלי גיבוי' : 'ייבוא צ\'אט מקובץ'}</button>
      <p class="note" style="margin-top:22px">🔒 הכל נשמר רק בטלפון הזה. האפליקציה חסומה לאינטרנט – היא לא יכולה לשלוח שום דבר לשום מקום.</p></div>`;
    return;
  }
  const sy = window.scrollY;
  app.innerHTML = topBar() + (S.view === 'data' ? viewData() : S.view === 'review' ? viewReview() : viewHome());
  nav.hidden = false; nav.innerHTML = navHtml();
  window.scrollTo(0, sy);
}
function minMonth() { return D.cfg.start.slice(0, 7); }
function topBar() {
  const m = S.month, li = D.meta.lastImport;
  return `<div class="top">
    <button class="iconbtn" data-act="settings" aria-label="תפריט">${I.menu}</button>
    <div class="pill monthnav">
      <button data-act="month" data-k="-1" ${m <= minMonth() ? 'disabled' : ''} aria-label="חודש קודם">${I.r}</button>
      <span>${mLabel(m)}</span>
      <button data-act="month" data-k="1" ${m >= addM(curMonth(), 1) ? 'disabled' : ''} aria-label="חודש הבא">${I.l}</button>
    </div>
    <button class="iconbtn" data-act="sync" aria-label="חיבור לאקסל – ייצוא וטעינה" title="חיבור לאקסל – ייצוא וטעינה" style="width:auto;padding:0 12px;border-radius:999px;gap:5px;display:flex;align-items:center;font-size:12.5px;font-weight:600"><svg width="19" height="19" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.1" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="4" y="4" width="16" height="16" rx="2.5"/><path d="M4 9.5h16M10 9.5V20"/><path d="M13 13.2h4M13 16.6h4"/></svg>אקסל</button>
  </div>
  ${S.waitingSW ? `<button class="banner" data-act="update" style="background:var(--green-soft)"><span>🆕 יש גרסה חדשה${S.newVersion ? ' (' + esc(S.newVersion.version) + ')' : ''}</span><b>מה חדש ←</b></button>` : ''}
  <div class="fresh">${li ? `מעודכן לפי ייצוא עד ${li.until ? dShort(li.until.slice(0, 10)) + ' ' + li.until.slice(11) : '—'}` : 'עוד לא יובא צ\'אט'}${D.meta.dirty ? ' · <a href="#" data-act="backup" style="color:var(--pink)">יש שינויים שלא גובו</a>' : ''}</div>`;
}
function navHtml() {
  const n = reviewCount();
  const b = (v, ic, t, extra = '') => `<button class="${S.view === v ? 'on' : ''}" data-act="view" data-v="${v}">${ic}<span>${t}</span>${extra}</button>`;
  return `<div class="tabs">${b('home', I.home, 'בית')}${b('data', I.flow, 'פירוט')}${b('review', I.check, 'לבדיקה', n ? `<span class="dot">${n}</span>` : '')}</div>
    <button class="fab" data-act="add" aria-label="הוספה">${I.plus}</button>`;
}

/* ---- dashboard ---- */
const isExpT = t => !t.pending && !isIncome(catById(t.cat));
const isIncT = t => !t.pending && isIncome(catById(t.cat));
const sumA = a => a.reduce((s, t) => s + t.amount, 0);
function dash(m) { return cached('dash|' + m, () => dash_(m)); }
function dash_(m) {
  const md = monthData(m), cur = m === curMonth(), td = todayStr(), future = m > curMonth();
  if (md.closed) {
    const byGroup = {}; let e = 0, i = 0;
    for (const id in md.byCat) { const c = catById(id), v = md.byCat[id]; if (!c) continue; if (isIncome(c)) i += v; else { e += v; byGroup[c.group] = (byGroup[c.group] || 0) + v; } }
    return {md, exp: md.list.filter(isExpT), inc: md.list.filter(isIncT), upcoming: [], spent: e, upc: 0, total: e, incExp: i, incIn: i, bal: i - e, byGroup, closed: true};
  }
  const exp = md.list.filter(isExpT), inc = md.list.filter(isIncT);
  const upcoming = exp.filter(t => t.src === 'fixed' && !t.ok && (future || (cur && t.date > td)));
  const soFar = exp.filter(t => !upcoming.includes(t));
  const byGroup = {}; for (const t of exp) { const g = catById(t.cat).group; byGroup[g] = (byGroup[g] || 0) + t.amount; }
  const incExp = sumA(inc), incIn = sumA(inc.filter(t => t.src !== 'income' || t.ok));
  const spent = sumA(soFar), upc = sumA(upcoming);
  return {md, exp, inc, upcoming, spent, upc, total: spent + upc, incExp, incIn, bal: incExp - spent - upc, byGroup};
}
function tips() {
  const out = [];
  if (!D.meta.guideSeen) out.push({id: 'guide', icon: I.star, act: 'guide', t: 'איך לכתוב בקבוצה כדי שהכל יזוהה לבד? קליק כאן ל-5 דוגמאות קצרות.'});
  if (!cats().some(c => isIncome(c) && num(c.budget) > 0)) out.push({id: 'inc', icon: I.star, act: 'income', t: 'עוד לא הוגדרו הכנסות צפויות (משכורות). קליק כאן כדי להגדיר, ואז נראה את היתרה הצפויה.'});
  if (!D.fixed.length) out.push({id: 'fixed', icon: I.ast, act: 'fixed', t: 'משכנתא, ביטוחים והוראות קבע לא נכתבים בקבוצה. קליק כאן כדי להגדיר תנועות קבועות.'});
  return out.filter(x => !S.dismissed.has(x.id));
}
// days left in the month as of the last export ("מעודכן לפי ייצוא עד …"), shown only for that month
function monthLeft(m) {
  const u = D.meta.lastImport && D.meta.lastImport.until;
  if (!u || u.slice(0, 7) !== m) return '';
  const [y, mo, dd] = u.slice(0, 10).split('-').map(Number), dim = new Date(y, mo, 0).getDate(), left = dim - dd;
  const t = left <= 0 ? 'יום אחרון בחודש' : left === 1 ? 'עוד יום אחד' : `עוד ${left} ימים`;
  return `<span class="mleft"><i><i style="width:${Math.round(dd / dim * 100)}%"></i></i>${t}</span>`;
}
// expenses by group up to a given day of the month (from the lines themselves, so a closed month compares fairly too)
function groupsUpTo(m, day) {
  const by = {}, lim = `${m}-${pad(Math.min(day, daysIn(m)))}`;
  for (const t of monthTx(m)) { if (!isExpT(t) || t.date > lim) continue; const c = catById(t.cat); if (!c) continue; by[c.group] = (by[c.group] || 0) + t.amount; }
  return by;
}
function viewHome() {
  const d = dash(S.month), pm = addM(S.month, -1), prev = dash(pm), cur = S.month === curMonth(), future = S.month > curMonth(), past = !cur && !future;
  const hero = cur ? `עד היום <span class="script">יצאו</span> החודש<span class="amt">${ils(d.spent, 0)}</span>` : future ? `ב${mName(S.month)} <span class="script">צפויות</span> הוצאות של<span class="amt">${ils(d.total, 0)}</span>` : `ב${mName(S.month)} <span class="script">יצאו</span><span class="amt">${ils(d.total, 0)}</span>`;
  const ti = d.closed ? '' : tips().map(x => `<div class="card insight"><button class="x" data-act="dismiss" data-id="${x.id}" aria-label="סגירה">✕</button><p data-act="${x.act}" data-v="${x.v || ''}">${x.t}</p>${x.icon}</div>`).join('');
  const groups = D.cfg.groups.filter(g => d.byGroup[g.id]).sort((a, b) => d.byGroup[b.id] - d.byGroup[a.id]);
  const mx = Math.max(1, ...groups.map(g => d.byGroup[g.id]));
  // this month is compared with the same days of last month, not with the whole of it
  const today = +todayStr().slice(8), pBy = cur ? groupsUpTo(pm, today) : prev.byGroup, pTot = cur ? Object.values(pBy).reduce((s, v) => s + v, 0) : prev.total;
  const vsTxt = `לעומת ${mName(pm)}`;
  const big = d.exp.filter(t => t.src !== 'fixed').sort((a, b) => b.amount - a.amount).slice(0, 5);
  const ms = []; for (let i = 5; i >= 0; i--) { const m = addM(S.month, -i); if (m >= minMonth()) ms.push(m); }
  const tr = ms.map(m => { const x = dash(m); return {m, e: x.total, i: x.incExp}; }), tmx = Math.max(1, ...tr.map(x => Math.max(x.e, x.i)));
  const pAll = pendingChat().length;
  const fxTot = d.closed ? cats().filter(c => c.kind === 'fixed').reduce((s, c) => s + (d.md.byCat[c.id] || 0), 0) : sumA(d.exp.filter(t => t.src === 'fixed' || (catById(t.cat) || {}).kind === 'fixed'));
  const upN = d.upcoming.length;
  const days = D.meta.dirtySince ? Math.floor((Date.now() - D.meta.dirtySince) / 864e5) : 0;
  return `<h1 class="hero">${hero}</h1>
    ${d.closed ? `<div style="text-align:center;margin:-8px 0 12px"><span class="tag ok">✓ מדויק – לפי האקסל (נסגר)</span></div>` : ''}
    ${days >= 7 && !S.dismissed.has('bk') ? `<button class="banner" data-act="backup" style="background:#FDEAF1"><span>💾 יש שינויים שלא גובו כבר ${days} ימים</span><b>לגבות ←</b></button>` : ''}
    ${pAll && !d.closed ? `<button class="banner" data-act="view" data-v="review"><span>${d.md.pend ? `${d.md.pend === 1 ? 'שורה אחת' : d.md.pend + ' שורות'} ב${mName(S.month)}${pAll > d.md.pend ? ` (${pAll} בסך הכל)` : ''}` : `${pAll === 1 ? 'שורה אחת' : pAll + ' שורות'} מחודשים אחרים`} עוד בלי סעיף, ולא נספרות</span><b>לסדר ←</b></button>` : ''}
    <div class="stat kpi">
      <div data-act="incdet"><span>הכנסות</span><b>${fmt(d.incExp, 0)}</b><span>${d.incIn < d.incExp ? `נכנס ${fmt(d.incIn, 0)}` : d.incExp ? 'הכל נכנס' : '—'}</span></div>
      ${past ? `<div data-act="view" data-v="review" data-rv="fixed"><span>מתוכן קבועות</span><b>${fmt(fxTot, 0)}</b><span>משכנתא, הו"ק וכד'</span></div>`
        : `<div data-act="view" data-v="review" data-rv="fixed"><span>עוד צפוי לרדת</span><b>${fmt(d.upc, 0)}</b><span>${upN ? `${upN === 1 ? 'חיוב קבוע אחד' : upN + ' חיובים קבועים'}` : 'הקבועות כבר ירדו'}</span></div>`}
      <div data-act="baldet"><span>${past ? (d.bal >= 0 ? 'עודף' : 'חוסר') : (d.bal >= 0 ? 'יתרה צפויה' : 'חוסר צפוי')}</span><b class="${d.bal < 0 ? 'red' : ''}">${fmt(Math.abs(d.bal), 0)}</b><span>${!d.incExp ? 'אין הכנסות מוגדרות' : !past && d.upc ? 'כולל הקבועות שצפויות לרדת' : 'הכנסות פחות הוצאות'}</span>${d.closed ? '' : monthLeft(S.month)}</div>
    </div>
    ${ti ? `<div class="panel">${ti}</div>` : ''}
    <div class="sec-title"><h2>לאן <span class="script blue">הלך</span> הכסף?</h2><div class="chip">${pTot ? (cur ? `עד ${today} ב${mName(pm)}: ${ils(pTot, 0)}` : `${mName(pm)}: ${ils(pTot, 0)}`) : ''}</div></div>
    <div class="card" style="padding:6px 14px">${groups.length ? groups.map(g => { const v = d.byGroup[g.id], pv = pBy[g.id] || 0, df = v - pv;
      const gs = groupStyle(g);
      return `<button class="catrow withicon" data-act="grp" data-id="${g.id}"><span class="gicon" style="background:${gs.bg}">${gs.icon}</span><div class="body"><div class="t"><b>${esc(g.name)}</b><b>${ils(v, 0)}</b></div>
        <div class="bar"><i style="width:${v / mx * 100}%;background:${gs.bar}"></i></div>
        <div class="foot"><span>${d.total ? Math.round(v / d.total * 100) : 0}% מההוצאות</span><span style="color:var(--muted)">${pv ? `${df >= 0 ? '▲' : '▼'} ${fmt(Math.abs(df), 0)} ${vsTxt}` : ''}</span></div></div></button>`; }).join('') : '<p class="note" style="text-align:center">אין עדיין הוצאות בחודש הזה</p>'}</div>
    ${big.length ? `<div class="label">ההוצאות הגדולות החודש</div><div class="card" style="padding:4px 14px">${big.map(t => txLine(t)).join('')}</div>` : ''}
    ${payCard(d.md)}
    ${tr.length > 1 ? `<div class="label">${tr.length === 6 ? 'חצי שנה אחרונה' : 'מתחילת התקציב'}</div>
    <div class="card"><div style="display:flex;align-items:flex-end;gap:8px;height:150px">${tr.map(x => `<button data-act="goto" data-m="${x.m}" style="flex:1;display:flex;flex-direction:column;justify-content:flex-end;align-items:center;height:100%">
        <div style="display:flex;gap:3px;align-items:flex-end;height:110px"><div style="width:12px;border-radius:5px;background:var(--green);height:${Math.max(2, x.i / tmx * 110)}px"></div><div style="width:12px;border-radius:5px;background:${x.m === S.month ? 'var(--blue)' : 'var(--blue-soft)'};height:${Math.max(2, x.e / tmx * 110)}px${x.m === curMonth() ? ';opacity:.55' : ''}"></div></div>
        <div style="font-size:12px;color:var(--muted);margin-top:4px;line-height:1.2">${sheetName(x.m)}${x.m === curMonth() ? '<br><span style="font-size:11px">עד היום</span>' : ''}</div></button>`).join('')}</div>
      <div class="foot" style="margin-top:8px;justify-content:center;gap:16px"><span><span class="pdotx" style="background:var(--green)"></span> הכנסות</span><span><span class="pdotx" style="background:var(--blue)"></span> הוצאות</span></div></div>` : ''}`;
}
// fixed charges without a payment method are shown as their own line, not as "לא צוין"
const payKey = t => t.pay || (t.src === 'fixed' ? 'קבועות (בלי אמצעי תשלום)' : t.src === 'income' ? 'הכנסה צפויה' : 'לא צוין');
function payCard(md) {
  const by = {};
  for (const t of md.list) { if (!isExpT(t)) continue; const p = payKey(t); by[p] = by[p] || {s: 0, n: 0}; by[p].s += t.amount; by[p].n++; }
  const rows = Object.entries(by).sort((a, b) => b[1].s - a[1].s);
  if (!rows.length) return '';
  return `<div class="label">הוצאות לפי אמצעי תשלום</div><div class="card" style="padding:4px 14px">${rows.map(([p, v]) => `<button class="txline" data-act="pay" data-p="${esc(p)}" style="align-items:center"><div><b>${esc(p)}</b><div class="d">${v.n} תנועות</div></div><div class="a">${ils(v.s)}</div></button>`).join('')}</div>`;
}

/* ---- data: every line of the month as a readable list ---- */
function viewData() {
  const f = S.f, d = dash(S.month);
  let list = d.md.list.filter(t => !t.pending || f.src === 'pending');
  if (f.src === 'pending') list = d.md.list.filter(t => t.pending);
  else if (f.src) list = list.filter(t => t.src === f.src || (f.src === 'fixed' && t.src === 'income'));
  if (f.g) list = list.filter(t => catById(t.cat) && catById(t.cat).group === f.g);
  if (f.pay) list = list.filter(t => payKey(t) === f.pay);
  if (f.who) list = list.filter(t => t.person === f.who);
  const q = f.q.trim().toLowerCase();
  if (q) list = list.filter(t => `${t.desc} ${t.text || ''} ${catById(t.cat)?.name || ''} ${t.pay} ${t.note || ''} ${Math.abs(t.amount)} ${t.person || ''}`.toLowerCase().includes(q));
  list = list.slice().sort(f.sort === 'amt' ? (a, b) => Math.abs(b.amount) - Math.abs(a.amount) : (a, b) => b.date.localeCompare(a.date) || String(b.time || '').localeCompare(String(a.time || '')));
  const expSum = sumA(list.filter(isExpT)), incSum = sumA(list.filter(isIncT));
  const days = []; for (const t of list) { if (!days.length || days[days.length - 1].date !== t.date) days.push({date: t.date, items: []}); days[days.length - 1].items.push(t); }
  const pays = [...new Set(d.md.list.map(payKey))];
  const chip = (k, v, label) => `<button class="${f[k] === v ? 'on' : ''}" data-act="f" data-k="${k}" data-v="${esc(v)}">${label}</button>`;
  const sel = (k, opts, all) => `<select class="inp" data-fsel="${k}" style="padding:9px 12px;font-size:13.5px"><option value="">${all}</option>${opts.map(([v, l]) => `<option value="${esc(v)}" ${f[k] === v ? 'selected' : ''}>${esc(l)}</option>`).join('')}</select>`;
  return `<h1 class="hero" style="margin-bottom:10px">כל <span class="script blue">התנועות</span><br>ב${mName(S.month)}</h1>
    <div class="toolbar"><input id="dq" placeholder="חיפוש…" value="${esc(f.q)}"></div>
    <div class="chips" style="margin-bottom:8px">${chip('src', '', 'הכל')}${chip('src', 'chat', '💬 ווטסאפ')}${chip('src', 'fixed', '🔁 קבועות')}${chip('src', 'manual', '✍️ ידני')}${d.md.pend ? chip('src', 'pending', `בלי סעיף (${d.md.pend})`) : ''}</div>
    <div class="row2" style="grid-template-columns:1fr 1fr 1fr;gap:6px;margin-bottom:10px">${sel('g', D.cfg.groups.map(g => [g.id, g.name]), 'כל הקבוצות')}${sel('pay', pays.map(p => [p, p]), 'כל אמצעי התשלום')}${sel('who', people().map(p => [p, p]), 'כולם')}</div>
    <div class="card fl-row" style="padding:12px 16px;margin-bottom:10px"><span class="note">${list.length} תנועות</span><span><b>${ils(expSum)}</b> הוצאות${incSum ? ` · <b class="green">${ils(incSum)}</b> הכנסות` : ''}</span></div>
    <div class="chips" style="margin-bottom:8px">${chip('sort', '', 'לפי תאריך')}${chip('sort', 'amt', 'לפי סכום')}</div>
    ${!list.length ? '<p class="note" style="text-align:center">אין תנועות</p>' : f.sort === 'amt' ? `<div class="card" style="padding:2px 14px">${list.map(t => txLine(t)).join('')}</div>`
      : days.map(dy => `<div class="gname" style="display:flex;justify-content:space-between"><span>${dShort(dy.date)}</span><span>${ils(sumA(dy.items.filter(isExpT)))}</span></div><div class="card" style="padding:2px 14px;margin-bottom:6px">${dy.items.map(t => txLine(t, {compact: true})).join('')}</div>`).join('')}`;
}
function txLine(t, o) {
  o = o && typeof o === 'object' ? o : {};
  const c = catById(t.cat);
  const st = t.src === 'income' ? (t.ok ? '<span class="tag ok">נכנס</span>' : '<span class="tag y">צפוי · לעדכן</span>') : t.src === 'fixed' ? (t.ok ? '<span class="tag ok">ירד</span>' : '<span class="tag y">צפוי</span>') : t.refundExp ? (t.refundDone ? '<span class="tag ok">החזר התקבל</span>' : '<span class="tag y">צפוי החזר</span>') : '';
  return `<button class="txline${o.compact ? ' cmp' : ''}" data-act="tx" data-id="${esc(t.id)}">
    <div><div class="a ${isIncome(c) ? 'green' : ''}">${t.amount < 0 ? '−' : ''}${ils(Math.abs(t.amount))} <span class="src">${SRC_ICON[t.src]}</span> ${st}</div>
      <div class="d">${esc(t.desc || '')}${c ? ` · ${esc(c.name)}` : ' · <b class="red">בלי סעיף</b>'}${t.note ? `<br>📝 ${esc(t.note)}` : ''}</div></div>
    <div class="r">${o.compact ? '' : dShort(t.date) + '<br>'}${esc(t.pay || '')}${t.person ? `${o.compact && t.pay ? ' · ' : '<br>'}${esc(t.person)}` : ''}</div>
  </button>`;
}

/* ---- review ---- */
function viewReview() {
  const pc = pendingChat(), of = dueFixed(), rf = openRefunds(), dp = dupPairs(), mn = multiNums(), ra = rfAsks();
  const fa0 = D.meta.fxAsk || [], fc0 = D.meta.fxChanges || [];
  const todo = pc.length + of.length + rf.length + dp.length + fa0.length + (fc0.length ? 1 : 0) + mn.length + ra.length;
  const seg = `<div class="seg rvseg"><button class="${S.rv !== 'fixed' ? 'on' : ''}" data-act="rvseg" data-v="todo">לטיפול${todo ? ` (${todo})` : ''}</button><button class="${S.rv === 'fixed' ? 'on' : ''}" data-act="rvseg" data-v="fixed">קבועות ${mName(S.month)}</button></div>`;
  if (S.rv === 'fixed') return `<h1 class="hero">הוצאות <span class="script blue">קבועות</span></h1>${seg}${fixedByGroup(S.month)}`;
  const fa = D.meta.fxAsk || [], fc = D.meta.fxChanges || [];
  const cm = [...new Set([...pc, ...ra].filter(t => isClosed(t.month)).map(t => sheetName(t.month)))];
  return `<h1 class="hero">מה <span class="script blue">לבדוק</span>?</h1>${seg}
    ${cm.length ? `<div class="card" style="padding:10px 14px;margin-bottom:10px;font-size:13.5px;line-height:1.55">יש כאן שורות של כסף שנכנס מחודשים שכבר נסגרו באקסל (${cm.join(', ')}) – הכלל של ההחזרים שינה אותן אחרי הסגירה. אחרי שבוחרים להן סעיף: ייצוא לאקסל שכולל את החודשים האלה, ובאקסל "ייבוא מאיפה הכסף" – הוא יראה מה עובר ויבקש אישור.</div>` : ''}
    ${mn.length ? `<div class="label">לבדוק סכום – יש בהודעה יותר ממספר אחד (${mn.length})</div>${mn.map(t => `<div class="card rv"><div class="meta"><span>${dShort(t.date)} ${t.time || ''} · ${esc(t.person || '')}</span><span class="amt">${ils(Math.abs(t.amount))}</span></div>
        <div class="t">${esc(t.text)}</div>
        <div class="chips"><button data-act="mn-ok" data-id="${esc(t.id)}">הסכום נכון: ${ils(Math.abs(t.amount))}</button><button data-act="mn-fix" data-id="${esc(t.id)}" style="color:var(--blue);font-weight:600">לתקן סכום</button></div></div>`).join('')}` : ''}
    ${fa.length ? `<div class="label">זה אותו חיוב? (${fa.length})</div>${fa.map(a => { const f = D.fixed.find(x => x.id === a.fid) || {}; return `<div class="card rv">
        <div class="t">מהאקסל: <b>${esc(a.name)}</b> – ${ils(a.amount)} (${esc(FX_TYPE[a.type] || a.type)})<br>תנועה קבועה שהגדרת: <b>${esc(f.name || '')}</b> – ${ils(num(f.amount))}</div>
        <div class="chips"><button data-act="fxa-yes" data-k="${esc(a.key)}">כן, אותו חיוב – לחבר</button><button data-act="fxa-no" data-k="${esc(a.key)}" style="color:var(--muted)">לא, חיוב אחר</button></div></div>`; }).join('')}` : ''}
    ${fc.length ? `<div class="label">מה השתנה בחיובים הקבועים</div><div class="card">${fc.map(x => `<div class="note" style="margin:4px 0">${fxChangeLine(x)}</div>`).join('')}
        <div class="chips" style="margin-top:6px"><button data-act="fxc-seen">ראיתי</button><button data-act="fixed" style="color:var(--blue)">לרשימת הקבועות</button></div></div>` : ''}
    ${dp.length ? `<div class="label">אולי נרשם פעמיים? (${dp.length})</div>${dp.map(([a, b]) => `<div class="card rv"><div class="meta"><span>${dShort(a.date)}</span><span class="amt">${ils(Math.abs(a.amount))}</span></div>
        <div class="t">• ${esc(a.person)}: ${esc(a.text)}<br>• ${esc(b.person)}: ${esc(b.text)}</div>
        <div class="chips"><button data-act="dup-skip" data-id="${esc(b.id)}">זה כפול – לספור פעם אחת</button><button data-act="dup-ok" data-a="${esc(a.id)}" data-b="${esc(b.id)}" style="color:var(--muted)">שתי הוצאות שונות</button></div></div>`).join('')}` : ''}
    ${ra.length ? `<div class="label">על איזו הוצאה ההחזר? (${ra.length})</div>${ra.map(t => `<div class="card rv"><div class="meta"><span>${dShort(t.date)} ${t.time || ''} · ${esc(t.person || '')}${closedTag(t)}</span><span class="amt green">+${ils(Math.abs(t.amount))}</span></div>
        <div class="t">${esc(t.text)}</div>
        <div class="chips">${t.rfAsk.map(e => `<button data-act="rf-link" data-id="${esc(t.id)}" data-e="${esc(e)}">${esc(rfLabel(e))}</button>`).join('')}
          ${incCats().map(c => `<button data-act="rf-inc" data-id="${esc(t.id)}" data-c="${esc(c.id)}" class="incchip">+ ${esc(c.name)}</button>`).join('')}
          <button data-act="rf-none" data-id="${esc(t.id)}" style="color:var(--muted)">לא קשור להוצאה – לבחור סעיף</button></div></div>`).join('')}` : ''}
    ${pc.length ? `<div class="label">לבחור סעיף (${pc.length}${pendGroups().length < pc.length ? ` שורות ב-${pendGroups().length} קבוצות` : ''})</div>${pendGroups().map(rvGroup).join('')}` : ''}
    ${of.length ? `<div class="label">קבועות והכנסות לאישור – כמה ירד / נכנס בפועל?</div><div class="card">${of.map(f => `<div class="txline" style="align-items:center"><div><div class="a">${ils(f.amount)}</div><div class="d">${esc(f.desc)} · ${esc(catById(f.cat)?.name || '')} · ${dShort(f.date)}</div></div>
        <div style="display:flex;gap:6px"><button class="smallbtn" data-act="fx-ok" data-k="${esc(f.key)}">ירד ✓</button><button class="smallbtn" data-act="tx" data-id="${esc(f.id)}">סכום אחר</button></div></div>`).join('')}</div>` : ''}
    ${rf.length ? `<div class="label">מחכות להחזר (${rf.length})</div><div class="card">${rf.map(t => `<div class="txline" style="align-items:center"><div><div class="a">${ils(t.refundAmt || t.amount)}${t.refundGot ? ` <span class="note">· חזרו ${ils(t.refundGot)}</span>` : ''}</div><div class="d">${esc(t.text)} · ${dShort(t.date)}</div></div>
        <button class="smallbtn" data-act="rf-done" data-id="${esc(t.id)}">התקבל ✓</button></div>`).join('')}</div>
        <p class="note" style="margin:6px">כשהחזר נכנס, כותבים אותו בקבוצה ("התקבל 445 החזר"). הוא מתחבר לבד להוצאה ויורד ממנה כמינוס, גם אם הגיע בכמה הודעות. כשחוזר כמעט כל הסכום ההוצאה נסגרת לבד; כאן מסמנים רק החזר חלקי שלא יגיע ממנו יותר.</p>` : ''}
    ${!todo ? `<div class="empty">${I.flower}<h2>הכל <span class="script">מסודר</span></h2><p class="note">אין כרגע מה לבדוק.</p><button class="ghost" data-act="rvseg" data-v="fixed" style="max-width:280px;margin:8px auto 0">לקבועות של ${mName(S.month)}</button></div>` : ''}`;
}
// every fixed expense of the month, grouped by category (group).
// What already went out is shown at its actual amount; what has not yet is shown as expected.
// A month closed in Excel shows the exact amounts from Excel (per section).
function fixedRows(m) {
  const fxCats = new Set(fixedFor(m).map(t => t.cat).filter(Boolean));
  const isFx = c => c && !isIncome(c) && (c.kind === 'fixed' || fxCats.has(c.id));
  const md = monthData(m), td = todayStr();
  if (md.closed) {
    return cats().filter(isFx).map(c => ({cat: c.id, desc: c.name, amount: md.byCat[c.id] || 0, st: 'xl', fid: ''}))
      .filter(r => r.amount || fxCats.has(r.cat));
  }
  return md.list.filter(t => !t.pending && isFx(catById(t.cat))).map(t => t.src === 'fixed'
    ? {cat: t.cat, desc: t.desc, amount: t.amount, st: t.ok ? 'ok' : t.date <= td ? 'due' : 'exp', day: +t.date.slice(8), fid: t.fid}
    : {cat: t.cat, desc: t.desc || t.text || '', amount: t.amount, st: 'ok', day: +t.date.slice(8), id: t.id});
}
function fixedByGroup(m) {
  const list = fixedRows(m), closed = isClosed(m);
  const head = `<div class="label">הוצאות קבועות – ${mName(m)}${closed ? ' · לפי האקסל (נסגר)' : ''}</div>`;
  if (!list.length) return `${head}<div class="card"><p class="note" style="margin:4px 0">אין הוצאות קבועות בחודש הזה.</p><div class="chips"><button data-act="fixed" style="color:var(--blue)">לרשימת הקבועות</button></div></div>`;
  const by = new Map();
  for (const t of list) { const c = catById(t.cat), g = c && D.cfg.groups.find(x => x.id === c.group); const k = g ? g.id : ''; if (!by.has(k)) by.set(k, {g, items: []}); by.get(k).items.push(t); }
  const blocks = [...by.values()].map(b => ({...b, sum: sumA(b.items)})).sort((a, b) => b.sum - a.sum);
  // the same rule as the home screen: a charge whose day has passed counts as gone out, even before it is confirmed
  const went = sumA(list.filter(t => t.st !== 'exp')), due = sumA(list.filter(t => t.st === 'due')), left = sumA(list.filter(t => t.st === 'exp'));
  const ST = {ok: '<span style="color:var(--green)">ירד ✓</span>', xl: '<span style="color:var(--green)">בפועל ✓</span>', due: 'לאישור', exp: 'צפוי'};
  return `${head}
    <div class="card" style="padding:10px 14px;margin-bottom:10px;display:flex;justify-content:space-between;font-size:13.5px"><span>${closed ? 'ירד בפועל' : 'ירד עד היום'} <b>${ils(went, 0)}</b>${due ? `<br><span class="note">מתוכם ${ils(due, 0)} לאישור</span>` : ''}</span>${closed || !left ? '' : `<span>צפוי בהמשך <b>${ils(left, 0)}</b></span><span style="color:var(--muted)">סה"כ ${ils(went + left, 0)}</span>`}</div>
    ${blocks.map(b => { const gs = b.g ? groupStyle(b.g) : null; return `<div class="card fxg">
      <div class="fxg-h">${gs ? `<span class="gicon" style="background:${gs.bg}">${gs.icon}</span>` : ''}<b>${b.g ? esc(b.g.name) : 'בלי סעיף'}</b><span>${ils(b.sum, 0)}</span></div>
      ${b.items.sort((x, y) => (x.day || 0) - (y.day || 0)).map(t => `<button class="txline" ${t.fid ? `data-act="fx-edit" data-id="${esc(t.fid)}"` : t.id ? `data-act="tx" data-id="${esc(t.id)}"` : 'data-act="cat" data-id="' + esc(t.cat) + '"'}><div><div class="a">${ils(t.amount)}</div><div class="d">${esc(t.desc)}${t.st === 'xl' ? '' : ' · ' + esc(catById(t.cat)?.name || '?')}</div></div>
        <div class="r">${ST[t.st]}${t.day ? `<br>ב-${t.day} לחודש` : ''}</div></button>`).join('')}</div>`; }).join('')}`;
}
// lines with the same description are sorted together: one choice for all of them
const descKey = t => (t.incoming ? '+' : '-') + String(t.desc || '').toLowerCase().replace(/[\d.,₪"'()\-–]/g, ' ').replace(/\s+/g, ' ').trim();
function pendGroups() {
  return cached('pgroups', () => {
    const by = new Map();
    for (const t of pendingChat()) { const k = descKey(t); if (!by.has(k)) by.set(k, {key: k, items: []}); by.get(k).items.push(t); }
    return [...by.values()].sort((a, b) => b.items.length - a.items.length || a.items[0].date.localeCompare(b.items[0].date));
  });
}
const learnWords = t => String(t.desc || '').replace(/[()"]/g, ' ').split(/\s+/).filter(w => w.length >= 2 && !/^\d/.test(w) && !/^(התקבל|החזר|העברה|בביט|ביט|לפייבוקס|בפייבוקס|פייבוקס|מזומן|עבור|של|על|את|לבנק)$/.test(w));
function rvGroup(g) {
  const t = g.items[0], n = g.items.length, word = learnWords(t).slice(0, 2).join(' '), gk = esc(g.key);
  const lines = g.items.slice(0, 3).map(x => `<div class="rvl"><span>${dShort(x.date)} · ${esc(x.person || '')}${closedTag(x)}</span><span>${x.incoming ? '+' : ''}${ils(Math.abs(x.amount))}</span></div>`).join('');
  return `<div class="card rv">
    ${n === 1 ? `<div class="meta"><span>${dShort(t.date)} ${t.time || ''} · ${esc(t.person || '')}${closedTag(t)}</span><span class="amt ${t.incoming ? 'green' : ''}">${t.incoming ? '+' : ''}${ils(Math.abs(t.amount))}</span></div><div class="t">${esc(t.text)}</div>`
      : `<div class="meta"><span><b style="font-size:16px;color:var(--ink)">${esc(t.desc || t.text)}</b> ×${n}</span><span class="amt">${ils(Math.abs(sumA(g.items)))}</span></div>${lines}${n > 3 ? `<div class="note" style="margin:2px 0 0">ועוד ${n - 3}</div>` : ''}`}
    <div class="chips">${quickCats(t).map(c => `<button data-act="rvg-cat" data-g="${gk}" data-c="${c.id}"${incChip(c)}>${isIncome(c) ? '+ ' : ''}${esc(c.name)}</button>`).join('')}
      <button data-act="rvg-more" data-g="${gk}" style="color:var(--blue);font-weight:600">סעיף אחר…</button>
      <button data-act="rvg-skip" data-g="${gk}" style="color:var(--muted)">לא לייבא</button></div>
    ${word ? `<label class="rvlearn"><input type="checkbox" data-learn="${gk}" ${S.noLearn && S.noLearn.has(g.key) ? '' : 'checked'}> לזכור: <input class="inp" data-lw="${gk}" value="${esc(S.lw && S.lw[g.key] != null ? S.lw[g.key] : word)}"></label>` : ''}
  </div>`;
}
function multiNums() { return cached('multi', () => D.chat.filter(it => it.multi && !(D.ov[it.id] || {}).numOk).map(it => chatTx(it)).filter(t => t && !t.skipped && t.date >= D.cfg.start && !isClosed(t.month))); }
// sections that fit this line: from keywords and from lines with shared words that already have a section, then the most used
// sections that fit this line: for money that came in – the likely expense sections (a refund) and also the income sections
function quickCats(t) {
  const ok = c => c && !c.archived && !isIncome(c);
  const idx = cached('wordIdx', () => {
    const w = new Map(), cnt = {};
    for (const it of D.chat) { const c = (D.ov[it.id] || {}).cat || it.cat; if (!c) continue; cnt[c] = (cnt[c] || 0) + 1;
      for (const x of learnWords(it)) { const k = x.toLowerCase(); if (k.length < 3) continue; const m = w.get(k) || w.set(k, {}).get(k); m[c] = (m[c] || 0) + 1; } }
    return {w, cnt};
  });
  const score = {};
  for (const x of learnWords(t)) { const k = x.toLowerCase(); const m = idx.w.get(k); if (m) for (const c in m) score[c] = (score[c] || 0) + 10 * m[c];
    for (const [kw, c] of D.kw) { const kk = String(kw).toLowerCase().trim(); if (kk.length >= 3 && (kk.includes(k) || k.includes(kk))) score[c] = (score[c] || 0) + 25; } }
  for (const c in idx.cnt) score[c] = (score[c] || 0) + idx.cnt[c] / 1000;
  const by = (a, b) => (score[b.id] || 0) - (score[a.id] || 0);
  const list = cats().filter(ok).sort(by);
  if (!t.incoming) return list.slice(0, 6);
  return [...list.slice(0, 4), ...incCats(score)];
}
const incCats = score => cats().filter(c => isIncome(c) && !c.archived).sort((a, b) => ((score || {})[b.id] || 0) - ((score || {})[a.id] || 0)).slice(0, 4);
const incChip = c => isIncome(c) ? ' class="incchip"' : '';
function groupPick(key, cat) {
  const g = pendGroups().find(x => x.key === key); if (!g) return;
  const learn = $(`[data-learn="${CSS.escape(key)}"]`), word = ($(`[data-lw="${CSS.escape(key)}"]`) || {}).value || '';
  const inc = cat !== SKIP && isIncome(catById(cat));
  for (const t of g.items) {
    D.ov[t.id] = {...(D.ov[t.id] || {}), ...(cat === SKIP ? {skip: true} : {cat})};
    if (inc && t.incoming) D.ov[t.id].rfOf = 'none';
    else if (t.rfPre && cat !== SKIP) D.ov[t.rfOf] = {...(D.ov[t.rfOf] || {}), cat};
  }
  let msg = cat === SKIP ? (g.items.length === 1 ? 'השורה לא תיספר' : `${g.items.length} שורות לא ייספרו`) : (g.items.length === 1 ? `← ${catById(cat).name}` : `${g.items.length} שורות ← ${catById(cat).name}`);
  if (learn && learn.checked && word.trim()) { const w = word.trim(); D.kw = D.kw.filter(([k]) => k !== w); D.kw.push([w, cat]); recategorize(); msg += ` · "${w}" נלמד`; }
  save(); render(); toast(msg);
}
function setChatCat(id, cat) {
  const it = chatItem(id); if (!it) return;
  D.ov[id] = {...(D.ov[id] || {}), cat};
  const L = refundLinks(), ex = it.incoming && L.of[id];
  // tied to an expense from before the budget start: the section goes to that expense (its other refunds follow);
  // tied to an expense inside the budget: a section chosen by hand unties it
  if (it.incoming && isIncome(catById(cat))) D.ov[id].rfOf = 'none';     // income: not a refund of anything
  else if (ex && L.pre[ex]) D.ov[ex] = {...(D.ov[ex] || {}), cat};
  else if (ex) D.ov[id].rfOf = 'none';
  save(); render(); refreshSheets();
  if (!it.incoming) learnSheet(it, cat);
}
function learnSheet(it, cat) {
  const words = it.desc.replace(/[()"]/g, ' ').split(/\s+/).filter(w => w.length >= 2 && !/^\d/.test(w) && !/^(התקבל|החזר|העברה|בביט|ביט|לפייבוקס|בפייבוקס|פייבוקס|מזומן|עבור|של|על)$/.test(w));
  const guess = words.slice(0, 2).join(' ');
  if (!guess || suggestCat(it.desc, false) === cat) return;
  openSheet(() => `${head('ללמוד לפעם הבאה?')}
    <p class="note">"${esc(it.text)}" ← <b>${esc(catById(cat).name)}</b></p>
    <label class="field"><span>איזו מילה תזהה בפעם הבאה את הסעיף הזה?</span><input class="inp" id="lw" value="${esc(guess)}"></label>
    <div class="chips">${words.slice(0, 6).map(w => `<button data-w="${esc(w)}">${esc(w)}</button>`).join('')}</div>
    <button class="primary" id="lw-ok">לזכור</button><button class="ghost" data-act="close">לא הפעם</button>`, {
    mount(sh) {
      sh.addEventListener('click', ev => { const b = ev.target.closest('[data-w]'); if (b) $('#lw', sh).value = b.dataset.w; if (ev.target.closest('#lw-ok')) {
        const w = $('#lw', sh).value.trim(); if (!w) return;
        D.kw = D.kw.filter(([k]) => k !== w); D.kw.push([w, cat]);
        recategorize(); save(); closeSheet(); render(); toast(`מעכשיו "${w}" ← ${catById(cat).name}`);
      } });
    }
  });
}
function learnSkipSheet(it) {
  const words = it.desc.replace(/[()"]/g, ' ').split(/\s+/).filter(w => w.length >= 2 && !/^\d/.test(w) && !/^(את|של|על|עבור|לבנק|העברה|ביט|בביט|מזומן)$/.test(w));
  openSheet(() => `${head('השורה לא תיספר')}
    <p class="note">"${esc(it.text)}"</p>
    <p style="line-height:1.5">לדלג לבד גם בפעם הבאה על שורות שיש בהן את המילה:</p>
    <input class="inp" id="sk-w" value="${esc(words[0] || '')}">
    <div class="chips" style="margin-top:8px">${words.slice(0, 6).map(w => `<button data-w="${esc(w)}">${esc(w)}</button>`).join('')}</div>
    <button class="primary" id="sk-ok">כן, לדלג תמיד</button><button class="ghost" data-act="close">רק הפעם</button>
    <p class="note">למשל "הפקדה" – כשמפקידים מזומן שכבר נרשם כהכנסה. אפשר לבטל ב"מילות מפתח".</p>`, {
    mount(sh) { sh.addEventListener('click', ev => { const b = ev.target.closest('[data-w]'); if (b) $('#sk-w', sh).value = b.dataset.w;
      if (ev.target.closest('#sk-ok')) { const w = $('#sk-w', sh).value.trim(); if (!w) return; D.kw = D.kw.filter(([k]) => k !== w); D.kw.push([w, SKIP]); recategorize(); save(); closeSheet(); render(); toast(`שורות עם "${w}" לא ייספרו`); } }); }
  });
}
function recategorize() {
  for (const it of D.chat) {
    if (it.incoming) it.cat = inCat(it.text, it.desc);
    else it.cat = suggestCat(it.desc, false) || '';
    it.autoSkip = it.cat === SKIP; if (it.autoSkip) it.cat = '';
  }
  for (const it of D.rfPre || []) { it.cat = suggestCat(it.desc, false) || ''; if (it.cat === SKIP) it.cat = ''; }
}

/* =============== sheets =============== */
const stack = [];
function openSheet(build, opts = {}) {
  const ov = document.createElement('div'); ov.className = 'ov';
  ov.innerHTML = `<div class="sheet" role="dialog"></div>`;
  document.body.appendChild(ov);
  const sh = ov.firstChild, e = {ov, sh, build, opts};
  stack.push(e); e.paint = () => { sh.innerHTML = build(e); };
  e.paint(); opts.mount && opts.mount(sh, e);
  requestAnimationFrame(() => ov.classList.add('on'));
  ov.addEventListener('click', ev => { if (ev.target === ov) closeSheet(e); });
  document.body.style.overflow = 'hidden';
  return e;
}
function closeSheet(e) {
  e = e || stack[stack.length - 1]; if (!e) return;
  const i = stack.indexOf(e); if (i >= 0) stack.splice(i, 1);
  e.ov.classList.remove('on'); setTimeout(() => e.ov.remove(), 200);
  if (!stack.length) document.body.style.overflow = '';
  e.opts.onClose && e.opts.onClose();
}
function refreshSheets() { for (const e of stack) if (e.opts.live && !e.sh.contains(document.activeElement)) e.paint(); }
const head = t => `<div class="sh-head"><h3>${t}</h3><button class="x-btn" data-act="close" aria-label="סגירה">✕</button></div>`;
function findTx(id) {
  if (id.startsWith('i:')) return entryByKey('inc:' + id.slice(2));
  if (id.startsWith('f:')) { const [fid, m] = id.slice(2).split('|'); return fixedFor(m).find(x => x.fid === fid); }
  const it = chatItem(id); if (it) return chatTx(it, true);
  const mx = D.manual.find(x => x.id === id); return mx ? manualTx(mx) : null;
}
function pickCat(title, income, cur, onPick, withSkip) {
  openSheet(() => `${head(title)}${withSkip ? `<div class="chips" style="margin-bottom:6px"><button data-pc="${SKIP}" style="color:var(--red)">⊘ לא לייבא</button></div>` : ''}${D.cfg.groups.map(g => {
    const gc = cats().filter(c => c.group === g.id && !c.archived && (income === null || isIncome(c) === income));
    return gc.length ? `<div class="gname">${esc(g.name)}</div><div class="chips">${gc.map(c => `<button data-pc="${esc(c.id)}" class="${cur === c.id ? 'on' : ''}">${esc(c.name)}</button>`).join('')}</div>` : '';
  }).join('')}`, {mount(sh) { sh.addEventListener('click', ev => { const b = ev.target.closest('[data-pc]'); if (b) { closeSheet(); onPick(b.dataset.pc); } }); }});
}
function askAmount(title, val, onOk) {
  openSheet(() => `${head(title)}<div class="amtin"><input id="aa" inputmode="decimal" value="${esc(val)}"><span>₪</span></div><button class="primary" id="aa-ok">שמירה</button>`, {
    mount(sh) { setTimeout(() => $('#aa', sh).select(), 250); const go = () => { const v = num($('#aa', sh).value); closeSheet(); onOk(v); };
      sh.addEventListener('click', ev => { if (ev.target.closest('#aa-ok')) go(); }); sh.addEventListener('keydown', ev => { if (ev.key === 'Enter') go(); }); }
  });
}
function askText(title, val, onOk) {
  openSheet(() => `${head(title)}<textarea class="inp" id="at" rows="3">${esc(val)}</textarea><button class="primary" id="at-ok">שמירה</button>`, {
    mount(sh) { setTimeout(() => $('#at', sh).focus(), 250); sh.addEventListener('click', ev => { if (ev.target.closest('#at-ok')) { const v = $('#at', sh).value.trim(); closeSheet(); onOk(v); } }); }
  });
}

/* ---- category detail ---- */
function catSheet(id) {
  openSheet(() => {
    const c = catById(id); if (!c) return head('הסעיף לא נמצא');
    const md = monthData(S.month), list = md.list.filter(t => t.cat === id && !t.pending), tot = md.closed ? (md.byCat[id] || 0) : sumA(list);
    const ms = []; for (let i = 5; i >= 0; i--) ms.push(addM(S.month, -i));
    const vals = ms.map(m => monthData(m).byCat[id] || 0), mx = Math.max(1, ...vals.map(Math.abs));
    const prev3 = [1, 2, 3].map(k => monthData(addM(S.month, -k)).byCat[id] || 0), avg = prev3.reduce((a, b) => a + b, 0) / 3;
    return `${head(esc(c.name))}
      <div class="card"><div class="chip">${esc(groupById(c.group)?.name || '')} · ${mLabel(S.month)}</div>
        <div class="fl-row" style="margin-top:10px"><div><div class="chip">${isIncome(c) ? 'נכנס' : 'יצא'}</div><div class="num ${isIncome(c) ? 'green' : ''}">${ils(tot)}</div></div>
        <div style="text-align:left"><div class="chip">ממוצע 3 חודשים קודמים</div><div class="num" style="font-size:17px">${avg ? ils(avg, 0) : '—'}</div></div></div></div>
      ${md.closed ? '<p class="note" style="margin:8px 4px">החודש נסגר באקסל: הסכום למעלה מדויק. הרשימה למטה היא מה שנרשם בקבוצה.</p>' : ''}
      <div class="card" style="margin-top:10px">${list.length ? list.map(txLine).join('') : `<p class="note" style="text-align:center">אין תנועות ב${mName(S.month)}</p>`}</div>
      <div class="row2" style="margin-top:6px"><button class="ghost" data-act="manual" data-id="${esc(id)}">+ הוספה ידנית</button><button class="ghost" data-act="catedit" data-id="${esc(id)}">שם והגדרות</button></div>
      <div class="label">ששת החודשים האחרונים</div><div class="card" style="display:flex;align-items:flex-end;gap:8px;height:150px;padding-bottom:10px">
      ${ms.map((m, i) => `<div style="flex:1;text-align:center;font-size:12px"><div style="font-weight:600;margin-bottom:3px">${vals[i] ? fmt(vals[i], 0) : ''}</div><div style="height:${Math.max(3, Math.abs(vals[i]) / mx * 80)}px;background:${m === S.month ? 'var(--blue)' : 'var(--blue-soft)'};border-radius:7px"></div><div style="margin-top:4px;color:var(--muted)">${sheetName(m)}</div></div>`).join('')}</div>`;
  }, {live: true, month: true});
}
function groupSheet(gid) {
  openSheet(() => {
    const g = groupById(gid), md = monthData(S.month);
    const gc = cats().filter(c => c.group === gid && md.byCat[c.id]).sort((a, b) => md.byCat[b.id] - md.byCat[a.id]);
    const tot = gc.reduce((s, c) => s + md.byCat[c.id], 0), mx = Math.max(1, ...gc.map(c => Math.abs(md.byCat[c.id])));
    const cnt = id => md.list.filter(t => t.cat === id && !t.pending).length;
    const all = dash(S.month).total, pc = (v, t) => t ? Math.round(v / t * 100) : 0;
    // first the share of all the month's expenses (same number base as the home screen), then the share of this group
    const pcts = c => isIncome(c) ? `${pc(md.byCat[c.id], tot)}% מ${esc(g.name)}` : `${pc(md.byCat[c.id], all)}% מכל ההוצאות · ${pc(md.byCat[c.id], tot)}% מ${esc(g.name)}`;
    return `${head(esc(g.name))}<div class="card fl-row"><span class="chip">${mLabel(S.month)}</span><span class="num">${ils(tot)}</span></div>
      <div class="card" style="margin-top:10px;padding:6px 14px">${gc.map(c => `<button class="catrow" data-act="cat" data-id="${esc(c.id)}"><div class="t"><b>${esc(c.name)}</b><b>${ils(md.byCat[c.id])}</b></div>
        <div class="bar ${isIncome(c) ? 'g' : ''}"><i style="width:${Math.abs(md.byCat[c.id]) / mx * 100}%"></i></div><div class="foot"><span>${cnt(c.id)} תנועות</span><span>${pcts(c)}</span></div></button>`).join('') || '<p class="note">אין תנועות</p>'}</div>`;
  }, {live: true, month: true});
}
function catEditSheet(id) {
  const c = catById(id); if (!c) return;
  openSheet(() => `${head('שם והגדרות')}
    <label class="field"><span>שם הסעיף (כמו בקובץ האקסל)</span><input class="inp" id="ce-name" value="${esc(c.name)}"></label>
    ${isIncome(c) ? `<label class="field"><span>הכנסה צפויה בכל חודש ₪</span><input class="inp" id="ce-bud" inputmode="decimal" value="${num(c.budget) || ''}"></label>` : ''}
    <button class="primary" id="ce-save">שמירה</button>
    <button class="ghost danger" id="ce-arch">${c.archived ? 'להחזיר את הסעיף' : 'להסתיר את הסעיף'}</button>
    <p class="note">שם הסעיף צריך להיות זהה לשם בגיליון "רשימות" באקסל, כדי שהייצוא ייכנס לעמודה הנכונה.</p>`, {
    mount(sh) {
      sh.addEventListener('click', ev => {
        if (ev.target.closest('#ce-save')) {
          const name = $('#ce-name', sh).value.trim(); if (!name) return toast('צריך שם');
          if (cats().some(x => x.id !== id && x.name === name)) return toast('כבר יש סעיף בשם הזה');
          c.name = name; if (isIncome(c)) c.budget = num($('#ce-bud', sh).value);
          save(); closeSheet(); render(); refreshSheets(); toast('נשמר');
        }
        if (ev.target.closest('#ce-arch')) { c.archived = !c.archived; save(); closeSheet(); render(); refreshSheets(); }
      });
    }
  });
}

// the income of the month by section: what came in and what is still expected; a month closed in Excel shows Excel's amounts
function incDetailSheet() {
  openSheet(() => {
    const m = S.month, d = dash(m), md = d.md, closed = md.closed;
    const ic = cats().filter(isIncome);
    // a closed month: the amount of every section is Excel's; under it – the lines recorded in the app that make it up,
    // and what is only in Excel (or only in the app) as a difference line
    const lines = c => md.list.filter(t => !t.pending && t.cat === c.id && (t.src !== 'income' || t.ok));
    const blocks = closed
      ? ic.filter(c => md.byCat[c.id] || lines(c).length).map(c => ({c, tot: md.byCat[c.id] || 0, items: lines(c)})).sort((a, b) => b.tot - a.tot)
          .map(b => { const df = b.tot - sumA(b.items);
            return `<div class="card fxg"><div class="fxg-h"><b>${esc(b.c.name)}</b><span>${ils(b.tot, 0)}</span></div>
              ${b.items.map(t => txLine(t)).join('')}
              ${Math.abs(df) >= 1 ? `<div class="txline" style="align-items:center"><div><div class="a">${df < 0 ? '−' : ''}${ils(Math.abs(df))}</div><div class="d">${b.items.length ? (df > 0 ? 'נרשם רק באקסל – לא באפליקציה' : 'נרשם באפליקציה ולא נכלל באקסל') : 'אין פירוט באפליקציה – הסכום מהאקסל'}</div></div><div class="r"><span class="tag">אקסל</span></div></div>` : ''}</div>`; })
      : ic.map(c => ({c, items: d.inc.filter(t => t.cat === c.id)})).filter(b => b.items.length).sort((a, b) => sumA(b.items) - sumA(a.items))
          .map(b => `<div class="card fxg"><div class="fxg-h"><b>${esc(b.c.name)}</b><span>${ils(sumA(b.items), 0)}</span></div>${b.items.map(t => txLine(t)).join('')}</div>`);
    const left = d.incExp - d.incIn;
    return `${head(`הכנסות – ${mName(m)}${closed ? ' · לפי האקסל' : ''}`)}
      <div class="card" style="padding:10px 14px;margin-bottom:10px;display:flex;justify-content:space-between;font-size:13.5px"><span>${closed ? 'נכנס בפועל' : 'נכנס עד היום'} <b>${ils(d.incIn, 0)}</b></span>${!closed && left > 0 ? `<span>עוד צפוי <b>${ils(left, 0)}</b></span><span style="color:var(--muted)">סה"כ ${ils(d.incExp, 0)}</span>` : ''}</div>
      ${closed && blocks.length ? '<p class="note" style="margin:0 4px 8px">החודש נסגר באקסל: הסכום של כל סעיף מדויק לפי האקסל, ומתחתיו מה שנרשם באפליקציה.</p>' : ''}
      ${blocks.length ? blocks.join('') : '<p class="note" style="text-align:center">אין הכנסות בחודש הזה</p>'}
      ${closed || !d.inc.some(t => t.src === 'income' && !t.ok) ? '' : '<p class="note" style="margin:8px 4px">הכנסה צפויה שעוד לא אושרה מסומנת "צפוי". לחיצה עליה – לאשר כמה נכנס בפועל.</p>'}
      <button class="ghost" data-act="income" style="margin-top:6px">הכנסות צפויות בכל חודש ⚙</button>`;
  }, {live: true, month: true});
}

// how the balance on the home screen is made: income (also what is still expected) − what went out − fixed charges still to come
function balSheet() {
  openSheet(() => {
    const m = S.month, d = dash(m), cur = m === curMonth(), future = m > curMonth(), past = !cur && !future;
    const rfOpen = cur || future ? openRefunds().filter(t => t.month <= m).reduce((s, t) => s + Math.max(0, (t.refundAmt || t.amount) - (t.refundGot || 0)), 0) : 0;
    const row = (label, v, sub, act) => `<button class="txline" ${act || ''} style="align-items:center"><div><b>${label}</b>${sub ? `<div class="d">${sub}</div>` : ''}</div><div class="a ${v > 0 ? 'green' : ''}">${v < 0 ? '− ' : v > 0 ? '+ ' : ''}${ils(Math.abs(v), 0)}</div></button>`;
    const title = past ? (d.bal >= 0 ? 'עודף' : 'חוסר') : (d.bal >= 0 ? 'יתרה צפויה' : 'חוסר צפוי');
    return `${head(`${title} – ${mName(m)}`)}
      <div class="card" style="padding:4px 14px">
        ${row('הכנסות', d.incExp, d.incIn < d.incExp ? `נכנס ${ils(d.incIn, 0)} · עוד צפוי ${ils(d.incExp - d.incIn, 0)}` : d.closed ? 'לפי האקסל' : 'הכל נכנס', 'data-act="incdet"')}
        ${row(past ? 'הוצאות' : 'יצא עד היום', -d.spent, d.closed ? 'לפי האקסל' : past ? 'כולל הקבועות' : 'כולל קבועות שהיום שלהן כבר עבר', 'data-act="view" data-v="data"')}
        ${past ? '' : row('קבועות שעוד צפויות לרדת', -d.upc, d.upc ? `${d.upcoming.length === 1 ? 'חיוב אחד' : d.upcoming.length + ' חיובים'} עד סוף החודש` : 'כל הקבועות כבר ירדו', 'data-act="view" data-v="review" data-rv="fixed"')}
        <div class="txline" style="align-items:center;border-top:2px solid var(--line, #eee)"><div><b>= ${title}</b></div><div class="a ${d.bal < 0 ? 'red' : 'green'}">${d.bal < 0 ? '−' : ''} ${ils(Math.abs(d.bal), 0)}</div></div>
      </div>
      ${past ? '' : `<p class="note" style="margin:10px 4px;line-height:1.6">לא נכללות: הוצאות משתנות שעוד לא נכתבו בקבוצה (קניות, דלק וכד' בהמשך החודש)${d.md.pend ? `, ושורות בלי סעיף (${d.md.pend})` : ''}.${rfOpen ? ` החזרים שעוד מחכים (${ils(rfOpen, 0)}) לא נכללים – הם ייכנסו כשיגיעו.` : ''}</p>`}`;
  }, {live: true, month: true});
}

/* ---- transaction detail ---- */
function txSheet(id) {
  openSheet(() => {
    const t = findTx(id); if (!t) return head('התנועה לא נמצאה');
    const c = catById(t.cat), o = D.ov[id] || {};
    const acts = t.src === 'chat' ? `
        <button data-act="t-cat" data-id="${esc(id)}">${c ? 'להעביר לסעיף אחר' : 'לבחור סעיף'} <span>⇄</span></button>
        <button data-act="t-amt" data-id="${esc(id)}">לתקן סכום <span>✎</span></button>
        <button data-act="t-pay" data-id="${esc(id)}">אמצעי תשלום <span class="v">${esc(t.pay || '—')}</span></button>
        <button data-act="t-note" data-id="${esc(id)}">${t.note ? 'לערוך הערה' : 'להוסיף הערה'} <span>💬</span></button>
        ${t.refundExp ? `<button data-act="rf-done" data-id="${esc(id)}">${t.refundDone ? 'לסמן שההחזר עוד לא התקבל' : 'ההחזר התקבל'} <span>✓</span></button>` : ''}
        ${t.rfOf ? `<button data-act="rf-none" data-id="${esc(id)}">ההחזר לא שייך להוצאה הזו <span>✂</span></button>` : rfIsIn(chatItem(id) || {}) && rfChoices(id).length ? `<button data-act="rf-pick" data-id="${esc(id)}">לחבר להוצאה שמחכה להחזר <span>↩︎</span></button>` : ''}
        ${o.cat != null || o.amount != null || o.pay != null || o.rfOf != null ? `<button data-act="t-reset" data-id="${esc(id)}">לבטל את התיקונים שלי <span>↺</span></button>` : ''}
        <button data-act="t-skip" data-id="${esc(id)}" class="danger">${t.skipped ? 'להחזיר לייבוא' : 'לא לייבא את השורה הזו'} <span>⊘</span></button>`
      : t.src === 'income' ? `
        <button data-act="fx-ok" data-k="${esc(t.key)}">${t.ok ? 'לסמן שעוד לא נכנס' : 'נכנס הסכום הצפוי'} <span>✓</span></button>
        <button data-act="fx-amt" data-k="${esc(t.key)}">נכנס סכום אחר <span>✎</span></button>
        <button data-act="catedit" data-id="${esc(t.cat)}">לשנות את ההכנסה הצפויה הקבועה <span>⚙</span></button>`
      : t.src === 'fixed' ? `
        <button data-act="fx-ok" data-k="${esc(t.key)}">${t.ok ? 'לסמן שעוד לא ירד' : 'ירד בפועל'} <span>✓</span></button>
        <button data-act="fx-amt" data-k="${esc(t.key)}">סכום אחר בחודש הזה <span>✎</span></button>
        <button data-act="fx-edit" data-id="${esc(t.fid)}">לערוך את התנועה הקבועה <span>🔁</span></button>`
      : `<button data-act="m-edit" data-id="${esc(id)}">לערוך <span>✎</span></button>
        <button data-act="m-del" data-id="${esc(id)}" class="danger">למחוק <span>🗑</span></button>`;
    return `${head('')}
      <div class="txhero"><div class="amt">${t.amount < 0 ? '−' : ''}${ils(Math.abs(t.amount))}</div>
        <div class="s">${c ? `<button class="catlink" data-act="cat" data-id="${esc(c.id)}">${esc(c.name)} ›</button>` : '<b><span class="red">בלי סעיף</span></b>'} · ${dShort(t.date)} ${t.time || ''}</div>
        <div class="s">${SRC_ICON[t.src]} ${t.src === 'chat' ? 'מהווטסאפ' : t.src === 'income' ? (t.ok ? 'הכנסה · אושרה' : 'הכנסה צפויה · עוד לא אושרה') : t.src === 'fixed' ? (t.ok ? 'קבועה · ירדה' : 'קבועה · צפויה') : 'הוזן ידנית'} | ${esc(t.pay || '—')}${t.person ? ' | ' + esc(t.person) : ''}</div>
        ${t.text ? `<div class="orig">${esc(t.text)}</div>` : `<div class="s">${esc(t.desc)}</div>`}
        ${t.note ? `<div class="s">📝 ${esc(t.note)}</div>` : ''}${t.refundExp ? `<div class="s">${t.refundDone ? '✓ ההחזר התקבל' : `צפוי החזר ${ils(t.refundAmt || t.amount)}`}${t.refundGot ? ` · חזרו ${ils(t.refundGot)}` : ''}</div>${(rfState(id).refunds).map(r => { const x = findTx(r); return x ? `<div class="s"><button class="catlink" data-act="tx" data-id="${esc(r)}">↩︎ ${ils(Math.abs(x.amount))} ב-${dShort(x.date)} ›</button></div>` : ''; }).join('')}` : ''}
        ${t.preBudget && t.refundExp ? `<div class="s">מלפני תחילת התקציב – לא נספרת. ההחזרים עליה יורדים מהסעיף שלה.</div>` : ''}
        ${t.rfOf ? `<div class="s"><button class="catlink" data-act="tx" data-id="${esc(t.rfOf)}">↩︎ החזר על: ${esc(rfLabel(t.rfOf))} ›</button></div>` : ''}</div>
      <div class="actlist">${c ? `<button data-act="cat" data-id="${esc(c.id)}">לסעיף "${esc(c.name)}" – כל התנועות בו <span>›</span></button>` : ''}${acts}</div>`;
  }, {live: true});
}
function setOv(id, patch) { D.ov[id] = {...(D.ov[id] || {}), ...patch}; for (const k in D.ov[id]) if (D.ov[id][k] == null) delete D.ov[id][k]; save(); render(); refreshSheets(); }
function setFixedM(k, patch) { D.fixedM[k] = {...(D.fixedM[k] || {}), ...patch}; save(); render(); refreshSheets(); }

/* ---- manual entry ---- */
function manualSheet(orig, presetCat) {
  const st = orig ? {...orig, amount: Math.abs(orig.amount)} : {date: todayStr(), amount: '', cat: presetCat || '', desc: '', pay: '', person: D.cfg.me || people()[0], note: ''};
  let mode = orig ? (isIncome(catById(orig.cat)) ? 'income' : orig.amount < 0 ? 'refund' : 'expense') : (isIncome(catById(presetCat)) ? 'income' : 'expense');
  const e = openSheet(() => `${head(orig ? 'עריכה' : 'הוספה ידנית')}
    <div class="seg"><button data-mode="expense" class="${mode === 'expense' ? 'on' : ''}">הוצאה</button><button data-mode="refund" class="${mode === 'refund' ? 'on' : ''}">החזר / זיכוי</button><button data-mode="income" class="${mode === 'income' ? 'on' : ''}">הכנסה</button></div>
    <div class="amtin"><input id="m-amt" inputmode="decimal" placeholder="0" value="${esc(st.amount)}"><span>₪</span></div>
    <label class="field"><span>תיאור</span><input class="inp" id="m-desc" value="${esc(st.desc)}"></label>
    <div class="field"><span>סעיף</span><button class="inp" id="m-cat" style="text-align:right">${st.cat && catById(st.cat) ? esc(catById(st.cat).name) : 'בחירה…'}</button></div>
    <div class="row2"><label class="field"><span>תאריך</span><input class="inp" type="date" id="m-date" value="${esc(st.date)}"></label>
      <label class="field"><span>אמצעי תשלום</span><select class="inp" id="m-pay"><option value="">—</option>${D.cfg.pays.map(p => `<option ${st.pay === p ? 'selected' : ''}>${esc(p)}</option>`).join('')}</select></label></div>
    <label class="field"><span>הערה</span><input class="inp" id="m-note" value="${esc(st.note || '')}"></label>
    <button class="primary" id="m-save">שמירה</button>`, {
    mount(sh) {
      sh.addEventListener('click', ev => {
        const b = ev.target.closest('button'); if (!b) return;
        if (b.dataset.mode) { read(); mode = b.dataset.mode; if (st.cat && isIncome(catById(st.cat)) !== (mode === 'income')) st.cat = ''; e.paint(); }
        if (b.id === 'm-cat') { read(); pickCat('סעיף', mode === 'income', st.cat, c => { st.cat = c; $('#m-cat', sh).textContent = catById(c).name; }); }
        if (b.id === 'm-save') {
          read(); const a = Math.abs(num(st.amount));
          if (!a) return toast('צריך סכום'); if (!catById(st.cat)) return toast('צריך לבחור סעיף');
          const rec = {id: orig ? orig.id : 'm' + uid(), date: st.date || todayStr(), amount: mode === 'refund' ? -a : a, cat: st.cat, desc: st.desc, pay: st.pay, person: st.person, note: st.note};
          if (orig) D.manual = D.manual.map(x => x.id === orig.id ? rec : x); else D.manual.push(rec);
          save(); closeSheet(e); render(); refreshSheets(); toast('נשמר');
        }
      });
      function read() { st.amount = $('#m-amt', sh).value; st.desc = $('#m-desc', sh).value.trim(); st.date = $('#m-date', sh).value; st.pay = $('#m-pay', sh).value; st.note = $('#m-note', sh).value.trim(); }
      if (!orig) setTimeout(() => $('#m-amt', sh).focus(), 250);
    }
  });
}

/* ---- fixed items ---- */
function fixedList() {
  openSheet(() => `${head('תנועות קבועות')}
    <p class="note">מה שיורד מהבנק או נכנס אליו בלי שנכתב בקבוצה: משכנתא, הוראות קבע, ביטוחים, משכורות. כל חודש הן מופיעות לבד כ"צפוי", ואפשר לאשר שירדו או לתקן סכום. מה שמסומן "מהאקסל" מתעדכן לבד בכל טעינת קובץ מהאקסל, לפי החודש האחרון באשראי ובעו"ש.</p>
    <div class="card" style="padding:4px 12px">${D.fixed.length ? D.fixed.map(f => `<button class="txline" data-act="fx-edit" data-id="${esc(f.id)}"><div><div class="a">${ils(f.amount)}</div><div class="d">${esc(f.name)} · ${esc(catById(f.cat)?.name || '?')}${f.xkey ? ` · <span style="color:var(--blue)">מהאקסל${f.xtype ? ' – ' + esc(FX_TYPE[f.xtype] || f.xtype) : ''}</span>` : ''}</div></div>
      <div class="r">${f.freq === '2m' ? 'כל חודשיים' : f.freq === 'list' ? 'חודשים: ' + (f.months || []).join(',') : 'כל חודש'}<br>${f.day ? 'ב-' + f.day + ' לחודש' : ''}${f.to ? '<br>עד ' + sheetName(f.to) : ''}</div></button>`).join('') : '<p class="note">אין עדיין</p>'}</div>
    <button class="primary" data-act="fx-edit" data-id="">+ תנועה קבועה</button>`, {live: true});
}
function fixedEdit(id) {
  const f0 = D.fixed.find(f => f.id === id);
  const st = f0 ? clone(f0) : {id: 'f' + uid(), name: '', cat: '', amount: '', day: 1, freq: 'm', months: [], from: curMonth(), to: '', pay: 'עו"ש (הו"ק / העברות)'};
  const e = openSheet(() => `${head(f0 ? 'עריכת תנועה קבועה' : 'תנועה קבועה חדשה')}
    <label class="field"><span>שם</span><input class="inp" id="x-name" value="${esc(st.name)}" placeholder="למשל: משכנתא, משכורת"></label>
    <div class="row2"><label class="field"><span>סכום ₪</span><input class="inp" id="x-amt" inputmode="decimal" value="${esc(st.amount)}"></label>
      <div class="field"><span>סעיף</span><button class="inp" id="x-cat" style="text-align:right">${catById(st.cat) ? esc(catById(st.cat).name) : 'בחירה…'}</button></div></div>
    <div class="row2"><label class="field"><span>תדירות</span><select class="inp" id="x-freq"><option value="m" ${st.freq === 'm' ? 'selected' : ''}>כל חודש</option><option value="2m" ${st.freq === '2m' ? 'selected' : ''}>כל חודשיים</option><option value="list" ${st.freq === 'list' ? 'selected' : ''}>חודשים מסוימים</option></select></label>
      <label class="field"><span>יום בחודש</span><input class="inp" id="x-day" inputmode="numeric" value="${esc(st.day || 1)}"></label></div>
    <div class="field" id="x-months-w" style="${st.freq === 'list' ? '' : 'display:none'}"><span>באילו חודשים</span><div class="chips" id="x-months">${HEB_M.map((n, i) => `<button data-mo="${i + 1}" class="${(st.months || []).includes(i + 1) ? 'on' : ''}">${n}</button>`).join('')}</div></div>
    <div class="row2"><label class="field"><span>מחודש</span><input class="inp" type="month" id="x-from" value="${esc(st.from)}"></label>
      <label class="field"><span>עד חודש (לא חובה)</span><input class="inp" type="month" id="x-to" value="${esc(st.to || '')}"></label></div>
    <label class="field"><span>אמצעי תשלום</span><select class="inp" id="x-pay">${D.cfg.pays.map(p => `<option ${st.pay === p ? 'selected' : ''}>${esc(p)}</option>`).join('')}</select></label>
    <p class="note">שינוי סכום משנה את כל החודשים שלא תוקנו בנפרד. אם המחיר השתנה רק מעכשיו, עדיף לסיים את הישנה ("עד חודש") ולפתוח חדשה.</p>
    <button class="primary" id="x-save">שמירה</button>${f0 ? '<button class="ghost danger" id="x-del">מחיקה</button>' : ''}`, {
    mount(sh) {
      const read = () => { st.name = $('#x-name', sh).value.trim(); st.amount = num($('#x-amt', sh).value); st.freq = $('#x-freq', sh).value; st.day = Math.min(31, Math.max(1, num($('#x-day', sh).value) || 1)); st.from = $('#x-from', sh).value || curMonth(); st.to = $('#x-to', sh).value || ''; st.pay = $('#x-pay', sh).value; };
      sh.addEventListener('change', ev => { if (ev.target.id === 'x-freq') $('#x-months-w', sh).style.display = ev.target.value === 'list' ? '' : 'none'; });
      sh.addEventListener('click', ev => {
        const b = ev.target.closest('button'); if (!b) return;
        if (b.dataset.mo) { const mo = +b.dataset.mo; st.months = st.months.includes(mo) ? st.months.filter(x => x !== mo) : [...st.months, mo].sort((a, c) => a - c); b.classList.toggle('on'); }
        if (b.id === 'x-cat') { read(); pickCat('סעיף', null, st.cat, c => { st.cat = c; $('#x-cat', sh).textContent = catById(c).name; }); }
        if (b.id === 'x-save') {
          read(); if (!st.name || !st.amount || !catById(st.cat)) return toast('צריך שם, סכום וסעיף');
          if (st.freq === 'list' && !st.months.length) return toast('צריך לבחור חודשים');
          if (f0) Object.assign(f0, st); else D.fixed.push(st);
          save(); closeSheet(e); render(); refreshSheets(); toast('נשמר');
        }
        if (b.id === 'x-del') { if (!b.dataset.sure) { b.dataset.sure = 1; b.textContent = 'בטוח? לחיצה נוספת מוחקת'; return; } if (f0 && f0.xkey) D.meta.fxDel = [...(D.meta.fxDel || []), f0.xkey]; D.fixed = D.fixed.filter(x => x.id !== id); save(); closeSheet(e); render(); refreshSheets(); }
      });
    }
  });
}

/* ---- settings ---- */
function settingsSheet() {
  openSheet(() => `${head('הגדרות')}
    <div class="actlist">
      <button data-act="sync">חיבור לאקסל (ייצוא וטעינה) <span>›</span></button>
      <button data-act="income">הכנסות צפויות <span>›</span></button>
      <button data-act="cats">סעיפים <span>›</span></button>
      <button data-act="fixed">תנועות קבועות <span class="v">${D.fixed.length}</span></button>
      <button data-act="kw">מילות מפתח <span class="v">${D.kw.length}</span></button>
      <button data-act="pays">אמצעי תשלום <span class="v">${D.cfg.pays.length}</span></button>
      <button data-act="importset">הגדרות ייבוא <span>›</span></button>
      <button data-act="import">ייבוא ידני של צ'אט מווטסאפ <span class="v">${D.meta.lastImport ? new Date(D.meta.lastImport.at).toLocaleDateString('he-IL') : '—'}</span></button>
      <button data-act="guide">איך כותבים בקבוצה <span>›</span></button>
      <button data-act="lock">נעילה בקוד / טביעת אצבע <span class="v">${lockOn() ? (L.bio ? 'קוד + טביעת אצבע' : 'קוד') : 'כבויה'}</span></button>
      <button data-act="backup">גיבוי <span class="v">${D.meta.lastBackup ? new Date(D.meta.lastBackup).toLocaleDateString('he-IL') : 'עוד לא'}</span></button>
      <button data-act="restore">שחזור מגיבוי <span>›</span></button>
      <button data-act="checkupd">בדיקת עדכון <span class="v">גרסה ${APP_VERSION}</span></button>
      <button data-act="history">היסטוריית גרסאות <span>›</span></button>
      <button data-act="wipe" class="danger">מחיקת כל הנתונים מהטלפון <span>🗑</span></button>
    </div>
    <p class="note">הכל נשמר רק בדפדפן בטלפון הזה. ניקוי נתוני הדפדפן מוחק אותו – לכן כדאי לגבות אחרי שינויים בתקציבים, בקבועות או בתיקונים. את ההוצאות עצמן תמיד אפשר לייבא שוב מהצ'אט.</p>`);
}
function importSettings() {
  const senders = [...new Set([...(D.meta.senders || []), ...Object.keys(D.cfg.senders || {})])];
  openSheet(() => `${head('הגדרות ייבוא')}
    <label class="field"><span>לקלוט הודעות מתאריך</span><input class="inp" type="date" id="is-start" value="${esc(D.cfg.start)}"></label>
    <label class="field"><span>סדר התאריך בקובץ</span><select class="inp" id="is-order"><option value="auto" ${D.cfg.dateOrder === 'auto' ? 'selected' : ''}>זיהוי אוטומטי</option><option value="dm" ${D.cfg.dateOrder === 'dm' ? 'selected' : ''}>יום/חודש</option><option value="md" ${D.cfg.dateOrder === 'md' ? 'selected' : ''}>חודש/יום</option></select></label>
    ${senders.length ? `<div class="gname">מי זה מי (לפי השם בצ'אט)</div><div class="card" style="padding:4px 12px">${senders.map(s => `<div class="setrow" style="grid-template-columns:1fr 1fr"><b style="font-size:14px">${esc(s)}</b><select class="inp" data-snd="${esc(s)}">${people().map(p => `<option ${senderToPerson(s) === p ? 'selected' : ''}>${esc(p)}</option>`).join('')}</select></div>`).join('')}</div>` : ''}
    <div class="gname">אמצעי תשלום כשלא כתוב בהודעה</div><div class="card" style="padding:4px 12px">${people().map(p => `<div class="setrow" style="grid-template-columns:1fr 1.4fr"><b style="font-size:14px">${esc(p)}</b><select class="inp" data-dp="${esc(p)}">${D.cfg.pays.map(x => `<option ${(D.cfg.defPay || {})[p] === x ? 'selected' : ''}>${esc(x)}</option>`).join('')}</select></div>`).join('')}</div>
    <button class="primary" id="is-save">שמירה</button>
    <p class="note">שינוי תאריך או שולחים נכנס לתוקף בייבוא הבא.</p>`, {
    mount(sh) {
      sh.addEventListener('click', ev => { if (!ev.target.closest('#is-save')) return;
        D.cfg.start = $('#is-start', sh).value || D.cfg.start; D.cfg.dateOrder = $('#is-order', sh).value;
        sh.querySelectorAll('[data-snd]').forEach(x => { D.cfg.senders[x.dataset.snd] = x.value; });
        D.cfg.defPay = D.cfg.defPay || {}; sh.querySelectorAll('[data-dp]').forEach(x => { D.cfg.defPay[x.dataset.dp] = x.value; });
        for (const it of D.chat) it.person = senderToPerson(it.sender);
        save(); closeSheet(); render(); toast('נשמר');
      });
    }
  });
}
function incomeSheet() {
  openSheet(() => `${head('הכנסות צפויות')}
    <p class="note">כמה צפוי להיכנס בכל חודש בכל סעיף הכנסה. הסכום נספר בדשבורד מה־1 לחודש, ובמסך "לבדיקה" מאשרים כמה נכנס בפועל.</p>
    <div class="card" style="padding:4px 12px">${cats().filter(c => isIncome(c) && !c.archived).map(c => `<div class="setrow" style="grid-template-columns:1fr 120px"><b style="font-size:14px">${esc(c.name)}</b><input class="bud" data-bd="${esc(c.id)}" inputmode="decimal" placeholder="0" value="${num(c.budget) || ''}"></div>`).join('')}</div>
    <p class="note">הכנסה שמוגדרת גם כתנועה קבועה נספרת רק פעם אחת.</p>`, {
    mount(sh) { sh.addEventListener('input', ev => { const t = ev.target; if (t.dataset.bd) { catById(t.dataset.bd).budget = num(t.value); save(); } }); },
    onClose() { render(); }
  });
}
function catsSheet() {
  openSheet(() => `${head('סעיפים')}<p class="note">שינוי שם או הסתרה. השמות צריכים להיות זהים לגיליון "רשימות" באקסל.</p>
    ${D.cfg.groups.map(g => { const gc = cats().filter(c => c.group === g.id); return `<div class="gname">${esc(g.name)}</div><div class="chips">${gc.map(c => `<button data-act="catedit" data-id="${esc(c.id)}" style="${c.archived ? 'opacity:.45' : ''}">${esc(c.name)}</button>`).join('')}</div>`; }).join('')}`, {live: true});
}
function kwSheet() {
  let cat = '';
  openSheet(() => `${head('מילות מפתח')}
    <p class="note">כשהתיאור בהודעה מכיל מילה מהרשימה, הסעיף נבחר לבד. אם כמה מילים מתאימות – המילה הארוכה (המדויקת) קובעת, למשל "סופר פארם" לפני "סופר". רווח לפני ואחרי (" חיה ") = רק מילה שלמה.</p>
    <div class="row2"><input class="inp" id="kw-w" placeholder="מילה"><button class="inp" id="kw-c" style="text-align:right">${cat ? esc(kwCatName(cat)) : 'סעיף…'}</button></div>
    <button class="primary" id="kw-add">הוספה</button>
    <div class="card" style="padding:4px 12px;margin-top:14px">${D.kw.map(([k, c], i) => [k, c, i]).reverse().map(([k, c, i]) => `<div class="setrow" style="grid-template-columns:1fr 1fr 34px"><b style="font-size:14px;white-space:pre">${esc(k)}</b><span class="note">${esc(kwCatName(c))}</span><button class="danger" data-del="${i}">✕</button></div>`).join('')}</div>`, {
    mount(sh, e) {
      sh.addEventListener('click', ev => {
        const d = ev.target.closest('[data-del]'); if (d) { D.kw.splice(+d.dataset.del, 1); recategorize(); save(); e.paint(); return; }
        if (ev.target.closest('#kw-c')) pickCat('לאיזה סעיף?', null, cat, c => { cat = c; $('#kw-c', sh).textContent = kwCatName(c); }, true);
        if (ev.target.closest('#kw-add')) { const w = $('#kw-w', sh).value; if (!w.trim() || !cat) return toast('צריך מילה וסעיף'); D.kw = D.kw.filter(([k]) => k !== w); D.kw.push([w, cat]); cat = ''; recategorize(); save(); e.paint(); render(); toast('נוסף'); }
      });
    }, onClose() { render(); }
  });
}

function paysSheet() {
  openSheet(() => `${head('אמצעי תשלום')}
    <p class="note">כרטיס מזוהה לפי השם שכתוב בהודעה, למשל "אשראי דני כאל" או "אשראי דני ישראכרט". כרטיס חדש בפורמט "אשראי + שם + חברה" נוסף לכאן לבד בייבוא. אפשר גם להוסיף ידנית – בדיוק כמו שכותבים בקבוצה.</p>
    <div class="card" style="padding:4px 12px">${D.cfg.pays.map((p, i) => `<div class="setrow" style="grid-template-columns:1fr auto 34px"><b style="font-size:14px">${esc(p)}</b><span class="note">${D.chat.filter(it => it.pay === p).length} שורות</span><button class="danger" data-del="${i}">✕</button></div>`).join('')}</div>
    <div class="row2" style="margin-top:10px;grid-template-columns:1fr auto"><input class="inp" id="py-new" placeholder="למשל: אשראי רותי מקס"><button class="smallbtn" id="py-add">הוספה</button></div>`, {
    mount(sh, e) {
      const redetect = () => { for (const it of D.chat) if (!(D.ov[it.id] || {}).pay) it.pay = detectPay(it.text, it.person, []).pay; };
      sh.addEventListener('click', ev => {
        const d = ev.target.closest('[data-del]');
        if (d) { D.cfg.pays.splice(+d.dataset.del, 1); redetect(); save(); e.paint(); render(); return; }
        if (ev.target.closest('#py-add')) { const v = $('#py-new', sh).value.trim(); if (!v) return; if (!D.cfg.pays.includes(v)) D.cfg.pays.push(v); redetect(); save(); e.paint(); render(); toast('נוסף'); }
      });
    }
  });
}
function paySheet(pay) {
  openSheet(() => { const list = monthTx(S.month).filter(t => payKey(t) === pay && !t.pending); const tot = list.filter(t => !isIncome(catById(t.cat))).reduce((s, t) => s + t.amount, 0);
    return `${head(esc(pay))}<p class="note">${mLabel(S.month)} · ${list.length} תנועות · הוצאות נטו ${ils(tot)}. נוח להשוואה מול פירוט הכרטיס.</p><div class="card">${list.map(txLine).join('') || '<p class="note">אין</p>'}</div>`; }, {live: true, month: true});
}

function syncSheet() {
  const xs = D.meta.excelSync, closed = Object.keys(D.closed || {}).sort();
  openSheet(() => `${head('חיבור לאקסל')}
    <div class="card" style="margin-bottom:10px"><b>1. מהאפליקציה לאקסל</b>
      <p class="note">קובץ עם כל התנועות (מווטסאפ, קבועות, ידני). באקסל: גיליון "סגירת חודש" ← "ייבוא מאיפה הכסף".${D.meta.lastExport ? ` ייצוא אחרון: ${new Date(D.meta.lastExport).toLocaleDateString('he-IL')}.` : ''}</p>
      <button class="primary" data-act="export" style="margin-top:6px">ייצוא לאקסל</button></div>
    <div class="card"><b>2. מהאקסל לאפליקציה</b>
      <p class="note">באקסל: "ייצוא לאיפה הכסף" יוצר קובץ קטן עם הסעיפים, מילות המפתח והסכומים המדויקים של החודשים שסגרת. טוענים אותו כאן.${xs ? ` טעינה אחרונה: ${new Date(xs).toLocaleDateString('he-IL')}.` : ''}</p>
      ${closed.length ? `<p class="note">חודשים סגורים: ${closed.map(sheetName).join(', ')}</p>` : ''}
      <button class="ghost" data-act="xlimport">טעינת קובץ מהאקסל</button></div>
    <p class="note">🔒 מעבירים את הקבצים בין הטלפון למחשב בכבל USB או ב-Quick Share – לא דרך ווטסאפ, מייל או דרייב.</p>`);
}
/* ---- fixed charges from the workbook (standing orders, subscriptions, bank debits, fees) ---- */
const FX_TYPE = {'הוראת קבע': 'הוראת קבע באשראי', 'מנוי': 'מנוי שחוזר כל חודש', 'עו"ש': 'הו"ק בעו"ש', 'עמלה': 'עמלה / דמי כרטיס'};
const fxWords = t => new Set(String(t || '').replace(/[()"'״׳,.\-*/]/g, ' ').split(/\s+/).filter(w => w.length >= 3 && !/^(בע"מ|בעמ|חברת|הוראת|קבע)$/.test(w)));
const fxShare = (a, b) => { const wa = fxWords(a); return [...fxWords(b)].some(w => wa.has(w)); };
function fxLooksSame(x, name, amount, cat) {
  if (x.to && x.to < curMonth()) return false;
  const a = num(x.amount); if (!(a > 0) || Math.abs(a - amount) > Math.max(a, amount) * 0.15) return false;
  return (cat && x.cat === cat) || fxShare(x.name, name);
}
// a new amount was seen in month m: the months before it keep the amount that was expected then
function fxFreeze(f, m) {
  const old = num(f.amount);
  for (let x = f.from || m; x < m; x = addM(x, 1)) {
    if (!fixedActive(f, x)) continue;
    const k = f.id + '|' + x; if ((D.fixedM[k] || {}).amount == null) D.fixedM[k] = {...(D.fixedM[k] || {}), amount: old};
  }
}
function fxNew(a) {
  return {id: 'f' + uid(), xkey: a.key, name: a.name, cat: a.cat, amount: a.amount, day: a.day, freq: 'm', months: [], from: addM(a.last, 1), to: '',
    pay: a.src === 'bank' ? 'עו"ש (הו"ק / העברות)' : '', xtype: a.type, xlast: a.last};
}
function syncFixed(rowsFx) {
  if (!rowsFx) return null;                                   // a workbook from before this feature: change nothing
  const M = D.meta; M.fxNo = M.fxNo || []; M.fxDel = M.fxDel || [];
  const ch = [], ask = [], seen = new Set();
  for (const [key0, name0, amt0, day0, cat0, type0, last0, src0] of rowsFx) {
    const key = String(key0).trim(), last = sheetToMonth(last0), amount = Math.round(num(amt0) * 100) / 100;
    if (!key || !last || !(amount > 0)) continue;
    seen.add(key);
    if (M.fxDel.includes(key)) continue;                        // deleted by hand: not brought back
    const name = String(name0).trim(), type = String(type0).trim(), src = String(src0).trim();
    const c = catByName(String(cat0).trim()), cat = c && !c.archived ? c.id : (suggestCat(name, false) || '');
    const a = {key, name, amount, day: Math.min(31, Math.max(1, Math.round(num(day0)) || 1)), cat: cat === SKIP ? '' : cat, type, last, src};
    const f = D.fixed.find(x => x.xkey === key);
    if (f) {
      if (f.to && f.to < addM(last, 1)) { f.to = ''; ch.push({t: 'back', name: f.name, amount}); }
      if (Math.abs(num(f.amount) - amount) > 0.005) { ch.push({t: 'up', name: f.name, old: num(f.amount), amount, month: last}); fxFreeze(f, last); f.amount = amount; }
      f.day = a.day; f.xtype = type; f.xlast = last;
      if (!catById(f.cat) && a.cat) f.cat = a.cat;
      if (fixedActive(f, last)) D.fixedM[f.id + '|' + last] = {...(D.fixedM[f.id + '|' + last] || {}), ok: true, amount};   // charged in that month: confirmed
      continue;
    }
    const was = (M.fxAsk || []).find(x => x.key === key);
    const cand = was ? D.fixed.find(x => x.id === was.fid && !x.xkey) : D.fixed.find(x => !x.xkey && !M.fxNo.includes(key + '|' + x.id) && fxLooksSame(x, name, amount, a.cat));
    if (cand) { ask.push({...a, fid: cand.id}); continue; }
    D.fixed.push(fxNew(a));
    ch.push({t: type === 'עמלה' ? 'fee' : 'new', name, amount, type});
  }
  // charges that stopped: no longer expected after the last month they were seen
  if (rowsFx.length) for (const f of D.fixed) if (f.xkey && !seen.has(f.xkey) && !f.to) { f.to = f.xlast || addM(curMonth(), -1); ch.push({t: 'gone', name: f.name, amount: num(f.amount)}); }
  M.fxAsk = ask;
  if (ch.length) M.fxChanges = [...(M.fxChanges || []), ...ch];
  return {ch, ask};
}
function fxChangeLine(x) {
  const n = `<b>${esc(x.name)}</b>`;
  const noCat = () => { const f = D.fixed.find(f => f.name === x.name && f.xkey); return f && !catById(f.cat) ? ' – <span class="red">צריך לבחור סעיף</span>' : ''; };
  if (x.t === 'new') return `חדש: ${n} – ${ils(x.amount)} (${esc(FX_TYPE[x.type] || x.type)})${noCat()}`;
  if (x.t === 'fee') return `עמלה / דמי כרטיס: ${n} – ${ils(x.amount)}${noCat()}`;
  if (x.t === 'up') return `${n}: ${ils(x.old)} ← ${ils(x.amount)} <span class="${x.amount > x.old ? 'red' : 'green'}">(${x.amount > x.old ? 'עלה' : 'ירד'} ב-${ils(Math.abs(x.amount - x.old))})</span>`;
  if (x.t === 'gone') return `${n}: לא חויב בחודש האחרון – הוסר מהצפי`;
  if (x.t === 'back') return `${n}: חזר לחייב – ${ils(x.amount)}`;
  return n;
}
function fxAnswer(key, same) {
  const M = D.meta, a = (M.fxAsk || []).find(x => x.key === key); if (!a) return;
  M.fxAsk = M.fxAsk.filter(x => x !== a);
  const f = D.fixed.find(x => x.id === a.fid);
  if (same && f) {
    if (Math.abs(num(f.amount) - a.amount) > 0.005) { M.fxChanges = [...(M.fxChanges || []), {t: 'up', name: f.name, old: num(f.amount), amount: a.amount, month: a.last}]; fxFreeze(f, a.last); }
    Object.assign(f, {xkey: a.key, amount: a.amount, day: a.day, xtype: a.type, xlast: a.last});
    if (!catById(f.cat) && a.cat) f.cat = a.cat;
    toast('חובר – מעכשיו הסכום מתעדכן מהאקסל');
  } else {
    M.fxNo = [...(M.fxNo || []), a.key + '|' + a.fid];
    D.fixed.push(fxNew(a));
    M.fxChanges = [...(M.fxChanges || []), {t: a.type === 'עמלה' ? 'fee' : 'new', name: a.name, amount: a.amount, type: a.type}];
    toast('נוסף כחיוב קבוע נפרד');
  }
  save(); render(); refreshSheets();
}
const sheetToMonth = n => { const m = String(n).trim().match(/^(\d{1,2})\.(\d{2})$/); return m ? `20${m[2]}-${pad(+m[1])}` : null; };
async function importExcel(file) {
  if (!file) return;
  let wb, XLSX;
  try { XLSX = await loadXLSX(); wb = XLSX.read(new Uint8Array(await file.arrayBuffer()), {type: 'array'}); } catch { toast('לא הצלחתי לקרוא את הקובץ', 4000); return; }
  const rows = n => wb.Sheets[n] ? XLSX.utils.sheet_to_json(wb.Sheets[n], {header: 1, defval: '', raw: true}).slice(1) : null;
  const sc = rows('סעיפים');
  if (!sc) { toast('זה לא קובץ "מהאקסל" של איפה הכסף (חסר גיליון "סעיפים")', 5000); return; }
  if (!D) D = blank();
  // groups & categories, in the workbook's order
  const info = Object.fromEntries((rows('מידע') || []).map(([k, v]) => [String(k).trim(), v]));
  if (num(info['פורמט']) > FILE_FORMAT) { toast('הקובץ נוצר בגרסה חדשה יותר של האקסל – כדאי לעדכן קודם את האפליקציה (תפריט ← בדיקת עדכון)', 7000); return; }
  const wasHidden = new Set(cats().filter(c => c.archived).map(c => c.id));
  const gBy = {}; D.cfg.groups.forEach(g => { gBy[g.name] = g; });
  const groups = [], seenG = new Set(), newCats = [], seenC = new Set();
  let added = 0;
  for (const [name0, gname0, type] of sc) {
    const name = String(name0).trim(), gname = String(gname0).trim(); if (!name || !gname) continue;
    let g = gBy[gname]; if (!g) { g = {id: 'g' + uid(), name: gname}; gBy[gname] = g; }
    if (!seenG.has(g.id)) { seenG.add(g.id); groups.push(g); }
    let c = catByName(name); if (!c) { c = {id: 'c' + uid(), name, budget: 0, focus: false, kind: 'variable'}; added++; }
    c.group = g.id; c.archived = false;
    if (String(type).trim() === 'הכנסה') c.kind = 'income'; else if (c.kind === 'income') c.kind = 'variable';
    newCats.push(c); seenC.add(c.id);
  }
  const hidden = cats().filter(c => !seenC.has(c.id));
  hidden.forEach(c => { c.archived = true; });
  D.cfg.groups = [...groups, ...D.cfg.groups.filter(g => !seenG.has(g.id))];
  D.cfg.cats = [...newCats, ...hidden];
  // keywords: the workbook list, then words learned here that the workbook doesn't have yet
  const kx = (rows('מילות מפתח') || []).map(([w, c]) => [String(w), catByName(String(c).trim())]).filter(([w, c]) => w.trim() && c).map(([w, c]) => [w, c.id]);
  if (kx.length) { const have = new Set(kx.map(([w]) => w)); D.kw = [...kx, ...D.kw.filter(([w, c]) => !have.has(w) && (c === SKIP || catById(c)))]; }
  // closed months: accurate totals per category
  const cm = rows('חודשים סגורים') || [];
  const closed = {};
  for (const [sn, cname, amt] of cm) { const m = sheetToMonth(sn); if (!m) continue; closed[m] = closed[m] || {}; const c = catByName(String(cname).trim()); if (c && num(amt)) closed[m][c.id] = (closed[m][c.id] || 0) + num(amt); }
  D.closed = closed;
  const fx = syncFixed(rows('חיובים קבועים'));
  D.meta.excelSync = Date.now();
  recategorize(); save(); stack.slice().forEach(e => closeSheet(e)); render();
  const cms = Object.keys(closed).sort();
  openSheet(() => `${head('נטען מהאקסל')}
    <div class="stat"><div><b>${newCats.length}</b><span>סעיפים${added ? ` (${added} חדשים)` : ''}</span></div><div><b>${kx.length}</b><span>מילות מפתח</span></div><div><b>${cms.length}</b><span>חודשים סגורים</span></div></div>
    ${cms.length ? `<p class="note">החודשים ${cms.map(sheetName).join(', ')} מוצגים עכשיו לפי האקסל – מדויק ✓, ולא יופיעו יותר במסך "לבדיקה".</p>` : ''}
    ${hidden.filter(c => !wasHidden.has(c.id)).length ? `<p class="note">סעיפים שלא קיימים באקסל הוסתרו: ${hidden.filter(c => !wasHidden.has(c.id)).map(c => esc(c.name)).join(', ')}</p>` : ''}
    ${fx && (fx.ch.length || fx.ask.length) ? `<div class="label">חיובים קבועים מהאשראי והעו"ש</div><div class="card">${fx.ch.map(x => `<div class="note" style="margin:4px 0">${fxChangeLine(x)}</div>`).join('')}${fx.ask.length ? `<div class="note" style="margin:6px 0">${fx.ask.length === 1 ? 'חיוב אחד דומה' : fx.ask.length + ' חיובים דומים'} לתנועה קבועה שהגדרת – השאלה מחכה במסך "לבדיקה".</div>` : ''}</div>` : fx ? '<p class="note">חיובים קבועים: אין שינוי מהפעם הקודמת.</p>' : ''}
    <button class="primary" data-act="close">מעולה</button>`);
}
function guideSheet() {
  D.meta.guideSeen = true; save(false);
  const n1 = people()[0] || 'דני', n2 = people()[1] || n1;
  const ex = [[`350 שופרסל אשראי ${n1} כאל`, 'סכום ראשון, אחריו מה קנית, ובסוף איך שילמת'], [`55 פיצה אשראי ${n2} ישראכרט`, 'כרטיס = "אשראי" + שם + חברה. כרטיס חדש נקלט לבד'],
    ['120 דלק לרכב\n30 קפה תחנת דלק', 'כמה הוצאות בהודעה אחת – שורה לכל אחת'], [`89 השתלמות למשרד אשראי ${n1} כאל יתקבל החזר`, '"יתקבל החזר" / "צריך להתקבל החזר" / "קופה קטנה" – מסומן כמחכה להחזר'],
    ['התקבל 85 השתלמות', '"התקבל" = החזר: מתחבר לבד להוצאה שחיכתה לו ויורד ממנה כמינוס (גם בכמה הודעות). הכנסה – רק עם מילה של הכנסה, כמו "התקבל משכורת"']];
  openSheet(() => `${head('איך כותבים בקבוצה')}
    ${ex.map(([a, b]) => `<div class="card" style="margin-bottom:8px;padding:12px 14px"><div style="font-weight:700;white-space:pre-line">${esc(a)}</div><div class="note">${esc(b)}</div></div>`).join('')}
    <p class="note">מילה שלא מזוהה? בוחרים לה סעיף פעם אחת במסך "לבדיקה", והיא נלמדת. כדאי לנעוץ את ההודעה הזו בקבוצה:</p>
    <button class="ghost" id="g-copy">העתקת הדוגמאות (לנעיצה בקבוצה)</button>`, {
    mount(sh) { sh.addEventListener('click', async ev => { if (!ev.target.closest('#g-copy')) return; const txt = 'איך כותבים הוצאה:\n' + ex.map(([a, b]) => `• ${a.replace(/\n/g, ' / ')} – ${b}`).join('\n');
      try { await navigator.clipboard.writeText(txt); toast('הועתק – אפשר להדביק בקבוצה'); } catch { toast('ההעתקה נחסמה בדפדפן'); } }); }
  });
  render();
}

/* ---- backup / restore / wipe ---- */
function download(name, blob) {
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 3000);
}
// ---- backup content: everything, optionally encrypted with a password (AES-GCM, key from PBKDF2) inside the phone ----
const ENC_ITER = 310000;
function b64big(u8) { let s = ''; for (let i = 0; i < u8.length; i += 0x8000) s += String.fromCharCode.apply(null, u8.subarray(i, i + 0x8000)); return btoa(s); }
const unb64big = t => Uint8Array.from(atob(t), c => c.charCodeAt(0));
async function pwKey(pw, salt, iter) {
  const k = await crypto.subtle.importKey('raw', new TextEncoder().encode(pw), 'PBKDF2', false, ['deriveKey']);
  return crypto.subtle.deriveKey({name: 'PBKDF2', hash: 'SHA-256', salt, iterations: iter}, k, {name: 'AES-GCM', length: 256}, false, ['encrypt', 'decrypt']);
}
async function encryptText(text, pw) {
  const salt = crypto.getRandomValues(new Uint8Array(16)), iv = crypto.getRandomValues(new Uint8Array(12));
  const ct = new Uint8Array(await crypto.subtle.encrypt({name: 'AES-GCM', iv}, await pwKey(pw, salt, ENC_ITER), new TextEncoder().encode(text)));
  return JSON.stringify({kesefBackup: 2, enc: 'AES-GCM', kdf: 'PBKDF2-SHA256', iter: ENC_ITER, salt: b64big(salt), iv: b64big(iv), data: b64big(ct)});
}
async function decryptText(o, pw) {
  const pt = await crypto.subtle.decrypt({name: 'AES-GCM', iv: unb64big(o.iv)}, await pwKey(pw, unb64big(o.salt), o.iter || ENC_ITER), unb64big(o.data));
  return new TextDecoder().decode(pt);
}
const encOn = () => !(D.meta && D.meta.encOff);
async function backupText() {
  const plain = JSON.stringify({...D, meta: {...D.meta, backupAt: Date.now()}});
  return encOn() ? encryptText(plain, S.bkPw) : plain;
}
function askPassword(title, twice, onOk) {
  openSheet(() => `${head(title)}
    <label class="field"><span>סיסמה (לפחות 6 תווים)</span><input class="inp" type="password" id="pw1" autocomplete="new-password"></label>
    ${twice ? '<label class="field"><span>שוב, לאימות</span><input class="inp" type="password" id="pw2" autocomplete="new-password"></label>' : ''}
    <button class="primary" id="pw-ok">המשך</button>
    <p class="note">${twice ? '⚠️ הסיסמה לא נשמרת בשום מקום. בלי הסיסמה אי אפשר לפתוח את הגיבוי – ואין דרך לשחזר סיסמה שנשכחה. כדאי לרשום אותה במקום בטוח.' : 'הסיסמה שבחרת כשיצרת את הגיבוי.'}</p>`, {
    mount(sh) {
      setTimeout(() => $('#pw1', sh).focus(), 250);
      const go = () => { const a = $('#pw1', sh).value, b = twice ? $('#pw2', sh).value : a;
        if (a.length < 6) return toast('הסיסמה קצרה מדי'); if (a !== b) return toast('הסיסמאות לא זהות'); closeSheet(); onOk(a); };
      sh.addEventListener('click', ev => { if (ev.target.closest('#pw-ok')) go(); });
      sh.addEventListener('keydown', ev => { if (ev.key === 'Enter') go(); });
    }
  });
}
// the password is kept only in memory until the app is closed, so a backup doesn't ask for it every time
function withPw(fn) { if (!encOn() || S.bkPw) return fn(); askPassword('סיסמה לגיבוי', true, pw => { S.bkPw = pw; fn(); }); }
function backupName() { const d = new Date(); return `גיבוי איפה הכסף ${pad(d.getDate())}.${pad(d.getMonth() + 1)}.${String(d.getFullYear()).slice(2)} ${pad(d.getHours())}-${pad(d.getMinutes())}.json`; }
function backupDone() { D.meta.lastBackup = Date.now(); D.meta.dirty = false; delete D.meta.dirtySince; save(false); render(); }
function canShareFile() { try { return !!(navigator.share && navigator.canShare && navigator.canShare({files: [new File(['x'], 'x.txt', {type: 'text/plain'})]})); } catch { return false; } }
// Android browsers can't overwrite a file or open a real "save as" dialog. Sharing the file lets you pick where it goes
// (e.g. Files → a "גיבויים" folder); a plain download gets a unique name (date + time) so no "download again?" prompt.
// One backup file that is updated every time: the first time the user picks where to save it (Android's save dialog),
// the file's handle is kept in IndexedDB on this phone, and later backups overwrite that same file.
const FS_OK = 'showSaveFilePicker' in window;
function fsDb() { return new Promise((res, rej) => { const r = indexedDB.open('kesef-fs', 1); r.onupgradeneeded = () => r.result.createObjectStore('h'); r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); }); }
async function fsGet() { try { const db = await fsDb(); return await new Promise(res => { const q = db.transaction('h').objectStore('h').get('backup'); q.onsuccess = () => res(q.result || null); q.onerror = () => res(null); }); } catch { return null; } }
async function fsSet(h) { try { const db = await fsDb(); await new Promise(res => { const t = db.transaction('h', 'readwrite'); h ? t.objectStore('h').put(h, 'backup') : t.objectStore('h').delete('backup'); t.oncomplete = res; t.onerror = res; }); } catch {} }
async function writeBackup(h) {
  if (h.queryPermission && (await h.queryPermission({mode: 'readwrite'})) !== 'granted') {
    if (!h.requestPermission || (await h.requestPermission({mode: 'readwrite'})) !== 'granted') throw Object.assign(new Error('perm'), {name: 'NotAllowedError'});
  }
  const data = await backupText();
  const w = await h.createWritable(); await w.write(data); await w.close();
}
async function pickBackupFile() {
  return await showSaveFilePicker({suggestedName: 'גיבוי איפה הכסף.json', types: [{description: 'גיבוי איפה הכסף', accept: {'application/json': ['.json']}}]});
}
async function backupSheet() {
  const share = canShareFile(), saved = FS_OK ? await fsGet() : null;
  let snapInfo = null; try { const s = await IDB.get('snap'); if (s) snapInfo = {at: JSON.parse(s).at}; } catch {}
  openSheet(() => `${head('גיבוי')}
    <p class="note">${D.meta.lastBackup ? `גיבוי אחרון: ${new Date(D.meta.lastBackup).toLocaleString('he-IL')}.` : 'עוד לא נוצר גיבוי.'} הגיבוי שומר הכל: הוצאות, סעיפים, קבועות, הזנות ידניות, תיקונים ומילות מפתח.</p>
    ${FS_OK ? (saved ? `<div class="card" style="padding:12px 14px;margin:6px 0"><span class="note">קובץ הגיבוי:</span><div style="margin-top:2px">${esc(saved.name.replace(/\.json$/i, ''))}</div></div>
        <button class="primary" id="bk-save">גיבוי עכשיו – מעדכן את אותו קובץ</button>
        <button class="ghost" id="bk-pick">שמירה במקום אחר…</button>`
      : `<button class="primary" id="bk-pick">שמירה בשם… (בחירת מקום)</button>
        <p class="note">ייפתח חלון השמירה של הטלפון: בוחרים תיקייה פרטית (למשל תיקייה בשם "גיבויים" – אפשר ליצור אותה מהחלון) ולוחצים שמירה. מהפעם הבאה כל גיבוי יעדכן את אותו קובץ – בלי קבצים כפולים.</p>`)
      : ''}
    <label class="field" style="display:flex;gap:10px;align-items:center;margin-top:10px"><input type="checkbox" id="bk-enc" ${encOn() ? 'checked' : ''}> להצפין את הגיבוי בסיסמה (מומלץ)</label>
    ${snapInfo ? `<button class="ghost" id="bk-unrestore">ביטול השחזור האחרון – חזרה למצב מ-${new Date(snapInfo.at).toLocaleString('he-IL')}</button>` : ''}
    <p class="note">🔒 לא לבחור בחלון השמירה את Google Drive או שירות ענן אחר, ולא את תיקיית ההורדות.${encOn() ? ' הגיבוי מוצפן: בלי הסיסמה אי אפשר לקרוא אותו.' : ''}</p>
    ${FS_OK ? '<details style="margin-top:8px"><summary class="note" style="cursor:pointer">אפשרויות נוספות</summary>' : '<p class="note">הדפדפן הזה לא תומך בעדכון של קובץ קיים, אז הגיבוי יוצא כקובץ חדש:</p>'}
      ${share ? '<button class="ghost" id="bk-share">שיתוף (למשל Quick Share למחשב)</button>' : ''}
      <button class="ghost" id="bk-dl">הורדה לתיקיית ההורדות</button>
    ${FS_OK ? '</details>' : ''}`, {
    mount(sh, e) {
      sh.addEventListener('change', ev => { if (ev.target.id === 'bk-enc') { D.meta.encOff = !ev.target.checked; if (!ev.target.checked) S.bkPw = null; save(false); } });
      sh.addEventListener('click', async ev => {
        if (ev.target.closest('#bk-unrestore')) return unRestore();
        if (ev.target.closest('#bk-save')) return withPw(async () => {
          const h = await fsGet(); if (!h) return e.paint();
          try { await writeBackup(h); backupDone(); stack.slice().forEach(x => closeSheet(x)); toast('הגיבוי עודכן ✓'); }
          catch (err) { if (err && err.name === 'AbortError') return; toast('לא הצלחתי לעדכן את הקובץ הקודם – בוחרים שוב מקום לגיבוי', 5000); await fsSet(null); stack.slice().forEach(x => closeSheet(x)); backupSheet(); }
        });
        if (ev.target.closest('#bk-pick')) return withPw(async () => {
          try { const h = await pickBackupFile(); await writeBackup(h); await fsSet(h); backupDone(); stack.slice().forEach(x => closeSheet(x)); toast('הגיבוי נשמר ✓ – מעכשיו כל גיבוי יעדכן את הקובץ הזה'); }
          catch (err) { if (!err || err.name !== 'AbortError') toast(`השמירה לא עבדה (${err && err.name || 'שגיאה'})`, 5000); }
        });
        if (ev.target.closest('#bk-dl')) return withPw(async () => { download(backupName(), new Blob([await backupText()], {type: 'application/json'})); backupDone(); stack.slice().forEach(x => closeSheet(x)); toast('הגיבוי ירד לתיקיית ההורדות'); });
        if (ev.target.closest('#bk-share')) return withPw(async () => {
          const f = new File([await backupText()], backupName().replace(/\.json$/, '.txt'), {type: 'text/plain'});
          try { await navigator.share({files: [f]}); backupDone(); stack.slice().forEach(x => closeSheet(x)); toast('הגיבוי נשלח'); }
          catch (err) { if (err && err.name !== 'AbortError') toast(`השיתוף לא עבד (${err && err.name || 'שגיאה'})`, 5000); }
        });
      });
    }
  });
}
// Restoring replaces everything in the app, so it first shows what is in the backup and asks.
// The current state is kept aside (on this phone) and can be brought back from the backup window.
async function restore(file) {
  if (!file) return;
  let o;
  try { o = JSON.parse(await file.text()); } catch { return toast('זה לא קובץ גיבוי של האפליקציה', 4000); }
  const go = d => {
    if (!d || !d.cfg || !Array.isArray(d.chat)) return toast('זה לא קובץ גיבוי של האפליקציה', 4000);
    confirmRestore(d, file);
  };
  if (o && o.kesefBackup && o.enc) {
    askPassword('הגיבוי מוצפן – סיסמה', false, async pw => {
      try { const d = JSON.parse(await decryptText(o, pw)); S.bkPw = pw; go(d); } catch { toast('סיסמה שגויה (או שהקובץ פגום)', 4000); }
    });
  } else go(o);
}
function confirmRestore(d, file) {
  const at = (d.meta && (d.meta.backupAt || d.meta.lastBackup)) || file.lastModified;
  const cnt = x => x ? `${x.chat.length} שורות מהקבוצה, ${(x.manual || []).length} ידניות, ${(x.fixed || []).length} קבועות` : '';
  const upto = x => x && x.meta && x.meta.lastImport && x.meta.lastImport.until ? dShort(x.meta.lastImport.until.slice(0, 10)) : '—';
  openSheet(() => `${head('שחזור מגיבוי')}
    <div class="card"><b>בגיבוי</b><p class="note" style="margin:4px 0">נוצר: ${at ? new Date(at).toLocaleString('he-IL') : 'לא ידוע'}<br>${cnt(d)}<br>הצ'אט מעודכן עד: ${upto(d)}</p></div>
    ${D ? `<div class="card" style="margin-top:8px"><b>עכשיו באפליקציה</b><p class="note" style="margin:4px 0">${cnt(D)}<br>הצ'אט מעודכן עד: ${upto(D)}</p></div>
    <p class="note">השחזור מחליף את כל מה שיש עכשיו באפליקציה. המצב הנוכחי נשמר בצד בטלפון, ואפשר לחזור אליו מחלון "גיבוי".</p>` : ''}
    <button class="primary" id="rs-go">לשחזר</button><button class="ghost" data-act="close">ביטול</button>`, {
    mount(sh) { sh.addEventListener('click', async ev => { if (!ev.target.closest('#rs-go')) return;
      if (D) { try { await IDB.put('snap', JSON.stringify({at: Date.now(), json: JSON.stringify(D)})); } catch {} }
      D = d; migrate(); migrate3(); migrate4(); migrate5(); migrate6(); D.meta.dirty = false; save(false); stack.slice().forEach(e => closeSheet(e)); render(); toast('הנתונים שוחזרו', 4000);
      if (S.pendingShare) setTimeout(importPendingShare, 600);
    }); }
  });
}
async function unRestore() {
  let s; try { s = JSON.parse(await IDB.get('snap')); } catch {} if (!s) return toast('אין מצב קודם לחזור אליו');
  D = JSON.parse(s.json); save(false); await IDB.del('snap'); stack.slice().forEach(e => closeSheet(e)); render(); toast('חזרת למצב שלפני השחזור ✓', 4000);
}
function wipeSheet() {
  openSheet(() => `${head('מחיקת כל הנתונים')}<p>זה מוחק מהטלפון את כל ההוצאות, ההגדרות, התקציבים והתנועות הקבועות. אי אפשר לבטל.</p>
    <button class="ghost" data-act="backup">קודם ליצור גיבוי</button><button class="primary" id="w-go" style="background:var(--red)">למחוק הכל</button>`, {
    mount(sh) { sh.addEventListener('click', ev => { if (!ev.target.closest('#w-go')) return; wipeStore(); fsSet(null); L = {}; D = null; stack.slice().forEach(e => closeSheet(e)); render(); toast('הכל נמחק'); }); }
  });
}

/* ---- Excel export: same layout as the month sheets in the budget workbook ---- */
function exportSheet() {
  const months = []; for (let m = minMonth(); m <= curMonth(); m = addM(m, 1)) months.push(m);
  const firstOpen = months.find(m => !isClosed(m)) || curMonth();
  openSheet(() => `${head('ייצוא לאקסל')}
    <p class="note">הקובץ נטען באקסל בכפתור "ייבוא מאיפה הכסף" – כל סכום נכנס לעמודה ולחודש שלו, עם הערה, ותנועה שכבר נכנסה לא תיכנס שוב (לפי המזהה). אפשר לייצא שוב ושוב את אותם חודשים.</p>
    <div class="row2"><label class="field"><span>מחודש</span><select class="inp" id="ex-from">${months.map(m => `<option value="${m}" ${m === firstOpen ? 'selected' : ''}>${sheetName(m)}${isClosed(m) ? ' (נסגר)' : ''}</option>`).join('')}</select></label>
    <label class="field"><span>עד חודש</span><select class="inp" id="ex-to">${months.map(m => `<option value="${m}" ${m === curMonth() ? 'selected' : ''}>${sheetName(m)}</option>`).join('')}</select></label></div>
    <label class="field" style="display:flex;gap:10px;align-items:center"><input type="checkbox" id="ex-fx"> לכלול קבועות והכנסות צפויות שעוד לא אושרו</label>
    <button class="primary" id="ex-go">הורדת קובץ Excel</button>`, {
    mount(sh) { sh.addEventListener('click', ev => { if (!ev.target.closest('#ex-go')) return; let a = $('#ex-from', sh).value, b = $('#ex-to', sh).value; if (a > b) [a, b] = [b, a]; doExport(a, b, $('#ex-fx', sh).checked); }); }
  });
}
async function doExport(from, to, inclOpen) {
  const X = await loadXLSX(), wb = X.utils.book_new(); wb.Workbook = {Views: [{RTL: true}]};
  const months = []; for (let m = from; m <= to; m = addM(m, 1)) months.push(m);
  const cols = []; for (const g of D.cfg.groups) for (const c of cats()) if (c.group === g.id) cols.push(c);
  const flat = [['מזהה', 'תאריך', 'גיליון', 'סעיף', 'קבוצה', 'סכום', 'תיאור', 'הודעה מקורית', 'אמצעי תשלום', 'מי', 'מקור', 'הערה', 'צפוי החזר']];
  let pendSkipped = 0, overflow = [];
  for (const m of months) {
    const list = monthTx(m).filter(t => { if (t.pending) { pendSkipped++; return false; } return (t.src !== 'fixed' && t.src !== 'income') || t.ok || inclOpen; }).sort((a, b) => a.date.localeCompare(b.date));
    const [y, mo] = m.split('-').map(Number);
    const aoa = [[`${HEB_M[mo - 1]} ${y}`], [], [''], ['סעיף']];
    cols.forEach((c, i) => { aoa[2][i + 1] = (i === 0 || cols[i - 1].group !== c.group) ? groupById(c.group).name : ''; aoa[3][i + 1] = c.name; });
    const per = cols.map(c => list.filter(t => t.cat === c.id));
    const rows = Math.max(50, ...per.map(p => p.length));
    per.forEach((p, i) => { if (p.length > 50) overflow.push(`${sheetName(m)} ${cols[i].name}`); });
    for (let r = 0; r < rows; r++) aoa.push([]);
    const totRow = 5 + rows;
    const ws = X.utils.aoa_to_sheet(aoa);
    per.forEach((p, i) => p.forEach((t, r) => {
      const a = X.utils.encode_cell({r: 4 + r, c: i + 1});
      const who = t.src === 'fixed' || t.src === 'income' ? 'קבוע' : t.src === 'manual' ? 'ידני' : (t.person || '');
      ws[a] = {t: 'n', v: r2(t.amount), z: '#,##0.00', c: [{a: 'איפה הכסף', t: `${ddmm(t.date)} ${t.desc}${who ? ` (${who})` : ''}`}]};
      ws[a].c.hidden = true;
    }));
    ws[X.utils.encode_cell({r: totRow - 1, c: 0})] = {t: 's', v: 'סה"כ'};
    cols.forEach((c, i) => { const L = X.utils.encode_col(i + 1); ws[`${L}${totRow}`] = {t: 'n', f: `SUM(${L}5:${L}${totRow - 1})`, v: per[i].reduce((s, t) => s + t.amount, 0), z: '#,##0.00'}; });
    ws['!ref'] = X.utils.encode_range({s: {r: 0, c: 0}, e: {r: totRow - 1, c: cols.length}});
    ws['!cols'] = [{wch: 10}, ...cols.map(c => ({wch: Math.max(9, Math.min(18, c.name.length + 2))}))];
    X.utils.book_append_sheet(wb, ws, sheetName(m));
    for (const t of list) { const c = catById(t.cat); flat.push([t.id, t.date, sheetName(m), c.name, groupById(c.group).name, r2(t.amount), t.desc, t.text || '', t.pay || '', t.person || '', t.src === 'chat' ? 'ווטסאפ' : t.src === 'income' ? (t.ok ? 'הכנסה' : 'הכנסה (צפויה)') : t.src === 'fixed' ? (t.ok ? 'קבועה' : 'קבועה (צפויה)') : 'ידני', t.note || '', t.refundExp ? (t.refundDone ? 'התקבל' : r2(t.refundAmt || t.amount)) : '']); }
  }
  const sum = [['קבוצה', 'סעיף', ...months.map(sheetName), 'סה"כ']];
  if (flat.length === 1) { toast('אין תנועות לייצוא בטווח הזה'); return; }
  for (const c of cols) { const v = months.map(m => r2(flat.filter(r => r[2] === sheetName(m) && r[3] === c.name).reduce((s, r) => s + r[5], 0)));   /* [id, date, sheet, section, group, amount, …] */ if (v.some(x => x)) sum.push([groupById(c.group).name, c.name, ...v, r2(v.reduce((a, b) => a + b, 0))]); }
  const ws2 = X.utils.aoa_to_sheet(sum); ws2['!cols'] = [{wch: 20}, {wch: 24}, ...months.map(() => ({wch: 11})), {wch: 12}];
  X.utils.book_append_sheet(wb, ws2, 'סיכום');
  const ws3 = X.utils.aoa_to_sheet(flat); ws3['!cols'] = [{wch: 8}, {wch: 11}, {wch: 7}, {wch: 22}, {wch: 18}, {wch: 10}, {wch: 26}, {wch: 40}, {wch: 18}, {wch: 7}, {wch: 12}, {wch: 18}, {wch: 10}];
  ws3['!autofilter'] = {ref: `A1:M${flat.length}`};
  X.utils.book_append_sheet(wb, ws3, 'תנועות');
  const kws = [['מילת מפתח', 'סעיף'], ...D.kw.filter(([, c]) => catById(c)).map(([k, c]) => [k, catById(c).name])];
  X.utils.book_append_sheet(wb, X.utils.aoa_to_sheet(kws), 'מילות מפתח');
  X.utils.book_append_sheet(wb, X.utils.aoa_to_sheet([['שדה', 'ערך'], ['פורמט', FILE_FORMAT], ['גרסת אפליקציה', APP_VERSION], ['נוצר', new Date().toLocaleString('he-IL')]]), 'מידע');
  const buf = X.write(wb, {bookType: 'xlsx', type: 'array'});
  D.meta.lastExport = Date.now(); save(false);
  download(`איפה הכסף - לאקסל ${sheetName(from)}${from !== to ? '-' + sheetName(to) : ''}.xlsx`, new Blob([buf], {type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));
  toast(`הקובץ ירד${pendSkipped ? ` · ${pendSkipped} שורות בלי סעיף לא נכללו` : ''}${overflow.length ? ` · יותר מ-50 שורות ב: ${overflow.join(', ')}` : ''}`, 5000);
}

/* =============== events =============== */
const ACT = {
  view: el => { if (el.dataset.rv) S.rv = el.dataset.rv; S.view = el.dataset.v; stack.slice().forEach(e => closeSheet(e)); window.scrollTo(0, 0); render(); },
  'go-review': () => ACT.view({dataset: {v: 'review'}}),
  month: el => { S.month = addM(S.month, +el.dataset.k); render(); },
  import: () => $('#f-chat').click(),
  restore: () => $('#f-restore').click(),
  backup: () => backupSheet(),
  wipe: () => wipeSheet(),
  add: () => openSheet(() => `${head('מה לעשות?')}<div class="actlist">
      <button data-act="import">ייבוא צ'אט מווטסאפ <span>💬</span></button>
      <button data-act="manual">הוספה ידנית (הכנסה / הוצאה שלא בקבוצה) <span>✍️</span></button>
      <button data-act="fixed">תנועות קבועות <span>🔁</span></button></div>`),
  manual: el => { if (stack.length && !el.dataset.id) closeSheet(); manualSheet(null, el.dataset.id || ''); },
  fixed: () => { if (stack.length && stack[stack.length - 1].sh.querySelector('[data-act=fixed]')) closeSheet(); fixedList(); },
  'fx-edit': el => fixedEdit(el.dataset.id),
  'fx-ok': el => { const k = el.dataset.k, ok = !(D.fixedM[k] || {}).ok; if (k.startsWith('inc:') && ok && (D.fixedM[k] || {}).amount == null) { const f = entryByKey(k); setFixedM(k, {ok, amount: f.expected || f.amount}); } else setFixedM(k, {ok}); },
  'fx-amt': el => { const k = el.dataset.k, f = entryByKey(k); askAmount(k.startsWith('inc:') ? 'כמה נכנס בפועל' : 'סכום בחודש הזה', f.ok ? f.amount : (f.expected || f.amount), v => setFixedM(k, {amount: v, ok: true})); },
  cat: el => el.dataset.id && catSheet(el.dataset.id),
  catedit: el => catEditSheet(el.dataset.id),
  tx: el => { const id = el.dataset.id; if ((id.startsWith('f:') || id.startsWith('i:')) && S.view === 'review' && !stack.length) { const k = id.startsWith('i:') ? 'inc:' + id.slice(2) : id.slice(2), f = entryByKey(k); return askAmount(k.startsWith('inc:') ? 'כמה נכנס בפועל' : 'סכום שירד בפועל', f.expected || f.amount, v => setFixedM(k, {amount: v, ok: true})); } txSheet(id); },
  't-cat': el => { const t = findTx(el.dataset.id); pickCat('לאיזה סעיף?', null, t.cat, c => setChatCat(el.dataset.id, c)); },
  't-amt': el => { const t = findTx(el.dataset.id); askAmount('הסכום הנכון', Math.abs(t.amount), v => setOv(el.dataset.id, {amount: v})); },
  't-pay': el => { const id = el.dataset.id; openSheet(() => `${head('אמצעי תשלום')}<div class="chips">${D.cfg.pays.map(p => `<button data-p="${esc(p)}">${esc(p)}</button>`).join('')}</div>`, {mount(sh) { sh.addEventListener('click', ev => { const b = ev.target.closest('[data-p]'); if (b) { closeSheet(); setOv(id, {pay: b.dataset.p}); } }); }}); },
  't-note': el => { const t = findTx(el.dataset.id); askText('הערה', t.note || '', v => setOv(el.dataset.id, {note: v || null})); },
  't-reset': el => setOv(el.dataset.id, {cat: null, amount: null, pay: null, rfOf: null}),
  't-skip': el => { const t = findTx(el.dataset.id); if (t.skipped) { setOv(el.dataset.id, {skip: t.autoSkip ? false : null}); toast('השורה הוחזרה'); } else { setOv(el.dataset.id, {skip: true}); closeSheet(); toast('השורה לא תיספר'); } },
  'rf-done': el => { const id = el.dataset.id, st = rfState(id), v = !st.done; setOv(id, {refundDone: v === st.auto ? null : v}); },
  'rf-link': el => { setOv(el.dataset.id, {rfOf: el.dataset.e, cat: null}); toast('ההחזר חובר להוצאה'); },
  'rf-inc': el => { setOv(el.dataset.id, {rfOf: 'none', cat: el.dataset.c}); toast(`← ${catById(el.dataset.c).name}`); },
  // not a refund of this expense: choose right away from every section (income too); nothing changes until one is chosen
  'rf-none': el => { const id = el.dataset.id, t = findTx(id);
    pickCat('לאיזה סעיף?', null, t && t.rfOf ? '' : (t && t.cat) || '', c => {
      if (c === SKIP) { setOv(id, {rfOf: 'none', skip: true}); toast('השורה לא תיספר'); return; }
      setOv(id, {rfOf: 'none', cat: c}); toast(`← ${catById(c).name}`);
    }, true); },
  'rf-pick': el => { const id = el.dataset.id; openSheet(() => `${head('לאיזו הוצאה שייך ההחזר?')}<div class="actlist">${rfChoices(id).map(e => `<button data-e="${esc(e)}">${esc(rfLabel(e))} <span>↩︎</span></button>`).join('')}</div>`,
    {mount(sh) { sh.addEventListener('click', ev => { const b = ev.target.closest('[data-e]'); if (b) { closeSheet(); setOv(id, {rfOf: b.dataset.e, cat: null}); toast('ההחזר חובר להוצאה'); } }); }}); },
  'rv-cat': el => setChatCat(el.dataset.id, el.dataset.c),
  'rvg-cat': el => groupPick(el.dataset.g, el.dataset.c),
  'rvg-more': el => { const g = el.dataset.g, t = (pendGroups().find(x => x.key === g) || {items: [{}]}).items[0]; pickCat('לאיזה סעיף?', null, '', c => groupPick(g, c)); },
  'rvg-skip': el => groupPick(el.dataset.g, SKIP),
  rvseg: el => { S.rv = el.dataset.v; window.scrollTo(0, 0); render(); },
  'mn-ok': el => setOv(el.dataset.id, {numOk: true}),
  'mn-fix': el => { const t = findTx(el.dataset.id); askAmount('הסכום הנכון', Math.abs(t.amount), v => setOv(el.dataset.id, {amount: v, numOk: true})); },
  'rv-more': el => pickCat('לאיזה סעיף?', null, '', c => setChatCat(el.dataset.id, c)),
  'dup-skip': el => { setOv(el.dataset.id, {skip: true}); toast('נספר פעם אחת'); },
  'dup-ok': el => { for (const id of [el.dataset.a, el.dataset.b]) if (D.chat.some(x => x.id === id)) D.ov[id] = {...(D.ov[id] || {}), notDup: true}; save(); render(); },
  'fxa-yes': el => fxAnswer(el.dataset.k, true),
  'fxa-no': el => fxAnswer(el.dataset.k, false),
  'fxc-seen': () => { D.meta.fxChanges = []; save(); render(); },
  'rv-skip': el => { const it = D.chat.find(x => x.id === el.dataset.id); setOv(el.dataset.id, {skip: true}); if (it) learnSkipSheet(it); else toast('השורה לא תיספר'); },
  'm-edit': el => { const x = D.manual.find(m => m.id === el.dataset.id); closeSheet(); manualSheet(x); },
  'm-del': el => { if (!el.dataset.sure) { el.dataset.sure = 1; el.innerHTML = 'בטוח? לחיצה נוספת מוחקת <span>🗑</span>'; return; } D.manual = D.manual.filter(m => m.id !== el.dataset.id); save(); closeSheet(); render(); refreshSheets(); },
  dismiss: el => { S.dismissed.add(el.dataset.id); render(); },
  settings: () => settingsSheet(),
  income: () => incomeSheet(),
  incdet: () => incDetailSheet(),
  baldet: () => balSheet(),
  cats: () => catsSheet(),
  budgets: () => incomeSheet(),
  grp: el => groupSheet(el.dataset.id),
  goto: el => { S.month = el.dataset.m; render(); window.scrollTo(0, 0); },
  f: el => { S.f[el.dataset.k] = el.dataset.v; render(); },
  kw: () => kwSheet(),
  guide: () => guideSheet(),
  sync: () => syncSheet(),
  xlimport: () => $('#f-xl').click(),
  pays: () => paysSheet(),
  pay: el => paySheet(el.dataset.p),
  importset: () => importSettings(),
  export: () => exportSheet(),
  close: () => closeSheet(),
  undo: () => doUndo(),
  update: () => updateSheet(),
  lock: () => lockSheet(),
  checkupd: () => checkUpdate(),
  history: () => historySheet(),
  'share-now': () => { D = blank(); save(false); importPendingShare(); },
};
document.addEventListener('click', ev => {
  const el = ev.target.closest('[data-act]'); if (!el) return;
  const f = ACT[el.dataset.act]; if (!f) return;
  ev.preventDefault(); f(el, ev);
});
$('#f-chat').addEventListener('change', async ev => { const f = ev.target.files[0]; ev.target.value = ''; stack.slice().forEach(e => closeSheet(e)); await importFile(f); });
$('#f-xl').addEventListener('change', async ev => { const f = ev.target.files[0]; ev.target.value = ''; await importExcel(f); });
$('#f-restore').addEventListener('change', async ev => { const f = ev.target.files[0]; ev.target.value = ''; await restore(f); });
document.addEventListener('input', ev => { if (ev.target.id === 'dq') { S.f.q = ev.target.value; const pos = ev.target.selectionStart; render(); const i = $('#dq'); if (i) { i.focus(); i.setSelectionRange(pos, pos); } } });
document.addEventListener('change', ev => { const k = ev.target.dataset && ev.target.dataset.fsel; if (k) { S.f[k] = ev.target.value; render(); } });
document.addEventListener('keydown', ev => { if (ev.key === 'Escape' && stack.length) closeSheet(); });
document.addEventListener('input', ev => { const k = ev.target.dataset && ev.target.dataset.lw; if (k != null) { S.lw = S.lw || {}; S.lw[k] = ev.target.value; } });
document.addEventListener('change', ev => { const k = ev.target.dataset && ev.target.dataset.learn; if (k != null) { S.noLearn = S.noLearn || new Set(); ev.target.checked ? S.noLearn.delete(k) : S.noLearn.add(k); } });
// the screen always fits the phone: no pinch zoom, no double-tap zoom, no sideways drag (iOS ignores the viewport tag)
document.addEventListener('gesturestart', ev => ev.preventDefault(), {passive: false});
document.addEventListener('touchmove', ev => { if (ev.touches.length > 1) ev.preventDefault(); }, {passive: false});
// Swipe between months from every screen, and in every window that shows a month (סעיף, קבוצה, הכנסות, יתרה, אמצעי תשלום).
// Experimental direction: swipe right = next month, swipe left = previous month. To switch back: SWIPE_RIGHT_NEXT = false.
const SWIPE_RIGHT_NEXT = true;
let swipe = null;
const topSheet = () => stack[stack.length - 1];
document.addEventListener('touchstart', ev => {
  swipe = null;
  if (!D || ev.touches.length !== 1 || document.body.classList.contains('locked')) return;
  if (stack.length && !(topSheet().opts.month && topSheet().sh.contains(ev.target))) return;   // a window that is not about a month (a line, a form) stays put
  if (ev.target.closest('input, textarea, select, nav')) return;
  const t = ev.touches[0]; swipe = {x: t.clientX, y: t.clientY, at: Date.now()};
}, {passive: true});
// long pages (פירוט, לבדיקה) scroll while the finger moves, and the browser may then end the touch with
// "touchcancel" instead of "touchend" – so the last position is kept and both endings count.
document.addEventListener('touchmove', ev => { if (swipe && ev.touches.length === 1) { swipe.lx = ev.touches[0].clientX; swipe.ly = ev.touches[0].clientY; } }, {passive: true});
document.addEventListener('touchcancel', ev => swipeEnd(swipe && swipe.lx != null ? {clientX: swipe.lx, clientY: swipe.ly} : null), {passive: true});
document.addEventListener('touchend', ev => swipeEnd(ev.changedTouches[0]), {passive: true});
function swipeEnd(t) {
  if (!swipe || !t) { swipe = null; return; }
  const dx = t.clientX - swipe.x, dy = t.clientY - swipe.y, dt = Date.now() - swipe.at;
  swipe = null;
  if (Math.abs(dx) < 70 || Math.abs(dx) < Math.abs(dy) * 1.8 || dt > 700) return;
  const k = (dx > 0) === SWIPE_RIGHT_NEXT ? 1 : -1, m = addM(S.month, k);
  if (m < minMonth() || m > addM(curMonth(), 1)) return toast(k < 0 ? 'אין חודשים קודמים' : 'אין חודשים נוספים', 1400);
  S.month = m; render();
  const top = topSheet(); if (top && top.opts.month) top.paint();
  // the new month comes in from the side the finger came from
  const el = top && top.opts.month ? top.sh : $('#app');
  el.style.transition = 'none'; el.style.opacity = '.4'; el.style.transform = `translateX(${dx > 0 ? -24 : 24}px)`;
  requestAnimationFrame(() => { el.style.transition = 'opacity .18s, transform .18s'; el.style.opacity = ''; el.style.transform = ''; });
  if (top && top.opts.month) toast(mName(m), 900);
}
// Back button: closes an open window, then returns to the home screen; leaving needs two presses in a row.
let guardActive = false;          // a history entry pushed during a tap (browsers honour those)
function pushGuard() { try { history.pushState({kesef: 1}, ''); } catch {} }
const onUser = () => { if (!guardActive) { pushGuard(); guardActive = true; } };
document.addEventListener('pointerdown', onUser, true);
document.addEventListener('keydown', onUser, true);
window.addEventListener('popstate', () => {
  guardActive = false;
  if (stack.length) { closeSheet(); pushGuard(); return; }
  if (S.view !== 'home') { S.view = 'home'; render(); window.scrollTo(0, 0); pushGuard(); return; }
  toast('כדי לצאת יש ללחוץ שוב "הקודם"', 2000);   // no entry pushed now, so a second press leaves
  setTimeout(() => { if (!guardActive) pushGuard(); }, 2200);
});
/* =============== installed app: share target + updates =============== */
const APP_VERSION = '3.0';
async function importPendingShare() {
  let c; try { c = await caches.open('shared'); } catch { return; }
  const fr = await c.match('./shared-file'), ir = await c.match('./shared-info');
  if (!fr) { S.pendingShare = false; if (ir) { await c.delete('./shared-info'); toast('השיתוף הגיע בלי קובץ. נסו שוב: ייצוא צ\'אט ← ללא מדיה', 5000); } return; }
  if (!D) { S.pendingShare = true; render(); toast('קודם שחזור מגיבוי – הצ\'אט ששותף מחכה ויובא מיד אחר כך', 6000); return; }
  const info = ir ? await ir.json().catch(() => ({})) : {}, f0 = (info.files || [])[0] || {};
  const file = new File([await fr.blob()], f0.name || 'WhatsApp.txt');
  await c.delete('./shared-file'); await c.delete('./shared-info');
  S.pendingShare = false;
  stack.slice().forEach(e => closeSheet(e));
  await importFile(file, true);
}
let swReg = null, reloading = false;
async function showUpdate(w) {
  S.waitingSW = w;
  try { S.newVersion = await (await fetch('version.json', {cache: 'no-store'})).json(); } catch { S.newVersion = null; }
  if (D) render();
}
function watchUpdate(reg) {
  if (reg.waiting && navigator.serviceWorker.controller) showUpdate(reg.waiting);
  reg.addEventListener('updatefound', () => { const w = reg.installing; if (w) w.addEventListener('statechange', () => { if (w.state === 'installed' && navigator.serviceWorker.controller) showUpdate(w); }); });
  setInterval(() => reg.update().catch(() => {}), 6 * 3600e3);
}
// what changed in every version, newest first (the list comes with the app, so it works offline)
async function historySheet() {
  let h = [];
  try { h = await (await fetch('changelog.json')).json(); } catch {}
  openSheet(() => `${head('היסטוריית גרסאות')}
    ${h.length ? h.map((v, i) => `<div class="card hist${v.version === APP_VERSION ? ' cur' : ''}">
      <div class="hist-h"><b>גרסה ${esc(v.version)}</b><span>${v.date ? dShort(v.date) : ''}${v.version === APP_VERSION ? ' · מותקנת' : ''}</span></div>
      ${i < 3 ? v.notes.map(n => `<p>• ${esc(n)}</p>`).join('') : `<details><summary>${v.notes.length === 1 ? esc(v.notes[0]).slice(0, 60) + (v.notes[0].length > 60 ? '…' : '') : v.notes.length + ' שינויים'}</summary>${v.notes.map(n => `<p>• ${esc(n)}</p>`).join('')}</details>`}
    </div>`).join('') : '<p class="note">לא הצלחתי לטעון את הרשימה. היא תופיע אחרי העדכון הבא.</p>'}`);
}
function updateSheet() {
  const v = S.newVersion || {};
  openSheet(() => `${head('גרסה חדשה' + (v.version ? ' ' + esc(v.version) : ''))}
    ${(v.notes || []).length ? `<div class="card" style="padding:10px 16px">${v.notes.map(n => `<p style="margin:6px 0;line-height:1.5">• ${esc(n)}</p>`).join('')}</div>` : '<p class="note">אין פירוט שינויים.</p>'}
    <p class="note">העדכון משנה רק את האפליקציה. הנתונים שלך נשארים בטלפון ולא משתנים. כרגע מותקנת גרסה ${APP_VERSION}.</p>
    <button class="primary" id="up-go">לעדכן עכשיו</button><button class="ghost" data-act="close">אחר כך</button>`, {
    mount(sh) { sh.addEventListener('click', ev => { if (!ev.target.closest('#up-go')) return; if (!S.waitingSW) return toast('העדכון כבר הותקן'); S.updating = true; S.waitingSW.postMessage('skip-waiting'); toast('מעדכן…'); }); }
  });
}
// pull-to-refresh (or any reload) = check for an update
const IS_RELOAD = (() => { try { const n = performance.getEntriesByType('navigation')[0]; return n ? n.type === 'reload' : !!(performance.navigation && performance.navigation.type === 1); } catch { return false; } })();
async function checkUpdate(auto) {
  if (!swReg) return auto ? 0 : toast('אפשר לבדוק עדכון רק באפליקציה המותקנת');
  if (!navigator.onLine) return auto ? 0 : toast('צריך אינטרנט כדי לבדוק אם יש עדכון');
  toast('בודק…');
  try { await swReg.update(); } catch {}
  setTimeout(() => { if (S.waitingSW || swReg.waiting || swReg.installing) { if (swReg.waiting && !S.waitingSW) showUpdate(swReg.waiting); toast('יש גרסה חדשה – ההודעה מופיעה למעלה'); } else toast(`אין גרסה חדשה. מותקנת גרסה ${APP_VERSION}`); }, 2500);
}
if ('serviceWorker' in navigator) {
  navigator.serviceWorker.register('sw.js').then(reg => { swReg = reg; watchUpdate(reg); if (IS_RELOAD) checkUpdate(true); }).catch(() => {});
  navigator.serviceWorker.addEventListener('controllerchange', () => { if (S.updating && !reloading) { reloading = true; location.reload(); } });
}
/* =============== app lock: 4-digit code + fingerprint =============== */
// The code is never stored: only a PBKDF2 hash with a random salt. The fingerprint is checked by Android itself
// (WebAuthn, platform authenticator, user verification); the app only learns "verified" or not.
// This is a lock on opening the app, not encryption of the stored data.
const LOCK_KEY = 'kesef-lock-v1';
let L = (() => { try { return JSON.parse(localStorage.getItem(LOCK_KEY)) || {}; } catch { return {}; } })();
const saveLock = () => { try { localStorage.setItem(LOCK_KEY, JSON.stringify(L)); } catch {} };
const lockOn = () => !!(L.pin && L.pin.hash);
const b64 = buf => btoa(String.fromCharCode(...new Uint8Array(buf)));
const unb64 = t => Uint8Array.from(atob(t), c => c.charCodeAt(0));
const FP_ICON = '<svg width="32" height="32" viewBox="0 0 24 24" stroke="#9C4A33" fill="none" stroke-linecap="round" stroke-linejoin="round" stroke-width="1.5" aria-hidden="true"><path d="M5.2 7.2A8.2 8.2 0 0 1 20 11.5"/><path d="M4.2 10.6A8.6 8.6 0 0 0 4 12.5v1.5"/><path d="M7 17.5V12a5 5 0 0 1 10 0c0 3.6-.6 6.5-1.8 8.6"/><path d="M12 12c0 3.5-.6 6.4-2 8.6"/><path d="M9.6 12.2a2.4 2.4 0 0 1 4.8-.2c0 1.5-.1 2.9-.3 4.2"/><path d="M19.8 14.8c-.1 1-.3 2-.6 2.9"/></svg>';
async function pinHash(pin, salt, iter) {
  const k = await crypto.subtle.importKey('raw', new TextEncoder().encode(pin), 'PBKDF2', false, ['deriveBits']);
  return b64(await crypto.subtle.deriveBits({name: 'PBKDF2', hash: 'SHA-256', salt, iterations: iter}, k, 256));
}
async function setPin(pin) {
  const salt = crypto.getRandomValues(new Uint8Array(16)), iter = 150000;
  L.pin = {salt: b64(salt), iter, hash: await pinHash(pin, salt, iter)}; L.fails = 0; L.until = 0; if (L.grace == null) L.grace = 60; saveLock();
}
async function checkPin(pin) { return !!L.pin && (await pinHash(pin, unb64(L.pin.salt), L.pin.iter)) === L.pin.hash; }
async function bioSupported() { try { return !!(window.PublicKeyCredential && await PublicKeyCredential.isUserVerifyingPlatformAuthenticatorAvailable()); } catch { return false; } }
async function bioRegister() {
  const cred = await navigator.credentials.create({publicKey: {challenge: crypto.getRandomValues(new Uint8Array(32)), rp: {name: 'איפה הכסף'},
    user: {id: crypto.getRandomValues(new Uint8Array(16)), name: 'איפה הכסף', displayName: 'איפה הכסף'},
    pubKeyCredParams: [{type: 'public-key', alg: -7}, {type: 'public-key', alg: -257}],
    authenticatorSelection: {authenticatorAttachment: 'platform', userVerification: 'required', residentKey: 'discouraged'}, attestation: 'none', timeout: 60000}});
  L.bio = {id: b64(cred.rawId)}; saveLock();
}
async function bioVerify() {
  const a = await navigator.credentials.get({publicKey: {challenge: crypto.getRandomValues(new Uint8Array(32)), allowCredentials: [{type: 'public-key', id: unb64(L.bio.id)}], userVerification: 'required', timeout: 60000}});
  return !!(a && (new Uint8Array(a.response.authenticatorData)[32] & 0x04));   // "user verified" flag
}
function pinScreen(o) {
  const el = document.createElement('div'); el.className = 'lock'; el.setAttribute('role', 'dialog'); el.setAttribute('aria-modal', 'true');
  let pin = '', busy = false;
  const paint = err => { el.innerHTML = `<div class="lock-in">
      <img src="icon-192.png" alt="" class="lock-ic">
      <h2>${o.title}</h2><p class="note">${err ? `<span class="red">${esc(err)}</span>` : (o.sub || '')}</p>
      <div class="pdots ${err ? 'shake' : ''}" aria-label="${pin.length} מתוך 4 ספרות">${[0, 1, 2, 3].map(i => `<i class="${i < pin.length ? 'on' : ''}"></i>`).join('')}</div>
      <div class="keys">${[1, 2, 3, 4, 5, 6, 7, 8, 9].map(n => `<button type="button" data-k="${n}">${n}</button>`).join('')}
        ${o.bio ? `<button type="button" data-k="bio" aria-label="כניסה בטביעת אצבע">${FP_ICON}</button>` : '<span></span>'}<button type="button" data-k="0">0</button><button type="button" data-k="del" aria-label="מחיקת ספרה">⌫</button></div>
      ${o.forgot ? '<button type="button" class="linkbtn" data-k="forgot">שכחתי את הקוד</button>' : ''}${o.cancel ? '<button type="button" class="linkbtn" data-k="cancel">ביטול</button>' : ''}
    </div>`; };
  paint();
  el.addEventListener('click', async ev => {
    const b = ev.target.closest('[data-k]'); if (!b || busy) return;
    const k = b.dataset.k;
    if (k === 'bio') return o.bio();
    if (k === 'forgot') return o.forgot();
    if (k === 'cancel') { el.remove(); return o.cancel && o.cancel(); }
    if (k === 'del') { pin = pin.slice(0, -1); return paint(); }
    if (pin.length >= 4) return;
    pin += k; paint();
    if (pin.length === 4) { busy = true; const r = await o.onPin(pin); busy = false; if (r !== true) { pin = ''; if (el.isConnected) paint(r || 'קוד שגוי'); } }
  });
  document.body.appendChild(el);
  return el;
}
let locked = false, hiddenAt = 0;
// "unlocked in this session" survives a reload of the open app, but not closing it (sessionStorage is per app window).
// While the lock screen is up, or after the app sat in the background past the chosen time, a reload still asks.
const SES_KEY = 'kesef-session';
const ses = {get() { try { return JSON.parse(sessionStorage.getItem(SES_KEY)) || {}; } catch { return {}; } }, set(o) { try { sessionStorage.setItem(SES_KEY, JSON.stringify(o)); } catch {} }};
function requireUnlock(then) {
  if (!lockOn()) return then();
  if (locked) return;
  locked = true; document.body.classList.add('locked'); ses.set({});
  let el = null;
  const done = () => { locked = false; hiddenAt = 0; ses.set({ok: true}); document.body.classList.remove('locked', 'cover'); if (el) el.remove(); then(); };
  const tryBio = async () => { try { if (await bioVerify()) done(); } catch {} };
  el = pinScreen({title: 'איפה <span class="script">הכסף</span>', sub: L.bio ? 'קוד או טביעת אצבע' : 'הקלידו את הקוד',
    onPin: async pin => {
      if (L.until && Date.now() < L.until) return `יותר מדי ניסיונות. אפשר לנסות שוב בעוד ${Math.ceil((L.until - Date.now()) / 1000)} שניות`;
      if (await checkPin(pin)) { L.fails = 0; L.until = 0; saveLock(); done(); return true; }
      L.fails = (L.fails || 0) + 1;
      if (L.fails >= 5) L.until = Date.now() + 30000 * Math.pow(2, Math.min(5, L.fails - 5));
      saveLock();
      return L.fails >= 5 ? `יותר מדי ניסיונות. אפשר לנסות שוב בעוד ${Math.ceil((L.until - Date.now()) / 1000)} שניות` : 'קוד שגוי';
    },
    bio: L.bio ? tryBio : null, forgot: () => forgotPin()});
  if (L.bio) setTimeout(tryBio, 400);
}
function forgotPin() {
  const el = document.createElement('div'); el.className = 'lock'; el.style.zIndex = 1001;
  el.innerHTML = `<div class="lock-in"><h2>שכחתי את הקוד</h2>
    <p class="note" style="line-height:1.6">אין דרך לשחזר את הקוד – הוא לא נשמר בשום מקום, וזה מה ששומר על הנתונים. האפשרות היחידה היא למחוק את כל הנתונים מהטלפון ולשחזר אותם מקובץ הגיבוי האחרון.</p>
    <button type="button" class="primary" id="fg-wipe" style="background:var(--red)">מחיקת הנתונים מהטלפון</button>
    <button type="button" class="linkbtn" id="fg-back">חזרה</button></div>`;
  el.addEventListener('click', ev => {
    if (ev.target.closest('#fg-back')) return el.remove();
    const w = ev.target.closest('#fg-wipe'); if (!w) return;
    if (!w.dataset.sure) { w.dataset.sure = 1; w.textContent = 'בטוח? לחיצה נוספת מוחקת הכל'; return; }
    wipeStore();
    L = {}; D = null; locked = false; document.body.classList.remove('locked', 'cover');
    document.querySelectorAll('.lock').forEach(x => x.remove()); render(); toast('הנתונים נמחקו – אפשר לשחזר מגיבוי', 4000);
  });
  document.body.appendChild(el);
}
function newPinFlow(after) {
  pinScreen({title: 'בחירת קוד', sub: '4 ספרות שרק אתם יודעים', cancel: () => {}, onPin: async first => {
    document.querySelectorAll('.lock').forEach(x => x.remove());
    pinScreen({title: 'שוב, לאישור', sub: 'הקלידו את אותו קוד', cancel: () => {}, onPin: async second => {
      if (second !== first) return 'הקודים לא זהים – נסו שוב';
      await setPin(first); document.querySelectorAll('.lock').forEach(x => x.remove()); after(); return true;
    }});
    return true;
  }});
}
function verifyPinThen(title, after) {
  pinScreen({title, sub: 'הקלידו את הקוד הנוכחי', cancel: () => {}, onPin: async pin => {
    if (!(await checkPin(pin))) return 'קוד שגוי';
    document.querySelectorAll('.lock').forEach(x => x.remove()); after(); return true;
  }});
}
async function lockSheet() {
  const bioOk = await bioSupported();
  const graceOpt = [[0, 'מיד'], [60, 'אחרי דקה'], [300, 'אחרי 5 דקות'], [900, 'אחרי רבע שעה']];
  openSheet(() => `${head('נעילת האפליקציה')}
    <p class="note">כשהנעילה פעילה, פתיחת האפליקציה דורשת קוד בן 4 ספרות${bioOk ? ' או טביעת אצבע' : ''}. </p>
    <div class="actlist">
      ${lockOn() ? '<button data-l="change">שינוי הקוד <span>›</span></button>' : '<button data-l="set">הפעלת נעילה בקוד <span>›</span></button>'}
      ${lockOn() && bioOk ? (L.bio ? '<button data-l="bio-off">טביעת אצבע <span class="v">פעילה · לכיבוי</span></button>' : '<button data-l="bio-on">הוספת טביעת אצבע <span>›</span></button>') : ''}
      ${lockOn() ? `<label style="cursor:default">לבקש קוד כשחוזרים לאפליקציה <select class="inp" id="l-grace" style="width:auto;padding:8px 10px">${graceOpt.map(([v, t]) => `<option value="${v}" ${(L.grace ?? 60) === v ? 'selected' : ''}>${t}</option>`).join('')}</select></label>` : ''}
      ${lockOn() ? '<button data-l="off" class="danger">כיבוי הנעילה <span>›</span></button>' : ''}
    </div>
    ${lockOn() && !bioOk ? '<p class="note">טביעת אצבע לא זמינה בדפדפן הזה – אפשר להשתמש רק בקוד.</p>' : ''}
    <p class="note">🔒 הקוד לא נשמר – רק "טביעה" חד־כיוונית שלו, ובטלפון בלבד. טביעת האצבע נבדקת על ידי אנדרואיד עצמו, והאפליקציה מקבלת רק "אושר" או "לא אושר". חשוב לזכור: זו נעילה של פתיחת האפליקציה, לא הצפנה של הנתונים בטלפון. אם שוכחים את הקוד – מוחקים ומשחזרים מגיבוי.</p>`, {
    live: true,
    mount(sh, e) {
      sh.addEventListener('change', ev => { if (ev.target.id === 'l-grace') { L.grace = +ev.target.value; saveLock(); toast('נשמר'); } });
      sh.addEventListener('click', async ev => {
        const b = ev.target.closest('[data-l]'); if (!b) return;
        const a = b.dataset.l;
        if (a === 'set') newPinFlow(() => { e.paint(); render(); toast('הנעילה הופעלה'); });
        if (a === 'change') verifyPinThen('שינוי הקוד', () => newPinFlow(() => { e.paint(); toast('הקוד הוחלף'); }));
        if (a === 'off') verifyPinThen('כיבוי הנעילה', () => { L = {}; saveLock(); e.paint(); render(); toast('הנעילה כובתה'); });
        if (a === 'bio-on') { try { await bioRegister(); e.paint(); render(); toast('טביעת האצבע נוספה'); } catch (err) { toast('לא הצלחתי להוסיף טביעת אצבע' + (err && err.name === 'NotAllowedError' ? ' (בוטל)' : ''), 4000); } }
        if (a === 'bio-off') { delete L.bio; saveLock(); e.paint(); render(); toast('טביעת האצבע כובתה – נשאר הקוד'); }
      });
    }
  });
}
// leaving the app hides its content at once; coming back after the chosen time asks for the code
// a reload also turns the page "hidden" for a moment (after pagehide) – that is not leaving the app
let unloading = false; addEventListener('pagehide', () => { unloading = true; });
document.addEventListener('visibilitychange', () => {
  if (!lockOn() || locked || unloading) return;
  if (document.hidden) { hiddenAt = Date.now(); ses.set({...ses.get(), hiddenAt}); document.body.classList.add('cover'); return; }
  document.body.classList.remove('cover');
  if (hiddenAt && Date.now() - hiddenAt >= (L.grace ?? 60) * 1000) requireUnlock(() => {}); else { hiddenAt = 0; ses.set({...ses.get(), hiddenAt: 0}); }
});
// privacy shield – always on, with or without the lock: the moment the app stops being in front
// (app switcher, home, another app, notification shade) the screen turns plain white, so no amounts
// show in the recent-apps screen. "blur" comes before "hidden", so it catches the switcher earlier.
// An open dropdown or file picker also blurs the page – that is not leaving the app, so it is skipped.
const shieldEl = document.createElement('div'); shieldEl.id = 'shield'; shieldEl.setAttribute('aria-hidden', 'true'); document.body.appendChild(shieldEl);
const shieldOn = () => document.body.classList.add('shield');
const shieldOff = () => { if (!document.hidden) document.body.classList.remove('shield'); };
addEventListener('blur', () => { const a = document.activeElement; if (a && (a.tagName === 'SELECT' || (a.tagName === 'INPUT' && a.type === 'file'))) return; shieldOn(); });
addEventListener('focus', shieldOff);
addEventListener('pagehide', shieldOn);
addEventListener('pageshow', shieldOff);
document.addEventListener('visibilitychange', () => { if (document.hidden) shieldOn(); else shieldOff(); });
addEventListener('pointerdown', shieldOff, true);   // safety net: touching the screen always brings the app back
(async () => {
  D = await loadStore();
  migrate(); migrate3(); migrate4(); migrate5(); migrate6();
  if (D) lastJSON = JSON.stringify(D);
  const sharedNow = new URLSearchParams(location.search).has('shared');
  if (sharedNow) history.replaceState(null, '', location.pathname);
  const s0 = ses.get(), awayTooLong = s0.hiddenAt && Date.now() - s0.hiddenAt >= (L.grace ?? 60) * 1000;
  if (IS_RELOAD && s0.ok && !awayTooLong) { ses.set({ok: true}); render(); if (sharedNow) importPendingShare(); }
  else requireUnlock(() => { render(); if (sharedNow) importPendingShare(); });
})();
