/* =====================================================================
   الجرس المدرسي الذكي — يعمل محلياً على جهاز واحد بدون خادم.
   الإعدادات تُحفظ في localStorage والملفات (PDF/صوت/فيديو/صور) في IndexedDB.
   ===================================================================== */
'use strict';
const APP_VERSION = '2026.11.24-e';

/* ------------------------------ أدوات ------------------------------ */
const $ = (s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => [...r.querySelectorAll(s)];
function h(tag, attrs, ...kids) {
  const el = document.createElement(tag);
  let value;
  for (const [k, v] of Object.entries(attrs || {})) {
    if (v == null || v === false) continue;
    if (k === 'class') el.className = v;
    else if (k === 'style') el.style.cssText = v;
    else if (k === 'html') el.innerHTML = v;
    else if (k === 'value') value = v;
    else if (k === 'checked' || k === 'selected' || k === 'disabled' || k === 'muted') el[k] = !!v;
    else if (k.startsWith('on') && typeof v === 'function') el.addEventListener(k.slice(2), v);
    else el.setAttribute(k, v === true ? '' : v);
  }
  for (const kid of kids.flat(Infinity)) {
    if (kid == null || kid === false) continue;
    el.append(kid.nodeType ? kid : String(kid));
  }
  if (value !== undefined) el.value = value;
  return el;
}
const uid = () => Math.random().toString(36).slice(2, 8) + Date.now().toString(36).slice(-4);
const pad = n => String(n).padStart(2, '0');
const clone = o => JSON.parse(JSON.stringify(o));
const sleep = ms => new Promise(r => setTimeout(r, ms));
const toLatin = s => String(s ?? '').replace(/[٠-٩]/g, d => '٠١٢٣٤٥٦٧٨٩'.indexOf(d)).replace(/[۰-۹]/g, d => '۰۱۲۳۴۵۶۷۸۹'.indexOf(d));
const norm = s => toLatin(s).trim().toLowerCase()
  .replace(/[أإآٱ]/g, 'ا').replace(/ة/g, 'ه').replace(/ى/g, 'ي')
  .replace(/[ً-ْـ]/g, '').replace(/\s+/g, ' ');
const ymd = d => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
const toSec = t => { if (!t) return null; const [a, b, c] = String(t).split(':').map(Number); return a * 3600 + (b || 0) * 60 + (c || 0); };
const fromSec = s => `${pad(Math.floor(s / 3600) % 24)}:${pad(Math.floor(s / 60) % 60)}`;
function fmtDur(s) {
  s = Math.max(0, Math.round(s));
  const hh = Math.floor(s / 3600), mm = Math.floor(s / 60) % 60, ss = s % 60;
  return (hh ? hh + ':' + pad(mm) : mm) + ':' + pad(ss);
}
function fmt12(t) {
  const s = toSec(t); if (s == null) return '';
  let hh = Math.floor(s / 3600), mm = Math.floor(s / 60) % 60;
  if (!CFG.general.clock24) { hh = hh % 12 || 12; }
  return `${hh}:${pad(mm)}`;
}
const digits = s => CFG.general.arabicDigits ? String(s).replace(/\d/g, d => '٠١٢٣٤٥٦٧٨٩'[d]) : String(s);
function toast(msg, err) {
  const t = h('div', { class: 'toast' + (err ? ' err' : '') }, msg);
  document.body.append(t); setTimeout(() => t.remove(), err ? 5000 : 2600);
}
const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

const DAYS = ['الأحد', 'الاثنين', 'الثلاثاء', 'الأربعاء', 'الخميس', 'الجمعة', 'السبت'];
const ORD = ['الأول', 'الثاني', 'الثالث', 'الرابع', 'الخامس', 'السادس', 'السابع', 'الثامن'];

/* أنواع الفترات */
const TYPES = [
  { k: 'assembly', n: 'الطابور الصباحي', start: 'بدأ الآن الطابور الصباحي', end: 'انتهى الطابور الصباحي، يرجى التوجه إلى الفصول' },
  { k: 'p0', n: 'الحصة الصفرية', lesson: true },
  { k: 'p1', n: 'الحصة الأولى', lesson: true },
  { k: 'p2', n: 'الحصة الثانية', lesson: true },
  { k: 'p3', n: 'الحصة الثالثة', lesson: true },
  { k: 'p4', n: 'الحصة الرابعة', lesson: true },
  { k: 'p5', n: 'الحصة الخامسة', lesson: true },
  { k: 'p6', n: 'الحصة السادسة', lesson: true },
  { k: 'p7', n: 'الحصة السابعة', lesson: true },
  { k: 'p8', n: 'الحصة الثامنة', lesson: true },
  { k: 'break', n: 'الفسحة' },
  { k: 'break2', n: 'الفسحة الثانية' },
  { k: 'activity', n: 'حصة النشاط', lesson: true },
  { k: 'riyada', n: 'حصة الريادة', lesson: true },
  { k: 'prayer', n: 'الصلاة', start: 'حان الآن وقت الصلاة', end: 'انتهى وقت الصلاة' },
  { k: 'dismiss', n: 'الانصراف', start: 'حان الآن وقت الانصراف، نتمنى لكم يوماً سعيداً', end: '' },
  { k: 'custom', n: 'مخصص (اكتب الاسم)' },
];
const TYPE = Object.fromEntries(TYPES.map(t => [t.k, t]));
const periodName = p => (p.type === 'custom' || p.label) ? (p.label || 'فترة') : (TYPE[p.type]?.n || 'فترة');
const isLesson = p => !!TYPE[p.type]?.lesson;

/* تحويل أي نص (من جدول المناوبة/الاحتياط) إلى مفتاح فترة */
function slotKey(str) {
  const n = norm(str);
  if (!n || /طوال|كامل|جميع|اليوم كله|عام|all/.test(n)) return '*';
  if (/طابور/.test(n)) return 'assembly';
  if (/فسح|استراح|راحه|break/.test(n)) return /ثاني|2/.test(n) ? 'break2' : 'break';
  if (/نشاط/.test(n)) return 'activity';
  if (/رياد/.test(n)) return 'riyada';
  if (/صلا/.test(n)) return 'prayer';
  if (/انصراف|خروج|نهايه الدوام/.test(n)) return 'dismiss';
  if (/صفر/.test(n) || n === '0') return 'p0';
  const ords = ['اول', 'ثاني', 'ثالث', 'رابع', 'خامس', 'سادس', 'سابع', 'ثامن'];
  const nn = n.replace(/تاني/g, 'ثاني').replace(/تالت|ثالت|تالث/g, 'ثالث').replace(/تامن/g, 'ثامن');
  for (let i = ords.length - 1; i >= 0; i--) if (nn.includes(ords[i])) return 'p' + (i + 1);
  const m = n.match(/\d+/); if (m) return 'p' + parseInt(m[0], 10);
  return 'x:' + n;
}
const periodKeys = p => new Set([p.type, slotKey(periodName(p))]);
function dayIdx(str) {
  const n = norm(str);
  const map = [['احد', 0], ['اثنين', 1], ['ثنين', 1], ['ثلاثا', 2], ['اربعا', 3], ['اربع', 3], ['خميس', 4], ['جمع', 5], ['سبت', 6],
    ['sun', 0], ['mon', 1], ['tue', 2], ['wed', 3], ['thu', 4], ['fri', 5], ['sat', 6]];
  for (const [w, i] of map) if (n.includes(w)) return i;
  return -1;
}
function weekNum(str) {
  const n = norm(str);
  if (!n || /كل|جميع/.test(n)) return 0;
  const m = n.match(/\d+/); if (m) return parseInt(m[0], 10);
  const ords = ['اول', 'ثاني', 'ثالث', 'رابع', 'خامس', 'سادس'];
  for (let i = 0; i < ords.length; i++) if (n.includes(ords[i])) return i + 1;
  return 0;
}
function parseDate(str) {
  const s = toLatin(str).trim();
  let m;
  if ((m = s.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})/))) return `${m[1]}-${pad(m[2])}-${pad(m[3])}`;
  if ((m = s.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{2,4})/))) { let y = +m[3]; if (y < 100) y += 2000; return `${y}-${pad(m[2])}-${pad(m[1])}`; }
  if (/^\d{5}(\.\d+)?$/.test(s)) { const d = new Date(Date.UTC(1899, 11, 30) + parseFloat(s) * 864e5); return d.toISOString().slice(0, 10); }
  return '';
}

/* ------------------------------ التخزين ------------------------------ */
const DB = {
  db: null,
  open() {
    return new Promise((res, rej) => {
      const r = indexedDB.open('school-bell', 1);
      r.onupgradeneeded = () => r.result.createObjectStore('files');
      r.onsuccess = () => { this.db = r.result; res(); };
      r.onerror = () => rej(r.error);
    });
  },
  _tx(mode, fn) {
    return new Promise((res, rej) => {
      const tx = this.db.transaction('files', mode); const st = tx.objectStore('files');
      const r = fn(st); tx.oncomplete = () => res(r && r.result); tx.onerror = () => rej(tx.error);
    });
  },
  put(id, blob) { return this._tx('readwrite', st => st.put(blob, id)); },
  get(id) { return this._tx('readonly', st => st.get(id)); },
  del(id) { return this._tx('readwrite', st => st.delete(id)); },
};
const urlCache = {};
async function fileURL(id) {
  if (urlCache[id]) return urlCache[id];
  const b = await DB.get(id); if (!b) return null;
  return (urlCache[id] = URL.createObjectURL(b));
}
async function addFile(file, kind) {
  const id = 'f_' + uid();
  await DB.put(id, file);
  CFG.files.push({ id, name: file.name || (kind + '-' + id), type: file.type, kind, size: file.size });
  save();
  return id;
}
async function removeFile(id) {
  await DB.del(id).catch(() => { });
  if (urlCache[id]) { URL.revokeObjectURL(urlCache[id]); delete urlCache[id]; }
  CFG.files = CFG.files.filter(f => f.id !== id);
}
function kindOf(file) {
  const t = file.type || '', n = (file.name || '').toLowerCase();
  if (t.includes('pdf') || n.endsWith('.pdf')) return 'pdf';
  if (t.startsWith('image')) return 'image';
  if (t.startsWith('video') || /\.(mp4|webm|mkv|mov|ogv)$/.test(n)) return 'video';
  if (t.startsWith('audio') || /\.(mp3|wav|ogg|m4a|aac|opus)$/.test(n)) return 'audio';
  return 'other';
}
function pickFile(accept, multiple = false) {
  return new Promise(res => {
    const fp = $('#filePicker'); fp.value = ''; fp.accept = accept || ''; fp.multiple = multiple;
    fp.onchange = () => res(multiple ? [...fp.files] : fp.files[0]);
    fp.click();
  });
}

/* ------------------------------ الإعدادات الافتراضية ------------------------------ */
function P(type, start, end, extra = {}) {
  return { id: uid(), type, label: '', start, end, bellStart: 'classic', bellEnd: 'dingdong', annStart: true, annEnd: true, startText: '', endText: '', startAudio: '', endAudio: '', ...extra };
}
function defaultConfig() {
  const base = {
    id: 'pr_main', name: 'التوقيت الشتوي', icon: '❄️', days: [0, 1, 2, 3, 4],
    periods: [
      P('assembly', '07:00', '07:15'),
      P('p1', '07:15', '07:55'), P('p2', '07:55', '08:35'), P('p3', '08:35', '09:15'),
      P('break', '09:15', '09:40', { bellStart: 'triple', bellEnd: 'classic' }),
      P('p4', '09:40', '10:20'), P('p5', '10:20', '11:00'), P('p6', '11:00', '11:40'),
      P('p7', '11:40', '12:20'), P('p8', '12:20', '13:00'),
    ],
  };
  return {
    version: 1,
    general: {
      schoolName: 'مدرسة ......', schoolSub: 'نتمنى لكم يوماً دراسياً موفقاً', logo: '',
      accent: '#14b8a6', clock24: false, arabicDigits: false, clockOffset: 0,
      workDays: [0, 1, 2, 3, 4], cycleWeeks: 1, cycleStart: '', holidays: [],
      bellEnabled: true, volume: 0.9, bannerSec: 8,
      theme: 'oman2040', animBg: true, gender: 'f',
      custom: { on: false, accent: '#159EC4', accent2: '#B64D9C', accent3: '#8585BB' },
    },
    profiles: [base],
    override: '', // معرّف جدول يُفرض على كل الأيام (مثلاً جدول رمضان/الاختبارات)
    voice: { omani: true, voiceSet: 'layan', tts: true, voiceName: '', rate: 0.9, repeat: 1, startTpl: 'بدأت الآن {name}', endTpl: 'انتهت الآن {name}', joinTpl: '، و', gapMs: 400 },
    reminders: [
      { id: uid(), enabled: true, name: 'ذكر منتصف الحصة', mode: 'mid', minutes: 5, time: '', applies: 'lessons', sound: 'none', text: 'أستغفر الله\nسبحان الله\nلا حول ولا قوة إلا بالله\nاللهم صلِّ وسلم على نبينا محمد', audio: '', vol: 0.25, show: 'soft' },
      { id: uid(), enabled: false, name: 'قبل نهاية الحصة', mode: 'beforeEnd', minutes: 5, time: '', applies: 'lessons', sound: 'soft', text: 'تبقى خمس دقائق على نهاية الحصة', audio: '' },
    ],
    slides: [
      { id: uid(), type: 'text', enabled: true, title: '', text: 'أهلاً وسهلاً بكم في مدرستنا\nالعلم نور', duration: 12 },
    ],
    autoSlides: { welcome: true, now: true, classday: true, schedule: false, duty: true, subs: true, absent: true, duration: 120, nowDuration: 120, perPage: 0,
      durs: { now: 30, classday: 30, welcome: 12, schedule: 15, duty: 20, subs: 20, absent: 20, phrase: 10 }, trDur: 1.6, lock: '' },
    ticker: { enabled: true, label: 'أخبار المدرسة', texts: ['أهلاً بكم في مدرستنا', 'النظافة من الإيمان', 'الالتزام بالوقت أساس النجاح'], speed: 90, includeDuty: true, includeSubs: true },
    display: { layout: 'hero', tickerPos: 'bottom', transition: 'fade', showStatus: true, showDuty: true, showSubs: true, nowStyle: 'grades', subsStyle: 'grouped', dutyAllWeeks: true, dutyStrip: true, dutyStripSec: 7 },
    absent: { source: 'manual', url: '', sheet: '', refreshMin: 10, lastFetch: '', lastError: '', rows: [], showScene: false, showNames: true, datedOnly: false },
    timetable: { source: 'manual', url: '', sheet: '', refreshMin: 60, lastFetch: '', lastError: '', rows: [], classes: [] },
    duty: { source: 'manual', url: '', sheet: '', refreshMin: 15, lastFetch: '', lastError: '', rows: [] },
    subs: { source: 'manual', url: '', sheet: '', refreshMin: 10, lastFetch: '', lastError: '', rows: [] },
    phrases: {
      enabled: true, showCard: true, cardSec: 15, slide: true, ticker: false,
      dhikr: ['سبحان الله', 'الحمد لله', 'لا إله إلا الله', 'الله أكبر', 'أستغفر الله', 'سبحان الله وبحمده، سبحان الله العظيم', 'لا حول ولا قوة إلا بالله', 'اللهم صلِّ وسلم على نبينا محمد', 'اذكروا الله يذكركم', 'حسبي الله ونعم الوكيل'],
      motivation: ['الانضباط طريقك إلى النجاح', 'احترامك للوقت احترام لنفسك', 'كن منظماً تكن متميزاً', 'النظام سر التفوق', 'الالتزام بالدوام عنوان الطالب المتميز', 'حافظ على نظافة مدرستك فهي بيتك الثاني', 'ابتسم وابدأ يومك بنشاط', 'كل يوم فرصة جديدة لتتعلم', 'احترم معلمك تنل علماً', 'الهدوء في الممرات دليل الرقي', 'ثق بنفسك، أنت قادر', 'العلم نور والجهل ظلام'],
      speak: false, speakWhen: 'breaks', speakEvery: 5, speakVolume: 0.3, speakRate: 1.2, speakDhikr: true, speakMotivation: true, audios: [],
    },
    quick: [
      { id: uid(), text: 'حان الآن وقت صلاة الظهر', sound: 'chime', audio: '' },
      { id: uid(), text: 'يرجى من جميع الطلاب التوجه إلى الساحة', sound: 'chime', audio: '' },
      { id: uid(), text: 'نرجو من الطلاب الالتزام بالهدوء والنظام', sound: 'soft', audio: '' },
      { id: uid(), text: 'يرجى من المعلمين المناوبين التوجه إلى مواقع المناوبة', sound: 'chime', audio: '' },
      { id: uid(), text: 'بدأ الآن السلام السلطاني', sound: 'none', audio: '' },
      { id: uid(), text: 'تنبيه: تدريب إخلاء، يرجى التوجه إلى نقطة التجمع بهدوء', sound: 'beep', audio: '' },
    ],
    files: [],
    shifts: { enabled: false, list: [{ id: 's1', name: 'الفترة الصباحية' }, { id: 's2', name: 'الفترة المسائية' }] },
  };
}
function mergeDefaults(def, cur) {
  if (Array.isArray(def)) return Array.isArray(cur) ? cur : def;
  if (def && typeof def === 'object') {
    const out = {};
    for (const k of Object.keys(def)) out[k] = mergeDefaults(def[k], cur ? cur[k] : undefined);
    if (cur) for (const k of Object.keys(cur)) if (!(k in out)) out[k] = cur[k];
    return out;
  }
  return cur === undefined ? def : cur;
}
const LS_KEY = 'schoolBell.config.v1';
let CFG;
/* ---------- مدرسة بفترتين: لكل فترة بياناتها (التوقيت، جدول المعلمين، المناوبة، الاحتياط، الغياب) ---------- */
const SHIFT_KEYS = ['profiles', 'override', 'timetable', 'duty', 'subs', 'absent'];
let CUR_SHIFT = 's1';
function shiftDefaults(id) {
  const d = defaultConfig(), o = {};
  SHIFT_KEYS.forEach(k => o[k] = d[k]);
  if (id === 's2') o.profiles = [{ id: 'pr_' + uid(), name: 'توقيت الفترة المسائية', icon: '🌇', days: [0, 1, 2, 3, 4], periods: genPeriods({ start: '13:15', assembly: 10, count: 6, len: 35, breakAfter: 3, breakLen: 15 }).map(p => p.type === 'assembly' ? { ...p, label: 'الطابور', startText: 'بدأ الآن الطابور', endText: 'انتهى الطابور، يرجى التوجه إلى الفصول' } : p) }];
  return o;
}
function loadConfig() {
  let cur = null;
  try { cur = JSON.parse(localStorage.getItem(LS_KEY)); } catch { }
  CFG = cur ? mergeDefaults(defaultConfig(), cur) : defaultConfig();
  // ترحيل: البيانات القديمة تصبح بيانات الفترة الأولى
  if (!CFG.shiftData) { CFG.shiftData = { s1: {} }; SHIFT_KEYS.forEach(k => CFG.shiftData.s1[k] = CFG[k]); }
  SHIFT_KEYS.forEach(k => delete CFG[k]);
  for (const id of Object.keys(CFG.shiftData)) CFG.shiftData[id] = mergeDefaults(shiftDefaults(id), CFG.shiftData[id]);
  if (CFG.shifts.enabled && !CFG.shiftData.s2) CFG.shiftData.s2 = shiftDefaults('s2');
  // CFG.profiles وغيرها تشير لبيانات الفترة الحالية (ولا تُحفظ مرتين)
  SHIFT_KEYS.forEach(k => Object.defineProperty(CFG, k, {
    get() { return (CFG.shiftData[CUR_SHIFT] || CFG.shiftData.s1)[k]; },
    set(v) { (CFG.shiftData[CUR_SHIFT] || CFG.shiftData.s1)[k] = v; },
    enumerable: false, configurable: true,
  }));
  if (!CFG.shiftData[CUR_SHIFT]) CUR_SHIFT = 's1';
  // تصحيح: "الأقسام المناوبة اليوم / قسم X" ← "مناوبة اليوم / X"
  for (const sd of Object.values(CFG.shiftData)) (sd.duty?.rows || []).forEach(r => {
    if (r.location === 'الأقسام المناوبة اليوم') { r.location = 'مناوبة اليوم'; r.teacher = String(r.teacher || '').replace(/^قسم\s+/, ''); }
  });
  // ذكر منتصف الحصة: أذكار متنوعة بصوت خافت وبطاقة هادئة — لمرة واحدة
  if (!CFG.display.v4) {
    CFG.display.v4 = 1;
    const r = CFG.reminders.find(x => x.mode === 'mid');
    const d = { sound: 'none', name: 'ذكر منتصف الحصة', text: 'أستغفر الله\nسبحان الله\nلا حول ولا قوة إلا بالله\nاللهم صلِّ وسلم على نبينا محمد', vol: 0.25, show: 'soft', sound: 'soft', applies: 'lessons', enabled: true };
    if (r) Object.assign(r, d); else CFG.reminders.unshift({ id: uid(), mode: 'mid', minutes: 5, time: '', audio: '', ...d });
  }
  // الذكر بدون نغمة قبله — لمرة واحدة
  if (!CFG.display.v5) { CFG.display.v5 = 1; CFG.reminders.filter(x => /ذكر/.test(x.name || '')).forEach(x => x.sound = 'none'); }
  // وضع مشاركة الشاشة مفعّل افتراضياً (المدرسة تعرض عبر ميراكاست) — لمرة واحدة
  if (!CFG.display.v6) { CFG.display.v6 = 1; CFG.display.cast = true; }
  if (!CFG.display.castTicker) CFG.display.castTicker = 'scroll';
  // كل مشهد يبقى دقيقتين قبل أن يتغير (بطلب المدرسة) — لمرة واحدة
  if (!CFG.display.v7) { CFG.display.v7 = 1; const A = CFG.autoSlides; A.duration = 120; A.nowDuration = 120; A.durs = A.durs || {}; for (const k of ['now', 'classday', 'welcome', 'schedule', 'duty', 'subs', 'absent', 'phrase']) A.durs[k] = 120; }
  // إلغاء مشهد "توقيت اليوم" (بطلب المدرسة) — لمرة واحدة
  if (!CFG.display.v3) { CFG.display.v3 = 1; CFG.autoSlides.schedule = false; }
  // ترقية لمرة واحدة: جدول كل الصفوف + حركة هادئة
  if (!CFG.display.v2) { CFG.display.v2 = 1; CFG.display.nowStyle = 'grades'; CFG.display.transition = 'fade'; if (!CFG.autoSlides.trDur || +CFG.autoSlides.trDur < 1.6) CFG.autoSlides.trDur = 1.6; }
}
const shiftIds = () => CFG.shifts.enabled ? CFG.shifts.list.map(s => s.id).filter(id => CFG.shiftData[id]) : ['s1'];
const shiftName = id => CFG.shifts.list.find(s => s.id === id)?.name || '';
function inShift(id, fn) { const prev = CUR_SHIFT; CUR_SHIFT = id; try { return fn(); } finally { CUR_SHIFT = prev; } }
/* الفترة المعروضة الآن: التي يقع الوقت الحالي ضمن دوامها (مع 45 دقيقة قبل بدايتها)، وإلا القادمة */
function displayShift(d = now()) {
  const ids = shiftIds(); if (ids.length < 2) return 's1';
  const sec = secOfDay(d);
  const spans = ids.map(id => inShift(id, () => { const ps = sortedPeriods(profileFor(d)); return ps.length ? { id, a: toSec(ps[0].start) - 45 * 60, b: toSec(ps[ps.length - 1].end) } : null; })).filter(Boolean);
  const inside = spans.find(x => sec >= x.a && sec < x.b); if (inside) return inside.id;
  const next = spans.filter(x => x.a > sec).sort((a, b) => a.a - b.a)[0]; if (next) return next.id;
  return spans.length ? spans[spans.length - 1].id : 's1';
}
let EDIT_SHIFT = 's1';
function syncShift() { CUR_SHIFT = $('#settings').hidden ? displayShift() : EDIT_SHIFT; }
let saveT;
function save(rerender = true) {
  clearTimeout(saveT);
  saveT = setTimeout(() => {
    try { localStorage.setItem(LS_KEY, JSON.stringify(CFG)); } catch (e) { toast('تعذر الحفظ: ' + e.message, true); }
  }, 150);
  if (rerender) onConfigChanged();
}

/* ------------------------------ الصوت ------------------------------ */
const Audio_ = {
  ctx: null, master: null, current: null, stopFlag: false,
  ensure() {
    if (!this.ctx) {
      this.ctx = new (window.AudioContext || window.webkitAudioContext)();
      this.master = this.ctx.createGain(); this.master.connect(this.ctx.destination);
    }
    this.master.gain.value = CFG.general.volume;
    if (this.ctx.state === 'suspended') this.ctx.resume();
    return this.ctx;
  },
  tone(freq, t, dur, vol = 0.5, partials = [[1, 1], [2, 0.45], [2.76, 0.3], [5.4, 0.12]]) {
    const c = this.ctx;
    for (const [m, g] of partials) {
      const o = c.createOscillator(), gn = c.createGain();
      o.type = 'sine'; o.frequency.value = freq * m;
      gn.gain.setValueAtTime(0.0001, t);
      gn.gain.exponentialRampToValueAtTime(vol * g, t + 0.01);
      gn.gain.exponentialRampToValueAtTime(0.0001, t + dur / (m > 2 ? 2 : 1));
      o.connect(gn); gn.connect(this.master); o.start(t); o.stop(t + dur + 0.05);
    }
  },
  /* أجراس مدمجة تُولَّد برمجياً (لا تحتاج ملفات) */
  bells: {
    classic: { n: 'جرس مدرسة كلاسيكي', play(A, t) {
      const c = A.ctx, dur = 3.2;
      const o1 = c.createOscillator(), o2 = c.createOscillator(), g = c.createGain(), lfo = c.createOscillator(), lg = c.createGain();
      o1.type = 'square'; o1.frequency.value = 880; o2.type = 'triangle'; o2.frequency.value = 2350;
      lfo.type = 'square'; lfo.frequency.value = 22; lg.gain.value = 0.12; g.gain.value = 0.12;
      lfo.connect(lg); lg.connect(g.gain);
      const env = c.createGain(); env.gain.setValueAtTime(0.0001, t); env.gain.exponentialRampToValueAtTime(0.6, t + 0.05);
      env.gain.setValueAtTime(0.6, t + dur - 0.3); env.gain.exponentialRampToValueAtTime(0.0001, t + dur);
      const f = c.createBiquadFilter(); f.type = 'lowpass'; f.frequency.value = 4000;
      o1.connect(g); o2.connect(g); g.connect(f); f.connect(env); env.connect(A.master);
      [o1, o2, lfo].forEach(o => { o.start(t); o.stop(t + dur + 0.1); });
      return dur;
    } },
    dingdong: { n: 'دينغ دونغ', play(A, t) { A.tone(659, t, 1.6, 0.6); A.tone(523, t + 0.75, 2.2, 0.6); return 3; } },
    triple: { n: 'رنين ثلاثي', play(A, t) { for (let i = 0; i < 3; i++) A.tone(880, t + i * 0.55, 1.2, 0.5); return 2.8; } },
    chime: { n: 'نغمة تنبيه (أربع نغمات)', play(A, t) { [523, 659, 784, 1047].forEach((f, i) => A.tone(f, t + i * 0.32, 1.6, 0.45)); return 2.6; } },
    soft: { n: 'نغمة هادئة', play(A, t) { [784, 659, 523].forEach((f, i) => A.tone(f, t + i * 0.4, 1.8, 0.35, [[1, 1], [3, 0.08]])); return 2.8; } },
    westminster: { n: 'ساعة (وستمنستر)', play(A, t) { [659, 523, 587, 392, 392, 587, 659, 523].forEach((f, i) => A.tone(f, t + i * 0.55 + (i > 3 ? 0.4 : 0), 1.6, 0.45)); return 5.4; } },
    beep: { n: 'صفارة قصيرة', play(A, t) {
      const c = A.ctx;
      for (let i = 0; i < 3; i++) {
        const o = c.createOscillator(), g = c.createGain(); o.type = 'square'; o.frequency.value = 1200;
        g.gain.setValueAtTime(0.0001, t + i * 0.35); g.gain.exponentialRampToValueAtTime(0.25, t + i * 0.35 + 0.01);
        g.gain.setValueAtTime(0.25, t + i * 0.35 + 0.2); g.gain.exponentialRampToValueAtTime(0.0001, t + i * 0.35 + 0.24);
        o.connect(g); g.connect(A.master); o.start(t + i * 0.35); o.stop(t + i * 0.35 + 0.3);
      }
      return 1.1;
    } },
  },
  async playSound(id) {
    if (!id || id === 'none' || this.stopFlag) return;
    this.ensure();
    if (this.bells[id]) {
      const d = this.bells[id].play(this, this.ctx.currentTime + 0.05);
      await sleep(d * 1000 + 150);
      return;
    }
    await this.playFile(id);
  },
  async playFile(id, vol) {
    const url = await fileURL(id);
    if (!url || this.stopFlag) return;
    await mediaPlay(url, vol, 5 * 60 * 1000);
  },
  /* عنصر صوت واحد يُفتح بلمسة "بدء التشغيل" — الآيباد/الآيفون لا يسمح بتشغيل صوت جديد بدون لمسة،
     لكنه يسمح بإعادة استخدام عنصر سبق تشغيله. وعلى الآيباد يُتحكم بمستواه عبر Web Audio لأن volume لا يعمل هناك */
  el: null, elGain: null,
  media() {
    if (!this.el) { this.el = new Audio(); this.el.preload = 'auto'; this.el.setAttribute('playsinline', ''); }
    if (!this.elGain && /^https?:/.test(location.protocol) && this.ctx) {
      try { const src = this.ctx.createMediaElementSource(this.el); this.elGain = this.ctx.createGain(); src.connect(this.elGain); this.elGain.connect(this.ctx.destination); } catch (e) { console.warn(e); }
    }
    return this.el;
  },
  unlock() {
    this.ensure();
    const a = this.media();
    a.src = 'data:audio/wav;base64,UklGRiQAAABXQVZFZm10IBAAAAABAAEAQB8AAIA+AAACABAAZGF0YQAAAAA=';
    a.play().catch(() => { });
    try { const u = new SpeechSynthesisUtterance(' '); u.volume = 0; speechSynthesis.speak(u); } catch { }
  },
  voices() { return (window.speechSynthesis?.getVoices() || []); },
  arabicVoices() { return this.voices().filter(v => /^ar/i.test(v.lang)); },
  speak(text, opt = {}) {
    return new Promise(res => {
      if (!text || !window.speechSynthesis || this.stopFlag) return res();
      const u = new SpeechSynthesisUtterance(text);
      const vs = this.voices();
      const v = vs.find(v => v.name === CFG.voice.voiceName) || this.arabicVoices()[0];
      if (v) { u.voice = v; u.lang = v.lang; } else u.lang = 'ar-SA';
      u.rate = opt.rate || +CFG.voice.rate || 0.9; u.volume = Math.min(1, opt.volume ?? CFG.general.volume + 0.1);
      const done = () => { clearTimeout(tm); res(); };
      const tm = setTimeout(done, 2500 + text.length * 180 / u.rate);
      u.onend = u.onerror = done;
      if (speechSynthesis.speaking) speechSynthesis.cancel();
      speechSynthesis.speak(u);
    });
  },
  stop() {
    this.stopFlag = true; Queue.items = [];
    try { speechSynthesis.cancel(); } catch { }
    if (this.current) { this.current.pause(); this.current = null; }
    if (this.ctx) { this.master.disconnect(); this.master = this.ctx.createGain(); this.master.connect(this.ctx.destination); this.master.gain.value = CFG.general.volume; }
    hideBanner();
    setTimeout(() => this.stopFlag = false, 300);
  },
};

/* التسجيلات العُمانية الجاهزة: نص ← ملف */
const clipKey = t => norm(t).replace(/[،,.!؟?\s]+/g, ' ').trim();
let CLIP_MAP = null;
function clipFor(text) {
  if (!CFG.voice.omani || !window.VOICE_CLIPS) return null;
  if (!CLIP_MAP) { CLIP_MAP = {}; for (const [k, v] of Object.entries(VOICE_CLIPS)) CLIP_MAP[clipKey(k)] = v; }
  const f = CLIP_MAP[clipKey(text)]; if (!f) return null;
  const set = (window.VOICE_SETS || {})[CFG.voice.voiceSet] || Object.values(window.VOICE_SETS || {})[0];
  return set ? set.dir + '/' + f + '.mp3' : null;
}
/* نص مدمج ("انتهت الآن الأولى، وبدأت الآن الثانية") ← قائمة تسجيلات، أو null إن نقص أحدها */
function clipsFor(text) {
  const one = clipFor(text); if (one) return [one];
  const segs = String(text).split(/،\s*و?|\s+و(?=بدأ|انته|حان)/).map(x => x.trim()).filter(Boolean);
  if (segs.length < 2) return null;
  const list = segs.map(clipFor);
  return list.every(Boolean) ? list : null;
}
function mediaPlay(url, vol, max = 60000) {
  return new Promise(res => {
    if (Audio_.stopFlag) return res();
    const a = Audio_.media(), v = Math.min(1, vol ?? CFG.general.volume);
    if (Audio_.elGain) { Audio_.ensure(); Audio_.elGain.gain.value = v; a.volume = 1; } else a.volume = v;
    let tm;
    const fin = () => { clearTimeout(tm); a.onended = a.onerror = null; Audio_.current = null; res(); };
    a.onended = a.onerror = fin; tm = setTimeout(fin, max);
    a.src = url; Audio_.current = a;
    a.play().catch(fin);
  });
}
const playURL = (url, vol) => mediaPlay(url, vol);
/* طابور الإعلانات: كل عنصر = {sound, parts:[{text, audio}], banner} */
const Queue = {
  items: [], busy: false,
  push(item) { this.items.push(item); this.run(); },
  async run() {
    if (this.busy) return; this.busy = true;
    while (this.items.length) {
      const it = this.items.shift();
      Audio_.stopFlag = false;
      if (it.banner) showBanner(it.banner, it.bannerIcon);
      if (it.soft) showSoftNote(it.soft);
      if (!it.quiet) Slides.duck(true);
      try {
        // صوت خافت: نخفض مستوى الجرس/النغمة لهذا التنبيه فقط
        if (it.vol != null && Audio_.ctx) { Audio_.ensure(); Audio_.master.gain.value = CFG.general.volume * it.vol; }
        await Audio_.playSound(it.sound);
        if (it.vol != null && Audio_.master) Audio_.master.gain.value = CFG.general.volume;
        for (let r = 0; r < (it.quiet ? 1 : Math.max(1, +CFG.voice.repeat || 1)); r++) {
          for (const p of it.parts || []) {
            if (Audio_.stopFlag) break;
            const cl = !p.audio && p.text ? clipsFor(p.text) : null;
            if (p.audio) await Audio_.playFile(p.audio, p.volume);
            else if (cl) { for (const u of cl) { if (Audio_.stopFlag) break; await playURL(u, p.volume); await sleep(250); } }
            else if (p.text && CFG.voice.tts) await Audio_.speak(p.text, { volume: p.volume, rate: p.rate });
            await sleep(+CFG.voice.gapMs || 300);
          }
        }
      } catch (e) { console.error(e); }
      if (!it.quiet) Slides.duck(false);
      if (it.banner) setTimeout(hideBanner, Math.max(1000, CFG.general.bannerSec * 1000 - 3000));
    }
    this.busy = false;
  },
};
/* بطاقة صغيرة هادئة أسفل الشاشة (للأذكار) */
let softT;
function showSoftNote(msg) {
  let el = $('#softNote');
  if (!el) { el = h('div', { id: 'softNote', class: 'soft-note' }); $('#display').append(el); }
  el.innerHTML = ''; el.append(h('span', { class: 'sn-ic' }, '🤲'), h('b', {}, msg));
  el.classList.remove('show'); void el.offsetWidth; el.classList.add('show');
  clearTimeout(softT); softT = setTimeout(() => el.classList.remove('show'), 12000);
}
let bannerT;
function showBanner(msg, icon = '🔔') {
  clearTimeout(bannerT);
  $('#bannerMsg').textContent = msg; $('#banner .icon').textContent = icon; $('#banner').hidden = false;
  bannerT = setTimeout(hideBanner, CFG.general.bannerSec * 1000 + 6000);
}
function hideBanner() { $('#banner').hidden = true; }

/* ------------------------------ محرك الجدول ------------------------------ */
const now = () => new Date(Date.now() + (+CFG.general.clockOffset || 0) * 1000);
const secOfDay = d => d.getHours() * 3600 + d.getMinutes() * 60 + d.getSeconds();
function holidayOf(d) {
  const k = ymd(d);
  return CFG.general.holidays.find(x => x.from && k >= x.from && k <= (x.to || x.from));
}
function isWorkDay(d) { return CFG.general.workDays.includes(d.getDay()) && !holidayOf(d); }
function profileFor(d) {
  if (CFG.override) { const p = CFG.profiles.find(p => p.id === CFG.override); if (p) return p; }
  const day = ymd(d), dow = d.getDay();
  // جدول له فترة تاريخية (رمضان، الاختبارات...) يتقدّم تلقائياً في أيامه
  const dated = CFG.profiles.find(p => p.dFrom && day >= p.dFrom && (!p.dTo || day <= p.dTo) && (!p.days.length || p.days.includes(dow)));
  if (dated) return dated;
  return CFG.profiles.find(p => !p.dFrom && p.days.includes(dow)) || null;
}
/* توليد فترات حسب المدد */
function genPeriods(gen) {
  let t = toSec(gen.start); const out = [];
  const add = (type, min, extra) => { out.push(P(type, fromSec(t), fromSec(t + min * 60), extra)); t += min * 60; };
  if (gen.assembly > 0) add('assembly', +gen.assembly);
  if (gen.zero) { add('p0', +gen.len); t += (gen.gap || 0) * 60; }
  for (let i = 1; i <= Math.min(8, gen.count); i++) {
    add('p' + i, +gen.len);
    if (i === +gen.breakAfter) add('break', +gen.breakLen, { bellStart: 'triple', bellEnd: 'classic' });
    else if (i === +gen.break2After) add('break2', +gen.break2Len, { bellStart: 'triple', bellEnd: 'classic' });
    else if (i < gen.count) t += (gen.gap || 0) * 60;
  }
  return out;
}
const PROFILE_TEMPLATES = [
  { k: 'winter', n: 'التوقيت الشتوي', icon: '❄️', gen: { start: '07:15', assembly: 15, count: 7, len: 40, breakAfter: 3, breakLen: 25 } },
  { k: 'summer', n: 'التوقيت الصيفي', icon: '☀️', gen: { start: '06:45', assembly: 15, count: 7, len: 40, breakAfter: 3, breakLen: 25 } },
  { k: 'zero', n: 'توقيت الحصة الصفرية', icon: '🌅', gen: { start: '06:30', assembly: 15, zero: true, count: 7, len: 40, breakAfter: 3, breakLen: 25 } },
  { k: 'ramadan', n: 'توقيت شهر رمضان', icon: '🌙', gen: { start: '08:30', assembly: 10, count: 7, len: 30, breakAfter: 4, breakLen: 15 } },
  { k: 'exams', n: 'توقيت الاختبارات', icon: '📝', gen: { start: '07:30', assembly: 0, count: 2, len: 90, breakAfter: 1, breakLen: 30 } },
  { k: 'thursday', n: 'توقيت يوم الخميس', icon: '📅', gen: { start: '07:15', assembly: 15, count: 6, len: 35, breakAfter: 3, breakLen: 20 } },
  { k: 'empty', n: 'جدول فارغ', icon: '🗓️', gen: null },
];
function sortedPeriods(pr) { return (pr?.periods || []).filter(p => p.start && p.end).slice().sort((a, b) => toSec(a.start) - toSec(b.start)); }
function cycleWeek(d) {
  const n = Math.max(1, +CFG.general.cycleWeeks || 1);
  if (n === 1) return 1;
  const st = CFG.general.cycleStart ? new Date(CFG.general.cycleStart + 'T00:00:00') : null;
  if (!st || isNaN(st)) return 1;
  const s0 = new Date(st); s0.setDate(s0.getDate() - s0.getDay()); // بداية أسبوع البداية (الأحد)
  const d0 = new Date(d.getFullYear(), d.getMonth(), d.getDate()); d0.setDate(d0.getDate() - d0.getDay());
  const w = Math.round((d0 - s0) / (7 * 864e5));
  return ((w % n) + n) % n + 1;
}
function startText(p) { return p.startText || TYPE[p.type]?.start || CFG.voice.startTpl.replace('{name}', periodName(p)); }
function endText(p) { return p.endText || TYPE[p.type]?.end || CFG.voice.endTpl.replace('{name}', periodName(p)); }
function reminderApplies(r, p) {
  if (r.applies === 'all') return true;
  if (r.applies === 'lessons') return isLesson(p);
  return p.type === r.applies;
}
/* كل الأحداث الصوتية لليوم */
function eventsFor(d) {
  if (!isWorkDay(d)) return [];
  // الجرس يرن لكل الفترات (الصباحية والمسائية)
  const all = [];
  shiftIds().forEach((id, i) => inShift(id, () => all.push(...shiftEvents(d, i === 0).map(e => ({ ...e, key: id + ':' + e.key })))));
  return all.sort((a, b) => a.t - b.t || a.order - b.order);
}
function shiftEvents(d, withFixed) {
  const pr = profileFor(d); if (!pr) return withFixed ? fixedReminders() : [];
  const ev = [];
  for (const p of sortedPeriods(pr)) {
    const s = toSec(p.start), e = toSec(p.end), name = periodName(p);
    if (p.annStart || (p.bellStart && p.bellStart !== 'none'))
      ev.push({ key: p.id + ':s', t: s, kind: 'start', sound: p.bellStart, text: p.annStart ? startText(p) : '', audio: p.annStart ? p.startAudio : '', banner: p.annStart ? startText(p) : name, order: 1 });
    if (p.annEnd || (p.bellEnd && p.bellEnd !== 'none'))
      ev.push({ key: p.id + ':e', t: e, kind: 'end', sound: p.bellEnd, text: p.annEnd ? endText(p) : '', audio: p.annEnd ? p.endAudio : '', banner: p.annEnd ? endText(p) : 'انتهت ' + name, order: 0 });
    for (const r of CFG.reminders) {
      if (!r.enabled || r.mode === 'fixed' || !reminderApplies(r, p)) continue;
      let t = null;
      if (r.mode === 'mid') t = Math.round((s + e) / 2 / 60) * 60;
      else if (r.mode === 'afterStart') t = s + r.minutes * 60;
      else if (r.mode === 'beforeEnd') t = e - r.minutes * 60;
      if (t != null && t > s && t < e) ev.push(remEvent(r, p.id + ':r:' + r.id, t));
    }
  }
  if (withFixed) ev.push(...fixedReminders());
  return ev;
}
function fixedReminders() {
  return CFG.reminders.filter(r => r.enabled && r.mode === 'fixed' && r.time).map(r => remEvent(r, 'fix:' + r.id, toSec(r.time)));
}
/* تنبيه: إن كان النص عدة أسطر يُختار سطر عشوائي (مثل الأذكار) */
const remLines = r => String(r.text || '').split('\n').map(x => x.trim()).filter(Boolean);
function remEvent(r, key, t) {
  const L = remLines(r), text = L.length ? L[Math.floor(Math.random() * L.length)] : '';
  return { key, t, kind: 'rem', sound: r.sound, text: r.speak === false ? '' : text, audio: r.speak === false ? '' : r.audio, banner: text || r.name, icon: '🤲', order: 2, vol: r.vol ?? 1, show: r.show || 'banner' };
}

const FIRED_KEY = 'schoolBell.fired';
let fired = {};
try { fired = JSON.parse(localStorage.getItem(FIRED_KEY)) || {}; } catch { }
function markFired(k) {
  fired[k] = 1;
  const today = ymd(now());
  for (const x of Object.keys(fired)) if (!x.startsWith(today)) delete fired[x];
  localStorage.setItem(FIRED_KEY, JSON.stringify(fired));
}
function checkEvents() {
  if (!CFG.general.bellEnabled || !started) return;
  const d = now(), sec = secOfDay(d), day = ymd(d);
  const due = eventsFor(d).filter(e => sec >= e.t && sec - e.t < 120 && !fired[day + '|' + e.key]);
  if (!due.length) return;
  due.forEach(e => markFired(day + '|' + e.key));
  // دمج الأحداث المتزامنة (نهاية حصة + بداية التالية) في إعلان واحد
  const groups = [];
  for (const e of due) {
    const g = groups.find(g => Math.abs(g.t - e.t) < 5 && g.kind !== 'rem' && e.kind !== 'rem');
    if (g) g.list.push(e); else groups.push({ t: e.t, kind: e.kind, list: [e] });
  }
  for (const g of groups) {
    const start = g.list.find(e => e.kind === 'start');
    const sound = (start && start.sound !== 'none' ? start.sound : null) || g.list.find(e => e.sound && e.sound !== 'none')?.sound || 'none';
    const withText = g.list.filter(e => e.text || e.audio);
    const parts = [];
    if (withText.length > 1 && withText.every(e => !e.audio)) parts.push({ text: withText.map(e => e.text).join(CFG.voice.joinTpl || '، ') });
    else withText.forEach(e => parts.push({ text: e.text, audio: e.audio }));
    const banner = (start || g.list[g.list.length - 1]).banner;
    const e0 = g.list[0];
    if (g.kind === 'rem') { Queue.push({ sound, parts: parts.map(p => ({ ...p, volume: Math.min(1, CFG.general.volume * (e0.vol ?? 1)) })), vol: e0.vol ?? 1, quiet: (e0.vol ?? 1) < 0.7, soft: e0.show === 'soft' ? banner : '', banner: e0.show === 'banner' ? banner : '', bannerIcon: '📢' }); continue; }
    Queue.push({ sound, parts, banner, bannerIcon: e0.icon || '🔔' });
  }
}
function manualBell() {
  const pr = profileFor(now());
  Queue.push({ sound: 'classic', parts: [], banner: 'جرس يدوي' });
}

/* ------------------------------ الأذكار والعبارات التحفيزية ------------------------------ */
const Phrases = {
  flip: false, last: '', cardT: null, lastSpoken: Date.now(),
  pick(kind) {
    const P = CFG.phrases;
    const d = P.dhikr.filter(Boolean), m = P.motivation.filter(Boolean);
    let k = kind || ((this.flip = !this.flip) ? 'd' : 'm');
    let arr = k === 'd' ? d : m;
    if (!arr.length) { k = k === 'd' ? 'm' : 'd'; arr = k === 'd' ? d : m; }
    if (!arr.length) return null;
    let t = arr[Math.floor(Math.random() * arr.length)];
    if (t === this.last && arr.length > 1) t = arr[(arr.indexOf(t) + 1) % arr.length];
    this.last = t;
    return { t, k };
  },
  renderCard() {
    clearTimeout(this.cardT);
    const P = CFG.phrases, card = $('#phraseCard');
    card.hidden = !(P.enabled && P.showCard);
    if (card.hidden) return;
    const ph = this.pick();
    if (ph) {
      $('#phraseIcon').textContent = ph.k === 'd' ? '🤲 ذِكر' : '⭐ كلمة تحفيزية';
      const el = $('#phraseText'); el.textContent = ph.t;
      el.style.animation = 'none'; void el.offsetWidth; el.style.animation = '';
    }
    // يتبدل أسرع في أوقات الفسحة وبين الحصص
    const st = currentState();
    const free = !st.cur || !isLesson(st.cur);
    card.classList.toggle('free', free && st.work);
    this.cardT = setTimeout(() => this.renderCard(), Math.max(5, +P.cardSec || 15) * 1000 * (free ? 0.7 : 1));
  },
  /* نطق عبارة بصوت خافت وسريع على فترات (لا يتداخل مع الجرس) */
  tick() {
    const P = CFG.phrases;
    if (!P.enabled || !P.speak || !started || !CFG.general.bellEnabled || Queue.busy) return;
    if (Date.now() - this.lastSpoken < Math.max(1, +P.speakEvery || 5) * 60000) return;
    const st = currentState(); if (!st.work || !st.periods.length) return;
    const first = toSec(st.periods[0].start) - 30 * 60, last = toSec(st.periods[st.periods.length - 1].end);
    if (st.sec < first || st.sec > last) return;
    if (P.speakWhen === 'breaks' && st.cur && isLesson(st.cur)) return;
    // لا تنطق قبل جرس قريب بدقيقة
    if (eventsFor(now()).some(e => e.t - st.sec >= 0 && e.t - st.sec < 60)) return;
    this.lastSpoken = Date.now();
    this.speakOne();
  },
  speakOne() {
    const P = CFG.phrases;
    const pool = [];
    if (P.speakDhikr) P.dhikr.filter(Boolean).forEach(t => pool.push({ text: t }));
    if (P.speakMotivation) P.motivation.filter(Boolean).forEach(t => pool.push({ text: t }));
    P.audios.filter(id => CFG.files.some(f => f.id === id)).forEach(id => pool.push({ audio: id }));
    if (!pool.length) return;
    const it = pool[Math.floor(Math.random() * pool.length)];
    Queue.push({ sound: 'none', quiet: true, parts: [{ ...it, volume: +P.speakVolume || 0.3, rate: +P.speakRate || 1.2 }] });
  },
};
function quickAnnounce(q) {
  Audio_.ensure();
  Queue.push({ sound: q.sound, parts: q.speak === false ? [] : [{ text: q.text, audio: q.audio }], banner: q.text, bannerIcon: '📢' });
  $('#quickMenu').hidden = true;
}
function toggleQuickMenu() {
  const m = $('#quickMenu');
  if (!m.hidden) { m.hidden = true; return; }
  m.innerHTML = '';
  m.append(h('div', { class: 'qm-title' }, '📢 إعلان فوري'));
  CFG.quick.forEach((q, i) => m.append(h('button', { onclick: () => quickAnnounce(q) }, h('b', {}, digits(i + 1)), q.text)));
  if (!CFG.quick.length) m.append(h('div', { class: 'qm-empty' }, 'أضف إعلانات من الإعدادات ← تنبيهات وإعلانات'));
  m.hidden = false;
}

/* ------------------------------ الحالة الحالية ------------------------------ */
function currentState(d = now()) {
  const sec = secOfDay(d);
  const st = { work: isWorkDay(d), holiday: holidayOf(d), cur: null, next: null, sec, periods: [] };
  if (!st.work) return st;
  const pr = profileFor(d); if (!pr) return st;
  st.periods = sortedPeriods(pr);
  st.cur = st.periods.find(p => sec >= toSec(p.start) && sec < toSec(p.end)) || null;
  st.next = st.periods.find(p => toSec(p.start) > sec) || null;
  return st;
}
function dutyFor(d, slotPeriod) {
  const dow = d.getDay(), wk = cycleWeek(d), day = ymd(d);
  const keys = slotPeriod ? periodKeys(slotPeriod) : null;
  return CFG.duty.rows.filter(r => {
    if (r.date) { if (r.date !== day) return false; }
    else if (r.day !== -1 && r.day !== dow) return false;
    if (!weekOK(r, d)) return false;
    const k = slotKey(r.slot);
    if (!keys) return k === '*';
    return k === '*' || keys.has(k);
  });
}
/* المناوبة: إظهار كل الأسابيع لنفس اليوم، أو الأسبوع الحالي فقط */
const weekOK = (r, d) => !r.week || CFG.display.dutyAllWeeks || r.week === cycleWeek(d);
const weekTag = r => r.week ? 'الأسبوع ' + (ORD[r.week - 1] || r.week) : '';
const byWeek = (a, b) => (a.week || 0) - (b.week || 0) || String(a.location || '').localeCompare(String(b.location || ''), 'ar');
function subsFor(d) {
  const day = ymd(d), dow = d.getDay();
  return CFG.subs.rows.filter(r => r.date ? r.date === day : (r.day === dow))
    .sort((a, b) => (parseInt(slotKey(a.period).slice(1)) || 99) - (parseInt(slotKey(b.period).slice(1)) || 99));
}

/* ------------------------------ العرض ------------------------------ */
const fmtGreg = () => new Intl.DateTimeFormat('ar-u-nu-' + (CFG.general.arabicDigits ? 'arab' : 'latn'), { day: 'numeric', month: 'long', year: 'numeric' });
const fmtHijri = () => new Intl.DateTimeFormat('ar-SA-u-ca-islamic-umalqura-nu-' + (CFG.general.arabicDigits ? 'arab' : 'latn'), { day: 'numeric', month: 'long', year: 'numeric' });
let lastDay = '';
function renderClock() {
  const d = now();
  let hh = d.getHours(); const ap = hh < 12 ? 'ص' : 'م';
  if (!CFG.general.clock24) hh = hh % 12 || 12;
  $('#clock').innerHTML = digits(`${pad(hh)}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`) + (CFG.general.clock24 ? '' : `<small>${ap}</small>`);
  const day = ymd(d);
  if (day !== lastDay) {
    const changed = !!lastDay;
    lastDay = day;
    renderDates(d);
    if (changed) { renderSide(true); Slides.rebuild(); refreshAllSources(); }
  }
}
function renderDates(d = now()) {
  {
    $('#dayName').textContent = DAYS[d.getDay()];
    try { $('#gregDate').textContent = fmtGreg().format(d) + ' م'; } catch { $('#gregDate').textContent = day; }
    try { $('#hijriDate').textContent = fmtHijri().format(d); } catch { $('#hijriDate').textContent = ''; }
    const n = +CFG.general.cycleWeeks || 1;
    $('#weekLabel').textContent = n > 1 ? 'الأسبوع ' + (ORD[cycleWeek(d) - 1] || cycleWeek(d)) : '';
    $('#weekLabel').hidden = n <= 1;
    $('#shiftLabel').textContent = CFG.shifts.enabled ? (CUR_SHIFT === 's1' ? '☀️ ' : '🌇 ') + shiftName(CUR_SHIFT) : '';
    $('#shiftLabel').hidden = !CFG.shifts.enabled;
  }
}
let lastSideSig = '';
function renderSide(force) {
  const d = now(), st = currentState(d), sec = st.sec;
  // بطاقة الآن
  let label = 'الآن', title = '—', range = '', count = '', prog = 0;
  if (!st.work) {
    title = st.holiday ? (st.holiday.name || 'إجازة') : 'عطلة نهاية الأسبوع';
    label = 'اليوم';
  } else if (st.cur) {
    const s = toSec(st.cur.start), e = toSec(st.cur.end);
    title = periodName(st.cur); range = digits(`من ${fmt12(st.cur.start)} إلى ${fmt12(st.cur.end)} — متبقٍ`);
    count = digits(fmtDur(e - sec)); prog = (sec - s) / (e - s) * 100;
  } else if (st.next) {
    const first = st.periods[0] === st.next;
    title = first ? 'قبل بداية الدوام' : 'بين الفترات';
    range = 'يبدأ ' + periodName(st.next) + ' بعد'; label = 'استعداد';
    count = digits(fmtDur(toSec(st.next.start) - sec));
  } else if (st.periods.length) {
    title = 'انتهى الدوام الدراسي'; range = 'نراكم غداً بإذن الله'; label = 'اليوم';
  } else { title = 'لا يوجد جدول لهذا اليوم'; label = 'اليوم'; }
  $('#nowLabel').textContent = label; $('#nowTitle').textContent = title; $('#nowRange').textContent = range;
  $('#nowCount').textContent = count; $('#nowProg').style.width = prog + '%';
  $('#nowProg').parentElement.parentElement.style.visibility = st.cur ? '' : 'hidden';
  $('#nextLine').innerHTML = st.next && st.cur ? `التالي: <b>${esc(periodName(st.next))}</b> ${digits(fmt12(st.next.start))}` : '';
  // عند تغيّر الفترة: اقفز لمشهد "الآن"
  const pid = st.work ? (st.cur ? st.cur.id : 'gap-' + (st.next ? st.next.id : 'end')) : 'off';
  if (renderSide.pid !== undefined && pid !== renderSide.pid) setTimeout(() => Slides.jump('auto-now'), 1500);
  renderSide.pid = pid;

  // المناوبة والاحتياط (تُحدَّث عند تغيّر الفترة فقط)
  const slotP = st.cur || st.next;
  const sig = ymd(d) + (slotP ? slotP.id : '-') + CFG._rev;
  if (!force && sig === lastSideSig) return;
  lastSideSig = sig;
  const dl = $('#dutyList'); dl.innerHTML = '';
  $('#dutySlot').textContent = slotP ? periodName(slotP) : 'اليوم';
  if (st.work) {
    const cur = slotP ? dutyFor(d, slotP).filter(r => slotKey(r.slot) !== '*') : [];
    const all = dutyFor(d, null);
    for (const r of cur) dl.append(h('li', { class: 'cur' }, h('span', { class: 'tag' }, r.slot || periodName(slotP)), h('span', {}, r.teacher), r.location ? h('span', { class: 'loc' }, '— ' + r.location) : null));
    for (const r of all) dl.append(h('li', {}, h('span', { class: 'tag' }, r.slot || 'اليوم'), h('span', {}, r.teacher), r.location ? h('span', { class: 'loc' }, '— ' + r.location) : null));
    if (!cur.length && !all.length) {
      // عرض مناوبة اليوم كاملة إن لم توجد مناوبة للفترة الحالية
      const today = CFG.duty.rows.filter(r => (r.date ? r.date === ymd(d) : (r.day === -1 || r.day === d.getDay())) && weekOK(r, d));
      if (today.length) $('#dutySlot').textContent = 'مناوبة اليوم';
      if (today.length) today.forEach(r => dl.append(h('li', {}, h('span', { class: 'tag' }, r.slot || '—'), h('span', {}, r.teacher), r.location ? h('span', { class: 'loc' }, '— ' + r.location) : null)));
      else dl.append(h('li', { class: 'muted' }, 'لا توجد مناوبة مسجلة لهذا الوقت'));
    }
  } else dl.append(h('li', { class: 'muted' }, '—'));

  const sl = $('#subsList'); sl.innerHTML = '';
  const subs = st.work ? subsFor(d) : [];
  $('#subsCount').textContent = subs.length ? digits(subs.length) : '';
  const curKeys = st.cur ? periodKeys(st.cur) : new Set();
  for (const r of subs) {
    const isCur = curKeys.has(slotKey(r.period));
    sl.append(h('li', { class: isCur ? 'cur' : '' },
      h('span', { class: 'tag' }, r.period || '—'),
      h('span', {}, [r.cls, r.sub].filter(Boolean).join(' — ')),
      r.absent ? h('span', { class: 'loc' }, '(بدلاً من ' + r.absent + ')') : null));
  }
  if (!subs.length) sl.append(h('li', { class: 'muted' }, 'لا توجد حصص احتياط اليوم'));
  if (CFG.display.layout === 'side') [dl, sl].forEach(autoScroll);
  updateTicker();
}
/* تمرير تلقائي للقوائم الطويلة */
const scrollers = new Map();
function autoScroll(el) {
  if (scrollers.has(el)) { clearInterval(scrollers.get(el)); scrollers.delete(el); }
  el.scrollTop = 0;
  let pause = 60;
  const t = setInterval(() => {
    if (el.scrollHeight <= el.clientHeight + 2) return;
    if (pause > 0) { pause--; return; }
    el.scrollTop += 1;
    if (el.scrollTop + el.clientHeight >= el.scrollHeight - 1) { pause = 80; setTimeout(() => { el.scrollTop = 0; }, 2500); }
  }, 50);
  scrollers.set(el, t);
}

/* الشريط الإخباري */
let tickerPending = null, tickerSig = '';
function tickerHTML() {
  const items = [...CFG.ticker.texts.filter(Boolean)];
  const d = now(), st = currentState(d);
  if (st.work && CFG.ticker.includeDuty) {
    const p = st.cur || st.next;
    const list = p ? dutyFor(d, p) : dutyFor(d, null);
    if (list.length) items.push('المناوبة' + (p ? ' (' + periodName(p) + ')' : '') + ': ' + list.map(r => r.teacher + (r.location ? ' - ' + r.location : '')).join('، '));
  }
  if (st.work && CFG.ticker.includeSubs) {
    const s = subsFor(d);
    if (s.length) items.push('حصص الاحتياط اليوم: ' + s.map(r => `${r.period || ''} ${r.cls || ''}: ${r.sub || ''}`.trim()).join(' ◆ '));
  }
  if (st.work && CFG.ticker.includeAbsent) { const a = absentFor(d); if (a.length) items.push('الغياب اليوم: ' + digits(a.length) + (CFG.absent.showNames ? ' — ' + a.map(r => r.name + (r.cls ? ' (' + r.cls + ')' : '')).join('، ') : '')); }
  if (CFG.phrases.enabled && CFG.phrases.ticker) items.push(...CFG.phrases.dhikr.filter(Boolean).slice(0, 4), ...CFG.phrases.motivation.filter(Boolean).slice(0, 3));
  return items.map(esc).join('<span class="sep">◆</span>');
}
function updateTicker(force) {
  $('#ticker').hidden = !CFG.ticker.enabled;
  $('#tickerLabel').textContent = CFG.ticker.label || 'أخبار';
  const html = tickerHTML();
  if (!force && html === tickerSig) return;
  tickerSig = html;
  const inner = $('#tickerInner');
  if (!inner.innerHTML || force) applyTicker(html); else tickerPending = html;
}
let castT = null;
function applyTicker(html) {
  const inner = $('#tickerInner');
  clearInterval(castT);
  // وضع مشاركة الشاشة: لا تمرير مستمر (يسبب تأخراً وتشويشاً في ميراكاست/AirPlay) — عبارة ثابتة تتبدل كل بضع ثوانٍ
  if (CFG.display.cast && CFG.display.castTicker === 'pages') {
    const items = (html || '').split('<span class="sep">◆</span>').filter(Boolean);
    let i = 0; inner.style.animation = 'none';
    const show = () => { inner.innerHTML = items[i % (items.length || 1)] || '&nbsp;'; i++; };
    show(); if (items.length > 1) castT = setInterval(show, Math.max(4, +CFG.display.castTickerSec || 8) * 1000);
    return;
  }
  inner.innerHTML = html || '&nbsp;';
  inner.style.animation = 'none';
  requestAnimationFrame(() => {
    const w = inner.scrollWidth + inner.parentElement.clientWidth;
    inner.style.animation = `tickerMove ${Math.max(8, w / (+CFG.ticker.speed || 90))}s linear infinite`;
    inner.style.willChange = 'transform';
  });
}
$('#tickerInner').addEventListener('animationiteration', () => { if (tickerPending != null) { applyTicker(tickerPending); tickerPending = null; } });

/* ------------------------------ شريط المناوبة (أسفل الشاشة) ------------------------------ */
const DutyStrip = {
  pages: [], i: 0, sig: '', t: null,
  build() {
    const el = $('#dutyStrip'), d = now(), st = currentState(d);
    const p = st.work ? (st.cur || st.next) : null;
    let rows = CFG.display.dutyStrip && p ? dutyFor(d, p) : [];
    // الاحتياط في الحصة الحالية (إن وُجد)
    const sp = st.cur || st.next, keys = sp && CFG.display.subsStrip !== false ? periodKeys(sp) : null;
    const subsNow = keys ? subsFor(d).filter(r => keys.has(slotKey(r.period))) : [];
    const sig = (p ? p.id : '-') + '|' + rows.length + '|' + subsNow.length + '|' + CFG._rev + '|' + cycleWeek(d);
    if (sig === this.sig) return;
    this.sig = sig;
    if (!rows.length && !subsNow.length) { el.hidden = true; this.pages = []; return; }
    // صفحات: صفحة لكل أسبوع (الحالي أولاً)، والسجلات العامة في صفحة "كل الأسابيع"
    const cw = +CFG.general.cycleWeeks > 1 ? cycleWeek(d) : 0;
    const byW = new Map(); rows.forEach(r => { const k = r.week || 0; (byW.get(k) || byW.set(k, []).get(k)).push(r); });
    const weeks = [...byW.keys()].sort((a, b) => (a === cw ? -1 : b === cw ? 1 : (a || 99) - (b || 99)));
    const pages = [];
    // صفحات الاحتياط أولاً: الصف ← البديلة (بدلاً من الغائبة)
    const sc = subsNow.map(r => [r.cls + (r.period ? ' — الحصة ' + r.period : ''), [shortName(r.sub || '—') + ' بدلاً من ' + shortName(r.absent || '—')]]);
    for (let k = 0; k < sc.length; k += 3) pages.push({ subs: true, chips: sc.slice(k, k + 3) });
    for (const w of weeks) {
      const locs = new Map(); byW.get(w).forEach(r => { const k = r.location || r.slot || ''; (locs.get(k) || locs.set(k, []).get(k)).push(r.teacher); });
      const chips = [...locs.entries()];
      for (let k = 0; k < chips.length; k += 3) pages.push({ w, cur: w === cw && cw > 0, chips: chips.slice(k, k + 3) });
    }
    this.pages = pages; this.i = -1;
    this.slot = p ? periodName(p) + (st.cur === p ? '' : ' (القادمة)') : '';
    this.subsSlot = sp ? periodName(sp) : '';
    el.hidden = false;
    this.next();
  },
  next() {
    clearTimeout(this.t);
    if (!this.pages.length) return;
    this.i = (this.i + 1) % this.pages.length;
    const pg = this.pages[this.i], box = $('#dsPage');
    $('#dutyStrip').classList.toggle('subs-mode', !!pg.subs);
    $('.ds-ic').textContent = pg.subs ? '🔁' : '🛡️';
    $('.ds-label b').textContent = pg.subs ? 'الاحتياط الآن' : 'المناوبة الآن';
    $('#dsSlot').textContent = pg.subs ? this.subsSlot : this.slot;
    const page = h('div', { class: 'ds-in' },
      pg.subs ? h('span', { class: 'ds-week cur' }, digits(this.pages.filter(x => x.subs).reduce((n, x) => n + x.chips.length, 0)) + ' احتياط')
        : h('span', { class: 'ds-week' + (pg.cur ? ' cur' : '') }, pg.w ? (pg.cur ? '● ' : '') + 'الأسبوع ' + (ORD[pg.w - 1] || pg.w) : 'كل الأسابيع'),
      pg.chips.map(([loc, names], k) => h('span', { class: 'ds-chip', style: `--i:${k}` }, loc ? h('em', {}, loc) : null, h('b', {}, names.join(' • ')))));
    box.innerHTML = ''; box.append(page);
    $('#dsDots').innerHTML = '';
    if (this.pages.length > 1) $('#dsDots').append(...this.pages.map((_, k) => h('i', { class: k === this.i ? 'on' : '' })));
    if (this.pages.length > 1) this.t = setTimeout(() => this.next(), Math.max(3, +CFG.display.dutyStripSec || 7) * 1000);
  },
};

/* ------------------------------ الشرائح ------------------------------ */
/* ---------- جدول المعلمين: من يدرّس أي صف في كل حصة ---------- */
const normCls = s => norm(s).replace(/[\s\-_\\.]+/g, '/').replace(/\/+/g, '/').replace(/^\/|\/$/g, '');
function ttFor(d, p) {
  const dow = d.getDay(), wk = cycleWeek(d), keys = periodKeys(p);
  return CFG.timetable.rows.filter(r => (r.day === -1 || r.day === dow) && (!r.week || r.week === wk) && keys.has(slotKey(r.period)));
}
function clsSorter() {
  const order = CFG.timetable.classes.filter(Boolean).map(normCls);
  const o = c => { const i = order.indexOf(normCls(c)); return i < 0 ? 999 : i; };
  return (a, b) => (o(a) - o(b)) || String(a).localeCompare(String(b), 'ar', { numeric: true });
}
const shortName = n => { const w = String(n || '').trim().split(/\s+/).filter(x => !/^(بنت|بن)$/.test(x)); return w.length > 2 ? w[0] + ' ' + w[w.length - 1] : String(n || '').trim(); };
const absentWord = () => ({ f: 'غائبة', m: 'غائب' }[CFG.general.gender] || 'غائب/ة');
/* كلمات حسب نوع المدرسة: معلمات / معلمون */
const G = k => ({
  teacher: { f: 'المعلمة', m: 'المعلم' }, absent: { f: 'المعلمة الغائبة', m: 'المعلم الغائب' }, sub: { f: 'المعلمة البديلة', m: 'المعلم البديل' },
  absentShort: { f: 'الغائبة', m: 'الغائب' }, subShort: { f: 'البديلة', m: 'البديل' },
}[k][CFG.general.gender] || { teacher: 'المعلم/ة', absent: 'الغائب/ة', sub: 'البديل/ة', absentShort: 'الغائب/ة', subShort: 'البديل/ة' }[k]);
function classCards(d, p) {
  const keys = periodKeys(p);
  const subs = subsFor(d).filter(r => keys.has(slotKey(r.period)));
  const cards = ttFor(d, p).map(r => ({ cls: r.cls, subject: r.subject, teacher: r.teacher, sub: subs.find(x => normCls(x.cls) === normCls(r.cls)) }));
  for (const x of subs) if (!cards.some(c => c.sub === x)) cards.push({ cls: x.cls, subject: x.note, teacher: x.absent, sub: x });
  const cmp = clsSorter();
  return cards.sort((a, b) => cmp(a.cls, b.cls));
}

const pdfCache = {};
async function pdfPages(slide) {
  const key = slide.fileId || slide.url;
  if (pdfCache[key]) return pdfCache[key];
  let data;
  if (slide.fileId) { const b = await DB.get(slide.fileId); if (!b) return null; data = await b.arrayBuffer(); }
  else {
    try { const r = await fetch(slide.url); if (!r.ok) throw 0; data = await r.arrayBuffer(); } catch { return null; }
  }
  const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(data), isEvalSupported: false }).promise;
  const out = [];
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const vp0 = page.getViewport({ scale: 1 });
    const vp = page.getViewport({ scale: Math.min(3, 1800 / vp0.width) });
    const cv = document.createElement('canvas'); cv.width = vp.width; cv.height = vp.height;
    const cx = cv.getContext('2d'); cx.fillStyle = '#fff'; cx.fillRect(0, 0, cv.width, cv.height);
    await page.render({ canvasContext: cx, viewport: vp }).promise;
    const blob = await new Promise(r => cv.toBlob(r, 'image/jpeg', 0.9));
    out.push(URL.createObjectURL(blob));
  }
  return (pdfCache[key] = out);
}
function embedURL(url) {
  url = normURL(url);
  let m;
  if ((m = url.match(/(?:youtube\.com\/(?:watch\?v=|shorts\/|embed\/)|youtu\.be\/)([\w-]{6,})/))) return { kind: 'iframe', src: `https://www.youtube.com/embed/${m[1]}?autoplay=1&mute=1&loop=1&playlist=${m[1]}&controls=0` };
  if ((m = url.match(/drive\.google\.com\/(?:file\/d\/|open\?id=)([\w-]+)/))) return { kind: 'iframe', src: `https://drive.google.com/file/d/${m[1]}/preview` };
  if ((m = url.match(/canva\.com\/design\/([\w-]+)(?:\/([\w-]+))?/)) && !/embed/.test(url)) return { kind: 'iframe', src: `https://www.canva.com/design/${m[1]}/${m[2] ? m[2] + '/' : ''}view?embed` };
  if (/\.(mp4|webm|ogv)(\?|$)/i.test(url)) return { kind: 'video', src: url };
  if (/\.(png|jpe?g|gif|webp|svg)(\?|$)/i.test(url)) return { kind: 'image', src: url };
  return { kind: 'iframe', src: url };
}

/* جدولة ظهور الوسائط: وقت محدد / أيام / تواريخ / أثناء فترة معينة */
function slideActive(s, d = now()) {
  if (!s.enabled) return false;
  const day = ymd(d), sec = secOfDay(d);
  if (s.dFrom && day < s.dFrom) return false;
  if (s.dTo && day > s.dTo) return false;
  if (s.days && s.days.length && !s.days.includes(d.getDay())) return false;
  if (s.when === 'window') {
    const a = toSec(s.from), b = toSec(s.to);
    if (a != null && sec < a) return false;
    if (b != null && sec >= b) return false;
  } else if (s.when === 'period') {
    const st = currentState(d), p = st.cur;
    if (!p) return false;
    if (s.ptype === 'lessons') return isLesson(p);
    if (s.ptype === 'breaks') return !isLesson(p);
    return p.type === s.ptype;
  }
  return true;
}
/* وضع "لا يوجد دوام": إجازة/عطلة، أو بعد نهاية الدوام، أو قبل بدايته بوقت طويل */
function offState(d = now()) {
  const st = currentState(d);
  if (!st.work) return { off: true, why: st.holiday ? 'holiday' : 'weekend', st };
  if (!st.periods.length) return { off: true, why: 'nosched', st };
  const a = toSec(st.periods[0].start), b = toSec(st.periods[st.periods.length - 1].end);
  if (st.sec >= b + 15 * 60) return { off: true, why: 'ended', st };
  if (st.sec < a - 90 * 60) return { off: true, why: 'early', st };
  return { off: false, st };
}
/* أول يوم دوام قادم وموعد بدايته */
function nextSchoolStart(d = now()) {
  for (let k = 0; k <= 45; k++) {
    const x = new Date(d.getFullYear(), d.getMonth(), d.getDate() + k);
    if (!isWorkDay(x)) continue;
    const ps = sortedPeriods(profileFor(x)); if (!ps.length) continue;
    const t = new Date(x.getTime() + toSec(ps[0].start) * 1000);
    if (t > d) return t;
  }
  return null;
}
const OFF_TYPES = new Set(['auto-off', 'auto-phrase', 'auto-welcome', 'text']);
/* مدة كل مشهد من الإعدادات */
const durOf = k => Math.max(3, +(CFG.autoSlides.durs || {})[k] || (k === 'now' ? CFG.autoSlides.nowDuration : CFG.autoSlides.duration) || 15);
const SCENE_FN = { classday: 'classDayScene', now: 'nowScene', subs: 'subsScene', duty: 'dutyScene', absent: 'absentScene', schedule: 'scheduleScene', welcome: 'welcomeScene' };
const SCENE_NAMES = { classday: '🗓️ جدول اليوم', now: '📚 الحصة الآن', subs: '🔁 الاحتياط', duty: '🛡️ المناوبة', absent: '📋 الغياب', schedule: '⏰ التوقيت', welcome: '🏫 الترحيب' };
const TRANS = ['tr-fade', 'tr-up', 'tr-zoom', 'tr-side', 'tr-flip'];
const Slides = {
  list: [], i: -1, page: 0, more: false, timer: null, building: false, video: null, skips: 0,
  async rebuild() {
    if (this.building) { this.again = true; return; }
    this.building = true;
    const media = [];
    const takeovers = [];
    for (const s of CFG.slides) {
      if (!s.enabled) continue;
      const dest = s.mode === 'takeover' ? takeovers : media;
      if (s.type === 'pdf') {
        let pages = null;
        try { pages = await pdfPages(s); } catch (e) { console.warn(e); }
        if (pages && pages.length) pages.forEach((src, pi) => dest.push({ ...s, _key: s.id + ':' + pi, fromPdf: true, type: 'image', src, title: (s.title || '') + (pages.length > 1 ? ` (${pi + 1}/${pages.length})` : '') }));
        else if (s.url) dest.push({ ...s, _key: s.id, type: 'url' });
      } else dest.push({ ...s, _key: s.id });
    }
    this.takeovers = takeovers;
    const A = CFG.autoSlides, auto = [];
    for (const k of ['welcome', 'classday', 'schedule', 'duty', 'subs', 'absent']) if (A[k]) auto.push({ type: 'auto-' + k, duration: durOf(k) });
    for (const [w, t] of [['duty', 'المناوبة'], ['subs', 'الاحتياط'], ['absent', 'الغياب'], ['timetable', 'جدول الحصص']])
      if (CFG[w].showScene && CFG[w].url) media.push({ type: 'url', url: CFG[w].url, title: t, duration: 30 });
    // تبادل: مشهد معلومات ثم وسائط
    let out = [];
    for (let k = 0; k < Math.max(auto.length, media.length); k++) { if (auto[k]) out.push(auto[k]); if (media[k]) out.push(media[k]); }
    // مشهد "الآن" هو الأهم: يتكرر كل 3 مشاهد
    if (A.now) {
      const res = [];
      out.forEach((s, k) => { if (k % 3 === 0) res.push({ type: 'auto-now', duration: durOf('now') }); res.push(s); });
      if (!out.length) res.push({ type: 'auto-now', duration: durOf('now') });
      out = res;
    }
    if (CFG.phrases.enabled && CFG.phrases.slide) {
      const res = [];
      out.forEach((s, k) => { res.push(s); if (k % 3 === 2) res.push({ type: 'auto-phrase', duration: durOf('phrase') }); });
      if (out.length < 3) res.push({ type: 'auto-phrase', duration: durOf('phrase') });
      out = res;
    }
    this.list = out;
    if (this.takeover) this.takeover = takeovers.find(t => t._key === this.takeover._key) || null;
    this.building = false;
    if (this.again) { this.again = false; return this.rebuild(); }
    this.i = -1; this.page = 0; this.more = false;
    this.next();
  },
  jump(type) {
    const k = this.list.findIndex(s => s.type === type);
    if (this.takeover) return;
    if (this.pinned()) { this.more = false; this.next(true); return; }
    if (k < 0) return;
    this.i = k - 1; this.page = 0; this.more = false; this.next();
  },
  skip() {
    if (++this.skips > this.list.length + 1) { this.skips = 0; this.showEmpty(); this.timer = setTimeout(() => this.next(true), 30000); return; }
    setTimeout(() => this.next(true), 10);
  },
  showEmpty() {
    this.mount(h('div', { class: 'slide welcome' }, ...this.welcomeKids()), 0);
  },
  /* ---------- وضع الاستحواذ: فيديو/صورة تملأ الشاشة في وقتها مع مقاطعات للمعلومات ---------- */
  activeTakeovers() { const d = now(); return (this.takeovers || []).filter(t => slideActive(t, d)); },
  watch() {
    if (this.building) return;
    const off = CFG.display.offMode !== false && offState().off;
    if (this.isOff !== undefined && off !== this.isOff && !this.takeover) { this.next(true); return; }
    const tk = this.activeTakeovers();
    if (tk.length && !this.takeover) this.next(true);
    else if (this.takeover && !this.interrupting && !tk.some(t => t._key === this.takeover._key)) { this.takeover = null; this.next(true); }
  },
  runTakeover(tk) {
    let s = this.takeover && tk.find(t => t._key === this.takeover._key);
    if (!s || (!s._resume && tk.length > 1 && this.takeoverDone)) s = tk[(tk.indexOf(s) + 1) % tk.length] || tk[0];
    this.takeoverDone = false;
    const resume = s._resume; s._resume = false;
    this.takeover = s; this.interrupting = false;
    this.render(s, resume ? s._pos : 0).then(dur => {
      const every = +s.intEvery || 0;
      if (every > 0 && (s.intScenes || []).length) this.timer = setTimeout(() => this.interrupt(s), every * 1000);
      else if (dur > 0) this.timer = setTimeout(() => { this.takeoverDone = true; this.next(); }, dur * 1000);
    });
  },
  interrupt(s) {
    clearTimeout(this.timer);
    if (this.video) s._pos = this.video.currentTime;
    s._resume = true; this.interrupting = true;
    const queue = [...(s.intScenes || [])];
    const step = () => {
      while (queue.length) {
        const k = queue.shift(), fn = SCENE_FN[k];
        this.page = 0; this.pages = 1;
        const el = fn === 'welcomeScene' ? h('div', { class: 'slide welcome' }, ...this.welcomeKids()) : this[fn]();
        if (el) { const dd = +s.intDur || 12; this.mount(el, dd); this.video = null; this.timer = setTimeout(step, dd * 1000); return; }
      }
      this.interrupting = false;
      if (this.activeTakeovers().some(t => t._key === s._key)) this.runTakeover(this.activeTakeovers());
      else { this.takeover = null; this.next(true); }
    };
    step();
  },
  /* محتوى محدد لفترة معينة (من جدول التوقيت) */
  pinned() { const st = currentState(); const c = st.cur?.content; return c && c !== 'rotate' ? c : null; },
  runPin(kind, advance) {
    if (kind === 'media') {
      const media = this.list.filter(s => !s.type.startsWith('auto-') && (!s._key || slideActive(s)));
      if (!media.length) return false;
      this.mi = ((this.mi ?? -1) + 1) % media.length;
      this.render(media[this.mi]).then(dur => { if (dur > 0) this.timer = setTimeout(() => this.next(), dur * 1000); });
      return true;
    }
    const fn = SCENE_FN[kind];
    if (fn && fn !== 'welcomeScene') { const pg = this.page, mo = this.more; this.page = 0; const ok = this[fn](); this.page = pg; this.more = mo; if (!ok) return false; }
    if (this.more && !advance && this.pinKind === kind) this.page++; else this.page = 0;
    this.more = false; this.pinKind = kind;
    const map = { classday: 'auto-classday', now: 'auto-now', duty: 'auto-duty', subs: 'auto-subs', absent: 'auto-absent', schedule: 'auto-schedule', welcome: 'auto-welcome', phrase: 'auto-phrase' };
    const s = { type: map[kind], duration: durOf(kind) };
    this.render(s).then(dur => {
      if (dur === null) return;
      this.timer = setTimeout(() => this.next(), Math.max(5, dur || 15) * 1000);
    });
    return true;
  },
  next(advance) {
    clearTimeout(this.timer);
    // 📌 الشاشة مثبتة على مشهد: نعيد رسمه كل فترة ليبقى محدّثاً (والجرس يعمل كالمعتاد)
    if (this.frozen && this.frozenSlide && !advance) { this.render(this.frozenSlide).then(() => { this.timer = setTimeout(() => this.next(), 30000); }); return; }
    const lock = CFG.autoSlides.lock;
    if (lock && !(CFG.display.offMode !== false && offState().off)) {
      if (this.more && !advance && this.pinKind === lock) this.page++; else if (this.pinKind !== lock) this.page = 0;
      if (lock !== 'media' && this.runPin(lock, advance)) return;
    }
    const tk = this.activeTakeovers();
    if (tk.length) { if (!this.interrupting) this.runTakeover(tk); return; }
    if (this.takeover) { this.takeover = null; advance = true; }
    // يوم بلا دوام: مشهد "لا يوجد دوام" مع الأذكار والترحيب والنصوص (وفيديو/صور مضبوطة لتظهر دائماً)
    const off = CFG.display.offMode !== false && offState().off;
    this.isOff = off;
    if (off) {
      const offList = [{ type: 'auto-off', duration: 15 }, ...this.list.filter(s => OFF_TYPES.has(s.type) || ((s.type === 'image' || s.type === 'video') && !s.fromPdf && s._key && s.when && s.when !== 'always' && slideActive(s)))];
      if (CFG.phrases.enabled) offList.push({ type: 'auto-phrase', duration: 10 });
      this.oi = ((this.oi ?? -1) + 1) % offList.length;
      this.render(offList[this.oi]).then(dur => { if (dur > 0) this.timer = setTimeout(() => this.next(), dur * 1000); });
      return;
    }
    const pin = this.pinned();
    if (pin && this.runPin(pin, advance)) return;
    this.pinKind = null;
    if (!this.list.length) { this.showEmpty(); return; }
    if (this.more && !advance) this.page++;
    else { this.page = 0; this.i = (this.i + 1) % this.list.length; }
    this.more = false;
    const s = this.list[this.i];
    if (s._key && !slideActive(s)) { this.skip(); return; }
    this.render(s).then(dur => {
      if (dur !== null) this.skips = 0;
      if (dur > 0) this.timer = setTimeout(() => this.next(), dur * 1000);
    });
  },
  toggleFreeze() {
    this.frozen = !this.frozen;
    const cur = [...$('#stage').children].filter(x => x.classList.contains('slide')).pop();
    this.frozenSlide = this.frozen ? this.lastSlide : null;
    clearTimeout(this.timer);
    $('#cPin').classList.toggle('off', this.frozen);
    $('#pinFlag').hidden = !this.frozen;
    toast(this.frozen ? '📌 تم تثبيت الشاشة على هذا المشهد — اضغط 📌 مرة أخرى للإلغاء' : 'تم إلغاء التثبيت');
    if (this.frozen) this.timer = setTimeout(() => this.next(), 30000); else this.next(true);
  },
  /* تركيب المشهد مع حركة الدخول والخروج */
  mount(el, dur) {
    const stage = $('#stage');
    const t = CFG.display.transition;
    if (!CFG.display.cast) el.classList.add(t === 'random' || !t ? TRANS[Math.floor(Math.random() * TRANS.length)] : 'tr-' + t);
    const td = Math.max(0.2, +CFG.autoSlides.trDur || 1.6);
    el.style.animationDuration = td + 's';
    document.documentElement.style.setProperty('--tr', td + 's');
    for (const old of [...stage.children]) {
      if (old.classList.contains('slide-dots') || old.classList.contains('scene-timer')) { old.remove(); continue; }
      old.classList.add('tr-out'); old.style.animationDuration = Math.max(0.2, (+CFG.autoSlides.trDur || 1.6) * 0.7) + 's';
      old.querySelectorAll('video').forEach(v => v.pause());
      if (CFG.display.cast) old.remove(); else setTimeout(() => old.remove(), Math.max(200, (+CFG.autoSlides.trDur || 1.6) * 700));
    }
    stage.append(el);
    // جدول الصفوف: مساحة كاملة (إخفاء شريط الحصة وشريط المناوبة وتصغير الأعلى)
    $('#display').classList.toggle('focus-table', !!el.querySelector('.gtables') && CFG.display.focusTable !== false);
    // تصغير الخط حتى يظهر كل الجدول داخل الشاشة (حسب حجم شاشة الجهاز الفعلي)
    el.querySelectorAll('.fit').forEach(f => {
      const go = () => { let n = 0, fs = parseFloat(f.style.fontSize) || 24; while (f.scrollHeight > f.clientHeight + 1 && fs > 10 && n++ < 40) { fs -= 1; f.style.fontSize = fs + 'px'; } };
      requestAnimationFrame(go); setTimeout(go, 400);
    });
    if (this.isOff) { /* لا عدّاد في وضع لا يوجد دوام */ }
    else if (this.list.length > 14) stage.append(h('div', { class: 'slide-dots count' }, digits(`${this.i + 1} / ${this.list.length}`)));
    else if (this.list.length > 1 && !this.isOff) stage.append(h('div', { class: 'slide-dots' }, this.list.map((_, k) => h('i', { class: k === this.i ? 'on' : '' }))));
    if (dur > 0 && !this.frozen) stage.append(h('div', { class: 'scene-timer', style: `animation-duration:${dur}s` }));
  },
  async render(s, startAt = 0) {
    this.lastSlide = s;
    let el, dur = +s.duration || 12;
    this.video = null;
    const title = s.title ? h('div', { class: 'slide-title' }, s.title) : null;
    if (s.type === 'image') {
      const src = s.src || await fileURL(s.fileId);
      if (!src) { this.skip(); return null; }
      const img = h('img', { src, alt: '' });
      el = h('div', { class: 'slide media' + (s.fit === 'scroll' ? ' scroll' : '') }, img, title);
      if (s.fit === 'scroll') img.onload = () => {
        const over = img.clientHeight - $('#stage').clientHeight;
        if (over > 0) img.animate([{ transform: 'translateY(0)' }, { transform: 'translateY(0)', offset: 0.12 }, { transform: `translateY(${-over}px)`, offset: 0.88 }, { transform: `translateY(${-over}px)` }], { duration: dur * 1000, fill: 'forwards' });
      };
    } else if (s.type === 'video') {
      const src = s.fileId ? await fileURL(s.fileId) : s.url;
      if (!src) { this.skip(); return null; }
      const v = h('video', { src, autoplay: true, playsinline: true });
      v.muted = s.muted !== false; v.volume = CFG.general.volume;
      const tk = s.mode === 'takeover';
      v.loop = !!(s.duration > 0) || (tk && (this.activeTakeovers().length <= 1));
      if (startAt > 0) v.addEventListener('loadedmetadata', () => { try { v.currentTime = Math.min(startAt, (v.duration || startAt + 1) - 0.5); } catch { } }, { once: true });
      if (!(s.duration > 0)) {
        v.onended = () => { s._pos = 0; if (tk) { this.takeoverDone = true; } this.next(true); };
        v.onerror = () => setTimeout(() => this.next(true), 3000); dur = 0;
      }
      v.play().catch(() => { v.muted = true; v.play().catch(() => { }); });
      this.video = v;
      el = h('div', { class: 'slide media' }, v, title);
    } else if (s.type === 'url') {
      const e = embedURL(s.url || '');
      if (e.kind === 'video') return this.render({ ...s, type: 'video', url: e.src });
      if (e.kind === 'image') return this.render({ ...s, type: 'image', src: e.src });
      el = h('div', { class: 'slide' }, h('iframe', { src: e.src, allow: 'autoplay' }), title);
    } else if (s.type === 'text') {
      el = h('div', { class: 'slide text-slide plain' }, h('div', { class: 'ps-icon' }, '📣'), h('div', { class: 'ps-text' }, s.text || ''));
    } else if (s.type === 'auto-phrase') {
      const ph = Phrases.pick();
      if (ph) el = h('div', { class: 'slide text-slide' }, h('div', { class: 'ps-icon' }, ph.k === 'd' ? '🤲' : '⭐'),
        h('div', { class: 'ps-kind' }, ph.k === 'd' ? 'ذِكر' : 'كلمة اليوم'), h('div', { class: 'ps-text' }, ph.t));
    } else if (s.type === 'auto-welcome') {
      el = h('div', { class: 'slide welcome' }, ...this.welcomeKids());
    } else if (s.type === 'auto-off') { el = this.offScene(); this.offTick(el); }
    else if (s.type === 'auto-now') el = this.nowScene();
    else if (s.type === 'auto-schedule') el = this.scheduleScene();
    else if (s.type === 'auto-duty') el = this.dutyScene();
    else if (s.type === 'auto-subs') el = this.subsScene();
    else if (s.type === 'auto-absent') el = this.absentScene();
    else if (s.type === 'auto-classday') el = this.classDayScene();
    if (!el) { this.skip(); return null; }
    // المشاهد ذات الصفحات: وزّع المدة على الصفحات
    if (this.pages > 1) dur = Math.max(8, Math.round(dur / Math.min(this.pages, 2)));
    this.pages = 1;
    this.mount(el, dur);
    return dur;
  },

  /* ---------- أدوات المشاهد ---------- */
  welcomeKids() {
    const g = CFG.general, d = now(), st = currentState(d);
    const logo = g.logo && urlCache[g.logo] ? h('img', { class: 'wlogo', src: urlCache[g.logo] }) : h('div', { class: 'wlogo-ph' }, '🏫');
    let line = '';
    if (!st.work) line = st.holiday ? '🌴 ' + (st.holiday.name || 'إجازة') : '🌙 عطلة نهاية الأسبوع';
    else if (st.cur) line = '📚 ' + periodName(st.cur) + ' الآن';
    else if (st.next) line = '⏳ يبدأ ' + periodName(st.next) + ' الساعة ' + digits(fmt12(st.next.start));
    else if (st.periods.length) line = '👋 انتهى الدوام، نراكم غداً بإذن الله';
    let date = DAYS[d.getDay()];
    try { date += ' — ' + fmtHijri().format(d); } catch { }
    return [logo, h('h1', {}, g.schoolName), g.schoolSub ? h('p', {}, g.schoolSub) : null, h('p', {}, date), line ? h('div', { class: 'wnow' }, line) : null];
  },
  head(icon, title, chips = []) {
    return h('div', { class: 'scene-head' }, h('span', { class: 'ic' }, icon), h('h2', {}, title),
      chips.filter(Boolean).map(c => h('span', { class: 'chip' + (c.hot ? ' hot' : '') }, c.t || c)),
      this.pages > 1 ? h('span', { class: 'pg' }, digits(`${this.page + 1} / ${this.pages}`)) : null);
  },
  /* توزيع البطاقات على شبكة تملأ المساحة + تقسيمها إلى صفحات */
  grid(items, mk, opt = {}) {
    const stage = $('#stage');
    const W = stage.clientWidth || 1200, H = (stage.clientHeight || 600) - Math.max(90, stage.clientHeight * 0.16);
    const maxCols = Math.max(1, Math.min(6, Math.floor(W / (opt.minW || 270))));
    const rowH = Math.max(110, H * (opt.rowFrac || 0.24));
    const maxRows = Math.max(1, Math.floor(H / rowH));
    const per = +CFG.autoSlides.perPage || maxCols * maxRows;
    this.pages = Math.max(1, Math.ceil(items.length / per));
    if (this.page >= this.pages) this.page = 0;
    this.more = this.page < this.pages - 1;
    const slice = items.slice(this.page * per, this.page * per + per);
    const n = Math.max(1, slice.length);
    let cols = Math.min(maxCols, Math.max(1, Math.ceil(Math.sqrt(n * (W / H) / 1.7))));
    const rows = Math.ceil(n / cols);
    cols = Math.ceil(n / rows);
    return h('div', { class: 'cards stagger', style: `grid-template-columns:repeat(${cols},1fr);grid-template-rows:repeat(${rows},minmax(0,${n <= 4 ? '1fr' : '1fr'}))` },
      slice.map((it, k) => { const c = mk(it); c.style.setProperty('--i', k); return c; }));
  },
  classCard(c, curP) {
    if (c.sub) return h('div', { class: 'ccard sub' }, h('span', { class: 'flag' }, '🔁 احتياط'),
      h('span', { class: 'cls' }, c.cls || '—'), c.subject ? h('div', { class: 'subj' }, c.subject) : null,
      (c.sub.absent || c.teacher) ? h('div', { class: 'absent' }, h('span', { class: 'abs-tag' }, absentWord()), ' ', h('s', { class: 'abs-name' }, c.sub.absent || c.teacher)) : null,
      h('div', { class: 'tch sub-name' }, '↩ ' + (c.sub.sub || '—')));
    return h('div', { class: 'ccard' }, h('span', { class: 'cls' }, c.cls || '—'), c.subject ? h('div', { class: 'subj' }, c.subject) : null, h('div', { class: 'tch' }, c.teacher || '—'));
  },
  /* كل الصفوف في شاشة واحدة: جدول لكل صف دراسي (الخامس، السادس…) بأعمدة: الصف | المعلمة | المادة */
  gradeTable(cards) {
    const GR = ['', 'الأول', 'الثاني', 'الثالث', 'الرابع', 'الخامس', 'السادس', 'السابع', 'الثامن', 'التاسع', 'العاشر', 'الحادي عشر', 'الثاني عشر'];
    const by = new Map();
    const num = c => { const m = normCls(c.cls).match(/^(\d{1,2})\/(\d{1,2})/); return m ? +m[1] * 100 + +m[2] : 9999; };
    cards = cards.slice().sort((a, b) => num(a) - num(b));
    cards.forEach(c => { const m = normCls(c.cls).match(/^(\d{1,2})\//); const k = m ? +m[1] : 0; (by.get(k) || by.set(k, []).get(k)).push(c); });
    let groups = [...by.entries()].sort((a, b) => a[0] - b[0]).map(([k, list]) => ({ title: k ? 'الصف ' + (GR[k] || k) : 'صفوف أخرى', list }));
    if (groups.length === 1 && groups[0].list.length > 9) { const L = groups[0].list, half = Math.ceil(L.length / 2); groups = [{ title: groups[0].title, list: L.slice(0, half) }, { title: '', list: L.slice(half) }]; }
    this.pages = 1; this.page = 0; this.more = false;
    const stage = $('#stage'), H = (stage.clientHeight || 600) - Math.max(110, stage.clientHeight * 0.2);
    const rows = Math.max(...groups.map(g => g.list.length)) + 1;
    const fs = Math.max(12, Math.min(44, Math.floor(H / rows * 0.5)));
    return h('div', { class: 'gtables fit', style: `grid-template-columns:repeat(${groups.length},1fr);font-size:${fs}px` },
      groups.map((g, gi) => h('table', { class: 'gtable', style: `--i:${gi}` },
        h('tr', {}, h('th', { class: 'gc' }, g.title || 'الصف'), h('th', {}, G('teacher')), h('th', {}, 'المادة')),
        g.list.map((c, k) => h('tr', { class: c.sub ? 'sub' : '', style: `--i:${k}` },
          h('td', { class: 'gc' }, c.cls || '—'),
          h('td', { class: 'gt' }, c.sub
            ? [h('span', { class: 'sub-name' }, '↩ ' + shortName(c.sub.sub || '—')), ' ', h('small', { class: 'absent' }, h('span', { class: 'abs-tag' }, absentWord()), ' ', h('s', { class: 'abs-name' }, shortName(c.sub.absent || c.teacher)))]
            : shortName(c.teacher) || '—'),
          h('td', { class: 'gs' }, c.subject || (c.sub && c.sub.note) || ''))))));
  },
  /* عرض جدولي مضغوط: أعمدة (المعلم | الصف) مثل لوحات المدارس */
  tableView(items, cols, rowFn) {
    const stage = $('#stage');
    const W = stage.clientWidth || 1200, H = (stage.clientHeight || 600) - Math.max(100, stage.clientHeight * 0.17);
    const rowsPer = Math.max(3, Math.floor(H / 58) - 1);
    const groups = Math.max(1, Math.min(4, Math.floor(W / 380)));
    const per = rowsPer * groups;
    this.pages = Math.max(1, Math.ceil(items.length / per));
    if (this.page >= this.pages) this.page = 0;
    this.more = this.page < this.pages - 1;
    const slice = items.slice(this.page * per, this.page * per + per);
    const g = Math.min(groups, Math.max(1, Math.ceil(slice.length / rowsPer)));
    const n = Math.ceil(slice.length / g);
    return h('div', { class: 'tview stagger', style: `grid-template-columns:repeat(${g},1fr)` },
      Array.from({ length: g }, (_, k) => h('table', { class: 'tgroup', style: `--i:${k}` },
        h('tr', {}, cols.map(c => h('th', {}, c))),
        slice.slice(k * n, k * n + n).map(rowFn))));
  },
  /* المناوبة مجمّعة حسب الأسبوع: بطاقة لكل أسبوع فيها معلمات اليوم المحدد فقط */
  dutyWeekCards(rows, d) {
    const by = new Map();
    rows.forEach(r => { const k = r.week || 0; (by.get(k) || by.set(k, []).get(k)).push(r); });
    const groups = [...by.entries()].sort((a, b) => (a[0] || 99) - (b[0] || 99));
    const curW = +CFG.general.cycleWeeks > 1 ? cycleWeek(d) : -1;
    return this.grid(groups, ([w, list]) => h('div', { class: 'ccard wk' + (w === curW ? ' cur' : '') },
      h('span', { class: 'cls' }, w ? '🗓️ الأسبوع ' + (ORD[w - 1] || w) + (w === curW ? ' (الحالي)' : '') : '🛡️ كل الأسابيع'),
      // سطر لكل موقع: 📍 الموقع ثم أسماء المناوبات
      h('div', { class: 'names' }, (() => {
        const locs = new Map(); list.forEach(r => { const k = r.location || ''; (locs.get(k) || locs.set(k, []).get(k)).push(r.teacher); });
        return [...locs.entries()].map(([l, ns]) => h('div', { class: 'loc-line' }, l ? h('div', { class: 'meta' }, '📍 ' + l) : null, h('div', {}, ns.join('، '))));
      })())));
  },
  heroNow(p, st, sub) {
    const s = toSec(p.start), e = toSec(p.end), cur = st.cur === p;
    const pct = cur ? Math.round((st.sec - s) / (e - s) * 100) : 0;
    const left = cur ? e - st.sec : s - st.sec;
    return h('div', { class: 'slide hero-now' },
      h('div', { class: 'ring', style: `--p:${pct}` }, h('span', {}, digits(fmtDur(left)))),
      h('h1', {}, (cur ? '' : 'استعداد: ') + periodName(p)),
      h('p', {}, sub || (cur ? digits(`من ${fmt12(p.start)} إلى ${fmt12(p.end)}`) : 'تبدأ الساعة ' + digits(fmt12(p.start)))));
  },
  nowScene() {
    const d = now(), st = currentState(d);
    if (!st.work) return null;
    const p = st.cur || st.next;
    if (!p) return null;
    const cur = st.cur === p;
    const chips = [digits(`${fmt12(p.start)} — ${fmt12(p.end)}`), { t: (cur ? 'متبقٍ ' : 'بعد ') + digits(fmtDur(cur ? toSec(p.end) - st.sec : toSec(p.start) - st.sec)), hot: true }];
    if (isLesson(p)) {
      const cards = classCards(d, p);
      if (!cards.length) return this.heroNow(p, st);
      const g = CFG.display.nowStyle === 'grades' ? this.gradeTable(cards)
        : CFG.display.nowStyle === 'table' ? this.tableView(cards, [G('teacher'), 'الصف'], c => h('tr', { class: c.sub ? 'sub' : '' },
          h('td', { class: 'tn' }, c.sub ? [h('div', { class: 'sub-name' }, '↩ ' + shortName(c.sub.sub || '—')), h('div', { class: 'absent' }, h('span', { class: 'abs-tag' }, absentWord()), ' ', h('s', { class: 'abs-name' }, shortName(c.sub.absent || c.teacher)))] : [shortName(c.teacher) || '—', c.subject ? h('span', { class: 'ts' }, ' — ' + c.subject) : null]),
          h('td', { class: 'tc' }, c.cls || '—')))
        : this.grid(cards, c => this.classCard(c, p));
      return h('div', { class: 'slide scene' }, this.head(cur ? '📚' : '⏳', (cur ? '' : 'استعداد: ') + periodName(p) + (cur ? ' الآن' : ''), chips), g);
    }
    // فسحة / طابور / صلاة: اعرض المناوبين
    const duty = dutyFor(d, p).filter(r => slotKey(r.slot) !== '*').sort(byWeek);
    if (!duty.length) return this.heroNow(p, st, st.next && cur ? 'التالي: ' + periodName(st.next) : '');
    const g = CFG.display.dutyAllWeeks && duty.some(r => r.week)
      ? this.dutyWeekCards(duty, d)
      : this.grid(duty, r => h('div', { class: 'ccard' }, h('span', { class: 'cls' }, '🛡️ مناوبة'), h('div', { class: 'tch' }, r.teacher), r.location ? h('div', { class: 'meta' }, '📍 ' + r.location) : null));
    return h('div', { class: 'slide scene' }, this.head(p.type.startsWith('break') ? '☕' : (p.type === 'assembly' ? '🚩' : '🕌'), periodName(p) + (cur ? ' الآن' : '') + ' — مناوبة يوم ' + DAYS[d.getDay()], chips), g);
  },
  scheduleScene() {
    const d = now(), st = currentState(d), pr = profileFor(d);
    if (!st.work || !pr || !st.periods.length) return null;
    this.pages = 1; this.page = 0;
    const g = this.grid(st.periods, p => {
      const s = toSec(p.start), e = toSec(p.end), cur = st.cur === p, past = e <= st.sec;
      return h('div', { class: 'ccard' + (cur ? ' cur' : past ? ' past' : '') },
        h('div', { class: 'big' }, periodName(p)),
        h('div', { class: 'meta' }, digits(`${fmt12(p.start)} — ${fmt12(p.end)}`)),
        cur ? h('div', { class: 'bar' }, h('div', { style: `width:${Math.round((st.sec - s) / (e - s) * 100)}%` })) : null);
    });
    return h('div', { class: 'slide scene' }, this.head('⏰', 'توقيت اليوم', [pr.name, DAYS[d.getDay()]]), g);
  },
  dutyScene() {
    const d = now(); if (!isWorkDay(d)) return null;
    const rows = CFG.duty.rows.filter(r => (r.date ? r.date === ymd(d) : (r.day === -1 || r.day === d.getDay())) && weekOK(r, d));
    if (!rows.length) return null;
    const st = currentState(d), keys = st.cur ? periodKeys(st.cur) : (st.next ? periodKeys(st.next) : new Set());
    rows.sort(byWeek);
    // بطاقة لكل (موقع + أسبوع) فيها أسماء المناوبات، موزعة على صفحات
    const cw = +CFG.general.cycleWeeks > 1 ? cycleWeek(d) : -1;
    const grp = new Map();
    rows.forEach(r => { const k = (r.week || 0) + '|' + (r.location || '') + '|' + (r.location ? '' : r.slot); (grp.get(k) || grp.set(k, []).get(k)).push(r); });
    const groups = [...grp.values()].sort((a, b) => (a[0].week || 99) - (b[0].week || 99));
    const g = this.grid(groups, list => {
      const r = list[0], byPeriod = list.length > 2 && /الحصه/.test(norm(r.slot));
      return h('div', { class: 'ccard' + (r.week && r.week === cw ? ' cur' : '') },
        h('span', { class: 'cls' }, r.week ? weekTag(r) : 'كل الأسابيع'), h('div', { class: 'subj' }, '📍 ' + (r.location || r.slot || '') + (r.slot && !byPeriod ? ' — ' + r.slot : '')),
        byPeriod ? h('div', { class: 'names small' }, list.map(x => h('div', {}, x.slot.replace('الحصة ', '') + ': ', h('b', {}, x.teacher))))
          : h('div', { class: 'tch' }, list.map(x => x.teacher).join('، ')));
    }, { minW: 340, rowFrac: 0.31 });
    if (false) this.grid(rows, r => h('div', { class: 'ccard' + (keys.has(slotKey(r.slot)) ? ' cur' : '') },
      h('span', { class: 'cls' }, r.slot || 'طوال اليوم'), r.week ? h('div', { class: 'subj' }, weekTag(r)) : null, h('div', { class: 'tch' }, r.teacher), r.location ? h('div', { class: 'meta' }, '📍 ' + r.location) : null));
    return h('div', { class: 'slide scene' }, this.head('🛡️', 'المناوبة اليوم', [DAYS[d.getDay()], CFG.general.cycleWeeks > 1 && !CFG.display.dutyAllWeeks ? 'الأسبوع ' + cycleWeek(d) : (CFG.display.dutyAllWeeks ? 'كل الأسابيع' : null)]), g);
  },
  subsScene() {
    const d = now(); if (!isWorkDay(d)) return null;
    const rows = subsFor(d); if (!rows.length) return null;
    const st = currentState(d), keys = st.cur ? periodKeys(st.cur) : new Set();
    let g;
    if (CFG.display.subsStyle === 'grouped') {
      // مجمّع حسب المعلم الغائب
      const by = new Map(); rows.forEach(r => { const k = r.absent || '—'; (by.get(k) || by.set(k, []).get(k)).push(r); });
      const flat = []; for (const [abs, list] of by) list.forEach((r, i) => flat.push({ r, abs, first: i === 0, span: list.length }));
      g = this.tableView(flat, [G('absent'), G('sub'), 'الصف', 'الحصة'], x => h('tr', { class: keys.has(slotKey(x.r.period)) ? 'cur' : '' },
        h('td', { class: 'tabs' + (x.first ? '' : ' cont') }, x.first ? [h('span', { class: 'abs-tag' }, absentWord()), ' ', h('b', {}, x.abs)] : ''),
        h('td', { class: 'sub-name' }, x.r.sub || '—'), h('td', { class: 'tc' }, x.r.cls || ''), h('td', {}, x.r.period || '')));
    } else g = this.grid(rows, r => {
      const c = this.classCard({ cls: r.cls, subject: (r.period ? 'الحصة ' + r.period : '') + (r.note ? ' — ' + r.note : ''), teacher: r.absent, sub: r });
      if (keys.has(slotKey(r.period))) c.classList.add('cur');
      return c;
    });
    return h('div', { class: 'slide scene' }, this.head('🔁', 'حصص الاحتياط اليوم', [{ t: digits(rows.length) + ' حصة', hot: true }]), g);
  },
  offScene() {
    const o = offState(), d = now();
    const info = {
      weekend: ['🌙', 'لا يوجد دوام اليوم', 'عطلة نهاية الأسبوع — نتمنى لكم راحة سعيدة'],
      holiday: ['🌴', 'لا يوجد دوام اليوم', o.st.holiday?.name || 'إجازة'],
      nosched: ['📅', 'لا يوجد دوام اليوم', 'لا توجد حصص مسجلة لهذا اليوم'],
      ended: ['🌇', 'انتهى الدوام الدراسي', 'شكراً لجهودكم — نراكم بإذن الله'],
      early: ['🌅', 'صباح الخير', 'لم يبدأ الدوام بعد'],
    }[o.why];
    const nx = nextSchoolStart(d);
    let when = '';
    if (nx) {
      const sameDay = ymd(nx) === ymd(d), tmr = ymd(nx) === ymd(new Date(d.getFullYear(), d.getMonth(), d.getDate() + 1));
      let date = ''; try { date = fmtGreg().format(nx); } catch { }
      const hh = nx.getHours(), t12 = digits(fmt12(pad(hh) + ':' + pad(nx.getMinutes()))) + (CFG.general.clock24 ? '' : (hh < 12 ? ' ص' : ' م'));
      when = (sameDay ? 'اليوم' : tmr ? 'غداً ' + DAYS[nx.getDay()] : DAYS[nx.getDay()] + ' ' + date) + ' — الساعة ' + t12;
    }
    const ph = CFG.phrases.enabled ? Phrases.pick('d') : null;
    return h('div', { class: 'slide off-scene' },
      h('div', { class: 'off-ic' }, info[0]),
      h('h1', {}, info[1]), h('p', { class: 'off-sub' }, info[2]),
      nx ? h('div', { class: 'off-next' }, h('span', {}, '📚 الدوام القادم'), h('b', {}, when), h('div', { class: 'off-count', 'data-t': nx.getTime() })) : null,
      ph ? h('div', { class: 'off-ph' }, '🤲 ' + ph.t) : null);
  },
  offTick(el) {
    clearInterval(this.offT);
    const c = el.querySelector('.off-count'); if (!c) return;
    const upd = () => {
      const s = Math.max(0, Math.round((+c.dataset.t - now()) / 1000)), dd = Math.floor(s / 86400), hh = Math.floor(s / 3600) % 24, mm = Math.floor(s / 60) % 60;
      c.textContent = 'بعد ' + digits([dd ? dd + ' يوم' : '', hh ? hh + ' ساعة' : '', mm + ' دقيقة'].filter(Boolean).join(' و '));
    };
    upd(); this.offT = setInterval(() => { if (!el.isConnected) return clearInterval(this.offT); upd(); }, 20000);
  },
  /* جدول اليوم: الصفوف × الحصص، مع إبراز الحصة الحالية والاحتياط */
  classDayScene() {
    const d = now(), st = currentState(d); if (!st.work) return null;
    const lessons = st.periods.filter(isLesson); if (!lessons.length || !CFG.timetable.rows.length) return null;
    const dow = d.getDay(), wk = cycleWeek(d);
    const todayRows = CFG.timetable.rows.filter(r => (r.day === -1 || r.day === dow) && (!r.week || r.week === wk));
    if (!todayRows.length) return null;
    const classes = (CFG.timetable.classes.length ? CFG.timetable.classes : [...new Set(todayRows.map(r => r.cls))]).slice().sort(clsSorter());
    const subs = subsFor(d);
    const stage = $('#stage'), H = (stage.clientHeight || 600) - Math.max(110, stage.clientHeight * 0.2);
    const per = Math.max(4, Math.floor(H / 46));
    this.pages = Math.max(1, Math.ceil(classes.length / per)); if (this.page >= this.pages) this.page = 0;
    this.more = this.page < this.pages - 1;
    const slice = classes.slice(this.page * per, this.page * per + per);
    const curIdx = lessons.indexOf(st.cur);
    const cellFor = (cls, p) => {
      const keys = periodKeys(p), n = normCls(cls);
      const r = todayRows.find(x => normCls(x.cls) === n && keys.has(slotKey(x.period)));
      const sb = subs.find(x => normCls(x.cls) === n && keys.has(slotKey(x.period)));
      if (sb) return h('td', { class: 'dg-sub' }, h('b', {}, '↩ ' + (sb.sub || '')), h('s', {}, sb.absent || r?.teacher || ''));
      if (!r) return h('td', { class: 'dg-empty' }, '—');
      const nm = String(r.teacher || '').trim().split(/\s+/); const short = nm.length > 2 ? nm[0] + ' ' + nm[nm.length - 1] : r.teacher;
      return h('td', { title: r.teacher }, h('b', {}, short || r.subject), r.teacher && r.subject ? h('small', {}, r.subject) : null);
    };
    const table = h('table', { class: 'daygrid stagger' },
      h('tr', {}, h('th', { class: 'dg-cls' }, 'الصف'), lessons.map((p, i) => h('th', { class: i === curIdx ? 'dg-cur' : (curIdx >= 0 && i < curIdx ? 'dg-past' : '') }, periodName(p).replace('الحصة ', '')))),
      slice.map((cls, k) => h('tr', { style: `--i:${k}` }, h('td', { class: 'dg-cls' }, cls), lessons.map((p, i) => { const c = cellFor(cls, p); if (i === curIdx) c.classList.add('dg-cur'); else if (curIdx >= 0 && i < curIdx) c.classList.add('dg-past'); return c; }))));
    return h('div', { class: 'slide scene' }, this.head('🗓️', 'جدول حصص اليوم', [DAYS[dow], st.cur && isLesson(st.cur) ? { t: 'الآن: ' + periodName(st.cur), hot: true } : null]), h('div', { class: 'dg-wrap' }, table));
  },
  absentScene() {
    const d = now(); if (!isWorkDay(d)) return null;
    const rows = absentFor(d); if (!rows.length) return null;
    const by = new Map();
    rows.forEach(r => { const k = r.cls || '—'; (by.get(k) || by.set(k, []).get(k)).push(r); });
    const groups = [...by.entries()];
    const word = CFG.general.gender === 'm' ? 'الطلاب الغائبون' : CFG.general.gender === 'f' ? 'الطالبات الغائبات' : 'الغياب';
    const show = CFG.absent.showNames;
    const g = this.grid(groups, ([cls, list]) => h('div', { class: 'ccard' },
      h('span', { class: 'cls' }, cls),
      show ? h('div', { class: 'names' }, list.map(r => h('div', {}, r.name, r.reason ? h('span', { class: 'meta' }, ' — ' + r.reason) : null)))
        : h('div', { class: 'big' }, digits(list.length) + (list.length > 2 && list.length < 11 ? ' غياب' : ' غياب'))));
    return h('div', { class: 'slide scene' }, this.head('📋', word + ' اليوم', [{ t: digits(rows.length) + ' ' + (CFG.general.gender === 'm' ? 'طالب' : CFG.general.gender === 'f' ? 'طالبة' : 'غياب'), hot: true }, digits(groups.length) + ' صف']), g);
  },
  duck(on) {
    if (!this.video) return;
    if (on) { this._wasMuted = this.video.muted; this.video.muted = true; }
    else if (this._wasMuted === false) this.video.muted = false;
  },
};

/* ------------------------------ مصادر المناوبة/الاحتياط ------------------------------ */
function parseCSV(text) {
  const rows = []; let row = [], cell = '', q = false;
  text = text.replace(/^﻿/, '');
  const sep = (text.split('\n')[0].match(/;/g) || []).length > (text.split('\n')[0].match(/,/g) || []).length ? ';' : (text.includes('\t') && !text.includes(',') ? '\t' : ',');
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (q) { if (c === '"') { if (text[i + 1] === '"') { cell += '"'; i++; } else q = false; } else cell += c; }
    else if (c === '"') q = true;
    else if (c === sep) { row.push(cell); cell = ''; }
    else if (c === '\n') { row.push(cell); rows.push(row); row = []; cell = ''; }
    else if (c !== '\r') cell += c;
  }
  if (cell || row.length) { row.push(cell); rows.push(row); }
  return rows;
}
function aoaFromWorkbook(wb, sheetName) {
  let name = wb.SheetNames[0];
  if (sheetName) name = wb.SheetNames.find(n => norm(n) === norm(sheetName)) || wb.SheetNames.find(n => norm(n).includes(norm(sheetName))) || name;
  return XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: false, defval: '', dateNF: 'yyyy-mm-dd' });
}
async function aoaFromBlob(blob, sheetName) {
  const buf = new Uint8Array(await blob.arrayBuffer());
  if (buf[0] === 0x50 && buf[1] === 0x4B || (buf[0] === 0xD0 && buf[1] === 0xCF)) return aoaFromWorkbook(XLSX.read(buf, { type: 'array' }), sheetName);
  const text = new TextDecoder('utf-8').decode(buf);
  if (/^\s*<(!doctype|html)/i.test(text)) throw new Error('الرابط أعاد صفحة ويب وليس جدولاً. تأكد من مشاركة الملف للجميع أو نشره كـ CSV.');
  return parseCSV(text);
}
function sheetFetchURL(url) {
  let m;
  if ((m = url.match(/docs\.google\.com\/spreadsheets\/d\/([\w-]+)/)) && !/output=csv|format=csv|tqx=out:csv/.test(url)) {
    const gid = (url.match(/[#&?]gid=(\d+)/) || [])[1];
    return `https://docs.google.com/spreadsheets/d/${m[1]}/gviz/tq?tqx=out:csv${gid ? '&gid=' + gid : ''}`;
  }
  if ((m = url.match(/drive\.google\.com\/(?:file\/d\/|open\?id=)([\w-]+)/))) return `https://drive.google.com/uc?export=download&id=${m[1]}`;
  return url;
}
const cell = (row, i) => i >= 0 ? String(row[i] ?? '').trim() : '';
function findHeader(aoa, tests, min = 2) {
  for (let r = 0; r < Math.min(15, aoa.length); r++) {
    const hits = aoa[r].filter(c => tests.some(t => t.test(norm(c)))).length;
    if (hits >= min) return r;
  }
  return -1;
}
function parseDuty(aoa) {
  aoa = aoa.filter(r => r.some(c => String(c).trim()));
  const out = [];
  // شكل مصفوفة: الأيام في الأعمدة
  for (let r = 0; r < Math.min(15, aoa.length); r++) {
    const dayCols = aoa[r].map((c, j) => [j, dayIdx(c)]).filter(([j, d]) => d >= 0 && String(aoa[r][j]).length < 15);
    if (dayCols.length >= 2) {
      const hdr = aoa[r].map(norm);
      const wCol = hdr.findIndex(c => /اسبوع/.test(c));
      const lCol = hdr.findIndex(c => /مكان|موقع|منطق/.test(c));
      const dc = new Set(dayCols.map(x => x[0]));
      let sCol = hdr.findIndex((c, j) => !dc.has(j) && j !== wCol && j !== lCol && /حص|فتر|وقت|موعد/.test(c));
      if (sCol < 0) sCol = hdr.findIndex((c, j) => !dc.has(j) && j !== wCol && j !== lCol);
      let wk = 0, lastSlot = '';
      for (const row of aoa.slice(r + 1)) {
        if (wCol >= 0 && cell(row, wCol)) wk = weekNum(cell(row, wCol));
        const slot = cell(row, sCol) || lastSlot; lastSlot = slot;
        if (/اسبوع/.test(norm(slot))) { wk = weekNum(slot); continue; }
        if (dayIdx(slot) >= 0) continue;
        for (const [j, d] of dayCols) {
          const t = cell(row, j);
          if (t) out.push({ week: wk, day: d, date: '', slot, teacher: t.replace(/\n/g, '، '), location: cell(row, lCol) });
        }
      }
      return out;
    }
  }
  // شكل قائمة
  const hr = findHeader(aoa, [/يوم/, /معلم|مناوب|اسم|المكلف/, /حص|فتر|وقت/, /اسبوع/, /مكان|موقع/, /تاريخ/]);
  if (hr < 0) throw new Error('لم أتعرف على أعمدة الجدول. يجب أن يحتوي على أعمدة مثل: اليوم، الفترة/الحصة، المعلم المناوب، المكان، الأسبوع.');
  const hd = aoa[hr].map(norm);
  const col = re => hd.findIndex(c => re.test(c));
  const c = { week: col(/اسبوع/), day: col(/يوم/), date: col(/تاريخ/), slot: col(/حص|فتر|وقت|موعد/), teacher: col(/معلم|مناوب|اسم|المكلف/), loc: col(/مكان|موقع|منطق|ساح/) };
  let wk = 0, dy = -1;
  for (const row of aoa.slice(hr + 1)) {
    if (cell(row, c.week)) wk = weekNum(cell(row, c.week));
    if (cell(row, c.day)) dy = dayIdx(cell(row, c.day));
    const teacher = cell(row, c.teacher); if (!teacher) continue;
    out.push({ week: wk, day: c.day >= 0 ? dy : -1, date: parseDate(cell(row, c.date)), slot: cell(row, c.slot), teacher, location: cell(row, c.loc) });
  }
  return out;
}
function parseSubs(aoa) {
  aoa = aoa.filter(r => r.some(c => String(c).trim()));
  // يدعم العناوين العربية، والإنجليزية (مثل برنامج الاحتياط: date, period, classSection, absentTeacherId...)
  const hr = findHeader(aoa, [/تاريخ|^date$/, /يوم|^day$/, /حص|^period$/, /صف|فصل|شعب|class/, /غائب|غياب|absent/, /احتياط|بديل|منفذ|المكلف|القائم|substitute/, /معلم|اسم/]);
  if (hr < 0) throw new Error('لم أتعرف على أعمدة جدول الاحتياط. يجب أن يحتوي على أعمدة مثل: التاريخ أو اليوم، الحصة، الصف، المعلم الغائب، المعلم البديل.');
  const hd = aoa[hr].map(norm);
  const col = (re, not) => hd.findIndex(c => re.test(c) && !(not && not.test(c)));
  const c = {
    date: col(/تاريخ|^date$/, /created|updated|received/), day: col(/يوم|^day$/, /تاريخ/), period: col(/حص|^period$/, /احتياط/),
    cls: col(/صف|فصل|شعب|class/), absent: col(/غائب|غياب|absentteacher|^absent$/, /specialty|تخصص/),
    sub: col(/احتياط|بديل|منفذ|المكلف|القائم|substituteteacher|^substitute$/, /specialty|تخصص/),
    subject: col(/ماده|absentspecialty/), note: col(/ملاحظ|^notes?$/),
  };
  if (c.absent < 0) c.absent = col(/absent/);
  if (c.sub < 0) c.sub = hd.findIndex((x, j) => /معلم|اسم|substitute/.test(x) && j !== c.absent);
  let dt = '', dy = -1;
  const out = [];
  for (const row of aoa.slice(hr + 1)) {
    if (cell(row, c.date)) dt = parseDate(cell(row, c.date));
    if (cell(row, c.day)) dy = dayIdx(cell(row, c.day));
    const sub = cell(row, c.sub); if (!sub && !cell(row, c.cls)) continue;
    let period = cell(row, c.period);
    const pn = toLatin(period).match(/^\d$/); if (pn && TYPE['p' + pn[0]]) period = TYPE['p' + pn[0]].n.replace('الحصة ', '');
    out.push({ date: dt, day: dt ? -1 : dy, period, cls: cell(row, c.cls), absent: cell(row, c.absent), sub, note: cell(row, c.subject) || cell(row, c.note) });
  }
  return out;
}
/* بعض البرامج تحفظ رموز المعلمين (T1fc73278ed) بدل الأسماء: نحولها لأسماء من ورقة المعلمين */
const isIdCode = v => /^T[0-9a-f]{6,}$/i.test(String(v || '').trim());
function nameMapFromAoa(aoa) {
  aoa = aoa.filter(r => r.some(c => String(c).trim()));
  const hd = (aoa[0] || []).map(norm);
  let ci = hd.findIndex(x => /^id$|رمز|الرقم/.test(x)), ni = hd.findIndex(x => /^name$|الاسم|اسم/.test(x));
  if (ci < 0) ci = 0; if (ni < 0) ni = 1;
  const m = {}; aoa.slice(1).forEach(r => { const k = String(r[ci] ?? '').trim(); if (k) m[k] = String(r[ni] ?? '').trim(); });
  return m;
}
function applyNameMap(rows, map, keys = ['absent', 'sub']) {
  let n = 0;
  rows.forEach(r => keys.forEach(k => { const v = String(r[k] || '').trim(); if (map[v]) { r[k] = map[v]; n++; } }));
  return n;
}
async function resolveTeacherIds(rows, src) {
  if (!rows.some(r => isIdCode(r.absent) || isIdCode(r.sub))) return;
  let map = null;
  try {
    if (src.namesUrl) { const res = await smartRead(src.namesUrl); const sh = res.type === 'pdf' ? await pdfToSheets(res.blob) : res.sheets; map = nameMapFromAoa(sh[0].aoa); }
    else {
      const m = normURL(src.url).match(/docs\.google\.com\/spreadsheets\/d\/([\w-]+)/);
      if (m) {
        for (const sheet of ['Teachers', 'المعلمين', 'المعلمون', 'المعلمات']) {
          try {
            const mm = nameMapFromAoa(await gsheetJSONP(m[1], { sheet }));
            if (Object.keys(mm).some(isIdCode)) { map = mm; break; }
          } catch { }
        }
      }
    }
  } catch (e) { console.warn(e); }
  if (map) { applyNameMap(rows, map); src.namesCache = map; }
  else if (src.namesCache) applyNameMap(rows, src.namesCache);
}
/* قراءة كل أوراق الملف (لجدول المعلمين: ورقة لكل يوم مثلاً) */
async function sheetsFromBlob(blob) {
  const buf = new Uint8Array(await blob.arrayBuffer());
  if (buf[0] === 0x50 && buf[1] === 0x4B || (buf[0] === 0xD0 && buf[1] === 0xCF)) {
    const wb = XLSX.read(buf, { type: 'array' });
    return wb.SheetNames.map(name => ({ name, aoa: XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: false, defval: '' }) }));
  }
  const text = new TextDecoder('utf-8').decode(buf);
  if (/^\s*<(!doctype|html)/i.test(text)) throw new Error('الرابط أعاد صفحة ويب وليس جدولاً. تأكد من مشاركة الملف للجميع.');
  return [{ name: '', aoa: parseCSV(text) }];
}
const SUBJ_RE = /رياض|عربي|لغه|انجليز|english|علوم|اسلام|قران|تربيه|اجتماع|دراسات|فنون|موسيق|حاسوب|تقنيه|مهارات|فيزياء|كيمياء|احياء|جغراف|تاريخ|حياتيه|نشاط|رياده|math|science|بدنيه/;
const isSubject = t => SUBJ_RE.test(norm(t));
const isTeacher = t => /^(ا\.|ا\/|استاذ|المعلم|معلم)/.test(norm(t));
function splitCell(t) {
  t = String(t).trim();
  const parts = t.split(/\n|\s[-–—|]\s|\s*[\/،,]\s*(?=(?:ا|أ)\.|المعلم)/).map(x => x.trim()).filter(Boolean);
  if (parts.length < 2) return isSubject(t) && !isTeacher(t) ? { subject: t, teacher: '' } : { subject: '', teacher: t };
  let i = parts.findIndex(isTeacher);
  if (i >= 0) {
    let j = i + 1; while (j < parts.length && !isSubject(parts[j]) && !isTeacher(parts[j])) j++;
    return { teacher: parts.slice(i, j).join(' '), subject: parts.filter((_, k) => k < i || k >= j).join(' ') };
  }
  i = parts.findIndex(isSubject);
  if (i >= 0) return { subject: parts[i], teacher: parts.filter((_, k) => k !== i).join(' ') };
  return { subject: parts[0], teacher: parts.slice(1).join(' ') };
}
function parseTimetable(sheets, only) {
  const out = [];
  for (const sh of sheets) {
    if (only && !norm(sh.name).includes(norm(only))) continue;
    const aoa = sh.aoa.filter(r => r.some(c => String(c).trim()));
    const sheetDay = dayIdx(sh.name);
    // 1) شكل مصفوفة: الحصص في الأعمدة
    let mh = -1, pcols = [];
    for (let r = 0; r < Math.min(15, aoa.length); r++) {
      const cols = aoa[r].map((c, j) => [j, slotKey(c), String(c).trim()]).filter(([, k, raw]) => /^p\d$/.test(k) && raw.length < 25);
      if (cols.length >= 3) { mh = r; pcols = cols; break; }
    }
    if (mh >= 0) {
      const hd = aoa[mh].map(norm), pset = new Set(pcols.map(x => x[0]));
      const dCol = hd.findIndex((c, j) => !pset.has(j) && /يوم/.test(c));
      const wCol = hd.findIndex((c, j) => !pset.has(j) && /اسبوع/.test(c));
      let kCol = hd.findIndex((c, j) => !pset.has(j) && /صف|فصل|شعب|معلم|اسم/.test(c));
      if (kCol < 0) kCol = hd.findIndex((c, j) => !pset.has(j) && j !== dCol && j !== wCol);
      const byTeacher = kCol >= 0 && /معلم|اسم/.test(hd[kCol]) && !/صف|فصل|شعب/.test(hd[kCol]);
      let day = sheetDay, wk = 0;
      for (const row of aoa.slice(mh + 1)) {
        const filled = row.map(c => String(c).trim()).filter(Boolean);
        if (filled.length === 1 && dayIdx(filled[0]) >= 0) { day = dayIdx(filled[0]); continue; }
        if (dCol >= 0 && cell(row, dCol)) day = dayIdx(cell(row, dCol));
        if (wCol >= 0 && cell(row, wCol)) wk = weekNum(cell(row, wCol));
        const key = cell(row, kCol);
        if (!key || pcols.filter(([j]) => /^p\d$/.test(slotKey(cell(row, j)))).length >= 3) continue;
        for (const [j, , raw] of pcols) {
          const t = cell(row, j); if (!t) continue;
          if (byTeacher) {
            const m = toLatin(t).match(/\d+\s*[\/\-\\]\s*\d+|\d+\s*[أ-ي]\b/);
            const cls = m ? m[0].replace(/\s+/g, '') : t;
            out.push({ week: wk, day, period: raw, cls, subject: m ? toLatin(t).replace(m[0], '').replace(/[-–—،,]/g, ' ').trim() : '', teacher: key });
          }
          else { const x = splitCell(t); out.push({ week: wk, day, period: raw, cls: key, subject: x.subject, teacher: x.teacher }); }
        }
      }
      continue;
    }
    // 2) شكل قائمة
    const hr = findHeader(aoa, [/صف|فصل|شعب/, /معلم|اسم/, /حص/], 3);
    if (hr < 0) continue;
    const hd = aoa[hr].map(norm), col = re => hd.findIndex(c => re.test(c));
    const c = { day: col(/يوم/), week: col(/اسبوع/), period: col(/حص/), cls: col(/صف|فصل|شعب/), teacher: col(/معلم|اسم/), subject: col(/ماده/) };
    let day = sheetDay, wk = 0;
    for (const row of aoa.slice(hr + 1)) {
      if (cell(row, c.day)) day = dayIdx(cell(row, c.day));
      if (cell(row, c.week)) wk = weekNum(cell(row, c.week));
      if (!cell(row, c.cls) || !cell(row, c.period)) continue;
      out.push({ week: wk, day, period: cell(row, c.period), cls: cell(row, c.cls), subject: cell(row, c.subject), teacher: cell(row, c.teacher) });
    }
  }
  if (!out.length) throw new Error('لم أتعرف على جدول المعلمين. الصيغ المقبولة: (1) الصفوف في الصفوف والحصص (الأولى، الثانية...) في الأعمدة، وكل ورقة ليوم أو عمود "اليوم". (2) قائمة بأعمدة: اليوم، الحصة، الصف، المعلم، المادة.');
  return out;
}

/* ---------------------------------------------------------------------
   قراءة جدول الحصص من ملف PDF (نصي، وليس صورة ممسوحة)
   يتعرف على: صفحة لكل صف أو لكل معلم (الأيام صفوف والحصص أعمدة)،
   أو جدول عام (الصفوف/المعلمون صفوف والحصص أعمدة، مع الأيام فوقها أو صفحة لكل يوم).
   --------------------------------------------------------------------- */
const PERIOD_WORDS = ['صفريه', 'اولي', 'ثانيه', 'ثالثه', 'رابعه', 'خامسه', 'سادسه', 'سابعه', 'ثامنه', 'تاسعه'];
function isPeriodLabel(t) {
  const n = norm(t).replace(/^الحصه\s*/, '').replace(/^ح\s*/, '').replace(/^ال/, '');
  return /^\d{1,2}$/.test(n) || PERIOD_WORDS.includes(n) || PERIOD_WORDS.includes('ال' + n);
}
const isClassLike = t => /\d+\s*[\/\-\\]\s*\d+|\d+\s*[أ-ي]$|^(صف|الصف|فصل)\s/.test(toLatin(norm(t)));
const KNOWN_AR = ['الاحد', 'الاثنين', 'الثلاثاء', 'الاربعاء', 'الخميس', 'الحصه', 'الصف', 'الاولي', 'الثانيه', 'الثالثه', 'المعلم', 'جدول', 'اليوم'];
async function pdfTextPages(blob) {
  const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(await blob.arrayBuffer()), isEvalSupported: false }).promise;
  const pages = [];
  for (let i = 1; i <= pdf.numPages; i++) {
    const tc = await (await pdf.getPage(i)).getTextContent();
    pages.push(tc.items.filter(it => it.str && it.str.trim()).map(it => {
      const hgt = Math.abs(it.transform[3]) || Math.abs(it.transform[0]) || 10;
      return { s: it.str.normalize('NFKC'), xL: it.transform[4], xR: it.transform[4] + it.width, y: it.transform[5], h: hgt };
    }));
  }
  // بعض ملفات PDF تخزن العربي بترتيب معكوس
  const all = pages.flat().map(t => norm(t.s));
  const hits = arr => arr.reduce((n, s) => n + KNOWN_AR.filter(w => s.includes(w)).length, 0);
  const rev = s => [...s].reverse().join('');
  if (hits(all.map(rev)) > hits(all) * 1.5) pages.forEach(p => p.forEach(t => { t.s = rev(t.s); }));
  return pages;
}
/* دمج أجزاء الكلمات المتجاورة على نفس السطر */
function mergeLine(tokens) {
  const lines = [];
  for (const t of tokens.slice().sort((a, b) => b.y - a.y)) {
    const L = lines.find(l => Math.abs(l.y - t.y) < Math.min(l.h, t.h) * 0.45);
    if (L) L.items.push(t); else lines.push({ y: t.y, h: t.h, items: [t] });
  }
  const out = [];
  for (const L of lines) {
    const it = L.items.sort((a, b) => b.xR - a.xR);
    let cur = null;
    for (const t of it) {
      if (cur && cur.xL - t.xR < Math.max(cur.h, t.h) * 0.55) { cur.s = cur.s + (cur.xL - t.xR > cur.h * 0.15 ? ' ' : '') + t.s; cur.xL = Math.min(cur.xL, t.xL); cur.parts.push(t); }
      else { cur = { ...t, parts: [t] }; out.push(cur); }
    }
  }
  out.forEach(t => {
    // نص بلا حروف عربية (مثل 6\4 أو 8:00 - 8:45) يُقرأ من اليسار لليمين
    if (t.parts.length > 1 && !/[\u0600-\u06FF]/.test(t.s)) t.s = t.parts.slice().sort((a, b) => a.xL - b.xL).map((p, i, arr) => (i && p.xL - arr[i - 1].xR > p.h * 0.15 ? ' ' : '') + p.s).join('');
    t.s = t.s.replace(/\s+/g, ' ').trim(); t.x = (t.xL + t.xR) / 2;
  });
  return out.filter(t => t.s);
}
const GRADES = ['اول', 'ثاني', 'ثالث', 'رابع', 'خامس', 'سادس', 'سابع', 'ثامن', 'تاسع', 'عاشر', 'حادي عشر', 'ثاني عشر'];
function gradeOf(t) {
  const n = norm(t);
  if (/ثاني عشر/.test(n)) return 12; if (/حادي عشر/.test(n)) return 11;
  for (let i = 9; i >= 0; i--) if (new RegExp('(^|\\s)ال' + GRADES[i]).test(n)) return i + 1;
  return 0;
}
/* خلية جدول صف: أسطر المادة (قد تنكسر الكلمة) ثم اسم المعلمة في السطر الأخير */
function splitClassCell(ts) {
  const lines = [];
  ts.slice().sort((a, b) => b.y - a.y).forEach(t => { const L = lines.find(l => Math.abs(l.y - t.y) < t.h * 0.5); if (L) L.s += ' ' + t.s; else lines.push({ y: t.y, s: t.s }); });
  const txt = lines.map(l => l.s.trim()).filter(Boolean);
  if (!txt.length) return { subject: '', teacher: '' };
  let teacher = '';
  if (txt.length > 1 && !isSubject(txt[txt.length - 1])) teacher = txt.pop();
  let subject = '';
  for (const l of txt) subject += (subject && l.length > 3 ? ' ' : '') + l;
  return { subject: subject.replace(/\s+/g, ' ').trim(), teacher };
}
const isTimeTok = t => /^\d{1,2}[:.]\d{2}(\s*[-–]\s*\d{1,2}[:.]\d{2})?$/.test(toLatin(t).trim()) || /^[-–]$/.test(t.trim());
const isFooterTok = t => /\d{1,2}\/\d{1,2}\/\d{2,4}|asc|timetables|تم انشاء|تم إنشاء|اعداد|برنامج/i.test(norm(t));
function parseTimetablePDFPage(raw, pageNo) {
  const toks = mergeLine(raw).filter(t => !isTimeTok(t.s) && !isFooterTok(t.s));
  if (!toks.length) return { rows: [], reason: 'empty' };
  // 1) سطر رؤوس الحصص
  const lines = {};
  toks.forEach(t => { const k = Math.round(t.y / (t.h * 0.5)); (lines[k] = lines[k] || []).push(t); });
  let head = null;
  for (const L of Object.values(lines)) {
    const ps = L.filter(t => isPeriodLabel(t.s));
    if (ps.length >= 3 && (!head || ps.length > head.length || (ps.length === head.length && ps[0].y > head[0].y))) head = ps;
  }
  if (!head) return { rows: [], reason: 'nohead' };
  const hy = head.reduce((a, t) => a + t.y, 0) / head.length, hh = head[0].h;
  const anchors = head.map(t => ({ x: t.x, raw: t.s })).sort((a, b) => b.x - a.x);
  const gaps = anchors.slice(1).map((a, i) => anchors[i].x - a.x).sort((a, b) => a - b);
  const colW = gaps[Math.floor(gaps.length / 2)] || 40;
  const xMax = anchors[0].x + colW * 0.55, xMin = anchors[anchors.length - 1].x - colW * 0.55;
  // 2) أيام فوق رؤوس الحصص (جدول عام لكل الأسبوع)
  const dayToks = toks.filter(t => t.y > hy + hh * 0.3 && t.y < hy + hh * 4 && dayIdx(t.s) >= 0 && t.s.length < 15);
  if (dayToks.length >= 2) anchors.forEach(a => { a.day = dayIdx(dayToks.reduce((b, t) => Math.abs(t.x - a.x) < Math.abs(b.x - a.x) ? t : b).s); });
  // 3) عنوان الصفحة (فوق الجدول)
  const titleToks = toks.filter(t => t.y > hy + hh * 0.3 && !dayToks.includes(t) && /[\u0600-\u06FF]|\d+\s*[\/\\]\s*\d+/.test(t.s) && !/[a-z]{3,}/i.test(t.s)).sort((a, b) => b.h - a.h || b.y - a.y || b.x - a.x);
  const title = titleToks.map(t => t.s.replace(/\([^)]*\)/g, '')).join(' ').trim();
  // 4) عناوين الصفوف (خارج أعمدة الحصص)
  const body = toks.filter(t => t.y < hy - hh * 0.4);
  const side = body.filter(t => t.x > xMax || t.x < xMin);
  const sideR = side.filter(t => t.x > xMax), sideL = side.filter(t => t.x < xMin);
  let labels = (sideR.length >= sideL.length ? sideR : sideL);
  // اختر العمود الأقرب للجدول فقط (قد توجد أعمدة إضافية مثل رقم تسلسلي)
  labels = labels.filter(t => !/^\d{1,3}$/.test(t.s));
  const lineL = [];
  for (const t of labels.sort((a, b) => b.y - a.y)) {
    const r = lineL.find(r => Math.abs(r.y - t.y) < t.h * 0.6);
    if (r) r.s = sideR.includes(t) ? r.s + ' ' + t.s : t.s + ' ' + r.s; else lineL.push({ y: t.y, s: t.s, h: t.h });
  }
  // اسم مكتوب على عدة أسطر: نستعمل أسطر الخلايا كمراكز للصفوف ونضم كل سطر اسم لأقرب مركز
  let rowsL = lineL;
  const dayLabels = lineL.filter(l => dayIdx(l.s) >= 0).length >= lineL.length * 0.6;
  if (!dayLabels) {
    const anchors = [];
    body.filter(t => t.x <= xMax && t.x >= xMin).sort((a, b) => b.y - a.y).forEach(t => { const A = anchors.find(a => Math.abs(a.y - t.y) < t.h * 0.6); if (!A) anchors.push({ y: t.y }); });
    if (anchors.length && anchors.length < lineL.length) {
      const groups = anchors.map(a => ({ y: a.y, h: hh, parts: [] }));
      lineL.forEach(L => { const g = groups.reduce((b, c) => Math.abs(c.y - L.y) < Math.abs(b.y - L.y) ? c : b); g.parts.push(L); });
      rowsL = groups.filter(g => g.parts.length).map(g => ({ y: g.y, h: g.h, s: g.parts.sort((a, b) => b.y - a.y).map(p => p.s).join(' ') }));
    }
  }
  if (!rowsL.length) return { rows: [], reason: 'nolabels' };
  // 5) توزيع محتوى الخلايا على (صف، عمود)
  const cells = new Map();
  for (const t of body) {
    if (t.x > xMax || t.x < xMin) continue;
    const a = anchors.reduce((b, c) => Math.abs(c.x - t.x) < Math.abs(b.x - t.x) ? c : b);
    if (Math.abs(a.x - t.x) > colW * 0.65) continue;
    let ri = 0, best = Infinity;
    rowsL.forEach((r, i) => { const d = Math.abs(r.y - t.y); if (d < best) { best = d; ri = i; } });
    const k = ri + '|' + anchors.indexOf(a);
    (cells.get(k) || cells.set(k, []).get(k)).push(t);
  }
  // 6) معنى الصفوف: أيام أم صفوف دراسية أم معلمون
  const dayRows = rowsL.filter(r => dayIdx(r.s) >= 0).length >= rowsL.length * 0.6;
  const titleDay = dayIdx(title);
  let entity = '', entityIsClass = false;
  if (dayRows) {
    const m = toLatin(title).match(/\d+\s*[\/\\]\s*\d+/);
    const gw = gradeOf(title), sec = toLatin(title).match(/\d{1,2}/);
    if (gw && sec) { entity = gw + '/' + sec[0]; entityIsClass = true; }
    else if (m) { entity = m[0].replace(/\s+/g, '').replace(/\\/g, '/'); entityIsClass = true; }
    else {
      const tm = title.match(/(?:المعلم[ةه]?|الأستاذ[ةه]?|الاستاذ[ةه]?|أ\.)\s*[:：]?\s*([^\d:：]{3,40})/);
      entity = (tm ? tm[1] : title.replace(/جدول|حصص|المعلم[ةه]?|مدرس[ةه].*|العام.*|الفصل الدراسي.*/g, '')).trim().split(/\s{2,}/)[0];
      if (/صف|فصل/.test(title)) entityIsClass = true;
    }
  }
  const labelsAreClasses = !dayRows && rowsL.filter(r => isClassLike(r.s)).length >= rowsL.length * 0.5;
  const out = [];
  for (const [k, ts] of cells) {
    const [ri, ci] = k.split('|').map(Number);
    const text = ts.sort((a, b) => b.y - a.y).map(t => t.s).join('\n');
    const label = rowsL[ri].s, a = anchors[ci];
    const day = dayRows ? dayIdx(label) : (a.day ?? (titleDay >= 0 ? titleDay : -1));
    const row = { week: 0, day, period: a.raw, cls: '', subject: '', teacher: '' };
    const classOwner = dayRows ? entityIsClass : labelsAreClasses;
    const owner = dayRows ? entity : label;
    if (classOwner && dayRows) Object.assign(row, { cls: owner }, splitClassCell(ts));
    else if (classOwner) Object.assign(row, { cls: owner }, splitCell(text));
    else {
      const m = toLatin(text).match(/\d+\s*[\/\-\\]\s*\d+|\d+\s*[أ-ي](?=\s|$)/);
      row.teacher = owner;
      row.cls = m ? m[0].replace(/\s+/g, '').replace(/\\/g, '/') : text.split('\n')[0];
      row.subject = m ? toLatin(text).replace(m[0], '').replace(/[-–—،,\n]/g, ' ').replace(/\s+/g, ' ').trim() : text.split('\n').slice(1).join(' ');
    }
    if (row.cls || row.teacher) out.push(row);
  }
  return { rows: out, mode: dayRows ? (entityIsClass ? 'class-page' : 'teacher-page') : (labelsAreClasses ? 'classes' : 'teachers'), entity };
}
/* الصف مثل 6/4: رقم الصف (5، 6) قيمه قليلة ورقم الشعبة (1..8) قيمه كثيرة؛ إن انعكس الترتيب نصححه */
function fixClassOrder(rows) {
  const pairs = rows.map(r => normCls(r.cls).match(/^(\d{1,2})\/(\d{1,2})$/)).filter(Boolean);
  if (pairs.length < 4) return;
  const a = new Set(pairs.map(m => m[1])).size, b = new Set(pairs.map(m => m[2])).size;
  if (a > b) rows.forEach(r => { const m = normCls(r.cls).match(/^(\d{1,2})\/(\d{1,2})$/); if (m) r.cls = m[2] + '/' + m[1]; });
  else rows.forEach(r => { const m = normCls(r.cls).match(/^(\d{1,2})\/(\d{1,2})$/); if (m) r.cls = m[1] + '/' + m[2]; });
}
async function parseTimetablePDF(blob) {
  const pages = await pdfTextPages(blob);
  if (!pages.some(p => p.length)) throw new Error('هذا الـ PDF عبارة عن صورة ممسوحة ضوئياً ولا يحتوي نصاً يمكن قراءته. يمكنك عرضه كشريحة في الشاشة، أو تصديره من برنامج الجداول كـ PDF نصي أو Excel.');
  const rows = [], info = [];
  pages.forEach((p, i) => { const r = parseTimetablePDFPage(p, i + 1); rows.push(...r.rows); info.push(r); });
  fixClassOrder(rows);
  if (!rows.length) throw new Error('لم أستطع التعرف على شكل الجدول في الـ PDF. تأكد أن رؤوس الأعمدة فيها أرقام الحصص (1، 2، 3 أو الأولى، الثانية...). أرسل الملف للدعم لإضافة شكله.');
  return { rows, info };
}

/* ---------------------------------------------------------------------
   قارئ الروابط والملفات الذكي: Google Sheets/Docs/Drive، Excel، CSV، PDF،
   JSON (مثل Google Apps Script)، صفحات ويب فيها جداول، وأي صفحة نصية.
   --------------------------------------------------------------------- */
/* رابط بدون https:// → نضيفها تلقائياً */
const normURL = u => { u = String(u || '').trim(); return u && !/^[a-z][a-z0-9+.-]*:/i.test(u) ? 'https://' + u.replace(/^\/+/, '') : u; };
function fixLinkURL(url) {
  url = normURL(url);
  let m;
  if ((m = url.match(/docs\.google\.com\/spreadsheets\/d\/([\w-]+)/)) && !/output=|format=|tqx=/.test(url)) {
    if (/\/d\/e\//.test(url)) return url.replace(/\/pubhtml.*$/, '/pub?output=csv');
    const gid = (url.match(/[#&?]gid=(\d+)/) || [])[1];
    return `https://docs.google.com/spreadsheets/d/${m[1]}/gviz/tq?tqx=out:csv${gid ? '&gid=' + gid : ''}`;
  }
  if ((m = url.match(/docs\.google\.com\/document\/d\/([\w-]+)/)) && !/export/.test(url)) return `https://docs.google.com/document/d/${m[1]}/export?format=html`;
  if ((m = url.match(/drive\.google\.com\/(?:file\/d\/|open\?id=)([\w-]+)/))) return `https://drive.google.com/uc?export=download&id=${m[1]}`;
  return url;
}
function tableToAoa(t) {
  const grid = [];
  [...t.rows].forEach((tr, r) => {
    grid[r] = grid[r] || [];
    let c = 0;
    for (const td of tr.cells) {
      while (grid[r][c] !== undefined) c++;
      const txt = (td.innerText ?? td.textContent ?? '').replace(/[ \t]+/g, ' ').trim();
      const rs = Math.max(1, td.rowSpan || 1), cs = Math.max(1, td.colSpan || 1);
      for (let i = 0; i < rs; i++) for (let j = 0; j < cs; j++) { grid[r + i] = grid[r + i] || []; grid[r + i][c + j] = txt; }
      c += cs;
    }
  });
  return grid.map(r => Array.from(r, x => x ?? ''));
}
function textToAoa(text) {
  return String(text).split(/\r?\n/).map(l => l.trim()).filter(Boolean).map(l => l.split(/\t|\s{2,}|\s*\|\s*|\s+[:：]\s+/).map(x => x.trim()));
}
function docToSheets(doc, allowText = true) {
  const sheets = [...doc.querySelectorAll('table')].map(t => ({ name: t.caption?.textContent || '', aoa: tableToAoa(t) })).filter(s => s.aoa.length > 1);
  if (!sheets.length && allowText) {
    const text = doc.body ? (doc.body.innerText || doc.body.textContent || '') : '';
    if (text.trim()) sheets.push({ name: '', aoa: textToAoa(text), text: true });
  }
  return sheets;
}
function htmlToSheets(html) {
  // صفحات Google Apps Script (HtmlService) تضع المحتوى داخل نص مشفّر
  const m = html.match(/"userHtml":"((?:[^"\\]|\\.)*)"/);
  if (m) { try { html = JSON.parse('"' + m[1].replace(/\\x([0-9a-f]{2})/gi, '\\u00$1') + '"'); } catch { } }
  return docToSheets(new DOMParser().parseFromString(html, 'text/html'));
}
function jsonToAoa(j) {
  if (Array.isArray(j)) {
    if (!j.length) return [];
    if (Array.isArray(j[0])) return j.map(r => r.map(c => c ?? ''));
    if (typeof j[0] === 'object') {
      const keys = [...new Set(j.flatMap(o => Object.keys(o || {})))];
      return [keys, ...j.map(o => keys.map(k => o?.[k] ?? ''))];
    }
    return j.map(x => [String(x)]);
  }
  if (j && typeof j === 'object') {
    for (const k of ['values', 'data', 'rows', 'items', 'records', 'result', 'table']) if (j[k]) return jsonToAoa(j[k]);
    const arr = Object.values(j).find(Array.isArray); if (arr) return jsonToAoa(arr);
    return Object.entries(j).map(([k, v]) => [k, typeof v === 'object' ? JSON.stringify(v) : v]);
  }
  return [];
}
/* PDF عام (مناوبة/احتياط/غياب): أعمدة حسب سطر العناوين */
async function pdfToSheets(blob) {
  const pages = await pdfTextPages(blob);
  if (!pages.some(p => p.length)) throw new Error('هذا الـ PDF صورة ممسوحة ولا يحتوي نصاً يمكن قراءته.');
  const HEAD = /يوم|حص|صف|فصل|شعب|معلم|مناوب|غائب|غياب|احتياط|بديل|مكان|موقع|تاريخ|فتر|اسم|طالب|سبب|ملاحظ|اسبوع/;
  return pages.map((raw, pi) => {
    const toks = mergeLine(raw);
    const lines = [];
    toks.sort((a, b) => b.y - a.y).forEach(t => { const L = lines.find(l => Math.abs(l.y - t.y) < t.h * 0.5); if (L) L.t.push(t); else lines.push({ y: t.y, t: [t] }); });
    let hi = lines.findIndex(l => l.t.filter(t => HEAD.test(norm(t.s)) || dayIdx(t.s) >= 0).length >= 2);
    if (hi < 0) hi = lines.reduce((b, l, i) => l.t.length > lines[b].t.length ? i : b, 0);
    const anchors = lines[hi].t.slice().sort((a, b) => b.x - a.x);
    const aoa = [anchors.map(a => a.s)];
    for (const L of lines.slice(hi + 1)) {
      const row = anchors.map(() => '');
      for (const t of L.t) {
        let k = 0; anchors.forEach((a, i) => { if (Math.abs(a.x - t.x) < Math.abs(anchors[k].x - t.x)) k = i; });
        row[k] = row[k] ? row[k] + ' ' + t.s : t.s;
      }
      aoa.push(row);
    }
    return { name: 'صفحة ' + (pi + 1), aoa };
  });
}
/* قراءة الصفحات التي تُبنى بالجافاسكربت (مثل Canva) داخل إطار مخفي */
function readLinkIframe(url) {
  return new Promise((res, rej) => {
    const f = h('iframe', { src: url, style: 'position:fixed;left:-10000px;top:0;width:1400px;height:1000px;border:0' });
    let done = false;
    const finish = fn => { if (done) return; done = true; setTimeout(() => f.remove(), 50); fn(); };
    const grab = () => {
      try {
        const walk = (w, acc) => { acc.push(...docToSheets(w.document, false)); for (let i = 0; i < w.frames.length; i++) { try { walk(w.frames[i], acc); } catch { } } return acc; };
        const tables = walk(f.contentWindow, []);
        if (tables.length) return finish(() => res(tables));
        const texts = []; const tw = w => { texts.push(w.document.body?.innerText || ''); for (let i = 0; i < w.frames.length; i++) { try { tw(w.frames[i]); } catch { } } };
        tw(f.contentWindow);
        const text = texts.join('\n').trim();
        if (!text) throw new Error('فارغ');
        finish(() => res([{ name: '', aoa: textToAoa(text), text: true }]));
      } catch (e) { finish(() => rej(new Error('blocked'))); }
    };
    f.onload = () => setTimeout(grab, 6000);
    setTimeout(() => finish(() => rej(new Error('timeout'))), 30000);
    document.body.append(f);
  });
}
async function smartFromBlob(blob, name = '') {
  const head = new Uint8Array(await blob.slice(0, 8).arrayBuffer());
  const n = String(name).toLowerCase();
  if (String.fromCharCode(...head.slice(0, 4)) === '%PDF' || n.endsWith('.pdf')) return { type: 'pdf', blob };
  if ((head[0] === 0x50 && head[1] === 0x4B) || (head[0] === 0xD0 && head[1] === 0xCF)) return { type: 'sheets', sheets: await sheetsFromBlob(blob) };
  const text = new TextDecoder('utf-8').decode(await blob.arrayBuffer()).replace(/^﻿/, '');
  const t = text.trim();
  if (/^[\[{]/.test(t)) { try { return { type: 'sheets', sheets: [{ name: '', aoa: jsonToAoa(JSON.parse(t)) }] }; } catch { } }
  if (/^</.test(t) || /<html|<table|<body/i.test(t.slice(0, 3000))) return { type: 'html', sheets: htmlToSheets(text) };
  return { type: 'sheets', sheets: [{ name: '', aoa: parseCSV(text) }] };
}
/* Google Sheets: البرنامج يعمل من ملف على القرص، وجوجل لا يسمح بقراءته بالطريقة العادية من هناك،
   لذلك نقرأه بطريقة JSONP (تحميل البيانات كسكربت) التي لا تحتاج إذناً خاصاً */
let jsonpN = 0;
function gsheetJSONP(id, { gid, sheet } = {}) {
  return new Promise((res, rej) => {
    const cb = '__gs' + (++jsonpN) + Date.now();
    const sc = document.createElement('script');
    const done = (fn, v) => { clearTimeout(tm); delete window[cb]; sc.remove(); fn(v); };
    const tm = setTimeout(() => done(rej, new Error('انتهت مهلة قراءة Google Sheets — تحقق من الإنترنت ومن مشاركة الملف "لأي شخص لديه الرابط"')), 30000);
    window[cb] = data => {
      if (!data || data.status !== 'ok') return done(rej, new Error('Google Sheets: ' + (data?.errors?.[0]?.detailed_message || data?.errors?.[0]?.message || 'تعذر القراءة — تأكد من مشاركة الملف')));
      const t = data.table, cell = c => c == null ? '' : (c.f != null ? c.f : (c.v == null ? '' : String(c.v)));
      const head = t.cols.map(c => c.label || '');
      const rows = t.rows.map(r => (r.c || []).map(cell));
      done(res, head.some(Boolean) ? [head, ...rows] : rows);
    };
    sc.onerror = () => done(rej, new Error('تعذر الاتصال بـ Google Sheets — تحقق من الإنترنت ومن مشاركة الملف'));
    sc.src = `https://docs.google.com/spreadsheets/d/${id}/gviz/tq?tqx=out:json;responseHandler:${cb}&headers=1` + (gid ? '&gid=' + gid : '') + (sheet ? '&sheet=' + encodeURIComponent(sheet) : '') + '&_=' + Date.now();
    document.head.append(sc);
  });
}
function gsheetRef(url) {
  const m = normURL(url).match(/docs\.google\.com\/spreadsheets\/d\/([\w-]+)/);
  if (!m || m[1] === 'e') return null;
  return { id: m[1], gid: (url.match(/[#&?]gid=(\d+)/) || [])[1] };
}
async function smartRead(url) {
  url = normURL(url);
  const gs = gsheetRef(url);
  if (gs) {
    try { return { type: 'sheets', sheets: [{ name: '', aoa: await gsheetJSONP(gs.id, { gid: gs.gid }) }], via: 'gsheet', gs }; }
    catch (e) { console.warn(e); /* نجرب الطرق الأخرى */ }
  }
  let res = null, err = null;
  try {
    const r = await fetch(fixLinkURL(url), { cache: 'no-store', redirect: 'follow' });
    if (!r.ok) throw new Error('HTTP ' + r.status);
    res = await smartFromBlob(await r.blob(), url.split('?')[0]);
  } catch (e) { err = e; }
  // صفحة تُبنى بالجافاسكربت أو موقع يمنع القراءة المباشرة → جرّب الإطار المخفي
  if (!res || (res.type === 'html' && !res.sheets.some(s => !s.text))) {
    try { return { type: 'sheets', sheets: await readLinkIframe(url), via: 'iframe' }; }
    catch { if (res && res.sheets.length) return res; }
    throw new Error('تعذر قراءة محتوى الرابط' + (err ? ' (' + err.message + ')' : '') + '. بعض المواقع (مثل Canva) تمنع قراءة بياناتها — فعّل خيار "عرض الرابط كمشهد في الشاشة"، أو استخدم رابط Google Sheets أو Apps Script يُرجع البيانات.');
  }
  return res;
}
/* تحويل ما قُرئ إلى سجلات حسب نوع البيانات */
async function rowsFromSmart(which, res, sheetName) {
  if (which === 'timetable') return res.type === 'pdf' ? (await parseTimetablePDF(res.blob)).rows : parseTimetable(res.sheets, sheetName);
  let sheets = res.type === 'pdf' ? await pdfToSheets(res.blob) : res.sheets;
  if (sheetName) { const f = sheets.filter(s => norm(s.name).includes(norm(sheetName))); if (f.length) sheets = f; }
  const fn = PARSERS[which];
  // ملف بعدة أوراق: فضّل الورقة الأساسية (Assignments / الاحتياط) وتجاهل النسخ الاحتياطية والسجلات
  const pref = { subs: /^assignments$|^الاحتياط$|احتياط/i, duty: /مناوب/, absent: /غياب|غائب/ }[which];
  const usable = sheets.filter(s => !/backup|audit|log|session|admin|سجل|ارشيف|أرشيف/i.test(s.name || ''));
  const ordered = [...usable.filter(s => pref && pref.test(s.name || '')), ...usable.filter(s => !(pref && pref.test(s.name || ''))), ...sheets.filter(s => !usable.includes(s))];
  let best = [], lastErr = null;
  for (const s of ordered) { try { const r = fn(s.aoa); if (r.length > best.length) best = r; if (best.length && pref && pref.test(s.name || '')) break; } catch (e) { lastErr = e; } }
  // رابط Google Sheets: قد يشير الرابط إلى ورقة أخرى (تعليمات أو نسخة احتياطية قديمة لا تتحدث)
  // لذلك نقرأ أيضاً الورقة الأساسية بالاسم ونأخذ الأحدث (آخر تاريخ) ثم الأكثر سجلات
  if (res.via === 'gsheet' && which !== 'timetable') {
    const maxDate = rs => rs.reduce((m, r) => r.date && r.date > m ? r.date : m, '');
    for (const name of ({ subs: ['Assignments', 'الاحتياط'], duty: ['المناوبة', 'Duty'], absent: ['الغياب', 'Absent'] }[which] || [])) {
      try {
        const r = fn(await gsheetJSONP(res.gs.id, { sheet: name }));
        if (r.length && (maxDate(r) > maxDate(best) || (maxDate(r) === maxDate(best) && r.length > best.length))) best = r;
        if (r.length) break;
      } catch (e) { lastErr = e; }
    }
  }
  if (!best.length) throw lastErr || new Error('لم يُعثر على بيانات مناسبة في المصدر.');
  // أسماء المعلمين من ورقة Teachers في نفس الملف (عند وجود رموز بدل الأسماء)
  if (which === 'subs' && best.some(r => isIdCode(r.absent) || isIdCode(r.sub))) {
    const t = sheets.find(s => /teachers|المعلم/i.test(s.name || ''));
    if (t) applyNameMap(best, nameMapFromAoa(t.aoa));
  }
  return best;
}

/* ---------- الطلاب الغائبون ---------- */
function parseAbsent(aoa) {
  aoa = aoa.filter(r => r.some(c => String(c).trim()));
  const hr = findHeader(aoa, [/اسم|طالب/, /صف|فصل|شعب/, /تاريخ/, /سبب|ملاحظ|حاله/, /يوم/]);
  if (hr < 0) throw new Error('لم أتعرف على أعمدة جدول الغياب. يجب أن يحتوي على: اسم الطالب، الصف، والتاريخ (اختياري)، والسبب (اختياري).');
  const hd = aoa[hr].map(norm), col = (re, not) => hd.findIndex(c => re.test(c) && !(not && not.test(c)));
  const c = { name: col(/اسم|طالب/, /صف|فصل/), cls: col(/صف|فصل|شعب/), date: col(/تاريخ/), day: col(/يوم/, /تاريخ/), reason: col(/سبب|ملاحظ|حاله/), status: col(/حضور|غياب|الحاله/) };
  let dt = '', dy = -1, cl = '';
  const out = [];
  for (const row of aoa.slice(hr + 1)) {
    if (cell(row, c.date)) dt = parseDate(cell(row, c.date));
    if (cell(row, c.day)) dy = dayIdx(cell(row, c.day));
    if (cell(row, c.cls)) cl = cell(row, c.cls);
    const name = cell(row, c.name); if (!name) continue;
    if (c.status >= 0 && cell(row, c.status) && /حاضر|present/.test(norm(cell(row, c.status)))) continue;
    out.push({ date: dt, day: dt ? -1 : dy, name, cls: cl, reason: cell(row, c.reason) });
  }
  return out;
}
function absentFor(d) {
  const day = ymd(d), dow = d.getDay();
  const cmp = clsSorter();
  return CFG.absent.rows.filter(r => r.date ? r.date === day : (r.day === -1 ? !CFG.absent.datedOnly : r.day === dow)).sort((a, b) => cmp(a.cls, b.cls) || a.name.localeCompare(b.name, 'ar'));
}
const PARSERS = { duty: a => parseDuty(a), subs: a => parseSubs(a), absent: a => parseAbsent(a) };

async function refreshSource(which, manual) {
  const src = CFG[which];
  if (src.source !== 'link' || !src.url) return;
  try {
    src.linkRows = await rowsFromSmart(which, await smartRead(src.url), src.sheet);
    if (which === 'subs') await resolveTeacherIds(src.linkRows, src);
    src.rows = src.linkRows.concat(src.fileRows || []);
    src.lastFetch = new Date().toLocaleString('ar'); src.lastError = '';
    const last = src.linkRows.reduce((m, r) => r.date && r.date > m ? r.date : m, '');
    if (manual) toast(`تم التحديث: ${src.rows.length} سجل` + (last ? ' — آخر تاريخ: ' + last : ''));
  } catch (e) {
    src.lastError = e.message || String(e);
    if (manual) toast('فشل التحديث: ' + src.lastError, true);
  }
  CFG._rev = (CFG._rev || 0) + 1;
  save();
}
function refreshAllSources() { shiftIds().forEach(id => inShift(id, () => ['duty', 'subs', 'timetable', 'absent'].forEach(w => refreshSource(w)))); }
const lastRefresh = {};
function maybeRefresh() {
  for (const id of shiftIds()) inShift(id, () => {
    for (const w of ['duty', 'subs', 'timetable', 'absent']) {
      const s = CFG[w], k = id + ':' + w;
      if (s.source === 'link' && s.url && Date.now() - (lastRefresh[k] || 0) > Math.max(1, +s.refreshMin || 10) * 60000) {
        lastRefresh[k] = Date.now(); refreshSource(w);
      }
    }
  });
}

/* ------------------------------ تحديث عام ------------------------------ */
let slideRebuildT;
function onConfigChanged() {
  if (typeof DutyStrip !== 'undefined') DutyStrip.sig = '';
  Phrases.renderCard && setTimeout(() => Phrases.renderCard(), 0);
  applyTheme();
  CFG._rev = (CFG._rev || 0) + 1;
  renderDates(); renderClock();
  renderSide(true); updateTicker(true);
}
function applyTheme() {
  const g = CFG.general, D = CFG.display, B = document.body;
  B.dataset.theme = g.theme || 'oman2040';
  [['--accent', 'accent'], ['--accent2', 'accent2'], ['--accent3', 'accent3']].forEach(([v, k]) => {
    if (g.custom?.on && g.custom[k]) B.style.setProperty(v, g.custom[k]); else B.style.removeProperty(v);
  });
  B.classList.toggle('no-anim', !g.animBg);
  B.classList.toggle('cast', !!D.cast);
  $('#schoolName').textContent = g.schoolName; $('#schoolSub').textContent = g.schoolSub;
  if (g.logo) fileURL(g.logo).then(u => { if (u) { $('#logo').src = u; $('#logo').hidden = false; $('#logoPh').hidden = true; } });
  else { $('#logo').hidden = true; $('#logoPh').hidden = false; }
  $('#display').classList.toggle('ticker-top', D.tickerPos === 'top');
  $('#statusbar').hidden = !D.showStatus;
  $('#content').classList.toggle('with-side', D.layout === 'side');
  $('#dutyCard').hidden = !D.showDuty; $('#subsCard').hidden = !D.showSubs;
  $('#mutedFlag').hidden = g.bellEnabled; $('#cMute').classList.toggle('off', !g.bellEnabled);
  $('#cMute').textContent = g.bellEnabled ? '🔊' : '🔇';
}

/* ---------------------------------------------------------------------
   قراءة التوقيت من صورة (OCR) — يعمل بدون إنترنت، يُحمَّل عند الحاجة فقط
   --------------------------------------------------------------------- */
const OCR = {
  loaded: false,
  loadScript(src) {
    return new Promise((res, rej) => { const s = document.createElement('script'); s.src = src; s.onload = res; s.onerror = () => rej(new Error('تعذر تحميل ملف القراءة ' + src)); document.head.append(s); });
  },
  async read(file, progress) {
    if (!this.loaded) {
      progress && progress('تحميل محرك القراءة...');
      for (const f of ['tesseract.min.js', 'OCR_WORKER.js', 'OCR_CORE.js', 'OCR_ARA.js', 'OCR_ENG.js']) await this.loadScript('lib/ocr/' + f);
      this.loaded = true;
    }
    const blobURL = t => URL.createObjectURL(new Blob([t], { type: 'application/javascript' }));
    // بيانات اللغتين تُمرَّر للعامل مباشرة (بدل تحميلها من الإنترنت)
    const prelude = `self.__OCRL={ara:"${OCR_ARA}",eng:"${OCR_ENG}"};const __f=self.fetch;self.fetch=function(u,o){const m=String(u).match(/(ara|eng)\\.traineddata/);if(m){const b=Uint8Array.from(atob(self.__OCRL[m[1]]),c=>c.charCodeAt(0));return Promise.resolve(new Response(b));}return __f.call(self,u,o);};\n`;
    const worker = await Tesseract.createWorker('ara+eng', 1, {
      workerPath: blobURL(prelude + OCR_CORE + '\n;self.TesseractCore = self.TesseractCore || TesseractCore;\n' + OCR_WORKER),
      langPath: 'https://local.ocr', workerBlobURL: false, cacheMethod: 'none', gzip: true,
      logger: m => progress && m.status === 'recognizing text' && progress('جارٍ القراءة ' + Math.round(m.progress * 100) + '%'),
    });
    try {
      // تجهيز الصورة: تكبير، رمادي، عكس الخلفية الداكنة، رفع التباين
      const img = await createImageBitmap(file);
      const k = Math.max(1, Math.min(3, 2200 / img.width));
      const cv = document.createElement('canvas'); cv.width = Math.round(img.width * k); cv.height = Math.round(img.height * k);
      const cx = cv.getContext('2d', { willReadFrequently: true }); cx.fillStyle = '#fff'; cx.fillRect(0, 0, cv.width, cv.height); cx.drawImage(img, 0, 0, cv.width, cv.height);
      const id = cx.getImageData(0, 0, cv.width, cv.height), px = id.data;
      let sum = 0; for (let i = 0; i < px.length; i += 4) sum += px[i] * 0.3 + px[i + 1] * 0.59 + px[i + 2] * 0.11;
      const dark = sum / (px.length / 4) < 128;
      for (let i = 0; i < px.length; i += 4) {
        let g = px[i] * 0.3 + px[i + 1] * 0.59 + px[i + 2] * 0.11;
        if (dark) g = 255 - g;
        g = g < 150 ? Math.max(0, g - 60) : 255;
        px[i] = px[i + 1] = px[i + 2] = g;
      }
      cx.putImageData(id, 0, 0);
      // جدول بخطوط؟ اقرأ كل خلية لوحدها (أدق بكثير)
      const grid = detectGrid(px, cv.width, cv.height);
      if (grid) {
        progress && progress('تم اكتشاف جدول — قراءة الخلايا...');
        await worker.setParameters({ tessedit_pageseg_mode: '7' });
        const lines = [];
        for (let r = 0; r < grid.rows.length - 1; r++) {
          const cells = [];
          for (let c = grid.cols.length - 2; c >= 0; c--) {
            const left = grid.cols[c] + 4, top = grid.rows[r] + 4, width = grid.cols[c + 1] - grid.cols[c] - 8, height = grid.rows[r + 1] - grid.rows[r] - 8;
            if (width < 12 || height < 12) continue;
            const { data } = await worker.recognize(cv, { rectangle: { left, top, width, height } });
            cells.push(data.text.replace(/\s+/g, ' ').trim());
          }
          progress && progress(`قراءة الجدول ${Math.round((r + 1) / (grid.rows.length - 1) * 100)}%`);
          if (cells.some(Boolean)) lines.push({ text: cells.join(' | ') });
        }
        if (lines.length) return { text: lines.map(l => l.text).join('\n'), lines };
      }
      // تجربة أكثر من طريقة لقراءة الصفحة واختيار الأفضل (الأكثر أوقاتاً)
      const countTimes = d => (toLatin(d.text).match(/\d{1,2}\s*[:.]\s*\d{2}/g) || []).length;
      let best = null;
      for (const psm of ['6', '4', '11']) {
        progress && progress('جارٍ القراءة (طريقة ' + psm + ')...');
        await worker.setParameters({ tessedit_pageseg_mode: psm, preserve_interword_spaces: '1' });
        const { data } = await worker.recognize(cv);
        if (!best || countTimes(data) > countTimes(best)) best = data;
      }
      return best;
    } finally { await worker.terminate(); }
  },
};
/* اكتشاف خطوط الجدول (أفقية وعمودية) من الصورة بعد تحويلها للأبيض والأسود */
function detectGrid(px, W, H) {
  const dark = (x, y) => px[(y * W + x) * 4] < 128;
  const lines = (len, other, isRow) => {
    const out = [];
    for (let i = 0; i < len; i++) {
      let run = 0, best = 0;
      for (let j = 0; j < other; j++) { if (isRow ? dark(j, i) : dark(i, j)) { run++; if (run > best) best = run; } else run = 0; }
      if (best > other * 0.45) { if (out.length && i - out[out.length - 1] <= 4) out[out.length - 1] = i; else out.push(i); }
    }
    return out;
  };
  const rows = lines(H, W, true), cols = lines(W, H, false);
  if (rows.length < 3 || cols.length < 2) return null;
  return { rows, cols };
}
function typeFromLabel(label) {
  const k = slotKey(label);
  if (TYPE[k] && k !== 'custom') return { type: k, label: '' };
  return { type: 'custom', label: label.replace(/بدايه|بداية|نهاية|نهايه|بدء|انتهاء/g, '').trim() || 'فترة' };
}
/* تحويل نص الصورة إلى فترات (من/إلى) — يدعم سطر فيه وقتان، أو أحداث "بداية/نهاية" بوقت واحد */
function parseScheduleText(lines) {
  const TRE = /(\d{1,2})\s*[:.：٫؛;]\s*(\d{2})/g;
  const toT = (h, m) => { h = +h; m = +m; if (h > 23 || m > 59) return null; if (h >= 1 && h <= 5) h += 12; return pad(h) + ':' + pad(m); };
  const items = [];
  for (const raw of lines) {
    // أرقام بدون نقطتين داخل خلية (مثل 1225 = 12:25)
    const fixed = toLatin(raw).split('|').map(c => { const t = c.trim(); const m = t.match(/^(\d{1,2})(\d{2})$/); return m && +m[1] <= 14 && +m[2] < 60 ? m[1] + ':' + m[2] : c; }).join('|');
    const line = fixed.replace(/[|«»"]/g, ' ');
    const times = [...line.matchAll(TRE)].map(m => toT(m[1], m[2])).filter(Boolean);
    if (!times.length) continue;
    const label = line.replace(TRE, ' ').replace(/\b(ص|م|am|pm)\b/gi, ' ').replace(/[-–—:()\[\]ـ_.,،]/g, ' ').replace(/\s+/g, ' ').trim();
    items.push({ label, times: times.sort((a, b) => toSec(a) - toSec(b)) });
  }
  const periods = [];
  const twin = items.filter(it => it.times.length >= 2);
  if (twin.length >= items.length / 2) {
    for (const it of twin) periods.push({ ...typeFromLabel(it.label), start: it.times[0], end: it.times[it.times.length - 1] });
  } else {
    // أحداث بوقت واحد: "بداية الحصة الأولى 7:30" ، "نهاية الحصة الأولى 8:10"
    const evs = items.map(it => ({ ...it, t: it.times[0], isEnd: /نهايه|نهاية|انتهاء|انتهت/.test(it.label) })).sort((a, b) => toSec(a.t) - toSec(b.t));
    evs.forEach((e, i) => {
      if (e.isEnd) {
        const key = slotKey(e.label.replace(/نهايه|نهاية|انتهاء|انتهت/g, ''));
        const p = [...periods].reverse().find(p => !p.end && (p.type === key || slotKey(p.label) === key)) || [...periods].reverse().find(p => !p.end);
        if (p) p.end = e.t;
        return;
      }
      periods.push({ ...typeFromLabel(e.label.replace(/بدايه|بداية|بدء/g, '')), start: e.t, end: '' });
    });
    periods.forEach((p, i) => { if (!p.end) p.end = periods[i + 1]?.start || fromSec(toSec(p.start) + 40 * 60); });
  }
  const out = periods.filter(p => toSec(p.end) > toSec(p.start)).sort((a, b) => toSec(a.start) - toSec(b.start));
  // اسم غير مقروء: استنتجه من الترتيب (بعد الحصة الثالثة تأتي الرابعة...)
  let last = 0;
  for (const p of out) {
    const m = p.type.match(/^p(\d)$/);
    if (m) { last = +m[1]; continue; }
    if (p.type === 'custom' && (p.label === 'فترة' || !/[\u0621-\u064A]{3,}/.test(p.label)) && last < 8) { last++; p.type = 'p' + last; p.label = ''; }
  }
  // فجوة واضحة بين فترتين = فسحة لم تُقرأ
  for (let i = out.length - 1; i > 0; i--) {
    const gap = toSec(out[i].start) - toSec(out[i - 1].end);
    if (gap >= 10 * 60 && gap <= 60 * 60) out.splice(i, 0, { type: out.some(x => x.type === 'break') ? 'break2' : 'break', label: '', start: out[i - 1].end, end: out[i].start });
  }
  return out;
}
async function scheduleFromImage(pr) {
  const f = await pickFile('image/*'); if (!f) return;
  const st = h('div', { class: 'toast' }, '⏳ تحميل محرك القراءة...'); document.body.append(st);
  try {
    const data = await OCR.read(f, m => { st.textContent = '⏳ ' + m; });
    const lines = (data.lines || []).map(l => l.text).filter(Boolean);
    const found = parseScheduleText(lines.length ? lines : String(data.text).split('\n'));
    st.remove();
    if (!found.length) { alert('لم أجد أوقاتاً واضحة في الصورة.\n\nالنص المقروء:\n' + data.text.slice(0, 800)); return; }
    const list = found.map(p => `• ${p.type === 'custom' ? p.label : TYPE[p.type].n}: ${p.start} — ${p.end}`).join('\n');
    if (!confirm(`تمت قراءة ${found.length} فترة من الصورة:\n\n${list}\n\nاستبدال فترات "${pr.name}" بها؟ (يمكنك مراجعتها وتعديلها بعد ذلك)`)) return;
    pr.periods = found.map(p => P(p.type, p.start, p.end, { label: p.label, ...(p.type.startsWith('break') ? { bellStart: 'triple', bellEnd: 'classic' } : {}) }));
    save(); rerender(); toast('تم ملء التوقيت من الصورة — راجعه');
  } catch (e) { st.remove(); toast('تعذرت القراءة: ' + e.message, true); }
}

/* =====================================================================
                                 الإعدادات
   ===================================================================== */
let curTab = 'general', curProfile = null;
function openSettings(tab) {
  const vh = $('#appVer'); if (vh) vh.textContent = 'مِجْداف • الإصدار ' + APP_VERSION + (window.VOICE_SETS ? ' • الأصوات: ' + Object.keys(VOICE_SETS).length : ' • ⚠️ ملف الأصوات غير موجود');
  EDIT_SHIFT = CFG.shiftData[CUR_SHIFT] ? CUR_SHIFT : 's1'; CUR_SHIFT = EDIT_SHIFT;
  $('#settings').hidden = false; document.body.classList.add('mouse');
  showTab(tab || curTab);
}
function closeSettings() { $('#settings').hidden = true; syncShift(); Slides.rebuild(); onConfigChanged(); }
function showTab(tab) {
  curTab = tab;
  $$('#tabs button').forEach(b => b.classList.toggle('on', b.dataset.tab === tab));
  const body = $('#setBody'); const top = body.scrollTop; body.innerHTML = '';
  const sec = h('section'); body.append(sec);
  // اختيار الفترة التي تُعدَّل بياناتها
  if (CFG.shifts.enabled && ['schedule', 'timetable', 'duty', 'subs', 'absent'].includes(tab))
    sec.append(h('div', { class: 'shift-bar' }, h('span', {}, '🏫 تعديل بيانات:'),
      shiftIds().map(id => h('button', { class: 'btn' + (id === EDIT_SHIFT ? ' pri' : ''), onclick: () => { EDIT_SHIFT = CUR_SHIFT = id; curProfile = null; ttDay = null; rerender(); } }, (id === 's1' ? '☀️ ' : '🌇 ') + shiftName(id)))));
  TABS[tab](sec);
  body.scrollTop = top;
}
const rerender = () => showTab(curTab);

/* عناصر نموذج مرتبطة بالإعدادات */
function bindInput(obj, key, attrs = {}, after) {
  const type = attrs.type || 'text';
  const el = h('input', { ...attrs, type, value: obj[key] ?? '' });
  if (type === 'checkbox') el.checked = !!obj[key];
  el.addEventListener('change', () => {
    obj[key] = type === 'checkbox' ? el.checked : (type === 'number' || type === 'range') ? +el.value : el.value;
    save(); after && after();
  });
  return el;
}
function bindSelect(obj, key, options, after, attrs = {}) {
  const el = h('select', attrs, options.map(([v, t]) => h('option', { value: v }, t)));
  el.value = obj[key] ?? '';
  el.addEventListener('change', () => { obj[key] = el.value; save(); after && after(); });
  return el;
}
const field = (label, el, hint) => h('label', { class: 'field' }, h('span', {}, label), el, hint ? h('small', { class: 'hint' }, hint) : null);
const chk = (obj, key, label, after) => h('label', { class: 'chk' }, bindInput(obj, key, { type: 'checkbox' }, after), label);
const panel = (title, ...kids) => h('div', { class: 'panel' }, title ? h('h4', {}, title) : null, ...kids);
function soundOptions(withNone = true, withTTSOnly) {
  const o = [];
  if (withNone) o.push(['none', '— بدون —']);
  for (const [k, b] of Object.entries(Audio_.bells)) o.push([k, '🔔 ' + b.n]);
  for (const f of CFG.files.filter(f => f.kind === 'audio')) o.push([f.id, '🎵 ' + f.name]);
  return o;
}
/* اختيار صوت واحد لكل حدث: بدون / إعلان فقط / نغمة فقط / نغمة ثم إعلان / ملف مرفوع
   k = {bell, ann, audio} أسماء الحقول في الكائن (ann غير موجود = نعم) */
function soundModeOf(o, k) {
  const bell = o[k.bell] || 'none', ann = o[k.ann] !== false, audio = o[k.audio] || '';
  if (audio && ann) return 'file:' + audio;
  if (bell !== 'none' && !Audio_.bells[bell]) return CFG.files.some(f => f.id === bell) ? 'file:' + bell : (ann ? 'ann' : 'none');
  if (bell === 'none') return ann ? 'ann' : 'none';
  return (ann ? 'both:' : 'bell:') + bell;
}
function setSoundMode(o, k, v) {
  const i = v.indexOf(':'), m = i < 0 ? v : v.slice(0, i), x = i < 0 ? '' : v.slice(i + 1);
  o[k.bell] = (m === 'bell' || m === 'both') ? x : 'none';
  o[k.ann] = m === 'ann' || m === 'both' || m === 'file';
  o[k.audio] = m === 'file' ? x : '';
}
function soundModeSelect(o, k, annLabel, attrs = {}) {
  const opt = (v, t) => h('option', { value: v }, t);
  const grp = (label, kids) => { const g = document.createElement('optgroup'); g.label = label; g.append(...kids); return g; };
  const bells = Object.entries(Audio_.bells), files = CFG.files.filter(f => f.kind === 'audio');
  const el = h('select', { style: 'width:230px', ...attrs },
    opt('none', '🔕 بدون صوت'),
    opt('ann', '🗣️ ' + annLabel + ' فقط'),
    grp('نغمة ثم ' + annLabel, bells.map(([b, x]) => opt('both:' + b, '🔔 ' + x.n + ' ← 🗣️ ' + annLabel))),
    grp('نغمة فقط', bells.map(([b, x]) => opt('bell:' + b, '🔔 ' + x.n + ' فقط'))),
    files.length ? grp('ملف صوتي مرفوع', files.map(f => opt('file:' + f.id, '🎵 ' + f.name))) : null);
  el.value = soundModeOf(o, k);
  el.addEventListener('change', () => { setSoundMode(o, k, el.value); save(); });
  return el;
}
const KS = { bell: 'bellStart', ann: 'annStart', audio: 'startAudio' }, KE = { bell: 'bellEnd', ann: 'annEnd', audio: 'endAudio' }, KR = { bell: 'sound', ann: 'speak', audio: 'audio' };
const playPeriod = (p, end) => { const k = end ? KE : KS, t = end ? endText(p) : startText(p); Queue.push({ sound: p[k.bell], parts: p[k.ann] ? [{ text: t, audio: p[k.audio] }] : [], banner: t }); };
function audioOptions() {
  return [['', '🗣️ نطق آلي للنص'], ...CFG.files.filter(f => f.kind === 'audio').map(f => [f.id, '🎵 ' + f.name])];
}
const testBtn = (fn, label = '▶') => h('button', { class: 'btn sm', title: 'تجربة', onclick: e => { e.preventDefault(); Audio_.ensure(); fn(); } }, label);

/* محتوى الشاشة أثناء كل فترة */
const CONTENT_OPTS = [['', 'تلقائي'], ['rotate', 'المشاهد المتبدلة'], ['now', '📚 الحصة الآن (الصفوف والمعلمين)'], ['classday', '🗓️ جدول حصص اليوم'], ['duty', '🛡️ المناوبة'], ['subs', '🔁 الاحتياط'], ['absent', '📋 الغياب'], ['schedule', '⏰ التوقيت'], ['welcome', '🏫 الترحيب'], ['phrase', '🤲 أذكار وتحفيز'], ['media', '🎬 الصور والفيديو فقط']];

/* منتقي وقت سهل بدون كتابة */
function timeBtn(obj, key, title) {
  const show = () => obj[key] ? digits(fmt12(obj[key])) + (CFG.general.clock24 ? '' : (toSec(obj[key]) < 12 * 3600 ? ' ص' : ' م')) : '--:--';
  return h('button', { class: 'btn time-btn', onclick: () => timePicker(obj[key], title, v => { obj[key] = v; save(); rerender(); }) }, '🕐 ' + show());
}
function timePicker(val, title, done) {
  let sec = toSec(val || '07:00') ?? 7 * 3600;
  const box = h('div', { class: 'tp-box' }), ov = h('div', { class: 'tp-ov', onclick: e => { if (e.target === ov) ov.remove(); } }, box);
  const draw = () => {
    const H = Math.floor(sec / 3600), M = Math.floor(sec / 60) % 60, pm = H >= 12, h12 = H % 12 || 12;
    const set = (hh, mm, p = pm) => { hh = hh % 12 + (p ? 12 : 0); sec = hh * 3600 + mm * 60; draw(); };
    box.innerHTML = '';
    box.append(h('h3', {}, title),
      h('div', { class: 'tp-top' }, h('div', { class: 'tp-big' }, digits(`${h12}:${pad(M)}`)),
        h('div', { class: 'tp-ap' }, h('button', { class: !pm ? 'on' : '', onclick: () => set(h12, M, false) }, 'صباحاً'), h('button', { class: pm ? 'on' : '', onclick: () => set(h12, M, true) }, 'مساءً'))),
      h('div', { class: 'tp-l' }, 'الساعة'),
      h('div', { class: 'tp-grid' }, [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12].map(x => h('button', { class: x === h12 ? 'on' : '', onclick: () => set(x, M, (x >= 1 && x <= 5) || x === 12) }, digits(x)))),
      h('div', { class: 'tp-l' }, 'الدقائق'),
      h('div', { class: 'tp-grid' }, [0, 5, 10, 15, 20, 25, 30, 35, 40, 45, 50, 55].map(x => h('button', { class: x === M ? 'on' : '', onclick: () => set(h12, x) }, digits(pad(x))))),
      h('div', { class: 'tp-row' }, h('button', { onclick: () => { sec = (sec - 60 + 86400) % 86400; draw(); } }, '− دقيقة'), h('button', { onclick: () => { sec = (sec + 60) % 86400; draw(); } }, '+ دقيقة')),
      h('div', { class: 'tp-row' },
        h('button', { class: 'tp-ok', onclick: () => { ov.remove(); done(fromSec(sec)); } }, 'تم'),
        h('button', { onclick: () => ov.remove() }, 'إلغاء')));
  };
  draw(); document.body.append(ov);
}
function scheduleLabel(s) {
  const parts = [];
  if (s.mode === 'takeover') parts.push('📺 ملء الشاشة');
  if (s.when === 'window') parts.push(`${s.from || '..'} — ${s.to || '..'}`);
  else if (s.when === 'period') parts.push('أثناء ' + (s.ptype === 'lessons' ? 'الحصص' : s.ptype === 'breaks' ? 'الفسح والطابور' : (TYPE[s.ptype]?.n || '')));
  if (s.days?.length) parts.push(s.days.map(d => DAYS[d]).join('، '));
  if (s.dFrom || s.dTo) parts.push((s.dFrom || '..') + ' ← ' + (s.dTo || '..'));
  return (parts.length ? parts.join(' | ') : 'دائماً') + (slideActive(s) ? ' • ظاهر الآن' : '');
}
function scheduleEditor(s) {
  s.when = s.when || 'always'; s.mode = s.mode || 'rotate'; s.days = s.days || []; s.intScenes = s.intScenes || ['now', 'subs'];
  const box = h('div', { style: 'flex-basis:100%;border-top:1px dashed #334155;margin-top:6px;padding-top:10px' });
  box.append(
    h('div', { class: 'row' },
      field('متى يظهر', bindSelect(s, 'when', [['always', 'دائماً'], ['window', 'من وقت إلى وقت'], ['period', 'أثناء فترة معينة']], rerender)),
      s.when === 'window' ? [field('من', bindInput(s, 'from', { type: 'time' })), field('إلى', bindInput(s, 'to', { type: 'time' }))] : null,
      s.when === 'period' ? field('الفترة', bindSelect(s, 'ptype', [['breaks', 'كل الفسح والطابور'], ['lessons', 'كل الحصص'], ...TYPES.filter(t => t.k !== 'custom').map(t => [t.k, t.n])])) : null,
      field('من تاريخ (اختياري)', bindInput(s, 'dFrom', { type: 'date' })),
      field('إلى تاريخ (اختياري)', bindInput(s, 'dTo', { type: 'date' }))),
    h('div', { class: 'row' }, h('span', { class: 'hint', style: 'margin:0' }, 'الأيام (بدون تحديد = كل الأيام):'),
      DAYS.map((d, k) => h('label', { class: 'chk' }, h('input', { type: 'checkbox', checked: s.days.includes(k), onchange: e => { s.days = e.target.checked ? [...s.days, k] : s.days.filter(x => x !== k); save(); } }), d))),
    h('div', { class: 'row' },
      field('طريقة الظهور', bindSelect(s, 'mode', [['rotate', 'ضمن المشاهد المتبدلة'], ['takeover', '📺 ملء الشاشة طوال الوقت المحدد']], rerender))),
    s.mode === 'takeover' ? h('div', {},
      h('p', { class: 'hint' }, 'في وقته يملأ الشاشة ويتكرر. ويمكن أن يتوقف كل فترة لعرض معلومات (الحصة الآن، الاحتياط...) ثم يكمل الفيديو من حيث توقف.'),
      h('div', { class: 'row' },
        field('إيقاف لعرض المعلومات كل (ث، 0 = لا)', bindInput(s, 'intEvery', { type: 'number', min: 0, style: 'width:100px' })),
        field('مدة كل مشهد معلومات (ث)', bindInput(s, 'intDur', { type: 'number', min: 3, style: 'width:100px' }))),
      h('div', { class: 'row' }, h('span', { class: 'hint', style: 'margin:0' }, 'المعلومات التي تظهر أثناء التوقف:'),
        Object.entries(SCENE_NAMES).map(([k, n]) => h('label', { class: 'chk' }, h('input', { type: 'checkbox', checked: s.intScenes.includes(k), onchange: e => { s.intScenes = e.target.checked ? [...s.intScenes, k] : s.intScenes.filter(x => x !== k); save(); } }), n)))) : null);
  return box;
}
const THEMES = [
  { k: 'oman2040', n: 'رؤية عُمان 2040 — داكن', c: ['#0a0f2e', '#1a1146', '#B64D9C', '#8585BB', '#159EC4'] },
  { k: 'oman2040-light', n: 'رؤية عُمان 2040 — فاتح', c: ['#f4f3fb', '#e6f4fa', '#a8418d', '#6f6fae', '#128bb0'] },
  { k: 'flag', n: 'ألوان علم عُمان', c: ['#12080a', '#1f0d11', '#d0103a', '#f5f5f5', '#0a9a4a'] },
  { k: 'night', n: 'ليلي (فيروزي)', c: ['#0b1220', '#111a2e', '#14b8a6', '#f59e0b', '#6366f1'] },
  { k: 'sand', n: 'رمال (ذهبي)', c: ['#241a10', '#3a2917', '#d9a441', '#2a9d8f', '#e07a5f'] },
  { k: 'light', n: 'فاتح كلاسيكي', c: ['#eef2f7', '#ffffff', '#0e7490', '#d97706', '#7c3aed'] },
];
const READY = [
  { name: 'صلاة الظهر', mode: 'fixed', time: '12:15', text: 'حان الآن وقت صلاة الظهر، يرجى التوجه إلى المصلى' },
  { name: 'الصلاة على النبي', mode: 'mid', text: 'اللهم صلِّ وسلم على نبينا محمد' },
  { name: 'استغفار', mode: 'mid', applies: 'all', text: 'أستغفر الله العظيم وأتوب إليه' },
  { name: 'تسبيح', mode: 'mid', text: 'سبحان الله وبحمده، سبحان الله العظيم' },
  { name: 'قبل نهاية الفسحة', mode: 'beforeEnd', minutes: 3, applies: 'break', text: 'تبقى ثلاث دقائق على نهاية الفسحة، استعدوا للعودة إلى الفصول' },
  { name: 'الالتزام بالنظام', mode: 'afterStart', minutes: 2, applies: 'break', sound: 'soft', text: 'نرجو من الطلاب المحافظة على النظام والنظافة أثناء الفسحة' },
  { name: 'أذكار الصباح', mode: 'afterStart', minutes: 3, applies: 'assembly', text: 'أصبحنا وأصبح الملك لله، والحمد لله' },
  { name: 'تنبيه المعلمين', mode: 'beforeEnd', minutes: 5, text: 'تبقى خمس دقائق على نهاية الحصة' },
];
const TABS = {
  /* ---------------- عام ---------------- */
  general(sec) {
    const g = CFG.general;
    sec.append(panel('بيانات المدرسة',
      h('div', { class: 'row' },
        field('اسم المدرسة', bindInput(g, 'schoolName', { style: 'width:340px' })),
        field('سطر فرعي', bindInput(g, 'schoolSub', { style: 'width:340px' }))),
      h('div', { class: 'row' },
        h('button', { class: 'btn', onclick: async () => { const f = await pickFile('image/*'); if (!f) return; if (g.logo) await removeFile(g.logo); g.logo = await addFile(f, 'logo'); save(); rerender(); } }, '🖼️ رفع شعار المدرسة'),
        g.logo ? h('button', { class: 'btn dan', onclick: async () => { await removeFile(g.logo); g.logo = ''; save(); rerender(); } }, 'حذف الشعار') : null)));

    sec.append(panel('الساعة والتاريخ',
      h('div', { class: 'row' },
        chk(g, 'clock24', 'نظام 24 ساعة'),
        chk(g, 'arabicDigits', 'أرقام عربية (١٢٣)'),
        field('ضبط فرق الساعة (ثوانٍ)', bindInput(g, 'clockOffset', { type: 'number', style: 'width:120px' }), 'إذا كانت ساعة الجهاز متقدمة/متأخرة عن ساعة المدرسة. مثال: -30 أو 45'))));

    sec.append(panel('أيام الدوام ودورة الأسابيع',
      h('div', { class: 'row' }, DAYS.map((d, i) => h('label', { class: 'chk' },
        h('input', { type: 'checkbox', checked: g.workDays.includes(i), onchange: e => { g.workDays = e.target.checked ? [...new Set([...g.workDays, i])] : g.workDays.filter(x => x !== i); save(); } }), d))),
      h('p', { class: 'hint' }, 'إذا كانت المناوبة تتغير حسب الأسبوع (أسبوع أول / ثانٍ ...)، حدد عدد الأسابيع في الدورة وتاريخ أي يوم من "الأسبوع الأول".'),
      h('div', { class: 'row' },
        field('عدد الأسابيع في الدورة', bindInput(g, 'cycleWeeks', { type: 'number', min: 1, max: 8, style: 'width:100px' })),
        field('تاريخ يقع في الأسبوع الأول', bindInput(g, 'cycleStart', { type: 'date' })),
        h('div', { class: 'field' }, h('span', {}, 'الأسبوع الحالي'), h('b', { style: 'font-size:18px;color:var(--text)' }, 'الأسبوع ' + cycleWeek(now()))))));

    const hl = h('div', { class: 'list-items' });
    g.holidays.forEach((x, i) => hl.append(h('div', { class: 'item' },
      field('من', bindInput(x, 'from', { type: 'date' })), field('إلى', bindInput(x, 'to', { type: 'date' })),
      field('المناسبة', bindInput(x, 'name', { style: 'width:240px' })),
      h('button', { class: 'btn dan sm', onclick: () => { g.holidays.splice(i, 1); save(); rerender(); } }, 'حذف'))));
    sec.append(panel('الإجازات (لا يرن الجرس فيها)', hl,
      h('div', { class: 'row', style: 'margin-top:10px' }, h('button', { class: 'btn', onclick: () => { g.holidays.push({ from: ymd(now()), to: ymd(now()), name: 'إجازة' }); save(); rerender(); } }, '+ إضافة إجازة'))));

    const SH = CFG.shifts;
    sec.append(panel('مدرسة بفترتين (صباحية ومسائية)',
      h('p', { class: 'hint' }, 'عند التفعيل يصبح لكل فترة توقيتها وجدول معلميها ومناوبتها واحتياطها وغيابها. الجرس يرن لكلتا الفترتين، والشاشة تعرض بيانات الفترة الحالية تلقائياً (وتنتقل للفترة المسائية قبل بدايتها بـ 45 دقيقة). تختار الفترة التي تعدّلها من أعلى صفحات التوقيت والجداول.'),
      h('div', { class: 'row' },
        h('label', { class: 'chk' }, h('input', { type: 'checkbox', checked: SH.enabled, onchange: e => {
          SH.enabled = e.target.checked;
          if (SH.enabled && !CFG.shiftData.s2) CFG.shiftData.s2 = shiftDefaults('s2');
          if (!SH.enabled) { EDIT_SHIFT = CUR_SHIFT = 's1'; }
          save(); rerender(); toast(SH.enabled ? 'تم تفعيل الفترتين — اضبط توقيت الفترة المسائية من "الجداول والتوقيت"' : 'تم الرجوع لفترة واحدة (بيانات المسائية محفوظة)');
        } }), 'تفعيل نظام الفترتين'),
        SH.enabled ? SH.list.map(x => field('اسم ' + (x.id === 's1' ? 'الفترة الأولى' : 'الفترة الثانية'), bindInput(x, 'name', { style: 'width:200px' }))) : null)));
    sec.append(panel('الجرس',
      h('div', { class: 'row' },
        chk(g, 'bellEnabled', 'تشغيل الجرس التلقائي'),
        field('مستوى الصوت', bindInput(g, 'volume', { type: 'range', min: 0, max: 1, step: 0.05, style: 'width:200px' }, () => Audio_.ensure())),
        field('مدة ظهور رسالة الجرس على الشاشة (ثوانٍ)', bindInput(g, 'bannerSec', { type: 'number', min: 2, max: 60, style: 'width:100px' })))));
  },

  /* ---------------- الجداول ---------------- */
  schedule(sec) {
    if (!curProfile || !CFG.profiles.find(p => p.id === curProfile)) curProfile = CFG.profiles[0]?.id;
    const pr = CFG.profiles.find(p => p.id === curProfile);
    const today = profileFor(now());
    const tplSel = h('select', {}, PROFILE_TEMPLATES.map(t => h('option', { value: t.k }, t.icon + ' ' + t.n)));
    sec.append(panel('التواقيت',
      h('p', { class: 'hint' }, 'لكل توقيت أيامه وفتراته. اضغط "تعديل" لضبطه، أو "تطبيق الآن" لفرضه على كل الأيام. ويمكن ربط توقيت بفترة تاريخية (مثل رمضان أو الاختبارات) فيُطبَّق تلقائياً في تلك الأيام ثم يعود التوقيت العادي.'),
      h('div', { class: 'prof-grid' },
        CFG.profiles.map(p => h('div', { class: 'prof-card' + (p.id === curProfile ? ' sel' : '') + (today === p ? ' live' : '') },
          today === p ? h('span', { class: 'prof-live' }, '● مُطبّق اليوم') : null,
          h('div', { class: 'prof-ic' }, p.icon || '🗓️'),
          h('div', { class: 'prof-nm' }, p.name),
          h('div', { class: 'prof-meta' }, p.days.length ? p.days.map(d => DAYS[d]).join('، ') : 'بدون أيام'),
          h('div', { class: 'prof-meta' }, digits(p.periods.length) + ' فترة' + (p.dFrom ? ` • ${p.dFrom} ← ${p.dTo || '..'}` : '')),
          h('div', { class: 'row', style: 'justify-content:center;margin:6px 0 0' },
            h('button', { class: 'btn sm' + (p.id === curProfile ? ' pri' : ''), onclick: () => { curProfile = p.id; rerender(); } }, 'تعديل'),
            h('button', { class: 'btn sm' + (CFG.override === p.id ? ' pri' : ''), onclick: () => { CFG.override = CFG.override === p.id ? '' : p.id; save(); rerender(); toast(CFG.override ? 'تم تطبيق "' + p.name + '" على كل الأيام' : 'رجوع للتطبيق حسب الأيام'); } }, CFG.override === p.id ? '✔ مفروض — إلغاء' : 'تطبيق الآن')))),
        h('div', { class: 'prof-card add' },
          h('div', { class: 'prof-ic' }, '➕'), h('div', { class: 'prof-nm' }, 'إضافة توقيت'), tplSel,
          h('button', { class: 'btn pri sm', style: 'margin-top:6px', onclick: () => {
            const t = PROFILE_TEMPLATES.find(x => x.k === tplSel.value);
            const n = { id: 'pr_' + uid(), name: t.n, icon: t.icon, days: [], periods: t.gen ? genPeriods(t.gen) : [] };
            CFG.profiles.push(n); curProfile = n.id; save(); rerender(); toast('أُضيف "' + t.n + '" — حدد أيامه أو تاريخه ثم راجع الأوقات');
          } }, 'إضافة'))),
      CFG.override ? h('p', { class: 'hint', style: 'color:#fcd34d;margin-top:10px' }, '⚠️ التوقيت "' + (CFG.profiles.find(p => p.id === CFG.override)?.name || '') + '" مفروض على كل الأيام الآن.') : null));
    if (!pr) return;

    sec.append(panel('إعدادات "' + pr.name + '"',
      h('div', { class: 'row' }, field('اسم التوقيت', bindInput(pr, 'name', { style: 'width:240px' }, rerender)),
        field('الرمز', bindSelect(pr, 'icon', ['🗓️', '❄️', '☀️', '🌅', '🌙', '📝', '📅', '⭐', '🏫', '⏰'].map(x => [x, x]), rerender)),
        field('يُطبَّق تلقائياً من تاريخ (اختياري)', bindInput(pr, 'dFrom', { type: 'date' }, rerender)),
        field('إلى تاريخ', bindInput(pr, 'dTo', { type: 'date' }, rerender)),
        h('button', { class: 'btn', onclick: () => { const n = clone(pr); n.id = 'pr_' + uid(); n.name = pr.name + ' (نسخة)'; n.days = []; n.dFrom = n.dTo = ''; n.periods.forEach(p => p.id = uid()); CFG.profiles.push(n); curProfile = n.id; save(); rerender(); } }, 'نسخ'),
        CFG.profiles.length > 1 ? h('button', { class: 'btn dan', onclick: () => { if (!confirm('حذف التوقيت "' + pr.name + '"؟')) return; CFG.profiles = CFG.profiles.filter(p => p !== pr); if (CFG.override === pr.id) CFG.override = ''; curProfile = null; save(); rerender(); } }, 'حذف') : null),
      h('div', { class: 'row' }, h('span', { class: 'hint', style: 'margin:0' }, 'الأيام:'), DAYS.map((d, i) => h('label', { class: 'chk' },
        h('input', { type: 'checkbox', checked: pr.days.includes(i), onchange: e => {
          if (e.target.checked) { if (!pr.dFrom) CFG.profiles.forEach(p => { if (!p.dFrom) p.days = p.days.filter(x => x !== i); }); pr.days.push(i); }
          else pr.days = pr.days.filter(x => x !== i);
          save(); rerender();
        } }), d)))));

    const tb = h('tbody');
    const periods = pr.periods;
    periods.forEach((p, i) => {
      const labelIn = bindInput(p, 'label', { placeholder: TYPE[p.type]?.n || 'الاسم', style: 'width:130px' });
      const typeSel = h('select', { onchange: e => { p.type = e.target.value; save(); rerender(); } }, TYPES.map(t => h('option', { value: t.k, selected: t.k === p.type }, t.n)));
      const dur = toSec(p.end) - toSec(p.start);
      tb.append(h('tr', {},
        h('td', {}, digits(i + 1)),
        h('td', {}, typeSel),
        h('td', {}, labelIn),
        h('td', {}, timeBtn(p, 'start', 'وقت البداية — ' + periodName(p))),
        h('td', {}, timeBtn(p, 'end', 'وقت النهاية — ' + periodName(p))),
        h('td', { style: 'color:var(--muted);white-space:nowrap' }, dur > 0 ? Math.round(dur / 60) + ' د' : '⚠️'),
        h('td', {}, h('div', { class: 'mini' }, soundModeSelect(p, KS, 'إعلان'), testBtn(() => playPeriod(p)))),
        h('td', {}, h('div', { class: 'mini' }, soundModeSelect(p, KE, 'إعلان'), testBtn(() => playPeriod(p, true)))),
        h('td', {}, bindSelect(p, 'content', CONTENT_OPTS, null, { style: 'width:150px' })),
        h('td', {}, h('div', { class: 'mini' },
          h('button', { class: 'btn sm', title: 'نصوص وتسجيلات مخصصة', onclick: () => { p._open = !p._open; rerender(); } }, '✎'),
          h('button', { class: 'btn sm', title: 'أعلى', onclick: () => { if (i) { [periods[i - 1], periods[i]] = [periods[i], periods[i - 1]]; save(); rerender(); } } }, '↑'),
          h('button', { class: 'btn sm', title: 'أسفل', onclick: () => { if (i < periods.length - 1) { [periods[i + 1], periods[i]] = [periods[i], periods[i + 1]]; save(); rerender(); } } }, '↓'),
          h('button', { class: 'btn dan sm', onclick: () => { periods.splice(i, 1); save(); rerender(); } }, '✕')))));
      if (p._open) tb.append(h('tr', {}, h('td', {}), h('td', { colspan: 9 },
        h('div', { class: 'row' },
          field('نص إعلان البداية', bindInput(p, 'startText', { placeholder: startText({ ...p, startText: '' }), style: 'width:300px' })),
          testBtn(() => playPeriod(p), '▶ تجربة البداية')),
        h('div', { class: 'row' },
          field('نص إعلان النهاية', bindInput(p, 'endText', { placeholder: endText({ ...p, endText: '' }), style: 'width:300px' })),
          testBtn(() => playPeriod(p, true), '▶ تجربة النهاية')))));
    });
    const warn = [];
    const sp = sortedPeriods(pr);
    for (let i = 1; i < sp.length; i++) if (toSec(sp[i].start) < toSec(sp[i - 1].end)) warn.push(`تداخل بين "${periodName(sp[i - 1])}" و"${periodName(sp[i])}"`);
    sec.append(panel('فترات اليوم',
      h('p', { class: 'hint' }, 'اختر نوع الفترة (طابور، حصة صفرية، الأولى ... الثامنة، فسحة، نشاط، ريادة ...)، والوقت من/إلى، واختيار واحد للصوت في البداية وواحد في النهاية: بدون صوت، أو إعلان صوتي فقط ("بدأت الآن الحصة الأولى")، أو نغمة فقط، أو نغمة ثم إعلان، أو ملف صوتي مرفوع. زر ✎ لكتابة نص إعلان خاص. إذا انتهت حصة وبدأت التالية في نفس الدقيقة يُدمج الإعلانان في جرس واحد.'),
      warn.length ? h('p', { class: 'hint', style: 'color:#fca5a5' }, '⚠️ ' + warn.join(' — ')) : null,
      h('div', { class: 'tbl-wrap' }, h('table', { class: 'edit' },
        h('thead', {}, h('tr', {}, ['#', 'النوع', 'اسم مخصص', 'من', 'إلى', 'المدة', 'صوت البداية', 'صوت النهاية', 'ما يظهر في الشاشة', ''].map(t => h('th', {}, t)))), tb)),
      h('div', { class: 'row', style: 'margin-top:12px' },
        h('button', { class: 'btn pri', onclick: () => {
          const last = sortedPeriods(pr).slice(-1)[0];
          const st = last ? last.end : '07:00';
          const used = new Set(pr.periods.map(p => p.type));
          const nextType = ['p1', 'p2', 'p3', 'p4', 'p5', 'p6', 'p7', 'p8'].find(k => !used.has(k)) || 'custom';
          pr.periods.push(P(nextType, st, fromSec(toSec(st) + 40 * 60))); save(); rerender();
        } }, '+ إضافة فترة'),
        h('button', { class: 'btn', onclick: () => { pr.periods.sort((a, b) => toSec(a.start) - toSec(b.start)); save(); rerender(); } }, 'ترتيب حسب الوقت'),
        h('button', { class: 'btn pri', onclick: () => scheduleFromImage(pr) }, '📷 قراءة التوقيت من صورة')),
      h('p', { class: 'hint' }, '📷 صوّر جدول التوقيت أو ارفع صورته، فيملأ البرنامج الفترات والأوقات تلقائياً (بدون إنترنت). الأفضل صورة واضحة ومستقيمة لجدول فيه خطوط. راجع النتيجة بعدها وعدّل ما يلزم — زر 🕐 على أي وقت يفتح منتقي الوقت السهل.')));

    // مولّد تلقائي
    const gen = { start: '07:00', assembly: 15, zero: false, count: 7, len: 40, gap: 0, breakAfter: 3, breakLen: 25, break2After: 0, break2Len: 15 };
    sec.append(panel('توليد الجدول تلقائياً (اختياري)',
      h('p', { class: 'hint' }, 'يملأ فترات هذا الجدول تلقائياً حسب المدة — سيستبدل الفترات الحالية.'),
      h('div', { class: 'row' },
        field('بداية الدوام', bindInput(gen, 'start', { type: 'time' })),
        field('مدة الطابور (د، 0 = بدون)', bindInput(gen, 'assembly', { type: 'number', style: 'width:90px' })),
        h('label', { class: 'chk' }, bindInput(gen, 'zero', { type: 'checkbox' }), 'حصة صفرية'),
        field('عدد الحصص', bindInput(gen, 'count', { type: 'number', min: 1, max: 8, style: 'width:80px' })),
        field('مدة الحصة (د)', bindInput(gen, 'len', { type: 'number', style: 'width:80px' })),
        field('فاصل بين الحصص (د)', bindInput(gen, 'gap', { type: 'number', style: 'width:80px' }))),
      h('div', { class: 'row' },
        field('الفسحة بعد الحصة رقم', bindInput(gen, 'breakAfter', { type: 'number', style: 'width:80px' })),
        field('مدة الفسحة (د)', bindInput(gen, 'breakLen', { type: 'number', style: 'width:80px' })),
        field('فسحة ثانية بعد الحصة (0 = لا)', bindInput(gen, 'break2After', { type: 'number', style: 'width:80px' })),
        field('مدتها (د)', bindInput(gen, 'break2Len', { type: 'number', style: 'width:80px' })),
        h('button', { class: 'btn pri', onclick: () => {
          if (pr.periods.length && !confirm('سيتم استبدال فترات هذا الجدول. متابعة؟')) return;
          pr.periods = genPeriods(gen); save(); rerender(); toast('تم توليد الجدول');
        } }, 'توليد'))));
  },

  /* ---------------- الأصوات ---------------- */
  sounds(sec) {
    const v = CFG.voice;
    const ar = Audio_.arabicVoices();
    const all = Audio_.voices();
    const vc = Object.keys(window.VOICE_CLIPS || {});
    const sets = Object.entries(window.VOICE_SETS || {});
    sec.append(panel('🎙️ الأصوات الجاهزة (تسجيلات احترافية)',
      h('p', { class: 'hint' }, 'تسجيلات جاهزة لإعلانات بداية ونهاية الحصص والطابور والفسحة والانصراف والأذكار. عند تفعيلها تُستخدم بدل النطق الآلي كلما وُجد تسجيل مطابق للنص، وما لا يوجد له تسجيل يُنطق آلياً.'),
      h('div', { class: 'row' }, chk(v, 'omani', 'استخدام الأصوات الجاهزة'),
        field('الصوت', bindSelect(v, 'voiceSet', sets.map(([k, x]) => [k, x.name]), rerender)),
        sets.map(([k, x]) => h('button', { class: 'btn sm' + (v.voiceSet === k ? ' pri' : ''), onclick: () => { Audio_.ensure(); Audio_.stopFlag = false; playURL(x.dir + '/start_p1.mp3'); } }, '▶ ' + x.name.split(' — ')[0]))),
      h('p', { class: 'hint', style: 'margin-top:8px' }, 'استماع لكل العبارات بالصوت المختار:'),
      h('div', { class: 'row', style: 'gap:6px' }, vc.map(t => h('button', { class: 'btn sm', onclick: () => { Audio_.ensure(); Audio_.stopFlag = false; const u = clipFor(t); if (u) playURL(u); } }, '▶ ' + t)))));
    sec.append(panel('الإعلان الصوتي (نطق آلي)',
      h('p', { class: 'hint' }, ar.length ? `✅ يوجد ${ar.length} صوت عربي على هذا الجهاز.` :
        '⚠️ لم يُعثر على صوت عربي مثبت في الجهاز. في ويندوز: الإعدادات ← الوقت واللغة ← اللغة ← إضافة "العربية" مع ميزة "تحويل النص إلى كلام"، ثم أعد تشغيل البرنامج. أو استخدم تسجيلات صوتية بدل النطق الآلي (من الأسفل).'),
      h('div', { class: 'row' },
        chk(v, 'tts', 'تفعيل النطق الآلي'),
        field('الصوت', bindSelect(v, 'voiceName', [['', 'تلقائي (أول صوت عربي)'], ...ar.map(x => [x.name, x.name + ' — ' + x.lang]), ...all.filter(x => !/^ar/i.test(x.lang)).map(x => [x.name, x.name + ' — ' + x.lang])], null, { style: 'max-width:320px' })),
        field('السرعة', bindInput(v, 'rate', { type: 'range', min: 0.5, max: 1.5, step: 0.05 })),
        field('تكرار الإعلان', bindSelect(v, 'repeat', [['1', 'مرة'], ['2', 'مرتين'], ['3', '3 مرات']]))),
      h('div', { class: 'row' },
        field('قالب البداية ({name} = اسم الفترة)', bindInput(v, 'startTpl', { style: 'width:260px' })),
        field('قالب النهاية', bindInput(v, 'endTpl', { style: 'width:260px' })),
        testBtn(() => Queue.push({ sound: 'classic', parts: [{ text: v.startTpl.replace('{name}', 'الحصة الأولى') }], banner: v.startTpl.replace('{name}', 'الحصة الأولى') }), '▶ تجربة إعلان'))));

    const lib = h('div', { class: 'list-items' });
    for (const [k, b] of Object.entries(Audio_.bells)) lib.append(h('div', { class: 'item' }, h('span', { class: 't' }, '🔔 ' + b.n), h('span', { class: 'badge' }, 'مدمج'), testBtn(() => Audio_.playSound(k))));
    for (const f of CFG.files.filter(f => f.kind === 'audio')) {
      const nm = h('input', { type: 'text', value: f.name, style: 'flex:1;min-width:160px', onchange: e => { f.name = e.target.value; save(); } });
      lib.append(h('div', { class: 'item' }, '🎵', nm, h('span', { class: 'badge' }, Math.round(f.size / 1024) + ' KB'), testBtn(() => Audio_.playFile(f.id)),
        h('button', { class: 'btn dan sm', onclick: async () => { if (!confirm('حذف الصوت؟')) return; await removeFile(f.id); save(); rerender(); } }, 'حذف')));
    }
    const recState = h('span');
    sec.append(panel('مكتبة الأصوات',
      h('p', { class: 'hint' }, '🎬 لرفع فيديو أو صور وتحديد وقت ظهورها: تبويب "المشاهد والشريط". — الأجراس المدمجة تعمل بدون ملفات. يمكنك رفع أصوات أخرى (MP3/WAV/OGG): جرس خاص، آية، دعاء، نشيد، إعلان مسجّل بصوتك... ثم اختيارها في الجداول أو التنبيهات.'),
      lib,
      h('div', { class: 'row', style: 'margin-top:12px' },
        h('button', { class: 'btn pri', onclick: async () => {
          const fs = await pickFile('audio/*', true); if (!fs?.length) return;
          for (const f of fs) await addFile(f, 'audio');
          toast('تم رفع ' + fs.length + ' ملف'); rerender();
        } }, '⬆ رفع ملفات صوتية'),
        h('button', { class: 'btn', onclick: e => recordAudio(e.target, recState) }, '🎙️ تسجيل بالميكروفون'), recState)));
  },

  /* ---------------- التنبيهات ---------------- */
  reminders(sec) {
    const list = h('div', { class: 'list-items' });
    const modes = [['mid', 'منتصف الفترة'], ['afterStart', 'بعد بداية الفترة بـ'], ['beforeEnd', 'قبل نهاية الفترة بـ'], ['fixed', 'في وقت محدد يومياً']];
    const applies = [['lessons', 'كل الحصص'], ['all', 'كل الفترات'], ...TYPES.filter(t => t.k !== 'custom').map(t => [t.k, t.n])];
    CFG.reminders.forEach((r, i) => {
      if (!r.show) r.show = 'banner'; if (r.vol == null) r.vol = 1;
      list.append(h('div', { class: 'item' + (r.enabled ? '' : ' disabled') },
        h('label', { class: 'chk' }, bindInput(r, 'enabled', { type: 'checkbox' }, rerender)),
        field('الاسم', bindInput(r, 'name', { style: 'width:150px' })),
        field('متى', bindSelect(r, 'mode', modes, rerender)),
        r.mode === 'fixed' ? field('الوقت', bindInput(r, 'time', { type: 'time' })) :
          (r.mode !== 'mid' ? field('دقائق', bindInput(r, 'minutes', { type: 'number', min: 1, style: 'width:70px' })) : null),
        r.mode !== 'fixed' ? field('يطبق على', bindSelect(r, 'applies', applies)) : null,
        field('الصوت', soundModeSelect(r, KR, 'نطق النص')),
        field('النص (كل سطر عبارة — يُختار واحد عشوائياً كل مرة)', h('textarea', { style: 'width:260px;min-height:70px', onchange: e => { r.text = e.target.value; save(); } }, r.text || '')),
        field('مستوى الصوت 🔉', bindInput(r, 'vol', { type: 'range', min: 0.05, max: 1, step: 0.05, value: r.vol ?? 1 })),
        field('الظهور على الشاشة', bindSelect(r, 'show', [['soft', 'بطاقة صغيرة هادئة أسفل الشاشة'], ['banner', 'رسالة كبيرة (مثل الجرس)'], ['none', 'بدون ظهور (صوت فقط)']])),
        testBtn(() => { const e = remEvent(r, 'test', 0); Queue.push({ sound: r.sound, parts: [{ text: e.text, audio: e.audio, volume: Math.min(1, CFG.general.volume * (r.vol ?? 1)) }], vol: r.vol ?? 1, quiet: (r.vol ?? 1) < 0.7, soft: (r.show || 'banner') === 'soft' ? e.banner : '', banner: (r.show || 'banner') === 'banner' ? e.banner : '', bannerIcon: '📢' }); }),
        h('button', { class: 'btn dan sm', onclick: () => { CFG.reminders.splice(i, 1); save(); rerender(); } }, 'حذف')));
    });
    sec.append(panel('تنبيهات إضافية أثناء اليوم',
      h('p', { class: 'hint' }, 'أمثلة: في منتصف كل حصة "اللهم صلِّ على محمد"، قبل نهاية الحصة بـ 5 دقائق تنبيه للمعلم، أو في وقت محدد (مثلاً 10:00 أذكار / النشيد الوطني). يمكن استخدام النطق الآلي أو ملف صوتي مرفوع.'),
      list,
      h('div', { class: 'row', style: 'margin-top:12px' },
        h('button', { class: 'btn pri', onclick: () => { CFG.reminders.push({ id: uid(), enabled: true, name: 'تنبيه جديد', mode: 'mid', minutes: 5, time: '10:00', applies: 'lessons', sound: 'chime', text: '', audio: '' }); save(); rerender(); } }, '+ إضافة تنبيه')),
      h('p', { class: 'hint', style: 'margin-top:12px' }, 'إضافة سريعة من عبارات جاهزة (يمكن تعديل الوقت والنص بعد الإضافة، أو اختيار تسجيل صوتي ترفعه لاحقاً):'),
      h('div', { class: 'row' }, READY.map(t => h('button', { class: 'btn sm', onclick: () => { CFG.reminders.push({ id: uid(), enabled: true, sound: 'chime', audio: '', minutes: 5, applies: 'lessons', ...clone(t) }); save(); rerender(); } }, '+ ' + t.name)))));

    const ql = h('div', { class: 'list-items' });
    CFG.quick.forEach((q, i) => ql.append(h('div', { class: 'item' },
      h('b', { style: 'width:22px' }, digits(i + 1)),
      field('النص', bindInput(q, 'text', { style: 'width:380px' })),
      field('الصوت', soundModeSelect(q, KR, 'نطق النص')),
      testBtn(() => quickAnnounce(q)),
      h('button', { class: 'btn dan sm', onclick: () => { CFG.quick.splice(i, 1); save(); rerender(); } }, 'حذف'))));
    sec.append(panel('إعلانات فورية بضغطة زر',
      h('p', { class: 'hint' }, 'تظهر عند الضغط على زر 📢 في الشاشة الرئيسية (أو حرف A)، أو بالضغط على رقمها من 1 إلى 9 مباشرة.'),
      ql,
      h('div', { class: 'row', style: 'margin-top:12px' },
        h('button', { class: 'btn pri', onclick: () => { CFG.quick.push({ id: uid(), text: 'نص الإعلان', sound: 'chime', audio: '' }); save(); rerender(); } }, '+ إضافة إعلان'))));
  },

  /* ---------------- العرض والشريط ---------------- */
  slides(sec) {
    const list = h('div', { class: 'list-items' });
    const icons = { pdf: '📄', image: '🖼️', video: '🎬', url: '🔗', text: '📝' };
    CFG.slides.forEach((s, i) => {
      const f = s.fileId && CFG.files.find(f => f.id === s.fileId);
      list.append(h('div', { class: 'item' + (s.enabled ? '' : ' disabled') },
        h('label', { class: 'chk' }, bindInput(s, 'enabled', { type: 'checkbox' }, rerender)),
        h('span', { style: 'font-size:22px' }, icons[s.type] || '•'),
        s.type === 'text' ? h('textarea', { style: 'flex:1;min-height:60px;min-width:220px', onchange: e => { s.text = e.target.value; save(); } }, s.text) :
          h('div', { class: 't' }, field('العنوان (اختياري)', bindInput(s, 'title', { placeholder: f ? f.name : (s.url || '') }))),
        s.type === 'url' || (s.url && !s.fileId) ? field('الرابط', bindInput(s, 'url', { type: 'url', style: 'width:260px;direction:ltr' }, () => { delete pdfCache[s.url]; })) : null,
        field(s.type === 'video' ? 'المدة (ث، 0 = حتى نهاية الفيديو)' : (s.type === 'pdf' ? 'مدة كل صفحة (ث)' : 'المدة (ث)'), bindInput(s, 'duration', { type: 'number', min: 0, style: 'width:90px' })),
        (s.type === 'pdf' || s.type === 'image') ? field('طريقة العرض', bindSelect(s, 'fit', [['contain', 'ملاءمة الشاشة'], ['scroll', 'بعرض الشاشة مع تمرير تلقائي']])) : null,
        s.type === 'video' ? h('label', { class: 'chk' }, bindInput(s, 'muted', { type: 'checkbox' }), '🔇 بدون صوت') : null,
        h('span', { class: 'badge' + (slideActive(s) ? ' ok' : '') }, scheduleLabel(s)),
        h('div', { class: 'mini' },
          h('button', { class: 'btn sm' + (s._open ? ' pri' : ''), title: 'متى يظهر', onclick: () => { s._open = !s._open; rerender(); } }, '⏱ الجدولة'),
          h('button', { class: 'btn sm', onclick: () => { if (i) { [CFG.slides[i - 1], CFG.slides[i]] = [CFG.slides[i], CFG.slides[i - 1]]; save(); rerender(); } } }, '↑'),
          h('button', { class: 'btn sm', onclick: () => { if (i < CFG.slides.length - 1) { [CFG.slides[i + 1], CFG.slides[i]] = [CFG.slides[i], CFG.slides[i + 1]]; save(); rerender(); } } }, '↓'),
          h('button', { class: 'btn dan sm', onclick: async () => { if (!confirm('حذف هذه الشريحة؟')) return; if (s.fileId) await removeFile(s.fileId); CFG.slides.splice(i, 1); save(); rerender(); } }, '✕')),
        s._open ? scheduleEditor(s) : null));
    });
    const addUrl = h('input', { type: 'url', placeholder: 'https://... رابط PDF أو صورة أو فيديو أو يوتيوب أو صفحة', style: 'flex:1;min-width:280px;direction:ltr' });
    sec.append(panel('الشرائح المتحركة في الشاشة الرئيسية',
      h('p', { class: 'hint' }, 'ارفع فيديو أو صور أو PDF أو ضع رابطاً. زر ⏱ الجدولة يحدد متى يظهر كل عنصر ومتى يختفي: دائماً، أو من وقت إلى وقت، أو أثناء فترة (مثل الفسحة)، وفي أيام وتواريخ معينة. ويمكن جعل الفيديو يملأ الشاشة في وقته ويتوقف كل فترة لعرض الحصة الآن أو الاحتياط ثم يكمل من حيث توقف. الفيديو مكتوم افتراضياً، ويُكتم تلقائياً أثناء الجرس.'),
      list,
      h('div', { class: 'row', style: 'margin-top:12px' },
        h('button', { class: 'btn pri', onclick: async () => {
          const fs = await pickFile('application/pdf,image/*,video/*', true); if (!fs?.length) return;
          for (const f of fs) {
            const k = kindOf(f); if (!['pdf', 'image', 'video'].includes(k)) { toast('نوع غير مدعوم: ' + f.name, true); continue; }
            const id = await addFile(f, k);
            CFG.slides.push({ id: uid(), type: k, fileId: id, enabled: true, title: k === 'pdf' ? f.name.replace(/\.pdf$/i, '') : '', duration: k === 'video' ? 0 : (k === 'pdf' ? 20 : 12), muted: true, fit: 'contain' });
          }
          save(); rerender(); Slides.rebuild(); toast('تمت الإضافة');
        } }, '⬆ رفع ملفات (PDF / صور / فيديو)'),
        h('button', { class: 'btn', onclick: () => { CFG.slides.push({ id: uid(), type: 'text', enabled: true, text: 'نص الإعلان هنا', duration: 12 }); save(); rerender(); } }, '+ شريحة نصية'),
        CFG.slides.some(s => s.type === 'pdf') ? h('button', { class: 'btn dan', onclick: async () => {
          const pdfs = CFG.slides.filter(s => s.type === 'pdf'); if (!confirm(`حذف ${pdfs.length} ملف PDF من المشاهد؟ (جدول المعلمين لا يتأثر)`)) return;
          for (const s of pdfs) if (s.fileId) await removeFile(s.fileId);
          CFG.slides = CFG.slides.filter(s => s.type !== 'pdf'); save(); rerender(); Slides.rebuild();
        } }, '🗑️ حذف كل ملفات PDF من المشاهد') : null),
      h('div', { class: 'row' }, addUrl,
        h('button', { class: 'btn', onclick: () => {
          const u = addUrl.value.trim(); if (!u) return;
          const isPdf = /\.pdf(\?|$)/i.test(u);
          CFG.slides.push({ id: uid(), type: isPdf ? 'pdf' : 'url', url: u, enabled: true, title: '', duration: isPdf ? 20 : 30, muted: true, fit: 'contain' });
          save(); rerender(); Slides.rebuild();
        } }, '+ إضافة رابط')),
      h('p', { class: 'hint' }, 'ملاحظة: الروابط تحتاج إنترنت، وبعض المواقع تمنع عرضها داخل برامج أخرى. الأفضل رفع الملف نفسه ليعمل بدون إنترنت.')));

    const A = CFG.autoSlides;
    sec.append(panel('يوم بلا دوام',
      h('div', { class: 'row' }, h('label', { class: 'chk' }, h('input', { type: 'checkbox', checked: CFG.display.offMode !== false, onchange: e => { CFG.display.offMode = e.target.checked; save(); } }), '🌙 في الإجازات وبعد نهاية الدوام: عرض "لا يوجد دوام" وموعد الدوام القادم مع الأذكار، وإخفاء الجداول')),
      h('p', { class: 'hint' }, 'الصور والفيديو المضبوطة بزر ⏱ الجدولة على أيام أو أوقات معينة تظهر في هذا الوضع أيضاً إن وافق وقتها.')));
    sec.append(panel('المشاهد التلقائية (تتبدل مع الصور والفيديو في وسط الشاشة)',
      h('p', { class: 'hint' }, 'مشهد "الحصة الآن" يعرض كل صف ومعلمه في الحصة الحالية (من تبويب جدول المعلمين) مع إبراز حصص الاحتياط: المعلم الغائب ← المعلم البديل. وفي الفسحة يعرض المناوبين. يتكرر كل 3 مشاهد، ويظهر فوراً عند بداية كل حصة.'),
      h('div', { class: 'row' },
        chk(A, 'welcome', '🏫 ترحيب (الشعار واسم المدرسة)'), chk(A, 'now', '📚 الحصة الآن'), chk(A, 'classday', '🗓️ جدول حصص اليوم لكل الصفوف'), chk(A, 'schedule', '⏰ توقيت اليوم'),
        chk(A, 'duty', '🛡️ المناوبة'), chk(A, 'subs', '🔁 الاحتياط'), chk(A, 'absent', '📋 الغياب')),
      h('div', { class: 'row' },
        field('عدد البطاقات في الصفحة (0 = تلقائي)', bindInput(A, 'perPage', { type: 'number', min: 0, style: 'width:90px' })))));
    A.durs = A.durs || {};
    const DN = [['now', '📚 الحصة الآن'], ['classday', '🗓️ جدول اليوم'], ['welcome', '🏫 الترحيب'], ['schedule', '⏰ التوقيت'], ['duty', '🛡️ المناوبة'], ['subs', '🔁 الاحتياط'], ['absent', '📋 الغياب'], ['phrase', '🤲 الأذكار']];
    sec.append(panel('⏱️ مدة العرض وسرعة الحركة',
      h('p', { class: 'hint' }, 'كم ثانية يبقى كل مشهد على الشاشة قبل الانتقال للتالي. (مدة الصور والفيديو والـ PDF تُضبط لكل عنصر في القائمة أعلاه.)'),
      h('div', { class: 'row' }, DN.map(([k, n]) => field(n + ' (ثانية — 60 = دقيقة)', bindInput(A.durs, k, { type: 'number', min: 3, style: 'width:90px', value: durOf(k) })))),
      h('div', { class: 'row' },
        field('سرعة الانتقال بين المشاهد', bindSelect(A, 'trDur', [['0.6', 'سريعة'], ['1.6', 'عادية هادئة'], ['2.5', 'بطيئة'], ['4', 'بطيئة جداً']])),
        field('نوع الحركة', bindSelect(CFG.display, 'transition', [['fade', 'تلاشي ناعم (الأهدأ)'], ['random', 'متنوعة'], ['up', 'صعود'], ['zoom', 'تكبير'], ['side', 'انزلاق'], ['flip', 'دوران']])),
        chk(CFG.general, 'animBg', 'خلفية متحركة'))));
    sec.append(panel('📌 تثبيت الشاشة على مشهد واحد',
      h('p', { class: 'hint' }, 'لإيقاف التبديل وعرض شيء واحد فقط باستمرار (يتحدث تلقائياً مع تغيّر الحصة). أو اضغطي زر 📌 في الشاشة الرئيسية (أو حرف P) لتثبيت المشهد المعروض حالياً مؤقتاً.'),
      h('div', { class: 'row' }, field('تثبيت على', bindSelect(A, 'lock', [['', 'لا — تبديل تلقائي بين المشاهد'], ['now', '📚 الحصة الآن (كل الصفوف)'], ['classday', '🗓️ جدول حصص اليوم'], ['duty', '🛡️ المناوبة'], ['subs', '🔁 الاحتياط'], ['absent', '📋 الغياب'], ['schedule', '⏰ التوقيت'], ['welcome', '🏫 الترحيب']])))));

    const T = CFG.ticker;
    const ta = h('textarea', { onchange: e => { T.texts = e.target.value.split('\n').map(s => s.trim()).filter(Boolean); save(); } }, T.texts.join('\n'));
    sec.append(panel('الشريط الإخباري المتحرك (أسفل الشاشة)',
      h('div', { class: 'row' }, chk(T, 'enabled', 'إظهار الشريط'), field('العنوان', bindInput(T, 'label')), field('السرعة', bindInput(T, 'speed', { type: 'range', min: 30, max: 250 })),
        chk(T, 'includeDuty', 'إضافة المناوبة الحالية'), chk(T, 'includeSubs', 'إضافة حصص الاحتياط'), chk(T, 'includeAbsent', 'إضافة الغياب')),
      h('p', { class: 'hint' }, 'كل سطر = خبر:'), ta));

  },

  /* ---------------- المظهر ---------------- */
  look(sec) {
    const g = CFG.general, D = CFG.display;
    const grid = h('div', { class: 'theme-grid' });
    for (const t of THEMES) grid.append(h('button', { class: 'theme-card' + (g.theme === t.k ? ' on' : ''), onclick: () => { g.theme = t.k; save(); rerender(); } },
      h('div', { class: 'sw', style: `background:linear-gradient(135deg, ${t.c[0]}, ${t.c[1]})` },
        t.c.slice(2).map((c, i) => h('i', { style: `background:${c};right:${10 + i * 28}px` }))),
      h('div', { class: 'nm' }, t.n)));
    sec.append(panel('الثيم', h('p', { class: 'hint' }, 'ثيم "رؤية عُمان 2040" مستوحى من ألوان شعار الرؤية (البنفسجي الوردي، الأزرق السماوي، اللافندر). يمكنك تعديل الألوان بدقة من "ألوان مخصصة".'), grid));
    sec.append(panel('ألوان مخصصة',
      h('div', { class: 'row' },
        chk(g.custom, 'on', 'استخدام ألواني بدلاً من ألوان الثيم'),
        field('اللون الأساسي', bindInput(g.custom, 'accent', { type: 'color' })),
        field('اللون الثاني', bindInput(g.custom, 'accent2', { type: 'color' })),
        field('اللون الثالث', bindInput(g.custom, 'accent3', { type: 'color' })))));
    sec.append(panel('الشاشة الرئيسية',
      h('div', { class: 'row' },
        field('التخطيط', bindSelect(D, 'layout', [['hero', 'شاشة كاملة للمشاهد (موصى به)'], ['side', 'مشاهد + عمود جانبي للمناوبة والاحتياط']])),
        field('مكان الشريط المتحرك', bindSelect(D, 'tickerPos', [['bottom', 'أسفل الشاشة'], ['top', 'أعلى الشاشة']])),
        field('حركة الانتقال بين المشاهد', bindSelect(D, 'transition', [['random', 'متنوعة عشوائياً'], ['fade', 'تلاشي'], ['up', 'صعود'], ['zoom', 'تكبير'], ['side', 'انزلاق جانبي'], ['flip', 'دوران']]))),
      h('div', { class: 'row' },
        chk(g, 'animBg', 'خلفية متحركة'),
        chk(D, 'cast', '📡 وضع مشاركة الشاشة (ميراكاست / ميرا سكرين / AirPlay): بدون حركة مستمرة، أوضح وأسرع على التلفاز', () => { applyTheme(); updateTicker(true); Slides.rebuild(); }),
        field('شريط الأخبار في وضع المشاركة', bindSelect(D, 'castTicker', [['scroll', 'متحرك (انسيابي)'], ['pages', 'عبارة ثابتة تتبدل (الأوضح للبث الضعيف)']], () => updateTicker(true))),
        D.castTicker === 'pages' ? field('تبديل العبارة كل (ثانية)', bindInput(D, 'castTickerSec', { type: 'number', min: 4, style: 'width:70px', value: D.castTickerSec || 8 }, () => updateTicker(true))) : null,
        chk(D, 'showStatus', 'شريط "الحصة الآن" أعلى الشاشة'),
        h('label', { class: 'chk' }, h('input', { type: 'checkbox', checked: D.focusTable !== false, onchange: e => { D.focusTable = e.target.checked; save(); } }), 'تكبير جدول الحصة الآن (إخفاء الأشرطة أثناء عرضه)'),
        field('عرض الحصة الآن', bindSelect(D, 'nowStyle', [['grades', 'كل الصفوف في جدول واحد (الصف | المعلمة | المادة)'], ['table', 'أعمدة مضغوطة (المعلم | الصف)'], ['cards', 'بطاقات كبيرة']])),
        field('عرض الاحتياط', bindSelect(D, 'subsStyle', [['grouped', 'مجمّع حسب المعلم الغائب'], ['cards', 'بطاقات']])),
        field('المعلمون في المدرسة', bindSelect(g, 'gender', [['f', 'معلمات (المعلمة غائبة)'], ['m', 'معلمون (المعلم غائب)'], ['mixed', 'مختلط']]))),
      D.layout === 'side' ? h('div', { class: 'row' }, chk(D, 'showDuty', 'بطاقة المناوبة'), chk(D, 'showSubs', 'بطاقة الاحتياط'), chk(CFG.phrases, 'showCard', 'بطاقة الأذكار')) : null));
  },

  /* ---------------- جدول المعلمين ---------------- */
  timetable(sec) { timetableTab(sec); },

  /* ---------------- الأذكار والتحفيز ---------------- */
  phrases(sec) {
    const P = CFG.phrases;
    const lines = key => h('textarea', { style: 'min-height:220px', onchange: e => { P[key] = e.target.value.split('\n').map(s => s.trim()).filter(Boolean); save(); } }, P[key].join('\n'));
    sec.append(panel('عرض الأذكار والعبارات التحفيزية تلقائياً',
      h('p', { class: 'hint' }, 'تظهر تلقائياً على الشاشة بدون أي إعداد: بطاقة في العمود الجانبي تتبدل باستمرار (أسرع في الفسحة وبين الحصص)، وشرائح كبيرة بين شرائح العرض. عدّل القوائم كما تريد — كل سطر عبارة.'),
      h('div', { class: 'row' },
        chk(P, 'enabled', 'تفعيل الأذكار والعبارات'),
        chk(P, 'showCard', 'بطاقة في العمود الجانبي'),
        chk(P, 'slide', 'شرائح كبيرة بين شرائح العرض'),
        chk(P, 'ticker', 'إضافتها للشريط الإخباري'),
        field('مدة كل عبارة (ث)', bindInput(P, 'cardSec', { type: 'number', min: 5, style: 'width:90px' }))),
      h('div', { class: 'grid2' },
        field('🤲 الأذكار', lines('dhikr')),
        field('⭐ عبارات تعزيزية وتحفيزية (الانتظام، الانضباط، النظافة...)', lines('motivation')))));

    const audios = CFG.files.filter(f => f.kind === 'audio');
    sec.append(panel('نطقها بصوت خافت',
      h('p', { class: 'hint' }, 'ينطق البرنامج عبارة عشوائية بصوت منخفض وسرعة أعلى كل عدة دقائق. يُفضّل اختيار "الفسحة وبين الحصص فقط" حتى لا يزعج الحصص. لا تُنطق قبل الجرس بدقيقة ولا أثناء إعلان آخر.'),
      h('div', { class: 'row' },
        chk(P, 'speak', 'تفعيل النطق'),
        field('متى', bindSelect(P, 'speakWhen', [['breaks', 'الطابور والفسحة وبين الحصص فقط'], ['always', 'طوال الدوام (حتى أثناء الحصص)']])),
        field('كل (دقيقة)', bindInput(P, 'speakEvery', { type: 'number', min: 1, style: 'width:80px' })),
        field('مستوى الصوت', bindInput(P, 'speakVolume', { type: 'range', min: 0.05, max: 1, step: 0.05 })),
        field('السرعة', bindInput(P, 'speakRate', { type: 'range', min: 0.7, max: 1.8, step: 0.05 }))),
      h('div', { class: 'row' },
        chk(P, 'speakDhikr', 'نطق الأذكار'), chk(P, 'speakMotivation', 'نطق العبارات التحفيزية'),
        testBtn(() => Phrases.speakOne(), '▶ تجربة')),
      h('p', { class: 'hint' }, 'ملفات صوتية تُشغَّل ضمن العبارات (مثلاً أذكار مسجلة بصوت قارئ) — ارفعها من تبويب "الأصوات والإعلانات" ثم حددها هنا:'),
      audios.length ? h('div', { class: 'row' }, audios.map(f => h('label', { class: 'chk' },
        h('input', { type: 'checkbox', checked: P.audios.includes(f.id), onchange: e => { P.audios = e.target.checked ? [...P.audios, f.id] : P.audios.filter(x => x !== f.id); save(); } }), f.name)))
        : h('p', { class: 'hint' }, '(لا توجد ملفات صوتية مرفوعة بعد)')));
  },

  duty(sec) { sourceTab(sec, 'duty'); },
  absent(sec) { sourceTab(sec, 'absent'); },
  subs(sec) { sourceTab(sec, 'subs'); },

  /* ---------------- النسخ الاحتياطي ---------------- */
  backup(sec) {
    sec.append(panel('النسخ الاحتياطي والنقل لجهاز آخر',
      h('p', { class: 'hint' }, 'كل البيانات محفوظة على هذا الجهاز فقط. احفظ نسخة احتياطية (ملف واحد يشمل الإعدادات والأصوات والملفات) لنقلها لجهاز آخر أو استرجاعها عند الحاجة.'),
      h('div', { class: 'row' },
        h('button', { class: 'btn pri', onclick: exportBackup }, '⬇ حفظ نسخة احتياطية'),
        h('button', { class: 'btn', onclick: importBackup }, '⬆ استرجاع نسخة احتياطية'))));
    sec.append(panel('إعادة الضبط',
      h('button', { class: 'btn dan', onclick: async () => {
        if (!confirm('سيتم حذف كل الإعدادات والملفات والرجوع للوضع الافتراضي. هل أنت متأكد؟')) return;
        for (const f of [...CFG.files]) await removeFile(f.id);
        localStorage.removeItem(LS_KEY); location.reload();
      } }, 'حذف كل البيانات')));
    sec.append(panel('اختصارات لوحة المفاتيح',
      h('p', { class: 'hint' }, 'S الإعدادات — F ملء الشاشة — B جرس يدوي — M إيقاف/تشغيل الجرس التلقائي — Esc إيقاف الصوت الجاري — → الشريحة التالية')));
  },
};

/* إعدادات المصدر المشتركة (رابط أي نوع / ملف / يدوي) */
const SRC_INFO = {
  duty: { title: 'المناوبة', hint: 'يعرض البرنامج تلقائياً من المناوب الآن (حسب اليوم والحصة/الفسحة والأسبوع). الأعمدة المتوقعة: الأسبوع (اختياري)، اليوم، الفترة/الحصة، المعلم المناوب، المكان — أو جدول الأيام فيه أعمدة (الأحد، الاثنين...) والفترات صفوف.' },
  subs: { title: 'حصص الاحتياط', hint: 'يعرض حصص اليوم ويربطها بالصفوف في مشهد "الحصة الآن". الأعمدة المتوقعة: التاريخ (أو اليوم)، الحصة، الصف، المعلم الغائب، المعلم البديل/الاحتياط.' },
  absent: { title: 'الطلاب الغائبون', hint: 'يعرض غياب اليوم في مشهد خاص مرتباً حسب الصف. الأعمدة المتوقعة: اسم الطالب، الصف، التاريخ (أو اليوم)، السبب (اختياري). يمكن ربطه بشيت يُحدَّث يومياً، وتظهر فقط صفوف تاريخ اليوم.' },
};
function linkPanel(S, which) {
  const result = h('div');
  return h('div', {},
    h('div', { class: 'row' },
      field('الرابط — أي رابط', bindInput(S, 'url', { type: 'url', style: 'width:560px;max-width:100%;direction:ltr', placeholder: 'Google Sheets / Apps Script / Canva / PDF / Excel / صفحة ويب...' })),
      field('اسم الورقة (اختياري)', bindInput(S, 'sheet', { style: 'width:130px' })),
      field('التحديث كل (دقيقة)', bindInput(S, 'refreshMin', { type: 'number', min: 1, style: 'width:80px' })),
      which === 'subs' ? field('رابط أسماء المعلمين (اختياري — عند وجود رموز بدل الأسماء)', bindInput(S, 'namesUrl', { type: 'url', style: 'width:330px;direction:ltr', placeholder: 'يُكتشف تلقائياً من ورقة Teachers' })) : null),
    h('div', { class: 'row' },
      h('button', { class: 'btn pri', onclick: async () => { await refreshSource(which, true); rerender(); } }, '⟳ قراءة البيانات الآن'),
      h('button', { class: 'btn', onclick: async e => {
        result.innerHTML = ''; result.append(h('p', { class: 'hint' }, '⏳ جارٍ فحص الرابط...'));
        try {
          const res = await smartRead(S.url);
          const kind = { pdf: 'ملف PDF', html: 'صفحة ويب', sheets: res.via === 'iframe' ? 'صفحة ويب (بعد تحميلها)' : 'جدول بيانات' }[res.type];
          let rows = [], perr = '';
          try { rows = await rowsFromSmart(which, res, S.sheet); if (which === 'subs') await resolveTeacherIds(rows, S); } catch (x) { perr = x.message; }
          const first = res.type === 'pdf' ? (await pdfToSheets(res.blob))[0] : res.sheets[0];
          result.innerHTML = '';
          result.append(h('p', { class: 'hint' }, `✅ تمت قراءة الرابط: ${kind}. ` + (rows.length ? `تم التعرف على ${rows.length} سجل.` : '⚠️ لكن لم أتعرف على أعمدة البيانات: ' + perr)),
            first ? h('div', { class: 'preview-box' }, h('table', {}, first.aoa.slice(0, 12).map(r => h('tr', {}, r.slice(0, 8).map(c => h('td', {}, c)))))) : null);
        } catch (x) { result.innerHTML = ''; result.append(h('p', { class: 'hint', style: 'color:#fca5a5' }, '❌ ' + x.message)); }
      } }, '🔍 فحص الرابط'),
      h('label', { class: 'chk' }, bindInput(S, 'showScene', { type: 'checkbox' }), '🖥️ عرض الرابط نفسه كمشهد في الشاشة')),
    h('p', { class: 'hint' }, 'يقبل: Google Sheets (شارك "لأي شخص لديه الرابط")، Google Docs، ملف في Google Drive، رابط Apps Script (الأفضل أن يُرجع JSON أو CSV)، ملف PDF أو Excel أو CSV، أو أي صفحة ويب فيها جدول. صفحات Canva لا تسمح عادةً بقراءة بياناتها، ففعّل "عرض الرابط كمشهد" ليظهر كما هو في الشاشة. البيانات تُحفظ بعد كل قراءة فتبقى ظاهرة إن انقطع الإنترنت.'),
    h('div', { class: 'row' },
      S.lastFetch ? h('span', { class: 'badge ok' }, 'آخر تحديث: ' + S.lastFetch) : null,
      S.lastError ? h('span', { class: 'badge err' }, 'خطأ: ' + S.lastError) : null),
    result);
}
async function importFileTo(which, S) {
  const f = await pickFile('.pdf,.xlsx,.xls,.csv,.json,.html,.htm,.txt'); if (!f) return;
  try {
    toast('جارٍ قراءة الملف...');
    const rows = await rowsFromSmart(which, await smartFromBlob(f, f.name), S.sheet);
    if (!rows.length) throw new Error('لم يُعثر على بيانات');
    if (S.source === 'link') {
      // مع الرابط: يُضاف الملف إلى بيانات الرابط ولا يُستبدل بها عند التحديث
      S.fileRows = rows; S.rows = (S.linkRows || []).concat(rows);
      toast(`تمت إضافة ${rows.length} سجل من الملف إلى بيانات الرابط`);
    } else { S.rows = rows; toast('تم استيراد ' + rows.length + ' سجل'); }
    save(); rerender();
  } catch (e) { toast(e.message, true); }
}
function sourceTab(sec, which) {
  const S = CFG[which], I = SRC_INFO[which];
  const isDuty = which === 'duty', isAbs = which === 'absent';
  sec.append(panel('مصدر بيانات ' + I.title,
    h('p', { class: 'hint' }, I.hint),
    h('div', { class: 'row' },
      field('المصدر', bindSelect(S, 'source', [['manual', 'إدخال يدوي / استيراد ملف'], ['link', 'رابط (أي نوع)']], rerender))),
    S.source === 'link' ? linkPanel(S, which) : null,
    h('div', { class: 'row' },
      h('button', { class: 'btn', onclick: () => importFileTo(which, S) }, '📥 استيراد ملف (PDF / Excel / CSV)'),
      h('button', { class: 'btn', onclick: () => downloadTemplate(which) }, '⬇ تحميل نموذج فارغ'),
      S.fileRows?.length ? h('button', { class: 'btn dan', onclick: () => { S.rows = (S.linkRows || []).slice(); S.fileRows = []; save(); rerender(); } }, `حذف بيانات الملف المرفوع (${S.fileRows.length})`) : null,
      S.rows.length ? h('button', { class: 'btn dan', onclick: () => { if (confirm('مسح كل السجلات؟')) { S.rows = []; S.fileRows = []; S.linkRows = []; save(); rerender(); } } }, 'مسح الكل') : null)));
  if (isDuty) sec.append(panel('خيارات العرض',
    h('div', { class: 'row' }, chk(CFG.display, 'dutyStrip', '🛡️ شريط "المناوبة الآن" أسفل الشاشة'),
      h('label', { class: 'chk' }, h('input', { type: 'checkbox', checked: CFG.display.subsStrip !== false, onchange: e => { CFG.display.subsStrip = e.target.checked; save(); } }), '🔁 إظهار "الاحتياط الآن" في نفس الشريط'),
      field('مدة كل صفحة في الشريط (ث)', bindInput(CFG.display, 'dutyStripSec', { type: 'number', min: 3, style: 'width:80px' }))),
    h('div', { class: 'row' }, chk(CFG.display, 'dutyAllWeeks', 'إظهار مناوبات كل الأسابيع لنفس اليوم (مع اسم الأسبوع بجانب كل معلمة) — بدل الأسبوع الحالي فقط'))));
  if (isAbs) sec.append(panel('خيارات العرض',
    h('div', { class: 'row' },
      chk(S, 'showNames', 'إظهار أسماء الطلاب (إلغاؤه يعرض عدد الغائبين لكل صف فقط)'),
      chk(S, 'datedOnly', 'تجاهل الصفوف التي ليس لها تاريخ'),
      chk(CFG.autoSlides, 'absent', 'مشهد الغياب في الشاشة الرئيسية'))));

  // معاينة اليوم
  const d = now();
  const todayRows = isDuty
    ? S.rows.filter(r => (r.date ? r.date === ymd(d) : (r.day === -1 || r.day === d.getDay())) && weekOK(r, d))
    : isAbs ? absentFor(d) : subsFor(d);
  sec.append(panel(`ما سيظهر اليوم (${DAYS[d.getDay()]}${isDuty && CFG.general.cycleWeeks > 1 ? ' — الأسبوع ' + cycleWeek(d) : ''}) — ${todayRows.length}`,
    todayRows.length ? h('div', { class: 'preview-box' }, h('table', {},
      todayRows.map(r => h('tr', {}, (isDuty ? [r.slot || 'طوال اليوم', r.teacher, r.location] : isAbs ? [r.cls, r.name, r.reason] : [r.period, r.cls, r.sub, r.absent]).map(x => h('td', {}, x || ''))))))
      : h('p', { class: 'hint' }, 'لا يوجد شيء لليوم.')));

  // محرر يدوي
  const tb = h('tbody');
  const dayOpts = [['-1', 'كل الأيام'], ...DAYS.map((d, i) => [String(i), d])];
  const wkOpts = [['0', 'كل الأسابيع'], ...Array.from({ length: Math.max(1, +CFG.general.cycleWeeks || 1) }, (_, i) => [String(i + 1), 'الأسبوع ' + (i + 1)])];
  const shown = S.rows.slice(0, 400);
  shown.forEach((r, i) => {
    const numSel = (key, opts) => { const el = h('select', { onchange: e => { r[key] = +e.target.value; save(); } }, opts.map(([v, t]) => h('option', { value: v }, t))); el.value = String(r[key] ?? (key === 'day' ? -1 : 0)); return el; };
    const del = h('button', { class: 'btn dan sm', onclick: () => { S.rows.splice(i, 1); save(); rerender(); } }, '✕');
    if (isDuty) tb.append(h('tr', {},
      h('td', {}, numSel('week', wkOpts)), h('td', {}, numSel('day', dayOpts)),
      h('td', {}, bindInput(r, 'date', { type: 'date' })),
      h('td', {}, bindInput(r, 'slot', { placeholder: 'الفسحة / الحصة الأولى / طوال اليوم' })),
      h('td', {}, bindInput(r, 'teacher')), h('td', {}, bindInput(r, 'location')), h('td', {}, del)));
    else if (isAbs) tb.append(h('tr', {},
      h('td', {}, bindInput(r, 'date', { type: 'date' })), h('td', {}, numSel('day', dayOpts)),
      h('td', {}, bindInput(r, 'cls', { placeholder: '5/1' })), h('td', {}, bindInput(r, 'name')), h('td', {}, bindInput(r, 'reason')), h('td', {}, del)));
    else tb.append(h('tr', {},
      h('td', {}, bindInput(r, 'date', { type: 'date' })), h('td', {}, numSel('day', dayOpts)),
      h('td', {}, bindInput(r, 'period', { placeholder: 'الأولى / 1' })), h('td', {}, bindInput(r, 'cls')),
      h('td', {}, bindInput(r, 'absent')), h('td', {}, bindInput(r, 'sub')), h('td', {}, bindInput(r, 'note')), h('td', {}, del)));
  });
  const heads = isDuty ? ['الأسبوع', 'اليوم', 'تاريخ محدد (اختياري)', 'الفترة / الحصة', 'المعلم المناوب', 'المكان', '']
    : isAbs ? ['التاريخ', 'أو اليوم', 'الصف', 'اسم الطالب', 'السبب / ملاحظة', '']
      : ['التاريخ', 'أو اليوم', 'الحصة', 'الصف', 'المعلم الغائب', 'المعلم البديل', 'ملاحظة', ''];
  sec.append(panel(`السجلات (${S.rows.length})` + (S.source === 'link' ? ' — ستُستبدل عند التحديث من الرابط' : ''),
    h('div', { class: 'tbl-wrap' }, h('table', { class: 'edit' }, h('thead', {}, h('tr', {}, heads.map(t => h('th', {}, t)))), tb)),
    S.rows.length > shown.length ? h('p', { class: 'hint' }, `يُعرض أول ${shown.length} سجل فقط للتعديل.`) : null,
    h('div', { class: 'row', style: 'margin-top:10px' },
      h('button', { class: 'btn pri', onclick: () => {
        S.rows.push(isDuty ? { week: 0, day: now().getDay(), date: '', slot: 'الفسحة', teacher: '', location: '' }
          : isAbs ? { date: ymd(now()), day: -1, cls: '', name: '', reason: '' }
            : { date: ymd(now()), day: -1, period: '', cls: '', absent: '', sub: '', note: '' });
        save(); rerender();
      } }, '+ إضافة سجل'))));
}
let ttDay = null;
function timetableTab(sec) {
  const T = CFG.timetable;
  sec.append(panel('جدول المعلمين (من يدرّس كل صف في كل حصة)',
    h('p', { class: 'hint' }, 'هذا الجدول يجعل الشاشة تعرض في كل حصة: الصف ← المادة ← المعلم، ومعه الاحتياط (الغائب ← البديل). 📄 يمكن رفعه كملف PDF (المصدَّر من برنامج الجداول — صفحة لكل صف أو لكل معلم أو جدول عام)، ثم مراجعته في الجدول بالأسفل. أو من Excel: الصفوف في الأسطر والحصص (الأولى، الثانية...) في الأعمدة، وكل ورقة باسم يوم (الأحد، الاثنين...)، والخلية مثل "رياضيات - أ. فاطمة". أو جدول المعلمين (المعلم في السطر والصف في الخلية). أو أدخله يدوياً بالأسفل.'),
    h('div', { class: 'row' },
      field('المصدر', bindSelect(T, 'source', [['manual', 'إدخال يدوي / استيراد ملف'], ['link', 'رابط (أي نوع)']], rerender))),
    T.source === 'link' ? linkPanel(T, 'timetable') : null,
    h('div', { class: 'row' },
      h('button', { class: 'btn', onclick: async () => {
        const f = await pickFile('.pdf,.xlsx,.xls,.csv'); if (!f) return;
        const isPdf = kindOf(f) === 'pdf';
        try {
          toast('جارٍ قراءة الملف...');
          const rows = isPdf ? (await parseTimetablePDF(f)).rows : parseTimetable(await sheetsFromBlob(f), T.sheet);
          const nCls = new Set(rows.map(r => normCls(r.cls))).size, nT = new Set(rows.map(r => r.teacher).filter(Boolean)).size;
          if (!confirm(`تمت قراءة ${rows.length} حصة: ${nCls} صف و ${nT} معلم.\nهل تريد اعتماد هذا الجدول؟ (سيستبدل الجدول الحالي — يمكنك مراجعته وتعديله بالأسفل)`)) return;
          T.rows = rows;
          T.classes = [...new Set(rows.map(r => r.cls).filter(Boolean))].sort(clsSorter());
          T.lastImport = { name: f.name, date: new Date().toLocaleString('ar'), periods: rows.length, classes: nCls, teachers: nT };
          save(); toast('تم استيراد ' + rows.length + ' حصة'); rerender();
        } catch (e) { toast(e.message, true); }
        if (isPdf && confirm('هل تريد أيضاً عرض ملف الـ PDF نفسه كشريحة في الشاشة الرئيسية؟')) {
          const id = await addFile(f, 'pdf');
          CFG.slides.push({ id: uid(), type: 'pdf', fileId: id, enabled: true, title: 'جدول الحصص', duration: 20, muted: true, fit: 'contain' });
          save(); Slides.rebuild();
        }
      } }, '📥 استيراد جدول (PDF / Excel / CSV)'),
      h('button', { class: 'btn', onclick: () => downloadTemplate('timetable') }, '⬇ تحميل نموذج'),
      T.rows.length ? h('button', { class: 'btn dan', onclick: () => { if (confirm('مسح جدول المعلمين كاملاً؟')) { T.rows = []; save(); rerender(); } } }, 'مسح الكل') : null,
      h('span', { class: 'badge' }, digits(T.rows.length) + ' حصة مسجلة')),
    T.lastImport ? h('p', { class: 'hint', style: 'margin-top:8px' }, `📄 آخر جدول مرفوع: ${T.lastImport.name} — ${T.lastImport.date} (${digits(T.lastImport.periods)} حصة، ${digits(T.lastImport.classes)} صف، ${digits(T.lastImport.teachers)} معلمة). عند تغيّر الجدول ارفعي الملف الجديد فقط، فيحل محل القديم.`) : null));

  const clsTa = h('textarea', { style: 'min-height:90px', placeholder: '5/1\n5/2\n6/1', onchange: e => { T.classes = e.target.value.split('\n').map(x => x.trim()).filter(Boolean); save(); rerender(); } }, T.classes.join('\n'));
  sec.append(panel('الصفوف (كل سطر صف — بنفس الكتابة المستخدمة في جدول الاحتياط)', clsTa,
    h('div', { class: 'row', style: 'margin-top:8px' }, h('button', { class: 'btn sm', onclick: () => { T.classes = [...new Set(T.rows.map(r => r.cls))].sort(clsSorter()); save(); rerender(); } }, 'استخراج الصفوف من الجدول'))));

  if (ttDay == null) ttDay = CFG.general.workDays.includes(now().getDay()) ? now().getDay() : (CFG.general.workDays[0] ?? 0);
  const pr = CFG.override ? CFG.profiles.find(p => p.id === CFG.override) : (CFG.profiles.find(p => p.days.includes(ttDay)) || CFG.profiles[0]);
  const lessons = sortedPeriods(pr).filter(isLesson);
  const classes = T.classes.length ? T.classes : [...new Set(T.rows.map(r => r.cls))].sort(clsSorter());
  const find = (cls, p) => T.rows.find(r => (r.day === ttDay || r.day === -1) && slotKey(r.period) === slotKey(periodName(p)) && normCls(r.cls) === normCls(cls));
  const upd = (cls, p, key, val) => {
    let r = find(cls, p);
    if (!r) { r = { week: 0, day: ttDay, period: periodName(p), cls, subject: '', teacher: '' }; T.rows.push(r); }
    r[key] = val.trim();
    if (!r.teacher && !r.subject) T.rows.splice(T.rows.indexOf(r), 1);
    save(false);
  };
  const tb = h('tbody');
  for (const cls of classes) tb.append(h('tr', {}, h('td', { class: 'cl' }, cls), lessons.map(p => {
    const r = find(cls, p);
    return h('td', {},
      h('input', { type: 'text', value: r?.teacher || '', placeholder: 'المعلم', onchange: e => upd(cls, p, 'teacher', e.target.value) }),
      h('input', { type: 'text', class: 'sj', value: r?.subject || '', placeholder: 'المادة', onchange: e => upd(cls, p, 'subject', e.target.value) }));
  })));
  sec.append(panel('تعديل يدوي',
    h('div', { class: 'row' }, DAYS.map((d, i) => CFG.general.workDays.includes(i) ? h('button', { class: 'btn' + (i === ttDay ? ' pri' : ''), onclick: () => { ttDay = i; rerender(); } }, d) : null),
      h('button', { class: 'btn sm', onclick: () => {
        if (!confirm('نسخ جدول ' + DAYS[ttDay] + ' إلى باقي أيام الدوام؟ (سيستبدل الموجود)')) return;
        const src = T.rows.filter(r => r.day === ttDay);
        T.rows = T.rows.filter(r => r.day === ttDay || !CFG.general.workDays.includes(r.day));
        for (const d of CFG.general.workDays) if (d !== ttDay) src.forEach(r => T.rows.push({ ...r, day: d }));
        save(); rerender(); toast('تم النسخ');
      } }, 'نسخ هذا اليوم لكل الأيام')),
    !classes.length ? h('p', { class: 'hint' }, 'اكتب أسماء الصفوف أولاً في المربع أعلاه.') :
      !lessons.length ? h('p', { class: 'hint' }, 'لا توجد حصص في جدول التوقيت لهذا اليوم.') :
        h('div', { class: 'tbl-wrap' }, h('table', { class: 'edit tt-grid' },
          h('thead', {}, h('tr', {}, h('th', { class: 'cl' }, 'الصف'), lessons.map(p => h('th', {}, periodName(p))))), tb))));
}

function downloadTemplate(which) {
  if (which === 'absent') {
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['التاريخ', 'الصف', 'اسم الطالب', 'السبب'], [ymd(now()), '5/1', 'فاطمة سالم', 'مرض'], [ymd(now()), '6/2', 'مريم علي', '']]), 'الغياب');
    XLSX.writeFile(wb, 'نموذج-الغياب.xlsx'); return;
  }
  if (which === 'timetable') {
    const wb = XLSX.utils.book_new();
    for (const d of CFG.general.workDays) {
      const rows = [['الصف', 'الأولى', 'الثانية', 'الثالثة', 'الرابعة', 'الخامسة', 'السادسة', 'السابعة', 'الثامنة'],
        ['5/1', 'رياضيات - أ. فاطمة', 'لغة عربية - أ. مريم', 'علوم - أ. سارة', '', '', '', '', ''],
        ['5/2', 'لغة عربية - أ. مريم', 'رياضيات - أ. فاطمة', 'إنجليزي - أ. هدى', '', '', '', '', '']];
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), DAYS[d]);
    }
    XLSX.writeFile(wb, 'نموذج-جدول-المعلمين.xlsx');
    return;
  }
  const rows = which === 'duty'
    ? [['الأسبوع', 'اليوم', 'الفترة', 'المعلم المناوب', 'المكان'], ['1', 'الأحد', 'الطابور', 'أ. محمد', 'الساحة'], ['1', 'الأحد', 'الفسحة', 'أ. أحمد', 'المقصف'], ['1', 'الأحد', 'الحصة الأولى', 'أ. سالم', 'الممر'], ['2', 'الأحد', 'الفسحة', 'أ. خالد', 'الساحة']]
    : [['التاريخ', 'الحصة', 'الصف', 'المعلم الغائب', 'المعلم البديل'], [ymd(now()), 'الأولى', '5/1', 'أ. علي', 'أ. يوسف'], [ymd(now()), 'الثالثة', '7/2', 'أ. علي', 'أ. ناصر']];
  const ws = XLSX.utils.aoa_to_sheet(rows); const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, which === 'duty' ? 'المناوبة' : 'الاحتياط');
  XLSX.writeFile(wb, which === 'duty' ? 'نموذج-المناوبة.xlsx' : 'نموذج-الاحتياط.xlsx');
}

/* تسجيل صوت بالميكروفون */
let recorder = null;
async function recordAudio(btn, state) {
  if (recorder) { recorder.stop(); return; }
  try {
    const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
    const chunks = [];
    recorder = new MediaRecorder(stream);
    recorder.ondataavailable = e => chunks.push(e.data);
    recorder.onstop = async () => {
      stream.getTracks().forEach(t => t.stop());
      const blob = new Blob(chunks, { type: recorder.mimeType || 'audio/webm' });
      recorder = null;
      const name = prompt('اسم التسجيل:', 'تسجيل ' + new Date().toLocaleTimeString('ar')) || 'تسجيل';
      const f = new File([blob], name, { type: blob.type });
      await addFile(f, 'audio'); rerender(); toast('تم حفظ التسجيل');
    };
    recorder.start();
    btn.textContent = '⏹ إيقاف التسجيل'; state.innerHTML = '<span class="rec-dot"></span> جارٍ التسجيل...';
  } catch (e) { toast('تعذر الوصول للميكروفون: ' + e.message, true); recorder = null; }
}

/* نسخ احتياطي */
const blobToDataURL = b => new Promise(r => { const fr = new FileReader(); fr.onload = () => r(fr.result); fr.readAsDataURL(b); });
async function exportBackup() {
  toast('جارٍ تجهيز النسخة...');
  const files = {};
  for (const f of CFG.files) { const b = await DB.get(f.id); if (b) files[f.id] = await blobToDataURL(b); }
  const data = JSON.stringify({ app: 'school-bell', version: 1, date: new Date().toISOString(), config: CFG, files });
  const a = h('a', { href: URL.createObjectURL(new Blob([data], { type: 'application/json' })), download: `نسخة-الجرس-${ymd(new Date())}.json` });
  document.body.append(a); a.click(); a.remove();
}
async function importBackup() {
  const f = await pickFile('.json,application/json'); if (!f) return;
  try {
    const data = JSON.parse(await f.text());
    if (data.app !== 'school-bell') throw new Error('الملف ليس نسخة احتياطية للجرس');
    if (!confirm('سيتم استبدال الإعدادات الحالية بالنسخة. متابعة؟')) return;
    await applyBackup(data);
    location.reload();
  } catch (e) { toast('فشل الاسترجاع: ' + e.message, true); }
}
async function applyBackup(data) {
  for (const [id, du] of Object.entries(data.files || {})) await DB.put(id, await (await fetch(du)).blob());
  CFG = mergeDefaults(defaultConfig(), data.config);
  localStorage.setItem(LS_KEY, JSON.stringify(CFG));
}
/* عند النشر على رابط: إن وُضع ملف config.json بجانب البرنامج يُحمَّل تلقائياً في أول فتح لجهاز جديد */
async function autoConfig() {
  if (!/^https?:/.test(location.protocol)) return false;
  try {
    const r = await fetch('config.json', { cache: 'no-store' }); if (!r.ok) return false;
    const data = await r.json(); if (data.app !== 'school-bell') return false;
    await applyBackup(data); return true;
  } catch { return false; }
}

/* ------------------------------ التشغيل ------------------------------ */
let started = false;
function start() {
  started = true;
  $('#startOverlay').hidden = true;
  Audio_.unlock();
  keepAwake();
}
/* إبقاء الشاشة مضاءة (الآيباد والأجهزة اللوحية) */
let wakeLock = null;
async function keepAwake() {
  try { if ('wakeLock' in navigator && !wakeLock) { wakeLock = await navigator.wakeLock.request('screen'); wakeLock.addEventListener('release', () => wakeLock = null); } } catch { }
}
document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible' && started) { keepAwake(); Audio_.ensure(); } });
function toggleFull() {
  const d = document, el = d.documentElement;
  if (d.fullscreenElement || d.webkitFullscreenElement) (d.exitFullscreen || d.webkitExitFullscreen).call(d);
  else { const r = (el.requestFullscreen || el.webkitRequestFullscreen)?.call(el); r?.catch?.(() => { }); }
}
function toggleBell() { CFG.general.bellEnabled = !CFG.general.bellEnabled; save(); toast(CFG.general.bellEnabled ? 'تم تشغيل الجرس التلقائي' : 'تم إيقاف الجرس التلقائي'); }

async function init() {
  const fresh = !localStorage.getItem(LS_KEY);
  loadConfig();
  try { await DB.open(); } catch (e) { toast('تعذر فتح قاعدة البيانات المحلية: ' + e.message, true); }
  if (fresh && await autoConfig()) { location.reload(); return; }
  try { navigator.storage?.persist?.(); } catch { }
  if (window.pdfjsLib) pdfjsLib.GlobalWorkerOptions.workerSrc = 'lib/pdf.worker.min.js';
  if (window.speechSynthesis) speechSynthesis.onvoiceschanged = () => { if (curTab === 'sounds' && !$('#settings').hidden) rerender(); };

  applyTheme(); renderClock(); renderSide(true); updateTicker(true); Slides.rebuild();

  // هل يسمح المتصفح بالصوت تلقائياً؟ (عند التشغيل بملف start-bell.bat نعم)
  try {
    const c = Audio_.ensure();
    await sleep(300);
    if (c.state === 'running') start(); else $('#startOverlay').hidden = false;
  } catch { $('#startOverlay').hidden = false; }
  $('#startBtn').onclick = start;

  setInterval(() => {
    const prev = CUR_SHIFT; syncShift();
    if (prev !== CUR_SHIFT && $('#settings').hidden) { renderDates(); renderSide(true); Slides.rebuild(); }
    renderClock(); renderSide(); checkEvents(); Phrases.tick(); Slides.watch(); DutyStrip.build();
  }, 1000);
  Phrases.renderCard();
  setInterval(maybeRefresh, 30000); maybeRefresh();

  // أزرار التحكم
  $('#gear').onclick = () => openSettings();
  $('#cFull').onclick = toggleFull;
  $('#cBell').onclick = () => { Audio_.ensure(); manualBell(); };
  $('#cMute').onclick = toggleBell;
  $('#cStop').onclick = () => Audio_.stop();
  $('#cNext').onclick = () => { if (Slides.frozen) Slides.toggleFreeze(); else Slides.next(true); };
  $('#cPin').onclick = () => Slides.toggleFreeze();
  $('#cQuick').onclick = toggleQuickMenu;
  $('#closeSettings').onclick = closeSettings;
  $$('#tabs button').forEach(b => b.onclick = () => showTab(b.dataset.tab));
  $('#banner').onclick = hideBanner;

  let mt;
  document.addEventListener('mousemove', () => { document.body.classList.add('mouse'); clearTimeout(mt); mt = setTimeout(() => { if ($('#settings').hidden) document.body.classList.remove('mouse'); }, 3000); });
  document.addEventListener('keydown', e => {
    if (!started) { start(); return; }
    if (!$('#settings').hidden) { if (e.key === 'Escape') closeSettings(); return; }
    const k = e.key.toLowerCase();
    if (k === 's' || k === 'س') openSettings();
    else if (k === 'f' || k === 'ب') toggleFull();
    else if (k === 'b' || k === 'لا') manualBell();
    else if (k === 'm' || k === 'ة') toggleBell();
    else if (k === 'escape') Audio_.stop();
    else if (k === 'arrowright' || k === 'arrowleft') Slides.next(true);
    else if (k === 'a' || k === 'ش') toggleQuickMenu();
    else if (k === 'p' || k === 'ح') Slides.toggleFreeze();
    else if (/^[1-9]$/.test(toLatin(e.key)) && CFG.quick[+toLatin(e.key) - 1]) quickAnnounce(CFG.quick[+toLatin(e.key) - 1]);
  });
  window.addEventListener('storage', e => { if (e.key === LS_KEY) { loadConfig(); onConfigChanged(); Slides.rebuild(); } });
}
init();
