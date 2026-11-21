/* =====================================================================
   الجرس المدرسي الذكي — يعمل محلياً على جهاز واحد بدون خادم.
   الإعدادات تُحفظ في localStorage والملفات (PDF/صوت/فيديو/صور) في IndexedDB.
   ===================================================================== */
'use strict';

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
  for (let i = ords.length - 1; i >= 0; i--) if (n.includes(ords[i])) return 'p' + (i + 1);
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
    id: 'pr_main', name: 'الجدول الأساسي', days: [0, 1, 2, 3, 4],
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
    },
    profiles: [base],
    override: '', // معرّف جدول يُفرض على كل الأيام (مثلاً جدول رمضان/الاختبارات)
    voice: { tts: true, voiceName: '', rate: 0.9, repeat: 1, startTpl: 'بدأت الآن {name}', endTpl: 'انتهت الآن {name}', joinTpl: '، و', gapMs: 400 },
    reminders: [
      { id: uid(), enabled: true, name: 'منتصف الحصة', mode: 'mid', minutes: 5, time: '', applies: 'lessons', sound: 'chime', text: 'اللهم صلِّ وسلم على نبينا محمد', audio: '' },
      { id: uid(), enabled: false, name: 'قبل نهاية الحصة', mode: 'beforeEnd', minutes: 5, time: '', applies: 'lessons', sound: 'soft', text: 'تبقى خمس دقائق على نهاية الحصة', audio: '' },
    ],
    slides: [
      { id: uid(), type: 'text', enabled: true, title: '', text: 'أهلاً وسهلاً بكم في مدرستنا\nالعلم نور', duration: 12 },
    ],
    autoSlides: { schedule: true, duty: true, subs: true, duration: 15 },
    ticker: { enabled: true, label: 'أخبار المدرسة', texts: ['أهلاً بكم في مدرستنا', 'النظافة من الإيمان', 'الالتزام بالوقت أساس النجاح'], speed: 90, includeDuty: true, includeSubs: true },
    display: { showSide: true, showDuty: true, showSubs: true },
    duty: { source: 'manual', url: '', sheet: '', refreshMin: 15, lastFetch: '', lastError: '', rows: [] },
    subs: { source: 'manual', url: '', sheet: '', refreshMin: 10, lastFetch: '', lastError: '', rows: [] },
    files: [],
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
function loadConfig() {
  let cur = null;
  try { cur = JSON.parse(localStorage.getItem(LS_KEY)); } catch { }
  CFG = cur ? mergeDefaults(defaultConfig(), cur) : defaultConfig();
}
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
  async playFile(id) {
    const url = await fileURL(id);
    if (!url || this.stopFlag) return;
    const a = new Audio(url); a.volume = CFG.general.volume; this.current = a;
    await new Promise(res => {
      a.onended = a.onerror = res; a.play().catch(res);
      setTimeout(res, 5 * 60 * 1000);
    });
    this.current = null;
  },
  voices() { return (window.speechSynthesis?.getVoices() || []); },
  arabicVoices() { return this.voices().filter(v => /^ar/i.test(v.lang)); },
  speak(text) {
    return new Promise(res => {
      if (!text || !window.speechSynthesis || this.stopFlag) return res();
      const u = new SpeechSynthesisUtterance(text);
      const vs = this.voices();
      const v = vs.find(v => v.name === CFG.voice.voiceName) || this.arabicVoices()[0];
      if (v) { u.voice = v; u.lang = v.lang; } else u.lang = 'ar-SA';
      u.rate = +CFG.voice.rate || 0.9; u.volume = Math.min(1, CFG.general.volume + 0.1);
      const done = () => { clearTimeout(tm); res(); };
      const tm = setTimeout(done, 2500 + text.length * 180);
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
      Slides.duck(true);
      try {
        await Audio_.playSound(it.sound);
        for (let r = 0; r < Math.max(1, +CFG.voice.repeat || 1); r++) {
          for (const p of it.parts || []) {
            if (Audio_.stopFlag) break;
            if (p.audio) await Audio_.playFile(p.audio);
            else if (p.text && CFG.voice.tts) await Audio_.speak(p.text);
            await sleep(+CFG.voice.gapMs || 300);
          }
        }
      } catch (e) { console.error(e); }
      Slides.duck(false);
      if (it.banner) setTimeout(hideBanner, Math.max(1000, CFG.general.bannerSec * 1000 - 3000));
    }
    this.busy = false;
  },
};
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
  return CFG.profiles.find(p => p.days.includes(d.getDay())) || null;
}
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
  const pr = profileFor(d); if (!pr) return [];
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
      if (t != null && t > s && t < e) ev.push({ key: p.id + ':r:' + r.id, t, kind: 'rem', sound: r.sound, text: r.text, audio: r.audio, banner: r.text || r.name, icon: '📢', order: 2 });
    }
  }
  for (const r of CFG.reminders) {
    if (r.enabled && r.mode === 'fixed' && r.time) ev.push({ key: 'fix:' + r.id, t: toSec(r.time), kind: 'rem', sound: r.sound, text: r.text, audio: r.audio, banner: r.text || r.name, icon: '📢', order: 2 });
  }
  return ev.sort((a, b) => a.t - b.t || a.order - b.order);
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
    Queue.push({ sound, parts, banner, bannerIcon: g.list[0].icon || '🔔' });
  }
}
function manualBell() {
  const pr = profileFor(now());
  Queue.push({ sound: 'classic', parts: [], banner: 'جرس يدوي' });
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
    if (r.week && r.week !== wk) return false;
    const k = slotKey(r.slot);
    if (!keys) return k === '*';
    return k === '*' || keys.has(k);
  });
}
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
    title = periodName(st.cur); range = digits(`من ${fmt12(st.cur.start)} إلى ${fmt12(st.cur.end)}`);
    count = 'متبقٍ ' + digits(fmtDur(e - sec)); prog = (sec - s) / (e - s) * 100;
  } else if (st.next) {
    const first = st.periods[0] === st.next;
    title = first ? 'قبل بداية الدوام' : 'بين الفترات';
    count = 'يبدأ التالي بعد ' + digits(fmtDur(toSec(st.next.start) - sec));
  } else if (st.periods.length) {
    title = 'انتهى الدوام الدراسي'; range = 'نراكم غداً بإذن الله';
  } else title = 'لا يوجد جدول لهذا اليوم';
  $('#nowLabel').textContent = label; $('#nowTitle').textContent = title; $('#nowRange').textContent = range;
  $('#nowCount').textContent = count; $('#nowProg').style.width = prog + '%';
  $('#nextLine').innerHTML = st.next ? `${esc(periodName(st.next))} <span>${digits(fmt12(st.next.start))}</span>` : '<span>—</span>';

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
      const today = CFG.duty.rows.filter(r => (r.date ? r.date === ymd(d) : (r.day === -1 || r.day === d.getDay())) && (!r.week || r.week === cycleWeek(d)));
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
  [dl, sl].forEach(autoScroll);
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
function applyTicker(html) {
  const inner = $('#tickerInner');
  inner.innerHTML = html || '&nbsp;';
  inner.style.animation = 'none';
  requestAnimationFrame(() => {
    const w = inner.scrollWidth + inner.parentElement.clientWidth;
    inner.style.animation = `tickerMove ${Math.max(8, w / (+CFG.ticker.speed || 90))}s linear infinite`;
  });
}
$('#tickerInner').addEventListener('animationiteration', () => { if (tickerPending != null) { applyTicker(tickerPending); tickerPending = null; } });

/* ------------------------------ الشرائح ------------------------------ */
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
  let m;
  if ((m = url.match(/(?:youtube\.com\/(?:watch\?v=|shorts\/|embed\/)|youtu\.be\/)([\w-]{6,})/))) return { kind: 'iframe', src: `https://www.youtube.com/embed/${m[1]}?autoplay=1&mute=1&loop=1&playlist=${m[1]}&controls=0` };
  if ((m = url.match(/drive\.google\.com\/(?:file\/d\/|open\?id=)([\w-]+)/))) return { kind: 'iframe', src: `https://drive.google.com/file/d/${m[1]}/preview` };
  if (/\.(mp4|webm|ogv)(\?|$)/i.test(url)) return { kind: 'video', src: url };
  if (/\.(png|jpe?g|gif|webp|svg)(\?|$)/i.test(url)) return { kind: 'image', src: url };
  return { kind: 'iframe', src: url };
}
const Slides = {
  list: [], i: -1, timer: null, el: null, building: false, video: null,
  async rebuild() {
    if (this.building) { this.again = true; return; }
    this.building = true;
    const out = [];
    for (const s of CFG.slides) {
      if (!s.enabled) continue;
      if (s.type === 'pdf') {
        let pages = null;
        try { pages = await pdfPages(s); } catch (e) { console.warn(e); }
        if (pages && pages.length) pages.forEach((src, pi) => out.push({ ...s, type: 'image', src, title: s.title ? s.title + (pages.length > 1 ? ` (${pi + 1}/${pages.length})` : '') : '' }));
        else if (s.url) out.push({ ...s, type: 'url' });
      } else out.push(s);
    }
    const A = CFG.autoSlides;
    if (A.schedule) out.push({ type: 'auto-schedule', duration: A.duration });
    if (A.duty && CFG.duty.rows.length) out.push({ type: 'auto-duty', duration: A.duration });
    if (A.subs) out.push({ type: 'auto-subs', duration: A.duration });
    this.list = out;
    this.building = false;
    if (this.again) { this.again = false; return this.rebuild(); }
    if (this.i >= out.length || this.i < 0) this.i = -1;
    this.next();
  },
  skips: 0,
  skip() {
    if (++this.skips > this.list.length) { this.skips = 0; this.showEmpty(); this.timer = setTimeout(() => this.next(), 30000); return; }
    setTimeout(() => this.next(), 10);
  },
  showEmpty() {
    const stage = $('#stage'); stage.innerHTML = '';
    stage.append(h('div', { class: 'slide' }, h('div', { class: 'slide-empty' }, 'أضف جدول الحصص (PDF) أو صوراً أو فيديو من الإعدادات ← العرض والشريط')));
  },
  next() {
    clearTimeout(this.timer);
    const stage = $('#stage');
    if (!this.list.length) { stage.innerHTML = ''; stage.append(h('div', { class: 'slide' }, h('div', { class: 'slide-empty' }, 'أضف جدول الحصص (PDF) أو صوراً أو فيديو من الإعدادات ← العرض والشريط'))); return; }
    this.i = (this.i + 1) % this.list.length;
    const s = this.list[this.i];
    this.render(s).then(dur => {
      if (dur !== null) this.skips = 0;
      if (dur > 0) this.timer = setTimeout(() => this.next(), dur * 1000);
    });
  },
  async render(s) {
    const stage = $('#stage');
    let el, dur = +s.duration || 12;
    this.video = null;
    const title = s.title ? h('div', { class: 'slide-title' }, s.title) : null;
    if (s.type === 'image' || (s.type === 'file' && s.kind === 'image')) {
      const src = s.src || await fileURL(s.fileId);
      const img = h('img', { src, alt: '' });
      el = h('div', { class: 'slide' + (s.fit === 'scroll' ? ' scroll' : '') }, img, title);
      if (s.fit === 'scroll') img.onload = () => {
        const over = img.clientHeight - stage.clientHeight;
        if (over > 0) img.animate([{ transform: 'translateY(0)' }, { transform: 'translateY(0)', offset: 0.12 }, { transform: `translateY(${-over}px)`, offset: 0.88 }, { transform: `translateY(${-over}px)` }], { duration: dur * 1000, fill: 'forwards' });
      };
    } else if (s.type === 'video') {
      const src = s.fileId ? await fileURL(s.fileId) : s.url;
      const v = h('video', { src, autoplay: true, playsinline: true });
      v.muted = s.muted !== false; v.volume = CFG.general.volume;
      v.loop = !!(s.duration > 0);
      if (!(s.duration > 0)) { v.onended = () => this.next(); v.onerror = () => setTimeout(() => this.next(), 3000); dur = 0; }
      v.play().catch(() => { v.muted = true; v.play().catch(() => { }); });
      this.video = v;
      el = h('div', { class: 'slide' }, v, title);
    } else if (s.type === 'url') {
      const e = embedURL(s.url || '');
      if (e.kind === 'video') return this.render({ ...s, type: 'video', url: e.src });
      if (e.kind === 'image') return this.render({ ...s, type: 'image', src: e.src });
      el = h('div', { class: 'slide' }, h('iframe', { src: e.src, allow: 'autoplay' }), title);
    } else if (s.type === 'text') {
      el = h('div', { class: 'slide text-slide' }, s.text || '');
    } else if (s.type === 'auto-schedule') {
      el = this.scheduleSlide();
    } else if (s.type === 'auto-duty') {
      el = this.dutySlide();
    } else if (s.type === 'auto-subs') {
      el = this.subsSlide();
    }
    if (!el) { this.skip(); return null; }
    const dots = h('div', { class: 'slide-dots' }, this.list.map((_, i) => h('i', { class: i === this.i ? 'on' : '' })));
    stage.innerHTML = ''; stage.append(el, dots);
    return dur;
  },
  scheduleSlide() {
    const d = now(), st = currentState(d);
    const pr = profileFor(d);
    if (!st.work || !pr) return null;
    const rows = sortedPeriods(pr).map(p => {
      const cls = st.cur === p ? 'cur' : (toSec(p.end) <= st.sec ? 'past' : '');
      return h('tr', { class: cls }, h('td', {}, periodName(p)), h('td', {}, digits(fmt12(p.start))), h('td', {}, digits(fmt12(p.end))));
    });
    return h('div', { class: 'slide table-slide' }, h('h2', {}, '⏰ توقيت اليوم — ' + pr.name),
      h('table', { class: 'big-table' }, h('tr', {}, h('th', {}, 'الفترة'), h('th', {}, 'من'), h('th', {}, 'إلى')), rows));
  },
  dutySlide() {
    const d = now(); if (!isWorkDay(d)) return null;
    const rows = CFG.duty.rows.filter(r => (r.date ? r.date === ymd(d) : (r.day === -1 || r.day === d.getDay())) && (!r.week || r.week === cycleWeek(d)));
    if (!rows.length) return null;
    const st = currentState(d), keys = st.cur ? periodKeys(st.cur) : new Set();
    return h('div', { class: 'slide table-slide' }, h('h2', {}, '🛡️ المناوبة اليوم — ' + DAYS[d.getDay()]),
      h('table', { class: 'big-table' }, h('tr', {}, h('th', {}, 'الفترة'), h('th', {}, 'المناوب'), h('th', {}, 'المكان')),
        rows.map(r => h('tr', { class: keys.has(slotKey(r.slot)) ? 'cur' : '' }, h('td', {}, r.slot || 'طوال اليوم'), h('td', {}, r.teacher), h('td', {}, r.location || '')))));
  },
  subsSlide() {
    const d = now(); if (!isWorkDay(d)) return null;
    const rows = subsFor(d); if (!rows.length) return null;
    const st = currentState(d), keys = st.cur ? periodKeys(st.cur) : new Set();
    return h('div', { class: 'slide table-slide' }, h('h2', {}, '🔁 حصص الاحتياط اليوم'),
      h('table', { class: 'big-table' }, h('tr', {}, h('th', {}, 'الحصة'), h('th', {}, 'الصف'), h('th', {}, 'المعلم البديل'), h('th', {}, 'بدلاً من')),
        rows.map(r => h('tr', { class: keys.has(slotKey(r.period)) ? 'cur' : '' }, h('td', {}, r.period), h('td', {}, r.cls), h('td', {}, r.sub), h('td', {}, r.absent || '')))));
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
      let wk = 0;
      for (const row of aoa.slice(r + 1)) {
        if (wCol >= 0 && cell(row, wCol)) wk = weekNum(cell(row, wCol));
        const slot = cell(row, sCol);
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
  const hr = findHeader(aoa, [/تاريخ/, /يوم/, /حص/, /صف|فصل|شعب/, /غائب|غياب/, /احتياط|بديل|منفذ|المكلف|القائم/, /معلم|اسم/]);
  if (hr < 0) throw new Error('لم أتعرف على أعمدة جدول الاحتياط. يجب أن يحتوي على أعمدة مثل: التاريخ أو اليوم، الحصة، الصف، المعلم الغائب، المعلم البديل.');
  const hd = aoa[hr].map(norm);
  const col = (re, not) => hd.findIndex(c => re.test(c) && !(not && not.test(c)));
  const c = { date: col(/تاريخ/), day: col(/يوم/, /تاريخ/), period: col(/حص/, /احتياط/), cls: col(/صف|فصل|شعب/), absent: col(/غائب|غياب/), sub: col(/احتياط|بديل|منفذ|المكلف|القائم/), subject: col(/ماده/), note: col(/ملاحظ/) };
  if (c.sub < 0) c.sub = hd.findIndex((x, j) => /معلم|اسم/.test(x) && j !== c.absent);
  let dt = '', dy = -1;
  const out = [];
  for (const row of aoa.slice(hr + 1)) {
    if (cell(row, c.date)) dt = parseDate(cell(row, c.date));
    if (cell(row, c.day)) dy = dayIdx(cell(row, c.day));
    const sub = cell(row, c.sub); if (!sub && !cell(row, c.cls)) continue;
    out.push({ date: dt, day: dt ? -1 : dy, period: cell(row, c.period), cls: cell(row, c.cls), absent: cell(row, c.absent), sub, note: cell(row, c.subject) || cell(row, c.note) });
  }
  return out;
}
async function refreshSource(which, manual) {
  const src = CFG[which];
  if (src.source !== 'link' || !src.url) return;
  try {
    const r = await fetch(sheetFetchURL(src.url), { cache: 'no-store' });
    if (!r.ok) throw new Error('HTTP ' + r.status);
    const aoa = await aoaFromBlob(await r.blob(), src.sheet);
    src.rows = which === 'duty' ? parseDuty(aoa) : parseSubs(aoa);
    src.lastFetch = new Date().toLocaleString('ar'); src.lastError = '';
    if (manual) toast(`تم التحديث: ${src.rows.length} سجل`);
  } catch (e) {
    src.lastError = (e.message || String(e)) + (/fetch|NetworkError|Failed/i.test(e.message) ? ' — تحقق من الإنترنت ومن أن الرابط عام (منشور على الويب).' : '');
    if (manual) toast('فشل التحديث: ' + src.lastError, true);
  }
  CFG._rev = (CFG._rev || 0) + 1;
  save();
}
function refreshAllSources() { refreshSource('duty'); refreshSource('subs'); }
const lastRefresh = { duty: 0, subs: 0 };
function maybeRefresh() {
  for (const w of ['duty', 'subs']) {
    const s = CFG[w];
    if (s.source === 'link' && s.url && Date.now() - lastRefresh[w] > Math.max(1, +s.refreshMin || 10) * 60000) {
      lastRefresh[w] = Date.now(); refreshSource(w);
    }
  }
}

/* ------------------------------ تحديث عام ------------------------------ */
let slideRebuildT;
function onConfigChanged() {
  applyTheme();
  CFG._rev = (CFG._rev || 0) + 1;
  renderDates(); renderClock();
  renderSide(true); updateTicker(true);
}
function applyTheme() {
  const g = CFG.general;
  document.documentElement.style.setProperty('--accent', g.accent || '#14b8a6');
  $('#schoolName').textContent = g.schoolName; $('#schoolSub').textContent = g.schoolSub;
  if (g.logo) fileURL(g.logo).then(u => { if (u) { $('#logo').src = u; $('#logo').hidden = false; } });
  else $('#logo').hidden = true;
  $('#content').classList.toggle('no-side', !CFG.display.showSide);
  $('#dutyCard').hidden = !CFG.display.showDuty; $('#subsCard').hidden = !CFG.display.showSubs;
  $('#mutedFlag').hidden = g.bellEnabled; $('#cMute').classList.toggle('off', !g.bellEnabled);
  $('#cMute').textContent = g.bellEnabled ? '🔊' : '🔇';
}

/* =====================================================================
                                 الإعدادات
   ===================================================================== */
let curTab = 'general', curProfile = null;
function openSettings(tab) {
  $('#settings').hidden = false; document.body.classList.add('mouse');
  showTab(tab || curTab);
}
function closeSettings() { $('#settings').hidden = true; Slides.rebuild(); onConfigChanged(); }
function showTab(tab) {
  curTab = tab;
  $$('#tabs button').forEach(b => b.classList.toggle('on', b.dataset.tab === tab));
  const body = $('#setBody'); const top = body.scrollTop; body.innerHTML = '';
  const sec = h('section'); body.append(sec);
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
function audioOptions() {
  return [['', '🗣️ نطق آلي للنص'], ...CFG.files.filter(f => f.kind === 'audio').map(f => [f.id, '🎵 ' + f.name])];
}
const testBtn = (fn, label = '▶') => h('button', { class: 'btn sm', title: 'تجربة', onclick: e => { e.preventDefault(); Audio_.ensure(); fn(); } }, label);

const TABS = {
  /* ---------------- عام ---------------- */
  general(sec) {
    const g = CFG.general;
    sec.append(panel('بيانات المدرسة',
      h('div', { class: 'row' },
        field('اسم المدرسة', bindInput(g, 'schoolName', { style: 'width:340px' })),
        field('سطر فرعي', bindInput(g, 'schoolSub', { style: 'width:340px' })),
        field('اللون الرئيسي', bindInput(g, 'accent', { type: 'color' }))),
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
    sec.append(panel('الجداول الزمنية',
      h('p', { class: 'hint' }, 'يمكنك إنشاء أكثر من جدول (مثلاً: الجدول الأساسي، جدول يوم الخميس، جدول رمضان، جدول الاختبارات) وتحديد الأيام لكل جدول، أو فرض جدول معين على كل الأيام مؤقتاً.'),
      h('div', { class: 'row' },
        field('الجدول المعروض للتعديل', h('select', { onchange: e => { curProfile = e.target.value; rerender(); } }, CFG.profiles.map(p => h('option', { value: p.id, selected: p.id === curProfile }, p.name)))),
        h('button', { class: 'btn', onclick: () => { const n = { id: 'pr_' + uid(), name: 'جدول جديد', days: [], periods: [] }; CFG.profiles.push(n); curProfile = n.id; save(); rerender(); } }, '+ جدول جديد'),
        pr ? h('button', { class: 'btn', onclick: () => { const n = clone(pr); n.id = 'pr_' + uid(); n.name = pr.name + ' (نسخة)'; n.days = []; n.periods.forEach(p => p.id = uid()); CFG.profiles.push(n); curProfile = n.id; save(); rerender(); } }, 'نسخ الجدول') : null,
        pr && CFG.profiles.length > 1 ? h('button', { class: 'btn dan', onclick: () => { if (!confirm('حذف الجدول "' + pr.name + '"؟')) return; CFG.profiles = CFG.profiles.filter(p => p !== pr); if (CFG.override === pr.id) CFG.override = ''; save(); rerender(); } }, 'حذف الجدول') : null),
      h('div', { class: 'row' },
        field('فرض جدول على كل الأيام', bindSelect(CFG, 'override', [['', 'لا — حسب أيام كل جدول'], ...CFG.profiles.map(p => [p.id, p.name])])))));
    if (!pr) return;

    sec.append(panel('إعدادات "' + pr.name + '"',
      h('div', { class: 'row' }, field('اسم الجدول', bindInput(pr, 'name', { style: 'width:260px' }, rerender))),
      h('div', { class: 'row' }, h('span', { class: 'hint', style: 'margin:0' }, 'الأيام:'), DAYS.map((d, i) => h('label', { class: 'chk' },
        h('input', { type: 'checkbox', checked: pr.days.includes(i), onchange: e => {
          if (e.target.checked) { CFG.profiles.forEach(p => p.days = p.days.filter(x => x !== i)); pr.days.push(i); }
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
        h('td', {}, bindInput(p, 'start', { type: 'time' }, rerender)),
        h('td', {}, bindInput(p, 'end', { type: 'time' }, rerender)),
        h('td', { style: 'color:var(--muted);white-space:nowrap' }, dur > 0 ? Math.round(dur / 60) + ' د' : '⚠️'),
        h('td', {}, h('div', { class: 'mini' }, bindSelect(p, 'bellStart', soundOptions(), null, { style: 'width:150px' }), testBtn(() => Audio_.playSound(p.bellStart)))),
        h('td', {}, h('label', { class: 'chk' }, bindInput(p, 'annStart', { type: 'checkbox' }), 'إعلان')),
        h('td', {}, h('div', { class: 'mini' }, bindSelect(p, 'bellEnd', soundOptions(), null, { style: 'width:150px' }), testBtn(() => Audio_.playSound(p.bellEnd)))),
        h('td', {}, h('label', { class: 'chk' }, bindInput(p, 'annEnd', { type: 'checkbox' }), 'إعلان')),
        h('td', {}, h('div', { class: 'mini' },
          h('button', { class: 'btn sm', title: 'نصوص وتسجيلات مخصصة', onclick: () => { p._open = !p._open; rerender(); } }, '✎'),
          h('button', { class: 'btn sm', title: 'أعلى', onclick: () => { if (i) { [periods[i - 1], periods[i]] = [periods[i], periods[i - 1]]; save(); rerender(); } } }, '↑'),
          h('button', { class: 'btn sm', title: 'أسفل', onclick: () => { if (i < periods.length - 1) { [periods[i + 1], periods[i]] = [periods[i], periods[i + 1]]; save(); rerender(); } } }, '↓'),
          h('button', { class: 'btn dan sm', onclick: () => { periods.splice(i, 1); save(); rerender(); } }, '✕')))));
      if (p._open) tb.append(h('tr', {}, h('td', {}), h('td', { colspan: 10 },
        h('div', { class: 'row' },
          field('نص إعلان البداية', bindInput(p, 'startText', { placeholder: startText({ ...p, startText: '' }), style: 'width:300px' })),
          field('أو تسجيل صوتي للبداية', bindSelect(p, 'startAudio', audioOptions())),
          testBtn(() => Queue.push({ sound: p.bellStart, parts: [{ text: startText(p), audio: p.startAudio }], banner: startText(p) }), '▶ تجربة البداية')),
        h('div', { class: 'row' },
          field('نص إعلان النهاية', bindInput(p, 'endText', { placeholder: endText({ ...p, endText: '' }), style: 'width:300px' })),
          field('أو تسجيل صوتي للنهاية', bindSelect(p, 'endAudio', audioOptions())),
          testBtn(() => Queue.push({ sound: p.bellEnd, parts: [{ text: endText(p), audio: p.endAudio }], banner: endText(p) }), '▶ تجربة النهاية')))));
    });
    const warn = [];
    const sp = sortedPeriods(pr);
    for (let i = 1; i < sp.length; i++) if (toSec(sp[i].start) < toSec(sp[i - 1].end)) warn.push(`تداخل بين "${periodName(sp[i - 1])}" و"${periodName(sp[i])}"`);
    sec.append(panel('فترات اليوم',
      h('p', { class: 'hint' }, 'اختر نوع الفترة (طابور، حصة صفرية، الأولى ... الثامنة، فسحة، نشاط، ريادة ...)، والوقت من/إلى، وصوت الجرس في البداية والنهاية، وهل يُقال إعلان صوتي ("بدأت الآن الحصة الأولى"). زر ✎ لكتابة نص خاص أو اختيار تسجيل صوتي بدل النطق الآلي. إذا انتهت حصة وبدأت التالية في نفس الدقيقة يُدمج الإعلانان في جرس واحد.'),
      warn.length ? h('p', { class: 'hint', style: 'color:#fca5a5' }, '⚠️ ' + warn.join(' — ')) : null,
      h('div', { class: 'tbl-wrap' }, h('table', { class: 'edit' },
        h('thead', {}, h('tr', {}, ['#', 'النوع', 'اسم مخصص', 'من', 'إلى', 'المدة', 'جرس البداية', '', 'جرس النهاية', '', ''].map(t => h('th', {}, t)))), tb)),
      h('div', { class: 'row', style: 'margin-top:12px' },
        h('button', { class: 'btn pri', onclick: () => {
          const last = sortedPeriods(pr).slice(-1)[0];
          const st = last ? last.end : '07:00';
          const used = new Set(pr.periods.map(p => p.type));
          const nextType = ['p1', 'p2', 'p3', 'p4', 'p5', 'p6', 'p7', 'p8'].find(k => !used.has(k)) || 'custom';
          pr.periods.push(P(nextType, st, fromSec(toSec(st) + 40 * 60))); save(); rerender();
        } }, '+ إضافة فترة'),
        h('button', { class: 'btn', onclick: () => { pr.periods.sort((a, b) => toSec(a.start) - toSec(b.start)); save(); rerender(); } }, 'ترتيب حسب الوقت'))));

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
          let t = toSec(gen.start); const out = [];
          const add = (type, min, extra) => { out.push(P(type, fromSec(t), fromSec(t + min * 60), extra)); t += min * 60; };
          if (gen.assembly > 0) add('assembly', +gen.assembly);
          if (gen.zero) { add('p0', +gen.len); t += gen.gap * 60; }
          for (let i = 1; i <= Math.min(8, gen.count); i++) {
            add('p' + i, +gen.len);
            if (i === +gen.breakAfter) add('break', +gen.breakLen, { bellStart: 'triple', bellEnd: 'classic' });
            else if (i === +gen.break2After) add('break2', +gen.break2Len, { bellStart: 'triple', bellEnd: 'classic' });
            else if (i < gen.count) t += gen.gap * 60;
          }
          pr.periods = out; save(); rerender(); toast('تم توليد الجدول');
        } }, 'توليد'))));
  },

  /* ---------------- الأصوات ---------------- */
  sounds(sec) {
    const v = CFG.voice;
    const ar = Audio_.arabicVoices();
    const all = Audio_.voices();
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
      h('p', { class: 'hint' }, 'الأجراس المدمجة تعمل بدون ملفات. يمكنك رفع أصوات أخرى (MP3/WAV/OGG): جرس خاص، آية، دعاء، نشيد، إعلان مسجّل بصوتك... ثم اختيارها في الجداول أو التنبيهات.'),
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
      list.append(h('div', { class: 'item' + (r.enabled ? '' : ' disabled') },
        h('label', { class: 'chk' }, bindInput(r, 'enabled', { type: 'checkbox' }, rerender)),
        field('الاسم', bindInput(r, 'name', { style: 'width:150px' })),
        field('متى', bindSelect(r, 'mode', modes, rerender)),
        r.mode === 'fixed' ? field('الوقت', bindInput(r, 'time', { type: 'time' })) :
          (r.mode !== 'mid' ? field('دقائق', bindInput(r, 'minutes', { type: 'number', min: 1, style: 'width:70px' })) : null),
        r.mode !== 'fixed' ? field('يطبق على', bindSelect(r, 'applies', applies)) : null,
        field('الصوت', bindSelect(r, 'sound', soundOptions(), null, { style: 'width:150px' })),
        field('النص المنطوق / المعروض', bindInput(r, 'text', { style: 'width:240px' })),
        field('أو تسجيل', bindSelect(r, 'audio', audioOptions(), null, { style: 'width:150px' })),
        testBtn(() => Queue.push({ sound: r.sound, parts: [{ text: r.text, audio: r.audio }], banner: r.text || r.name, bannerIcon: '📢' })),
        h('button', { class: 'btn dan sm', onclick: () => { CFG.reminders.splice(i, 1); save(); rerender(); } }, 'حذف')));
    });
    sec.append(panel('تنبيهات إضافية أثناء اليوم',
      h('p', { class: 'hint' }, 'أمثلة: في منتصف كل حصة "اللهم صلِّ على محمد"، قبل نهاية الحصة بـ 5 دقائق تنبيه للمعلم، أو في وقت محدد (مثلاً 10:00 أذكار / النشيد الوطني). يمكن استخدام النطق الآلي أو ملف صوتي مرفوع.'),
      list,
      h('div', { class: 'row', style: 'margin-top:12px' },
        h('button', { class: 'btn pri', onclick: () => { CFG.reminders.push({ id: uid(), enabled: true, name: 'تنبيه جديد', mode: 'mid', minutes: 5, time: '10:00', applies: 'lessons', sound: 'chime', text: '', audio: '' }); save(); rerender(); } }, '+ إضافة تنبيه'))));
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
        h('div', { class: 'mini' },
          h('button', { class: 'btn sm', onclick: () => { if (i) { [CFG.slides[i - 1], CFG.slides[i]] = [CFG.slides[i], CFG.slides[i - 1]]; save(); rerender(); } } }, '↑'),
          h('button', { class: 'btn sm', onclick: () => { if (i < CFG.slides.length - 1) { [CFG.slides[i + 1], CFG.slides[i]] = [CFG.slides[i], CFG.slides[i + 1]]; save(); rerender(); } } }, '↓'),
          h('button', { class: 'btn dan sm', onclick: async () => { if (!confirm('حذف هذه الشريحة؟')) return; if (s.fileId) await removeFile(s.fileId); CFG.slides.splice(i, 1); save(); rerender(); } }, '✕'))));
    });
    const addUrl = h('input', { type: 'url', placeholder: 'https://... رابط PDF أو صورة أو فيديو أو يوتيوب أو صفحة', style: 'flex:1;min-width:280px;direction:ltr' });
    sec.append(panel('الشرائح المتحركة في الشاشة الرئيسية',
      h('p', { class: 'hint' }, 'ارفع جدول الحصص (PDF — تظهر كل صفحة كشريحة)، أو صوراً، أو فيديو (مكتوم الصوت افتراضياً حتى لا يصل الصوت للفصول)، أو ضع رابطاً. الفيديو يُكتم تلقائياً أثناء الجرس.'),
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
        h('button', { class: 'btn', onclick: () => { CFG.slides.push({ id: uid(), type: 'text', enabled: true, text: 'نص الإعلان هنا', duration: 12 }); save(); rerender(); } }, '+ شريحة نصية')),
      h('div', { class: 'row' }, addUrl,
        h('button', { class: 'btn', onclick: () => {
          const u = addUrl.value.trim(); if (!u) return;
          const isPdf = /\.pdf(\?|$)/i.test(u);
          CFG.slides.push({ id: uid(), type: isPdf ? 'pdf' : 'url', url: u, enabled: true, title: '', duration: isPdf ? 20 : 30, muted: true, fit: 'contain' });
          save(); rerender(); Slides.rebuild();
        } }, '+ إضافة رابط')),
      h('p', { class: 'hint' }, 'ملاحظة: الروابط تحتاج إنترنت، وبعض المواقع تمنع عرضها داخل برامج أخرى. الأفضل رفع الملف نفسه ليعمل بدون إنترنت.')));

    const A = CFG.autoSlides;
    sec.append(panel('شرائح تلقائية',
      h('div', { class: 'row' },
        chk(A, 'schedule', 'توقيت حصص اليوم'), chk(A, 'duty', 'جدول مناوبة اليوم'), chk(A, 'subs', 'حصص الاحتياط اليوم'),
        field('مدة كل منها (ث)', bindInput(A, 'duration', { type: 'number', style: 'width:90px' })))));

    const T = CFG.ticker;
    const ta = h('textarea', { onchange: e => { T.texts = e.target.value.split('\n').map(s => s.trim()).filter(Boolean); save(); } }, T.texts.join('\n'));
    sec.append(panel('الشريط الإخباري المتحرك (أسفل الشاشة)',
      h('div', { class: 'row' }, chk(T, 'enabled', 'إظهار الشريط'), field('العنوان', bindInput(T, 'label')), field('السرعة', bindInput(T, 'speed', { type: 'range', min: 30, max: 250 })),
        chk(T, 'includeDuty', 'إضافة المناوبة الحالية'), chk(T, 'includeSubs', 'إضافة حصص الاحتياط')),
      h('p', { class: 'hint' }, 'كل سطر = خبر:'), ta));

    const D = CFG.display;
    sec.append(panel('تخطيط الشاشة',
      h('div', { class: 'row' }, chk(D, 'showSide', 'إظهار العمود الجانبي (الحصة الحالية/المناوبة/الاحتياط)'), chk(D, 'showDuty', 'بطاقة المناوبة'), chk(D, 'showSubs', 'بطاقة الاحتياط'))));
  },

  duty(sec) { sourceTab(sec, 'duty'); },
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

/* تبويب مصدر بيانات (المناوبة / الاحتياط) */
function sourceTab(sec, which) {
  const S = CFG[which], isDuty = which === 'duty';
  const title = isDuty ? 'المناوبة' : 'حصص الاحتياط';
  sec.append(panel('مصدر بيانات ' + title,
    h('p', { class: 'hint' }, isDuty
      ? 'يقرأ البرنامج جدول المناوبة ويعرض تلقائياً من المناوب الآن (حسب اليوم والحصة/الفسحة والأسبوع). الصيغ المقبولة: (1) قائمة بأعمدة: الأسبوع (اختياري)، اليوم، الفترة/الحصة، المعلم المناوب، المكان. أو (2) جدول الأيام فيه أعمدة (الأحد، الاثنين...) والفترات صفوف. اكتب "طوال اليوم" في الفترة للمناوبة العامة.'
      : 'يقرأ البرنامج جدول الاحتياط ويعرض حصص اليوم. الأعمدة المقبولة: التاريخ (أو اليوم)، الحصة، الصف، المعلم الغائب، المعلم البديل/الاحتياط. تظهر فقط صفوف تاريخ اليوم.'),
    h('div', { class: 'row' },
      field('المصدر', bindSelect(S, 'source', [['manual', 'إدخال يدوي / استيراد ملف'], ['link', 'رابط (Google Sheets أو ملف CSV/Excel)']], rerender))),
    S.source === 'link' ? h('div', {},
      h('div', { class: 'row' },
        field('الرابط', bindInput(S, 'url', { type: 'url', style: 'width:520px;direction:ltr', placeholder: 'https://docs.google.com/spreadsheets/d/...' })),
        field('اسم الورقة (اختياري)', bindInput(S, 'sheet', { style: 'width:140px' })),
        field('التحديث كل (دقيقة)', bindInput(S, 'refreshMin', { type: 'number', min: 1, style: 'width:90px' })),
        h('button', { class: 'btn pri', onclick: async () => { await refreshSource(which, true); rerender(); } }, '⟳ تحديث الآن')),
      h('p', { class: 'hint' }, 'Google Sheets: شارك الملف "لأي شخص لديه الرابط" ثم الصق الرابط كما هو. البيانات تُحفظ محلياً بعد كل تحديث، فتبقى ظاهرة حتى لو انقطع الإنترنت.'),
      h('div', { class: 'row' },
        S.lastFetch ? h('span', { class: 'badge ok' }, 'آخر تحديث: ' + S.lastFetch) : null,
        S.lastError ? h('span', { class: 'badge err' }, 'خطأ: ' + S.lastError) : null)) : null,
    h('div', { class: 'row' },
      h('button', { class: 'btn', onclick: async () => {
        const f = await pickFile('.xlsx,.xls,.csv,text/csv'); if (!f) return;
        try {
          const aoa = await aoaFromBlob(f, S.sheet);
          const rows = isDuty ? parseDuty(aoa) : parseSubs(aoa);
          if (!rows.length) throw new Error('لم يُعثر على بيانات');
          S.rows = rows; save(); toast('تم استيراد ' + rows.length + ' سجل'); rerender();
        } catch (e) { toast(e.message, true); }
      } }, '📥 استيراد من ملف Excel / CSV'),
      h('button', { class: 'btn', onclick: () => downloadTemplate(which) }, '⬇ تحميل نموذج فارغ'),
      S.rows.length ? h('button', { class: 'btn dan', onclick: () => { if (confirm('مسح كل السجلات؟')) { S.rows = []; save(); rerender(); } } }, 'مسح الكل') : null)));

  // معاينة اليوم
  const d = now();
  const todayRows = isDuty
    ? S.rows.filter(r => (r.date ? r.date === ymd(d) : (r.day === -1 || r.day === d.getDay())) && (!r.week || r.week === cycleWeek(d)))
    : subsFor(d);
  sec.append(panel(`ما سيظهر اليوم (${DAYS[d.getDay()]}${isDuty && CFG.general.cycleWeeks > 1 ? ' — الأسبوع ' + cycleWeek(d) : ''})`,
    todayRows.length ? h('div', { class: 'preview-box' }, h('table', {},
      isDuty ? todayRows.map(r => h('tr', {}, h('td', {}, r.slot || 'طوال اليوم'), h('td', {}, r.teacher), h('td', {}, r.location)))
        : todayRows.map(r => h('tr', {}, h('td', {}, r.period), h('td', {}, r.cls), h('td', {}, r.sub), h('td', {}, r.absent)))))
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
    else tb.append(h('tr', {},
      h('td', {}, bindInput(r, 'date', { type: 'date' })), h('td', {}, numSel('day', dayOpts)),
      h('td', {}, bindInput(r, 'period', { placeholder: 'الأولى / 1' })), h('td', {}, bindInput(r, 'cls')),
      h('td', {}, bindInput(r, 'absent')), h('td', {}, bindInput(r, 'sub')), h('td', {}, bindInput(r, 'note')), h('td', {}, del)));
  });
  const heads = isDuty ? ['الأسبوع', 'اليوم', 'تاريخ محدد (اختياري)', 'الفترة / الحصة', 'المعلم المناوب', 'المكان', ''] : ['التاريخ', 'أو اليوم', 'الحصة', 'الصف', 'المعلم الغائب', 'المعلم البديل', 'ملاحظة', ''];
  sec.append(panel(`السجلات (${S.rows.length})` + (S.source === 'link' ? ' — ستُستبدل عند التحديث من الرابط' : ''),
    h('div', { class: 'tbl-wrap' }, h('table', { class: 'edit' }, h('thead', {}, h('tr', {}, heads.map(t => h('th', {}, t)))), tb)),
    S.rows.length > shown.length ? h('p', { class: 'hint' }, `يُعرض أول ${shown.length} سجل فقط للتعديل.`) : null,
    h('div', { class: 'row', style: 'margin-top:10px' },
      h('button', { class: 'btn pri', onclick: () => {
        S.rows.push(isDuty ? { week: 0, day: now().getDay(), date: '', slot: 'الفسحة', teacher: '', location: '' } : { date: ymd(now()), day: -1, period: '', cls: '', absent: '', sub: '', note: '' });
        save(); rerender();
      } }, '+ إضافة سجل'))));
}
function downloadTemplate(which) {
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
    for (const [id, du] of Object.entries(data.files || {})) await DB.put(id, await (await fetch(du)).blob());
    CFG = mergeDefaults(defaultConfig(), data.config);
    localStorage.setItem(LS_KEY, JSON.stringify(CFG));
    location.reload();
  } catch (e) { toast('فشل الاسترجاع: ' + e.message, true); }
}

/* ------------------------------ التشغيل ------------------------------ */
let started = false;
function start() {
  started = true;
  $('#startOverlay').hidden = true;
  Audio_.ensure();
}
function toggleFull() {
  if (document.fullscreenElement) document.exitFullscreen(); else document.documentElement.requestFullscreen().catch(() => { });
}
function toggleBell() { CFG.general.bellEnabled = !CFG.general.bellEnabled; save(); toast(CFG.general.bellEnabled ? 'تم تشغيل الجرس التلقائي' : 'تم إيقاف الجرس التلقائي'); }

async function init() {
  loadConfig();
  try { await DB.open(); } catch (e) { toast('تعذر فتح قاعدة البيانات المحلية: ' + e.message, true); }
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

  setInterval(() => { renderClock(); renderSide(); checkEvents(); }, 1000);
  setInterval(maybeRefresh, 30000); maybeRefresh();

  // أزرار التحكم
  $('#cSettings').onclick = () => openSettings();
  $('#cFull').onclick = toggleFull;
  $('#cBell').onclick = () => { Audio_.ensure(); manualBell(); };
  $('#cMute').onclick = toggleBell;
  $('#cStop').onclick = () => Audio_.stop();
  $('#cNext').onclick = () => Slides.next();
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
    else if (k === 'arrowright' || k === 'arrowleft') Slides.next();
  });
  window.addEventListener('storage', e => { if (e.key === LS_KEY) { loadConfig(); onConfigChanged(); Slides.rebuild(); } });
}
init();
