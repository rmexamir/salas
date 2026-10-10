/* =====================================================================
   Code.gs — بک‌اند کامل سامانه مدرسه (حضور و غیاب، نمرات، برنامه هفتگی، پیامک)
   نصب: همه‌ی کدهای قبلی را پاک کنید، این فایل را کامل جایگزین کنید،
   تابع setup را یک‌بار اجرا کنید، سپس Deploy ← Web app.
   ===================================================================== */
const MANAGER_PIN = '02276992';   // رمز تب مدیر
const ACCESS_KEY = '';                        // اختیاری؛ باید با ACCESS_KEY در index.html یکی باشد
const KAVENEGAR_KEY = '';                     // اختیاری؛ برای پیامک به اولیا (این فایل را با کسی به اشتراک نگذارید)
const SMS_SENDER = '';                        // شماره خط ارسال؛ خالی = خط پیش‌فرض حساب
const SCHOOL = 'دبیرستان صلاح‌الدین ایوبی';
const SMS_TEMPLATE = 'اولیای گرامی، {name} در تاریخ {date} در درس {course} {status}. {school}';

const SH = { att:'حضور و غیاب', gr:'نمرات', stu:'دانش‌آموزان', tch:'دبیران', sch:'برنامه', log:'پیامک‌ها', exc:'غیبت موجه', mock:'آزمون‌های آزمایشی' };
const HEADS = {
  att:['تاریخ','ساعت','رشته','کلاس','درس','دبیر','دانش‌آموز','وضعیت','توضیحات','زمان ثبت','نوع ثبت'],
  gr:['تاریخ','ساعت','رشته','کلاس','درس','دبیر','دانش‌آموز','نوع ارزیابی','سقف نمره','وضعیت','نمره','نمره از ۲۰','زمان ثبت'],
  stu:['رشته','کلاس','نام دانش‌آموز','کد ملی','نام پدر','تلفن پدر','نام مادر','تلفن مادر','تلفن دانش‌آموز'],
  tch:['نام دبیر','کد ورود','درس','ابلاغ (ساعت)','روزهای تدریس','کد پرسنلی','کد ملی','شماره تماس','مدرک تحصیلی'],
  sch:['کلاس','روز','زنگ','درس','دبیر','شیفت'],
  log:['زمان','تاریخ','کلاس','دانش‌آموز','وضعیت','درس','شماره','نتیجه','کلید'],
  exc:['تاریخ','کلاس','دانش‌آموز','علت','زمان ثبت'],
  mock:['زمان ثبت','دانش‌آموز','کلاس','رشته','مؤسسه','عنوان آزمون','تراز کل','رتبه کشوری','جزئیات دروس','شناسه فایل','نام فایل']
};

/* ---------- ابزارها ---------- */
function sheet_(k) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SH[k]);
  if (!sh) {
    sh = ss.insertSheet(SH[k]); sh.appendRow(HEADS[k]); sh.setFrozenRows(1);
    if (k === 'att' || k === 'gr') sh.getRange('A:B').setNumberFormat('@');   // تاریخ و ساعت به‌صورت متن بماند
    if (k === 'stu') sh.getRange('D:I').setNumberFormat('@');
    if (k === 'log') sh.getRange('G:G').setNumberFormat('@');
    if (k === 'tch') { sh.getRange('B:B').setNumberFormat('@'); sh.getRange('F:H').setNumberFormat('@'); }
    if (k === 'exc') sh.getRange('A:A').setNumberFormat('@');
  }
  return sh;
}
function setup() { Object.keys(SH).forEach(k => { const sh = sheet_(k); sh.getRange(1, 1, 1, HEADS[k].length).setValues([HEADS[k]]); }); }
const json_ = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
const err_ = m => ({ status:'error', message:m });
const norm_ = s => String(s == null ? '' : s).replace(/ي/g, 'ی').replace(/ك/g, 'ک').replace(/\u200c/g, ' ').replace(/\s+/g, ' ').trim();
const rowsOf_ = k => sheet_(k).getDataRange().getValues().slice(1);
const findRows_ = (sh, fn) => { const v = sh.getDataRange().getValues(), a = []; for (let i = 1; i < v.length; i++) if (fn(v[i])) a.push(i + 1); return a; };
const keyOk_ = k => !ACCESS_KEY || k === ACCESS_KEY;
const managerOk_ = pin => String(pin || '') !== '' && String(pin) === MANAGER_PIN;
function teacherOk_(p) {
  const t = rowsOf_('tch').filter(r => r[0]);
  if (!t.some(r => String(r[1]).trim())) return true;               // اگر هیچ کدی ثبت نشده، ورود آزاد است
  const pin = String(p.pin || '').trim();
  if (!pin) return false;
  return managerOk_(pin) || t.some(r => norm_(r[0]) === norm_(p.teacher) && String(r[1]).trim() === pin);
}

/* ---------- نقطه‌های ورود ---------- */
function doGet(e) {
  const q = e.parameter || {};
  if (!keyOk_(q.key)) return json_(err_('کلید دسترسی نادرست است.'));
  if (q.action === 'roster') return json_(roster_());
  if (q.action === 'schedule') return json_(getSchedule_());
  return json_({ status:'ok', message:'سامانه فعال است' });
}
function doPost(e) {
  let p; try { p = JSON.parse(e.postData.contents); } catch (x) { return json_(err_('درخواست نامعتبر است.')); }
  if (!keyOk_(p.key)) return json_(err_('کلید دسترسی نادرست است.'));
  try {
    if (p.action === 'login') return json_(teacherOk_(p) ? { status:'ok', manager:managerOk_(p.pin) } : err_('نام دبیر یا کد ورود نادرست است.'));
    if (p.action === 'inbox') return json_(teacherOk_(p) ? inbox_(p) : err_('نام دبیر یا کد ورود نادرست است.'));                 // کارتابل دبیر   // ورود دبیر
    if (p.action === 'student') return json_(student_(p));
    if (p.action === 'mockSubmit') return json_(mockSubmit_(p));                                                    // ثبت آزمون آزمایشی توسط دانش‌آموز
    if (p.action === 'excFor') return json_(teacherOk_(p) ? { status:'ok', list:excusedFor_(p.cls, p.date) } : err_('نام دبیر یا کد ورود نادرست است.'));                                                           // ورود دانش‌آموز با کد ملی
    if (['attendance', 'grades', 'sms'].indexOf(p.action) >= 0) {
      if (!teacherOk_(p)) return json_(err_('کد ورود یا نام دبیر نادرست است.'));
      if (!allowed_(p)) return json_(err_('این کلاس و درس در برنامه‌ی هفتگی شما نیست.'));   // دسترسی فقط به کلاس و درس خودش
      const lock = LockService.getScriptLock(); lock.waitLock(20000);
      try { return json_(p.action === 'attendance' ? saveAtt_(p) : p.action === 'grades' ? saveGr_(p) : sendSms_(p)); }
      finally { lock.releaseLock(); }
    }
    if (['report', 'saveSchedule', 'directory', 'studentCard', 'today', 'excAdd', 'excDel', 'excList', 'mockList', 'mockFile'].indexOf(p.action) >= 0) {
      if (!managerOk_(p.pin)) return json_(err_('رمز مدیر نادرست است.'));
      const H = { report:report_, directory:directory_, studentCard:student_, today:todayAtt_, saveSchedule:saveSchedule_, excAdd:excAdd_, excDel:excDel_, excList:excList_, mockList:mockList_, mockFile:mockFile_ };
      return json_(H[p.action](p));
    }
    return json_(err_('عملیات نامعتبر است.'));
  } catch (x) { return json_(err_(String(x))); }
}

/* ---------- فهرست کلاس‌ها و دبیران ---------- */
function roster_() {
  const classes = {};
  rowsOf_('stu').forEach(r => {
    const s = String(r[0]).trim(), c = String(r[1]).trim(), n = String(r[2]).trim(); if (!s || !c) return;
    classes[s] = classes[s] || {}; classes[s][c] = classes[s][c] || []; if (n) classes[s][c].push(n);
  });
  const t = rowsOf_('tch').filter(r => r[0]);
  const courses = t.map(r => String(r[2]).trim()).filter(String).filter((x, i, a) => a.indexOf(x) === i);
  return { status:'ok', classes:classes, teachers:t.map(r => String(r[0]).trim()), courses:courses, authRequired:t.some(r => String(r[1]).trim()) };
}

/* ---------- حضور و غیاب ---------- */
function saveAtt_(p) {
  const sh = sheet_('att');
  const idx = findRows_(sh, r => String(r[0]) === p.date && String(r[1]) === p.time && String(r[3]) === p.cls && String(r[4]) === p.course);
  if (idx.length && !p.force) return { status:'duplicate', message:'برای این کلاس، درس و زمان قبلاً حضور و غیاب ثبت شده است.' };
  idx.reverse().forEach(i => sh.deleteRow(i));
  const now = new Date(), b = [p.date, p.time, p.stream, p.cls, p.course, p.teacher];
  // غیبت موجه‌ی ثبت‌شده توسط مدیر ثابت است: دبیر نمی‌تواند آن را تغییر دهد یا برای کسی ثبت کند
  const exc = excusedFor_(p.cls, p.date), isMgr = managerOk_(p.pin);
  let list = (p.rows || []).filter(r => !exc.some(e => norm_(e.n) === norm_(r.name)));
  if (!isMgr) list = list.filter(r => String(r.status).indexOf('موجه') < 0);
  list = list.concat(exc.map(e => ({ name:e.n, status:'غیبت موجه', note:e.r, auto:true })));
  const rows = list.length ? list.map(r => b.concat([r.name, r.status, r.note || '', now, r.auto ? 'خودکار' : ''])) : [b.concat(['', 'همه حاضر', '', now, ''])];
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, 11).setValues(rows);
  return { status:'ok' };
}

/* ---------- نمرات ---------- */
function saveGr_(p) {
  const sh = sheet_('gr');
  const idx = findRows_(sh, r => String(r[0]) === p.date && String(r[3]) === p.cls && String(r[4]) === p.course && String(r[7]) === p.examType);
  if (idx.length && !p.force) return { status:'duplicate', message:'برای این کلاس، درس و نوع ارزیابی در همین تاریخ قبلاً نمره ثبت شده است.' };
  idx.reverse().forEach(i => sh.deleteRow(i));
  const now = new Date(), b = [p.date, p.time, p.stream, p.cls, p.course, p.teacher];
  const rows = (p.rows || []).map(r => r.absent
    ? b.concat([r.name, p.examType, p.max, 'غایب', '', '', now])
    : b.concat([r.name, p.examType, p.max, 'حاضر', r.score, Math.round(r.score / p.max * 2000) / 100, now]));
  if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, 13).setValues(rows);
  return { status:'ok' };
}

/* ---------- گزارش مدیر ---------- */
function report_() {
  const tz = Session.getScriptTimeZone();
  const s = a => a.map(x => x instanceof Date ? Utilities.formatDate(x, tz, 'yyyy/MM/dd HH:mm') : String(x));
  const a = sheet_('att').getDataRange().getValues(), g = sheet_('gr').getDataRange().getValues();
  return { status:'ok', attHead:s(a[0].slice(0, 10)), grHead:s(g[0]), att:a.slice(1).map(r => s(r.slice(0, 10))), gr:g.slice(1).map(s) };
}

/* ---------- برنامه هفتگی ---------- */
const DAYS_ = ['شنبه','یکشنبه','دوشنبه','سهشنبه','چهارشنبه','پنجشنبه'];
const pad6_ = r => [0, 1, 2, 3, 4, 5].map(i => r[i] === undefined ? '' : r[i]);
// ابلاغ هر دبیر: ساعت هفتگی و روزهای تدریس (ستون‌های D و E برگه‌ی دبیران)
function teacherInfo_() {
  const info = {};
  rowsOf_('tch').filter(r => r[0]).forEach(r => {
    const h = Number(String(r[3]).replace(/[۰-۹]/g, d => '۰۱۲۳۴۵۶۷۸۹'.indexOf(d)));
    const txt = norm_(r[4]).replace(/(یک|دو|سه|چهار|پنج)\s*شنبه/g, '$1شنبه');
    let ds = DAYS_.map((d, i) => txt.split(/[\s،,;؛\-\/]+/).indexOf(d) >= 0 ? i : -1).filter(i => i >= 0);
    if (ds.length === 2 && /(^|\s)تا(\s|$)/.test(txt)) { const lo = ds[0], hi = ds[1]; ds = []; for (let i = lo; i <= hi; i++) ds.push(i); }   // «شنبه تا چهارشنبه»
    info[String(r[0]).trim()] = { h: isNaN(h) ? 0 : h, d: ds.length ? ds : null };
  });
  return info;
}
function getSchedule_() {
  const rows = rowsOf_('sch').filter(r => r[0]).map(r => { r = pad6_(r); return [String(r[0]), Number(r[1]), Number(r[2]), String(r[3]), String(r[4]), String(r[5])]; });
  return { status:'ok', rows:rows, info:teacherInfo_() };
}
function saveSchedule_(p) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = sheet_('sch');
    const keep = rowsOf_('sch').filter(r => r[0] && String(r[0]) !== p.cls).map(pad6_);
    const out = keep.concat((p.rows || []).map(r => [p.cls, r[1], r[2], r[3], r[4], r[5] || '']));
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, sh.getMaxColumns()).clearContent();
    if (out.length) sh.getRange(2, 1, out.length, 6).setValues(out);
    sh.getRange(1, 1, 1, 6).setValues([HEADS.sch]);
    return { status:'ok' };
  } finally { lock.releaseLock(); }
}

/* ---------- پیامک به اولیا (کاوه‌نگار) ---------- */
function normPhone_(v) {
  let s = String(v).replace(/[۰-۹]/g, d => '۰۱۲۳۴۵۶۷۸۹'.indexOf(d)).replace(/[٠-٩]/g, d => '٠١٢٣٤٥٦٧٨٩'.indexOf(d)).replace(/\D/g, '');
  if (s.indexOf('98') === 0) s = '0' + s.slice(2);
  if (/^9\d{9}$/.test(s)) s = '0' + s;
  return /^09\d{9}$/.test(s) ? s : '';
}
function sendSms_(p) {
  if (!KAVENEGAR_KEY) return err_('کلید کاوه‌نگار در Code.gs تنظیم نشده است.');
  const stu = rowsOf_('stu'), log = sheet_('log');
  const done = new Set(log.getLastRow() > 1 ? log.getRange(2, 9, log.getLastRow() - 1, 1).getValues().map(r => r[0]) : []);
  let sent = 0, skipped = 0, failed = 0; const out = [];
  (p.rows || []).forEach(r => {
    const k = [p.date, p.cls, p.course, r.name, r.status].join('|');
    if (done.has(k)) return;                                          // جلوگیری از ارسال تکراری
    const row = stu.find(x => norm_(x[1]) === norm_(p.cls) && norm_(x[2]) === norm_(r.name));
    const phone = row ? normPhone_(row[5]) : '';   // مبنای ارسال پیامک: شماره‌ی پدر
    if (!phone) { skipped++; return; }
    const msg = SMS_TEMPLATE.replace('{name}', r.name).replace('{date}', p.date).replace('{course}', p.course || '')
      .replace('{status}', r.status === 'غایب' ? 'غایب بود' : 'با تأخیر وارد کلاس شد').replace('{school}', SCHOOL);
    const payload = { receptor:phone, message:msg }; if (SMS_SENDER) payload.sender = SMS_SENDER;
    let ok = false, res = '';
    try {
      const resp = UrlFetchApp.fetch('https://api.kavenegar.com/v1/' + KAVENEGAR_KEY + '/sms/send.json', { method:'post', payload:payload, muteHttpExceptions:true });
      const j = JSON.parse(resp.getContentText());
      ok = !!(j.return && j.return.status === 200); res = ok ? 'ارسال شد' : ((j.return && j.return.message) || 'خطا ' + resp.getResponseCode());
    } catch (e) { res = String(e); }
    if (ok) sent++; else failed++;
    out.push([new Date(), p.date, p.cls, r.name, r.status, p.course || '', phone, res, ok ? k : '']);
  });
  if (out.length) log.getRange(log.getLastRow() + 1, 1, out.length, 9).setValues(out);
  return { status:'ok', sent:sent, skipped:skipped, failed:failed };
}

/* ---------- بانک اطلاعات دانش‌آموزان و دبیران (فقط مدیر) ---------- */
// ستون‌های برگه‌ی «دانش‌آموزان»: A رشته، B کلاس، C نام، D کد ملی، E نام پدر، F تلفن پدر، G نام مادر، H تلفن مادر، I تلفن دانش‌آموز
// برای برگه‌ی قدیمی: این تابع ستون «تلفن اولیا» را حذف می‌کند (فقط یک‌بار و فقط اگر هنوز وجود داشته باشد) و قالب «متن» را می‌گذارد
function setupDirectory() {
  const s = sheet_('stu');
  if (norm_(s.getRange(1, 4).getValue()) === 'تلفن اولیا') s.deleteColumn(4);
  s.getRange(1, 4, 1, 6).setValues([['کد ملی', 'نام پدر', 'تلفن پدر', 'نام مادر', 'تلفن مادر', 'تلفن دانش‌آموز']]);
  s.getRange('D:I').setNumberFormat('@');
  const t = sheet_('tch'); t.getRange('F:H').setNumberFormat('@');
  t.getRange(1, 6, 1, 4).setValues([['کد پرسنلی', 'کد ملی', 'شماره تماس', 'مدرک تحصیلی']]);
}
const en_ = v => String(v == null ? '' : v).replace(/[۰-۹]/g, d => '۰۱۲۳۴۵۶۷۸۹'.indexOf(d)).replace(/[٠-٩]/g, d => '٠١٢٣٤٥٦٧٨٩'.indexOf(d)).trim();
function phoneRaw_(v) {           // فقط رقم؛ صفر ازدست‌رفته‌ی اول موبایل را برمی‌گرداند
  let s = en_(v).replace(/\D/g, '');
  if (/^9\d{9}$/.test(s)) s = '0' + s;
  if (s.indexOf('98') === 0 && s.length === 12) s = '0' + s.slice(2);
  return s;
}
function idRaw_(v) { const s = en_(v).replace(/\D/g, ''); return s && s.length < 10 ? ('0000000000' + s).slice(-10) : s; }   // کد ملی ۱۰ رقمی
function directory_() {
  const stu = rowsOf_('stu').filter(r => r[2]).map(r => ({
    n:String(r[2]).trim(), c:String(r[1]).trim(), id:idRaw_(r[3]),
    fn:String(r[4] || '').trim(), fp:phoneRaw_(r[5]), mn:String(r[6] || '').trim(), mp:phoneRaw_(r[7]), sp:phoneRaw_(r[8]) }));
  const m = {};
  rowsOf_('tch').filter(r => r[0]).forEach(r => {                     // هر دبیر چند ردیف (یک ردیف برای هر درس) دارد؛ ادغام می‌شود
    const n = String(r[0]).trim(), t = m[n] = m[n] || { n:n, courses:[], pc:'', id:'', ph:'', deg:'' }, c = String(r[2] || '').trim();
    if (c && t.courses.indexOf(c) < 0) t.courses.push(c);
    t.pc = t.pc || en_(r[5]); t.id = t.id || idRaw_(r[6]); t.ph = t.ph || phoneRaw_(r[7]); t.deg = t.deg || String(r[8] || '').trim();
  });
  return { status:'ok', students:stu, teachers:Object.keys(m).map(k => m[k]) };
}

/* ---------- پرتال دانش‌آموز: ورود با کد ملی؛ فقط اطلاعات خود دانش‌آموز و میانگین‌های کلاس برمی‌گردد ---------- */
function student_(p) {
  const id = idRaw_(p.nid);
  if (id.length !== 10) return err_('کد ملی باید ۱۰ رقم باشد.');
  const me = rowsOf_('stu').find(r => r[2] && idRaw_(r[3]) === id);
  if (!me) return err_('کد ملی در فهرست دانش‌آموزان پیدا نشد. با مدرسه تماس بگیرید.');
  const name = String(me[2]).trim(), cls = String(me[1]).trim(), K = norm_, isMe = r => K(r[6]) === K(name), sameC = r => K(r[3]) === K(cls);
  const num = v => (v === '' || v == null || isNaN(Number(v))) ? null : Number(v);
  const avg = a => a.length ? Math.round(a.reduce((x, y) => x + y, 0) / a.length * 100) / 100 : null;
  const gr = rowsOf_('gr').filter(sameC), ek = r => [r[0], r[4], r[7]].join('|'), ex = {}, byC = {};
  gr.forEach(r => {
    const n = num(r[11]); if (r[9] !== 'حاضر' || n === null) return;
    (ex[ek(r)] = ex[ek(r)] || []).push(n);
    const o = byC[r[4]] = byC[r[4]] || { all:[], me:[] }; o.all.push(n); if (isMe(r)) o.me.push(n);
  });
  const grades = gr.filter(isMe).map(r => ({ d:String(r[0]), c:String(r[4]), e:String(r[7]), m:num(r[8]), s:num(r[10]), n:num(r[11]), a:r[9] === 'غایب', ca:avg(ex[ek(r)] || []) }));
  const courses = Object.keys(byC).map(c => ({ c:c, me:avg(byC[c].me), cls:avg(byC[c].all) }));
  const ov = k => avg(courses.map(x => x[k]).filter(x => x !== null));
  const at = rowsOf_('att').filter(sameC), sess = {};
  at.forEach(r => { sess[[r[0], r[1], r[4]].join('|')] = 1; });
  const mine = at.filter(r => isMe(r) && String(r[7]) !== 'همه حاضر');
  const cnt = f => mine.filter(r => f(String(r[7]))).length;
  const absent = cnt(s => s === 'غایب'), late = cnt(s => s === 'تأخیر'), excused = cnt(s => s.indexOf('موجه') >= 0), sessions = Object.keys(sess).length;
  // نشان‌های تشویقی: فقط بر اساس پیشرفت خود دانش‌آموز، بدون مقایسه با دیگران
  const badges = [], ym = String(p.ym || ''), mAt = ym ? at.filter(r => String(r[0]).indexOf(ym) === 0) : [], myC = {};
  if (mAt.length && !mAt.some(r => isMe(r) && (String(r[7]) === 'غایب' || String(r[7]) === 'تأخیر'))) badges.push({ i:'🌟', t:'حضور کامل این ماه', d:'در این ماه هیچ غیبت یا تأخیری نداشتی.' });
  if (sessions >= 10 && !absent) badges.push({ i:'🏅', t:'انضباط درخشان', d:'تا امروز حتی یک غیبت نداشته‌ای.' });
  grades.forEach(x => { myC[x.c] = 1; });
  Object.keys(myC).forEach(c => {
    const ser = grades.filter(x => x.c === c && x.n !== null).map(x => x.n);
    if (ser.length >= 2 && ser[ser.length - 1] - ser[0] >= 2) badges.push({ i:'📈', t:'رشد در ' + c, d:'نمره‌ات در این درس از ابتدای سال بهتر شده است.' });
    if (ser.length >= 3 && Math.min.apply(null, ser) >= 15) badges.push({ i:'🎯', t:'ثبات در ' + c, d:'در همه‌ی آزمون‌های این درس نمره‌ی ۱۵ یا بیشتر گرفته‌ای.' });
  });
  const top = grades.filter(x => x.n !== null && x.n >= 18).length;
  if (top) badges.push({ i:'🏆', t:'نمره‌ی درخشان', d:top + ' بار نمره‌ی ۱۸ یا بالاتر گرفته‌ای.' });
  return { status:'ok', name:name, cls:cls, stream:String(me[0]).trim(), mocks:mocksOf_(name, cls).map(m => Object.assign(m, { f:'' })), badges:badges, grades:grades, courses:courses, overall:{ me:ov('me'), cls:ov('cls') },
    att:{ sessions:sessions, absent:absent, late:late, excused:excused, present:Math.max(0, sessions - absent - excused) },
    list:mine.slice(-100).reverse().map(r => ({ d:String(r[0]), c:String(r[4]), s:String(r[7]) })) };
}

/* ---------- دسترسی دبیر بر اساس برنامه‌ی هفتگی + کارتابل ---------- */
const pairs_ = rows => { const m = {}; rows.forEach(r => { if (r[3]) m[norm_(r[0]) + '|' + norm_(r[3])] = { c:String(r[0]).trim(), s:String(r[3]).trim() }; }); return Object.keys(m).map(k => m[k]); };
function assigned_(p) {           // ترکیب‌های «کلاس + درس» این دبیر؛ مدیر همه را دارد
  const sch = rowsOf_('sch').filter(r => r[0]);
  return pairs_(managerOk_(p.pin) ? sch : sch.filter(r => norm_(r[4]) === norm_(p.teacher)));
}
function allowed_(p) {            // مدیر، یا وقتی هنوز برنامه‌ای ثبت نشده، آزاد؛ وگرنه فقط کلاس و درس خود دبیر
  const sch = rowsOf_('sch').filter(r => r[0]);
  if (managerOk_(p.pin) || !sch.length) return true;
  return assigned_(p).some(x => norm_(x.c) === norm_(p.cls) && norm_(x.s) === norm_(p.course));
}
function inbox_(p) {
  const K = norm_, att = rowsOf_('att'), gr = rowsOf_('gr'), stu = rowsOf_('stu').filter(r => r[2]);
  const num = v => (v === '' || v == null || isNaN(Number(v))) ? null : Number(v);
  const avg = a => a.length ? Math.round(a.reduce((x, y) => x + y, 0) / a.length * 100) / 100 : null;
  const items = assigned_(p).map(x => {
    const inC = r => K(r[3]) === K(x.c) && K(r[4]) === K(x.s), A = att.filter(inC), G = gr.filter(inC), sess = {}, ex = {}, all = [];
    A.forEach(r => { sess[[r[0], r[1]].join('|')] = 1; });
    G.forEach(r => { const n = num(r[11]); if (r[9] !== 'حاضر' || n === null) return; all.push(n); const k = [r[0], r[7]].join('|'); (ex[k] = ex[k] || []).push(n); });
    const students = stu.filter(r => K(r[1]) === K(x.c)).map(r => {
      const nm = String(r[2]).trim(), mine = a => a.filter(q => K(q[6]) === K(nm));
      const a = mine(A).filter(q => String(q[7]) !== 'همه حاضر'), g2 = mine(G), cnt = f => a.filter(q => f(String(q[7]))).length;
      return { n:nm, ab:cnt(s => s === 'غایب'), lt:cnt(s => s === 'تأخیر'), ex:cnt(s => s.indexOf('موجه') >= 0),
        att:a.map(q => ({ d:String(q[0]), t:String(q[1]), s:String(q[7]) })).reverse(),
        gr:g2.map(q => ({ d:String(q[0]), e:String(q[7]), m:num(q[8]), s:num(q[10]), n:num(q[11]), a:q[9] === 'غایب', ca:avg(ex[[q[0], q[7]].join('|')] || []) })).reverse(),
        avg:avg(g2.map(q => num(q[11])).filter(v => v !== null)) };
    });
    return { c:x.c, s:x.s, sessions:Object.keys(sess).length, avg:avg(all), students:students };
  });
  return { status:'ok', items:items };
}

/* ---------- کلاس‌هایی که امروز حضور و غیابشان ثبت شده (برای هشدار مدیر) ---------- */
function todayAtt_(p) {
  const m = {};
  rowsOf_('att').forEach(r => { if (String(r[0]) === String(p.date)) m[norm_(r[3]) + '|' + norm_(r[4])] = { c:String(r[3]), s:String(r[4]), t:String(r[5]) }; });
  return { status:'ok', done:Object.keys(m).map(k => m[k]) };
}

/* ---------- غیبت موجه: ثبت توسط مدیر؛ در هر سه زنگ ثابت و برای دبیر غیرقابل تغییر ---------- */
function excusedFor_(cls, date) {
  return rowsOf_('exc').filter(r => String(r[0]) === String(date) && norm_(r[1]) === norm_(cls)).map(r => ({ n:String(r[2]).trim(), r:String(r[3] || '') }));
}
function excList_(p) {
  const d = String(p.date || '').trim();
  return { status:'ok', list:rowsOf_('exc').filter(r => r[0] && (!d || String(r[0]) === d)).map(r => ({ d:String(r[0]), c:String(r[1]), n:String(r[2]), r:String(r[3] || '') })).slice(-300).reverse() };
}
function excAdd_(p) {
  const date = String(p.date || '').trim(), cls = String(p.cls || '').trim(), names = (p.names || []).map(x => String(x).trim()).filter(String), why = String(p.reason || '').trim().slice(0, 80) || 'غیبت موجه (ثبت مدیر)';
  if (!/^\d{4}\/\d{2}\/\d{2}$/.test(date) || !cls || !names.length) return err_('تاریخ، کلاس و دست‌کم یک دانش‌آموز لازم است.');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const ex = sheet_('exc'), have = excusedFor_(cls, date).map(x => norm_(x.n)), now = new Date();
    const add = names.filter(n => have.indexOf(norm_(n)) < 0).map(n => [date, cls, n, why, now]);
    if (add.length) ex.getRange(ex.getLastRow() + 1, 1, add.length, 5).setValues(add);
    const at = sheet_('att'), v = at.getDataRange().getValues(), sess = {}, mine = {};   // حضور و غیاب‌های همان روز که قبلاً ثبت شده‌اند
    for (let i = 1; i < v.length; i++) {
      const r = v[i]; if (String(r[0]) !== date || norm_(r[3]) !== norm_(cls)) continue;
      const k = String(r[1]) + '|' + String(r[4]); if (!sess[k]) sess[k] = r;
      names.forEach(n => { if (norm_(r[6]) === norm_(n)) mine[k + '|' + norm_(n)] = i + 1; });
    }
    const app = [];
    Object.keys(sess).forEach(k => names.forEach(n => {
      const key = k + '|' + norm_(n), t = sess[k];
      if (mine[key]) at.getRange(mine[key], 8, 1, 2).setValues([['غیبت موجه', why]]);
      else app.push([t[0], t[1], t[2], t[3], t[4], t[5], n, 'غیبت موجه', why, now, 'خودکار']);
    }));
    if (app.length) at.getRange(at.getLastRow() + 1, 1, app.length, 11).setValues(app);
    return { status:'ok', added:add.length };
  } finally { lock.releaseLock(); }
}
function excDel_(p) {
  const date = String(p.date || '').trim(), cls = String(p.cls || '').trim(), name = String(p.name || '').trim();
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const ex = sheet_('exc'); findRows_(ex, r => String(r[0]) === date && norm_(r[1]) === norm_(cls) && norm_(r[2]) === norm_(name)).reverse().forEach(i => ex.deleteRow(i));
    const at = sheet_('att'), v = at.getDataRange().getValues(), del = [], cnt = {}, sk = r => [r[0], r[1], r[3], r[4]].map(String).join('|');
    for (let i = 1; i < v.length; i++) cnt[sk(v[i])] = (cnt[sk(v[i])] || 0) + 1;
    for (let i = 1; i < v.length; i++) {          // ردیف‌های ساخته‌شده‌ی خودکار حذف می‌شوند؛ ردیف‌هایی که دبیر «غایب» زده بود به «غایب» برمی‌گردند
      const r = v[i]; if (String(r[0]) !== date || norm_(r[3]) !== norm_(cls) || norm_(r[6]) !== norm_(name) || String(r[7]).indexOf('موجه') < 0) continue;
      if (String(r[10]) !== 'خودکار') at.getRange(i + 1, 8, 1, 2).setValues([['غایب', '']]);
      else if (cnt[sk(r)] > 1) { del.push(i + 1); cnt[sk(r)]--; }
      else at.getRange(i + 1, 7, 1, 5).setValues([['', 'همه حاضر', '', r[9], '']]);   // تنها ردیف آن جلسه: جلسه‌ی ثبت‌شده باقی بماند
    }
    del.reverse().forEach(i => at.deleteRow(i));
    return { status:'ok' };
  } finally { lock.releaseLock(); }
}

/* ---------- آزمون‌های آزمایشی: ثبت توسط دانش‌آموز، مشاهده توسط مدیر ---------- */
function authorizeDrive() { DriveApp.getRootFolder(); }   // یک‌بار اجرا شود تا دسترسی ذخیره‌ی کارنامه‌ها به گوگل درایو داده شود
function mockRow_(r) {
  let d = []; try { d = JSON.parse(r[8] || '[]'); } catch (e) {}
  const num = v => v === '' || v == null ? null : Number(v);
  return { t:r[0] instanceof Date ? Utilities.formatDate(r[0], Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm') : String(r[0]), n:String(r[1]), c:String(r[2]), g:String(r[3]), i:String(r[4]), e:String(r[5]), z:num(r[6]), k:num(r[7]), d:d, f:String(r[9] || '') };
}
const mocksOf_ = (name, cls) => rowsOf_('mock').filter(r => r[1] && norm_(r[1]) === norm_(name) && norm_(r[2]) === norm_(cls)).map(mockRow_).reverse();
function mockList_() { return { status:'ok', list:rowsOf_('mock').filter(r => r[1]).map(mockRow_).reverse() }; }
function mockFile_(p) {
  const id = String(p.id || ''); if (!id || !rowsOf_('mock').some(r => String(r[9]) === id)) return err_('فایل پیدا نشد.');
  const f = DriveApp.getFileById(id), b = f.getBlob();
  return { status:'ok', mime:b.getContentType(), name:f.getName(), data:Utilities.base64Encode(b.getBytes()) };
}
function mockSubmit_(p) {
  const id = idRaw_(p.nid), me = id.length === 10 && rowsOf_('stu').find(r => r[2] && idRaw_(r[3]) === id);
  if (!me) return err_('ورود نامعتبر است.');
  const inst = String(p.inst || ''), other = String(p.other || '').trim().slice(0, 40), title = String(p.title || '').trim().slice(0, 80), f = p.file || {};
  if (['ماز', 'قلمچی', 'راه موفقیت', 'خیلی سبز', 'گزینه دو', 'سایر'].indexOf(inst) < 0) return err_('مؤسسه را انتخاب کنید.');
  if (inst === 'سایر' && !other) return err_('نام مؤسسه را بنویسید.');
  if (!title) return err_('عنوان یا تاریخ آزمون را بنویسید.');
  const ext = { 'image/jpeg':'jpg', 'image/png':'png', 'image/webp':'webp', 'application/pdf':'pdf' }[f.mime];
  if (!f.data || !ext) return err_('آپلود کارنامه (JPG، PNG یا PDF) الزامی است.');
  if (String(f.data).length > 4200000) return err_('حجم فایل زیاد است (حداکثر ۳ مگابایت).');
  const nz = (v, lo, hi) => { if (v === '' || v == null) return ''; const x = Number(v); return isNaN(x) || x < lo || x > hi ? null : Math.round(x * 100) / 100; };
  const tz = nz(p.tz, 0, 20000), rk = nz(p.rank, 1, 2000000), subj = [];
  if (tz === null || rk === null) return err_('تراز کل یا رتبه نامعتبر است.');
  for (const q of (p.subj || []).slice(0, 12)) { const t = nz(q.t, 0, 20000), pc = nz(q.p, -34, 100); if (t === null || pc === null) return err_('تراز یا درصد یکی از دروس نامعتبر است.'); if (t !== '' || pc !== '') subj.push({ s:String(q.s).slice(0, 40), t:t, p:pc }); }
  const name = String(me[2]).trim(), cls = String(me[1]).trim(), it = DriveApp.getFoldersByName('کارنامه آزمون‌های آزمایشی'), folder = it.hasNext() ? it.next() : DriveApp.createFolder('کارنامه آزمون‌های آزمایشی');
  const file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(f.data), f.mime, cls + ' - ' + name + ' - ' + new Date().getTime() + '.' + ext));
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = sheet_('mock');
    sh.getRange(sh.getLastRow() + 1, 1, 1, 11).setValues([[new Date(), name, cls, String(me[0]).trim(), inst === 'سایر' ? 'سایر: ' + other : inst, title, tz, rk, JSON.stringify(subj), file.getId(), file.getName()]]);
  } finally { lock.releaseLock(); }
  return { status:'ok', mocks:mocksOf_(name, cls).map(m => Object.assign(m, { f:'' })) };
}
