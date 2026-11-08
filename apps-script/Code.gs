/**
 * نظام مخزن مفقودات مراسي — أبلكيشن ويب على Google Sheets
 * ملفين في Extensions ← Apps Script:
 *  - Code.gs: الملف ده.
 *  - Index.html: الواجهة (اعمله من + ← HTML واسمه Index بالظبط).
 *
 * الدخول باسم مستخدم وكلمة سر، والإدارة هي اللي بتعمل الحسابات وبتحدد دور كل واحد.
 * مش محتاج إيميلات ولا حسابات جوجل للمستخدمين.
 *
 * خطوات التشغيل:
 *  1) غيّر SETUP_CODE تحت لكود من عندك (ده كود تسجيل أول مدير).
 *  2) Deploy ← New deployment ← Web app:
 *       Execute as: Me
 *       Who has access: Anyone
 *     والتطبيق على الموبايل: https://ziedomar696.github.io/lost-found/ (بيفتح لينك الـ /exec ده جواه).
 *  3) افتح اللينك: هيطلب كود التسجيل واسمك واسم مستخدم وكلمة سر، وتبقى المدير.
 *     الكود بيشتغل بس لما مايكونش فيه أي مدير.
 *  4) من تاب «الحسابات» اعمل حساب لكل واحد وحدد دوره (إدارة، أمين مخزن، مشرف، كنترول روم).
 *
 *  كلمات السر بتتحفظ مشفّرة في تبويب «الحسابات» (محدش يقدر يقراها، والإدارة بتعمل إعادة تعيين بس).
 *  صور الهوية والتوقيعات في فولدر خاص بحسابك، وبتتعرض جوه الأبلكيشن بس للي داخل بحساب.
 */
const SETUP_CODE = 'MRS-7314';

/* ID الشيت (اختياري): لو السكريبت مش معمول من جوه الشيت (Extensions ← Apps Script)،
   حط هنا الـ ID بتاع الشيت، وهو الجزء اللي في اللينك بين /d/ و /edit.
   لو سبته فاضي والسكريبت مش مربوط بشيت، الأبلكيشن هيعمل شيت جديد لوحده في الدرايف بتاعك. */
const SHEET_ID = '';

/* نقل بيانات البرنامج القديم (اختياري): حط هنا الـ ID بتاع الشيت القديم (من لينكه بين /d/ و /edit)
   وشغّل الدالة importOld من المحرر. شرحها جنبها تحت. */
const OLD_SHEET_ID = '';

const FILES_FOLDER = 'مفقودات مراسي - ملفات';
const TZ = 'Africa/Cairo';
const ROLES = { admin: 'الإدارة', keeper: 'أمين مخزن', supervisor: 'مشرف', cr: 'الكنترول روم' };
const SESSION_SECONDS = 21600;   // 6 ساعات، وبتتجدد مع كل استخدام

/* ===================== الجداول ===================== */
const SHEETS = {
  accounts: { name: 'الحسابات', headers: ['اسم المستخدم','الاسم','الصلاحية','الحالة','كلمة السر','أضافه','تاريخ الإضافة'] },
  items:   { name: 'المفقودات', headers: ['رقم المفقود','تاريخ العثور','الوقت','النوع','الوصف','اللون والماركة','مكان العثور','تفاصيل المكان','المشرف','القطاع','اللي لقاه','مكان التخزين','الصورة','تأكيد الإدارة','ملاحظات','سجّله','أكّده'] },
  permits: { name: 'الأذونات', headers: ['رقم الإذن','رقم المفقود','الغرض','المستلم أو الجهة','الموبايل','صفة المستلم','الوحدة','تاريخ الطلب','قرار الإدارة','تاريخ ووقت القرار','الصلاحية','ملاحظة الإدارة','طلبه'] },
  deliv:   { name: 'التسليم', headers: ['رقم التسليم','رقم المفقود','رقم الإذن','تاريخ ووقت التسليم','اسم المستلم','علاقته بالمفقود','صفته في مراسي','الوحدة','نوع الهوية','رقم الهوية','الموبايل','وصفه قبل العرض','نتيجة المطابقة','إثبات الملكية','خرّجه من المخزن','الشاهد','حالة المفقود','صورة الهوية','صورة الأبلكيشن أو QR','توقيع المستلم','ملاحظات','سجّله'] },
  other:   { name: 'جهات_أخرى', headers: ['رقم العملية','رقم المفقود','رقم الإذن','التاريخ','الجهة','رقم المحضر','جهة الاستلام','خرّجه من المخزن','ملاحظات','سجّله'] },
  reports: { name: 'البلاغات', headers: ['رقم البلاغ','تاريخ البلاغ','الاسم','الموبايل','صفته في مراسي','الوحدة','النوع','الوصف','اللون والماركة','ضاعت من','لحد','المكان المتوقع','تفاصيل المكان','رقم المفقود المطابق','حالة البلاغ','متابعات','سجّله'] },
  cr:      { name: 'الكنترول_روم', headers: ['رقم البلاغ','تاريخ ووقت الإرسال','الأولوية','المطلوب','الأماكن والكاميرات','فترة المراجعة','خط سير صاحب البلاغ','أوصاف صاحب البلاغ','نتيجة المراجعة','الكاميرات اللي اتراجعت','الإفادة','الصور','تاريخ ووقت الرد','ملخص البلاغ','رد بواسطة'] },
  cats:    { name: 'الأنواع', headers: ['النوع','مدة الحفظ','ثمين'], seed: [['هواتف وأجهزة إلكترونية',90,'نعم'],['محافظ ونقود',90,'نعم'],['مجوهرات وساعات',90,'نعم'],['بطاقات هوية ومستندات',14,'نعم'],['مفاتيح',60,'لا'],['حقائب',60,'لا'],['نظارات',30,'لا'],['ملابس وأحذية',30,'لا'],['مستلزمات شاطئ ومناشف',14,'لا'],['ألعاب أطفال',30,'لا'],['مأكولات ومشروبات',1,'لا'],['أخرى',30,'لا']] },
  locs:    { name: 'أماكن_العثور', headers: ['المكان'], seed: [['البوابة الرئيسية'],['البوابات الفرعية'],['الشاطئ'],['المارينا'],['الممشى والكورنيش'],['المطاعم والكافيهات'],['الفلل والشاليهات'],['الفنادق'],['النوادي والملاعب'],['الجراجات ومواقف السيارات'],['أخرى']] },
  stores:  { name: 'أماكن_التخزين', headers: ['المكان'], seed: [['الخزنة (للثمين)'],['رف A'],['رف B'],['رف C'],['صندوق الملابس'],['صندوق المستلزمات']] },
  staff:   { name: 'الموظفين', headers: ['الاسم'], seed: [] },
  parties: { name: 'الجهات', headers: ['الجهة'], seed: [['تسليم لقسم الشرطة'],['تسليم للإدارة'],['تبرع لجهة خيرية'],['إعدام / تخلص']] },
  prios:   { name: 'الأولويات', headers: ['الأولوية','المهلة'], seed: [['عاجل جداً',1],['عاجل',4],['عادي',24]] }
};

/* ===================== الدخول ===================== */
function doGet() {
  setup_();
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('مخزن مفقودات مراسي')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);   // علشان يفتح جوه تطبيق GitHub
}

/** اختياري: من المحرر، بتجهّز التبويبات وفولدر الصور. */
function setup() {
  setup_();
  const folder = ensureFolder_();
  Logger.log('تم. الشيت: ' + ss_().getUrl() + ' · فولدر الصور: ' + folder.getUrl() + (hasAdmin_() ? '' : ' · مفيش مدير لسه: افتح الأبلكيشن وسجّل بكود التسجيل.'));
}

function setup_() {
  const ss = ss_();
  Object.keys(SHEETS).forEach(function (k) {
    const def = SHEETS[k];
    let sh = ss.getSheetByName(def.name);
    if (!sh) {
      sh = ss.insertSheet(def.name);
      sh.setRightToLeft(true);
      sh.getRange(1, 1, 1, def.headers.length).setValues([def.headers]).setFontWeight('bold').setBackground('#0B6B66').setFontColor('#FFFFFF');
      sh.setFrozenRows(1);
      if (def.seed && def.seed.length) sh.getRange(2, 1, def.seed.length, def.seed[0].length).setValues(def.seed);
      return;
    }
    const last = Math.max(1, sh.getLastColumn());
    const have = sh.getRange(1, 1, 1, last).getValues()[0].map(String);
    const missing = def.headers.filter(function (h) { return have.indexOf(h) < 0; });
    if (missing.length) {
      const start = have.filter(String).length + 1;
      sh.getRange(1, start, 1, missing.length).setValues([missing]).setFontWeight('bold').setBackground('#0B6B66').setFontColor('#FFFFFF');
    }
  });
}

function ensureFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('FILES_FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const folder = DriveApp.createFolder(FILES_FOLDER);
  try { folder.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE); } catch (e) {}
  props.setProperty('FILES_FOLDER_ID', folder.getId());
  return folder;
}

/* ===================== الحسابات وكلمات السر ===================== */
function normUser_(u) { return String(u || '').trim().toLowerCase(); }
function hash_(salt, pw) {
  const d = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + pw, Utilities.Charset.UTF_8);
  return Utilities.base64Encode(d);
}
function makePass_(pw) { const salt = Utilities.getUuid().replace(/-/g, ''); return salt + '$' + hash_(salt, pw); }
function checkPass_(stored, pw) {
  const p = String(stored || '').split('$');
  if (p.length === 3 && p[0] === 'old') return oldHash_(p[1], pw) === p[2];   // حسابات منقولة من البرنامج القديم
  return p.length === 2 && hash_(p[0], pw) === p[1];
}
function oldHash_(salt, pw) {
  let h = String(pw);
  for (let i = 0; i < 50; i++) h = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + h, Utilities.Charset.UTF_8));
  return h;
}
function validPass_(pw) { if (String(pw || '').length < 6) throw new Error('كلمة السر لازم تبقى 6 حروف أو أرقام على الأقل.'); }
function roleCode_(label) { label = String(label || '').trim(); let r = ''; Object.keys(ROLES).forEach(function (k) { if (ROLES[k] === label || k === label) r = k; }); return r; }

function accounts_() {
  const sh = sheet_('accounts');
  if (!sh) return [];
  return rows_('accounts').map(function (r) {
    return { user: normUser_(r['اسم المستخدم']), name: String(r['الاسم'] || ''), role: roleCode_(r['الصلاحية']),
      active: String(r['الحالة']).trim() !== 'موقوف', pass: String(r['كلمة السر'] || ''), by: String(r['أضافه'] || ''), at: iso_(r['تاريخ الإضافة']) };
  });
}
function hasAdmin_() { return accounts_().some(function (u) { return u.active && u.role === 'admin'; }); }
function putAccount_(u, by) {
  const sh = sheet_('accounts'), h = headers_(sh), last = sh.getLastRow();
  const keys = last < 2 ? [] : sh.getRange(2, 1, last - 1, 1).getValues().map(function (r) { return normUser_(r[0]); });
  const i = keys.indexOf(u.user);
  const patch = {};
  if (u.name !== undefined) patch['الاسم'] = u.name;
  if (u.role !== undefined) patch['الصلاحية'] = ROLES[u.role];
  if (u.active !== undefined) patch['الحالة'] = u.active ? 'نشط' : 'موقوف';
  if (u.password) patch['كلمة السر'] = makePass_(u.password);
  if (i >= 0) {
    Object.keys(patch).forEach(function (name) { const c = h.indexOf(name); if (c >= 0) sh.getRange(i + 2, c + 1).setValue(patch[name]); });
  } else {
    patch['اسم المستخدم'] = u.user; patch['أضافه'] = by; patch['تاريخ الإضافة'] = new Date();
    append_('accounts', patch);
  }
}

/* ===================== الجلسات ===================== */
function newSession_(user) {
  const t = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  CacheService.getScriptCache().put('s_' + t, user, SESSION_SECONDS);
  return t;
}
function me_(t) {
  const cache = CacheService.getScriptCache();
  const user = t ? cache.get('s_' + t) : null;
  if (!user) throw new Error('AUTH: الجلسة انتهت، ادخل تاني.');
  const u = accounts_().filter(function (x) { return x.user === user; })[0];
  if (!u || !u.active || !u.role) { cache.remove('s_' + t); throw new Error('AUTH: الحساب ده موقوف أو اتشال.'); }
  cache.put('s_' + t, user, SESSION_SECONDS);
  return { user: u.user, name: u.name, role: u.role };
}
function need_(t, roles) {
  const m = me_(t);
  if (roles.indexOf(m.role) < 0) throw new Error('العملية دي مش متاحة لصلاحيتك.');
  return m;
}

/** الواجهة بتسأل: محتاج تسجيل أول مدير؟ */
function needsSetup() { setup_(); return !hasAdmin_(); }

/** تسجيل أول مدير بكود التسجيل (بيشتغل بس لو مفيش مدير). */
function registerFirstAdmin(o) {
  o = clean_(o);
  const user = normUser_(o.user);
  if (!o.name) throw new Error('اكتب اسمك.');
  if (!/^\S{3,30}$/.test(user)) throw new Error('اسم المستخدم من 3 لـ 30 حرف ومن غير مسافات.');
  validPass_(o.password);
  return withLock_(function () {
    setup_();
    if (hasAdmin_()) throw new Error('فيه مدير متسجّل بالفعل. اطلب منه يعملك حساب.');
    if (String(o.code || '') !== SETUP_CODE) throw new Error('كود التسجيل غلط.');
    putAccount_({ user: user, name: o.name, role: 'admin', active: true, password: o.password }, user);
    ensureFolder_();
    return { token: newSession_(user) };
  });
}

function login(user, password) {
  user = normUser_(user);
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('f_' + user) || 0);
  if (fails >= 5) throw new Error('محاولات كتير غلط. استنى 10 دقايق وجرّب تاني.');
  const u = accounts_().filter(function (x) { return x.user === user; })[0];
  if (!u || !checkPass_(u.pass, String(password || ''))) {
    cache.put('f_' + user, String(fails + 1), 600);
    throw new Error('اسم المستخدم أو كلمة السر غلط.');
  }
  if (!u.active || !u.role) throw new Error('الحساب ده موقوف. كلّم الإدارة.');
  cache.remove('f_' + user);
  return { token: newSession_(user) };
}
function logout(t) { if (t) CacheService.getScriptCache().remove('s_' + t); return true; }

function changePassword(t, oldPw, newPw) {
  const m = me_(t);
  const u = accounts_().filter(function (x) { return x.user === m.user; })[0];
  if (!checkPass_(u.pass, String(oldPw || ''))) throw new Error('كلمة السر الحالية غلط.');
  validPass_(newPw);
  putAccount_({ user: m.user, password: newPw }, m.user);
  return true;
}

/** رقم الصف في الشيت للمفتاح ده في أول عمود (أو -1). */
function rowOf_(k, key, norm) {
  const sh = sheet_(k), last = sh.getLastRow();
  if (last < 2) return -1;
  const f = norm || function (x) { return String(x).trim(); };
  const keys = sh.getRange(2, 1, last - 1, 1).getValues().map(function (r) { return f(r[0]); });
  const i = keys.indexOf(f(key));
  return i < 0 ? -1 : i + 2;
}

/** إضافة أو تعديل حساب (الإدارة بس). o.oldUser = اسم المستخدم القديم لو بتعدّل. */
function saveUser(t, o) {
  const m = need_(t, ['admin']);
  o = clean_(o);
  const user = normUser_(o.user), oldUser = normUser_(o.oldUser || o.user);
  if (!/^\S{3,30}$/.test(user)) throw new Error('اسم المستخدم من 3 لـ 30 حرف ومن غير مسافات.');
  if (!ROLES[o.role]) throw new Error('اختار الدور.');
  if (o.password) validPass_(o.password);
  const active = o.active !== false;
  return withLock_(function () {
    const all = accounts_();
    const old = o.isNew ? null : all.filter(function (x) { return x.user === oldUser; })[0];
    if (!o.isNew && !old) throw new Error('الحساب مش موجود.');
    if ((o.isNew || user !== oldUser) && all.some(function (x) { return x.user === user; })) throw new Error('اسم المستخدم ده موجود بالفعل.');
    if (o.isNew && !o.password) throw new Error('اكتب كلمة سر للحساب.');
    if (!o.name && !(old && old.name)) throw new Error('اكتب الاسم.');
    const otherAdmins = all.filter(function (x) { return x.active && x.role === 'admin' && x.user !== oldUser; }).length;
    if (old && old.active && old.role === 'admin' && (o.role !== 'admin' || !active) && !otherAdmins) throw new Error('لازم يفضل مدير واحد على الأقل.');
    if (old && old.user === m.user && !active) throw new Error('ما ينفعش توقف حسابك وإنت داخل بيه.');
    if (old && user !== oldUser) sheet_('accounts').getRange(rowOf_('accounts', oldUser, normUser_), 1).setValue(user);
    putAccount_({ user: user, name: o.name || old.name, role: o.role, active: active, password: o.password || '' }, m.user);
    return { renamedSelf: !!(old && old.user === m.user && user !== oldUser) };
  });
}

/** مسح حساب نهائياً (الإدارة بس). العمليات القديمة بتفضل مسجّلة باسم المستخدم. */
function deleteUser(t, user) {
  const m = need_(t, ['admin']);
  user = normUser_(user);
  if (user === m.user) throw new Error('ما ينفعش تمسح حسابك وإنت داخل بيه.');
  return withLock_(function () {
    const all = accounts_();
    const u = all.filter(function (x) { return x.user === user; })[0];
    if (!u) throw new Error('الحساب مش موجود.');
    if (u.active && u.role === 'admin' && !all.some(function (x) { return x.active && x.role === 'admin' && x.user !== user; })) throw new Error('لازم يفضل مدير واحد على الأقل.');
    sheet_('accounts').deleteRow(rowOf_('accounts', user, normUser_));
    return true;
  });
}

/* ===================== القوائم (الإعدادات) ===================== */
// النوع والمكان لما يتغيّر اسمهم، الاسم الجديد بيتطبق على المفقودات والبلاغات القديمة كمان.
const LISTS = {
  cats:    { label: 'النوع',          rename: [['items', 'النوع'], ['reports', 'النوع']] },
  locs:    { label: 'مكان العثور',     rename: [['items', 'مكان العثور'], ['reports', 'المكان المتوقع']] },
  stores:  { label: 'مكان التخزين',    rename: [['items', 'مكان التخزين']] },
  staff:   { label: 'الموظف',          rename: [] },
  parties: { label: 'الجهة',           rename: [] }
};
function renameIn_(k, col, from, to) {
  const sh = sheet_(k), last = sh.getLastRow();
  if (last < 2) return 0;
  const c = headers_(sh).indexOf(col);
  if (c < 0) return 0;
  const rg = sh.getRange(2, c + 1, last - 1, 1), v = rg.getValues();
  let n = 0;
  v.forEach(function (r) { if (String(r[0]).trim() === from) { r[0] = to; n++; } });
  if (n) rg.setValues(v);
  return n;
}
/** إضافة (oldName فاضي) أو تعديل عنصر في قايمة. للأنواع: o.days و o.valuable. */
function listSave(t, kind, oldName, o) {
  need_(t, ['admin']);
  const L = LISTS[kind];
  if (!L) throw new Error('قايمة مش معروفة.');
  o = clean_(o);
  const name = String(o.name || '').trim();
  oldName = String(oldName || '').trim();
  if (!name) throw new Error('اكتب ' + L.label + '.');
  let row = [name];
  if (kind === 'cats') {
    const days = Math.round(Number(o.days));
    if (!(days >= 1 && days <= 3650)) throw new Error('مدة الحفظ لازم تبقى رقم أيام (من 1 لـ 3650).');
    row = [name, days, o.valuable ? 'نعم' : 'لا'];
  }
  return withLock_(function () {
    const exists = rowOf_(kind, name);
    if (exists > 0 && name !== oldName) throw new Error('«' + name + '» موجود بالفعل.');
    if (!oldName) { sheet_(kind).appendRow(row); return { renamed: 0 }; }
    const r = rowOf_(kind, oldName);
    if (r < 0) throw new Error('«' + oldName + '» مش موجود.');
    sheet_(kind).getRange(r, 1, 1, row.length).setValues([row]);
    let n = 0;
    if (name !== oldName) L.rename.forEach(function (x) { n += renameIn_(x[0], x[1], oldName, name); });
    return { renamed: n };
  });
}
function listDelete(t, kind, name) {
  need_(t, ['admin']);
  const L = LISTS[kind];
  if (!L) throw new Error('قايمة مش معروفة.');
  name = String(name || '').trim();
  return withLock_(function () {
    const r = rowOf_(kind, name);
    if (r < 0) throw new Error('«' + name + '» مش موجود.');
    if (['cats', 'locs', 'stores'].indexOf(kind) >= 0 && rows_(kind).length <= 1) throw new Error('لازم يفضل عنصر واحد على الأقل في القايمة.');
    sheet_(kind).deleteRow(r);
    return true;   // المفقودات القديمة بتفضل بالاسم المكتوب عليها
  });
}

/** عرض صورة من فولدر الأبلكيشن لأي حد داخل بحساب. */
function fileData(t, url) {
  me_(t);
  const mm = String(url || '').match(/\/d\/([\w-]+)/) || String(url || '').match(/[?&]id=([\w-]+)/);
  if (!mm) throw new Error('الرابط ده مش صورة.');
  const folderId = PropertiesService.getScriptProperties().getProperty('FILES_FOLDER_ID');
  const f = DriveApp.getFileById(mm[1]);
  let inFolder = false; const ps = f.getParents();
  while (ps.hasNext()) if (ps.next().getId() === folderId) inFolder = true;
  if (!inFolder) throw new Error('الملف ده مش من ملفات الأبلكيشن.');
  const b = f.getBlob();
  return 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes());
}

/* ===================== أدوات الشيت ===================== */
let SS_ = null;
function ss_() {
  if (SS_) return SS_;
  if (SHEET_ID) return (SS_ = SpreadsheetApp.openById(SHEET_ID));
  const active = SpreadsheetApp.getActive();
  if (active) return (SS_ = active);
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SHEET_ID');
  if (id) { try { return (SS_ = SpreadsheetApp.openById(id)); } catch (e) {} }
  SS_ = SpreadsheetApp.create('مخزن مفقودات مراسي - البيانات');
  props.setProperty('SHEET_ID', SS_.getId());
  return SS_;
}
function sheet_(k) { return ss_().getSheetByName(SHEETS[k].name); }
function headers_(sh) { return sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0].map(String); }
function rows_(k) {
  const sh = sheet_(k);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const h = headers_(sh);
  return sh.getRange(2, 1, last - 1, h.length).getValues()
    .filter(function (r) { return String(r[0]).trim() !== ''; })
    .map(function (r) { const o = {}; h.forEach(function (name, i) { o[name] = r[i]; }); return o; });
}
function append_(k, obj) {
  const sh = sheet_(k);
  const h = headers_(sh);
  sh.appendRow(h.map(function (name) { return obj[name] === undefined ? '' : obj[name]; }));
}
function update_(k, key, patch) {
  const sh = sheet_(k);
  const h = headers_(sh);
  const last = sh.getLastRow();
  if (last < 2) throw new Error('السجل مش موجود.');
  const keys = sh.getRange(2, 1, last - 1, 1).getValues().map(function (r) { return String(r[0]); });
  const i = keys.indexOf(String(key));
  if (i < 0) throw new Error('السجل مش موجود: ' + key);
  Object.keys(patch).forEach(function (name) {
    const c = h.indexOf(name);
    if (c >= 0) sh.getRange(i + 2, c + 1).setValue(patch[name]);
  });
}
function list_(k, col) { return rows_(k).map(function (r) { return String(r[col] || ''); }).filter(String); }

function pad_(n, l) { n = String(n); while (n.length < l) n = '0' + n; return n; }
function nextId_(k, prefix, width) {
  const re = new RegExp('^' + prefix + '(\\d+)$');
  let max = 0;
  rows_(k).forEach(function (r) { const m = String(r[SHEETS[k].headers[0]]).match(re); if (m) max = Math.max(max, Number(m[1])); });
  return prefix + pad_(max + 1, width);
}
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}
function toDate_(s) {               // 'yyyy-mm-dd' (الساعة 12 الضهر علشان التوقيت الصيفي) أو 'yyyy-mm-ddThh:mm'
  if (!s) return '';
  if (s instanceof Date) return s;
  const p = String(s).split(/[-T: ]/).map(Number);
  const hasTime = String(s).length > 10;
  const str = p[0] + '-' + pad_(p[1] || 1, 2) + '-' + pad_(p[2] || 1, 2) + ' ' + (hasTime ? pad_(p[3] || 0, 2) + ':' + pad_(p[4] || 0, 2) : '12:00');
  return Utilities.parseDate(str, TZ, 'yyyy-MM-dd HH:mm');
}
function iso_(v, withTime) {
  if (!(v instanceof Date) || isNaN(v)) return v === '' || v === null || v === undefined ? '' : String(v);
  return Utilities.formatDate(v, TZ, withTime ? "yyyy-MM-dd'T'HH:mm" : 'yyyy-MM-dd');
}
function time_(v) { return v instanceof Date ? Utilities.formatDate(v, TZ, 'HH:mm') : String(v || ''); }
function hours_(v) {                // 48 أو "048:00:00"
  if (typeof v === 'number') return v;
  const s = String(v || '');
  if (s.indexOf(':') >= 0) { const p = s.split(':').map(Number); return (p[0] || 0) + (p[1] || 0) / 60; }
  return Number(s) || 0;
}
function saveFile_(dataUrl, name) {  // صورة أو توقيع ← ملف في درايف برابط
  if (!dataUrl) return '';
  const m = String(dataUrl).match(/^data:(image\/[a-z]+);base64,(.+)$/);
  if (!m) return '';
  const folder = ensureFolder_();
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], name + (m[1] === 'image/png' ? '.png' : '.jpg'));
  return folder.createFile(blob).getUrl();   // الملف بياخد صلاحيات الفولدر (خاص بالمستخدمين المسجلين بس)
}
function clean_(o) { const r = {}; Object.keys(o || {}).forEach(function (k) { r[k] = typeof o[k] === 'string' ? o[k].trim() : o[k]; }); return r; }

/* ===================== حالة المفقود ===================== */
function catMap_() {
  const m = {};
  rows_('cats').forEach(function (r) { m[String(r['النوع'])] = { days: Number(r['مدة الحفظ']) || 30, valuable: String(r['ثمين']) === 'نعم' }; });
  return m;
}
function itemsWithStatus_() {
  const cats = catMap_();
  const delivered = {}; rows_('deliv').forEach(function (r) { delivered[String(r['رقم المفقود'])] = true; });
  const disposed = {};  rows_('other').forEach(function (r) { disposed[String(r['رقم المفقود'])] = true; });
  const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  return rows_('items').map(function (r) {
    const id = String(r['رقم المفقود']);
    const c = cats[String(r['النوع'])] || { days: 30, valuable: false };
    const found = r['تاريخ العثور'] instanceof Date ? r['تاريخ العثور'] : toDate_(r['تاريخ العثور']);
    const due = found ? new Date(found.getTime() + c.days * 86400000) : '';
    let status = 'في المخزن';
    if (delivered[id]) status = 'تم التسليم';
    else if (disposed[id]) status = 'سُلّم لجهة أخرى';
    else if (r['تأكيد الإدارة'] === 'مرفوض') status = 'مرفوض';
    else if (r['تأكيد الإدارة'] !== 'مؤكد') status = 'منتظر تأكيد الإدارة';
    else if (due && iso_(due) < today) status = 'متأخر عن مدة الحفظ';
    return {
      id: id, found: iso_(found), time: time_(r['الوقت']), cat: String(r['النوع']), desc: String(r['الوصف']), brand: String(r['اللون والماركة']),
      loc: String(r['مكان العثور']), locd: String(r['تفاصيل المكان']), sup: String(r['المشرف']), sector: String(r['القطاع']), finder: String(r['اللي لقاه']),
      store: String(r['مكان التخزين']), photo: String(r['الصورة']), confirm: String(r['تأكيد الإدارة']), notes: String(r['ملاحظات']),
      days: c.days, due: iso_(due), valuable: c.valuable, status: status
    };
  });
}
function permitsWithStatus_() {
  const used = {};
  rows_('deliv').forEach(function (r) { used[String(r['رقم الإذن'])] = true; });
  rows_('other').forEach(function (r) { used[String(r['رقم الإذن'])] = true; });
  const now = new Date();
  return rows_('permits').map(function (r) {
    const id = String(r['رقم الإذن']);
    const dec = String(r['قرار الإدارة'] || 'منتظر');
    const at = r['تاريخ ووقت القرار'] instanceof Date ? r['تاريخ ووقت القرار'] : (r['تاريخ ووقت القرار'] ? toDate_(r['تاريخ ووقت القرار']) : '');
    const h = hours_(r['الصلاحية']) || 48;
    const until = at ? new Date(at.getTime() + h * 3600e3) : '';
    let status = 'منتظر موافقة الإدارة';
    if (dec === 'مرفوض') status = 'مرفوض';
    else if (dec === 'موافق') status = used[id] ? 'اتستخدم' : (until && now > until ? 'منتهي' : 'ساري');
    return {
      id: id, item: String(r['رقم المفقود']), purpose: String(r['الغرض']), who: String(r['المستلم أو الجهة']), phone: String(r['الموبايل']),
      res: String(r['صفة المستلم']), unit: String(r['الوحدة']), asked: iso_(r['تاريخ الطلب']), decision: dec, decidedAt: iso_(at, true),
      hours: h, until: iso_(until, true), note: String(r['ملاحظة الإدارة']), status: status
    };
  });
}

/* ===================== تحميل الأبلكيشن ===================== */
function boot(t) {
  const m = need_(t, ['admin', 'keeper', 'supervisor', 'cr']);
  const lists = {
    cats: rows_('cats').map(function (r) { return { name: String(r['النوع']), days: Number(r['مدة الحفظ']) || 30, valuable: String(r['ثمين']) === 'نعم' }; }),
    locs: list_('locs', 'المكان'), stores: list_('stores', 'المكان'), staff: list_('staff', 'الاسم'), parties: list_('parties', 'الجهة'),
    prios: rows_('prios').map(function (r) { return { name: String(r['الأولوية']), hours: hours_(r['المهلة']) || 24 }; })
  };
  const crRows = rows_('cr').map(function (r) {
    return { id: String(r['رقم البلاغ']), sent: iso_(r['تاريخ ووقت الإرسال'], true), prio: String(r['الأولوية']), req: String(r['المطلوب']),
      places: String(r['الأماكن والكاميرات']), window: String(r['فترة المراجعة']), route: String(r['خط سير صاحب البلاغ']), look: String(r['أوصاف صاحب البلاغ']),
      result: String(r['نتيجة المراجعة']), cams: String(r['الكاميرات اللي اتراجعت']), text: String(r['الإفادة']), photos: String(r['الصور']),
      replied: iso_(r['تاريخ ووقت الرد'], true), summary: String(r['ملخص البلاغ']) };
  });
  const out = { me: m, lists: lists, cr: crRows, now: iso_(new Date(), true), today: iso_(new Date()) };
  if (m.role === 'admin') out.users = accounts_().filter(function (u) { return u.role; }).map(function (u) { return { user: u.user, name: u.name, role: u.role, active: u.active, by: u.by, at: u.at }; });
  if (m.role === 'cr') return out;                       // الكنترول روم: بلاغاتهم بس
  out.items = itemsWithStatus_();
  if (m.role === 'supervisor') { out.items = out.items.map(function (i) { return { id: i.id, found: i.found, cat: i.cat, desc: i.desc, loc: i.loc, store: i.store, status: i.status }; }); return out; }
  out.permits = permitsWithStatus_();
  out.deliveries = rows_('deliv').map(function (r) {
    return { id: String(r['رقم التسليم']), item: String(r['رقم المفقود']), permit: String(r['رقم الإذن']), at: iso_(r['تاريخ ووقت التسليم'], true),
      name: String(r['اسم المستلم']), rel: String(r['علاقته بالمفقود']), res: String(r['صفته في مراسي']), unit: String(r['الوحدة']),
      idType: String(r['نوع الهوية']), idNo: String(r['رقم الهوية']), phone: String(r['الموبايل']), blind: String(r['وصفه قبل العرض']),
      match: String(r['نتيجة المطابقة']), proofs: String(r['إثبات الملكية']), releasedBy: String(r['خرّجه من المخزن']), witness: String(r['الشاهد']),
      cond: String(r['حالة المفقود']), idPhoto: String(r['صورة الهوية']), appPhoto: String(r['صورة الأبلكيشن أو QR']), sig: String(r['توقيع المستلم']), notes: String(r['ملاحظات']) };
  });
  out.others = rows_('other').map(function (r) {
    return { id: String(r['رقم العملية']), item: String(r['رقم المفقود']), permit: String(r['رقم الإذن']), date: iso_(r['التاريخ']), party: String(r['الجهة']),
      ref: String(r['رقم المحضر']), to: String(r['جهة الاستلام']), releasedBy: String(r['خرّجه من المخزن']), notes: String(r['ملاحظات']) };
  });
  out.reports = rows_('reports').map(function (r) {
    return { id: String(r['رقم البلاغ']), date: iso_(r['تاريخ البلاغ']), name: String(r['الاسم']), phone: String(r['الموبايل']), res: String(r['صفته في مراسي']),
      unit: String(r['الوحدة']), cat: String(r['النوع']), desc: String(r['الوصف']), brand: String(r['اللون والماركة']), from: iso_(r['ضاعت من']), to: iso_(r['لحد']),
      loc: String(r['المكان المتوقع']), locd: String(r['تفاصيل المكان']), item: String(r['رقم المفقود المطابق']), status: String(r['حالة البلاغ'] || 'مفتوح'), notes: String(r['متابعات']) };
  });
  return out;
}

/* ===================== العمليات ===================== */
function saveItem(t, o) {
  const m = need_(t, ['admin', 'keeper', 'supervisor']);
  o = clean_(o);
  if (!o.desc || !o.cat || !o.found) throw new Error('اكتب النوع والوصف وتاريخ العثور.');
  return withLock_(function () {
    const y = String(o.found).slice(0, 4);
    const id = nextId_('items', 'LF-' + y + '-', 5);
    append_('items', {
      'رقم المفقود': id, 'تاريخ العثور': toDate_(o.found), 'الوقت': o.time ? "'" + o.time : '', 'النوع': o.cat, 'الوصف': o.desc, 'اللون والماركة': o.brand || '',
      'مكان العثور': o.loc || '', 'تفاصيل المكان': o.locd || '', 'المشرف': o.sup || '', 'القطاع': o.sector || '', 'اللي لقاه': o.finder || '',
      'مكان التخزين': m.role === 'supervisor' ? '' : (o.store || ''), 'الصورة': saveFile_(o.photo, id),
      'تأكيد الإدارة': m.role === 'admin' ? 'مؤكد' : 'منتظر', 'ملاحظات': o.notes || '', 'سجّله': m.user, 'أكّده': m.role === 'admin' ? m.user : ''
    });
    return { id: id, confirmed: m.role === 'admin' };
  });
}
function confirmItem(t, id, decision, store) {
  const m = need_(t, ['admin']);
  const patch = { 'تأكيد الإدارة': decision === 'مرفوض' ? 'مرفوض' : 'مؤكد', 'أكّده': m.user };
  if (store) patch['مكان التخزين'] = store;
  update_('items', id, patch);
  return true;
}
function moveItem(t, id, store) { need_(t, ['admin', 'keeper']); update_('items', id, { 'مكان التخزين': store }); return true; }

function requestPermit(t, o) {
  const m = need_(t, ['admin', 'keeper']);
  o = clean_(o);
  if (['تسليم لصاحبه', 'تسليم لجهة أخرى'].indexOf(o.purpose) < 0) throw new Error('اختار غرض الإذن.');
  if (o.purpose === 'تسليم لصاحبه' && (!o.who || !o.phone)) throw new Error('اكتب اسم المستلم والموبايل.');
  return withLock_(function () {
    const item = itemsWithStatus_().filter(function (i) { return i.id === o.item; })[0];
    if (!item) throw new Error('رقم المفقود مش موجود.');
    if (['في المخزن', 'متأخر عن مدة الحفظ'].indexOf(item.status) < 0) throw new Error('المفقود ده مش في المخزن (' + item.status + ').');
    const open = permitsWithStatus_().filter(function (p) { return p.item === o.item && (p.status === 'منتظر موافقة الإدارة' || p.status === 'ساري'); })[0];
    if (open) throw new Error('فيه إذن ' + open.status + ' للمفقود ده بالفعل: ' + open.id);
    const y = Utilities.formatDate(new Date(), TZ, 'yyyy');
    const id = nextId_('permits', 'EX-' + y + '-', 5);
    const isAdmin = m.role === 'admin';
    append_('permits', {
      'رقم الإذن': id, 'رقم المفقود': o.item, 'الغرض': o.purpose, 'المستلم أو الجهة': o.who || '', 'الموبايل': o.phone || '', 'صفة المستلم': o.res || '',
      'الوحدة': o.unit || '', 'تاريخ الطلب': new Date(), 'قرار الإدارة': isAdmin ? 'موافق' : 'منتظر', 'تاريخ ووقت القرار': isAdmin ? new Date() : '',
      'الصلاحية': Number(o.hours) || 48, 'ملاحظة الإدارة': o.note || '', 'طلبه': m.user
    });
    return { id: id, approved: isAdmin };
  });
}
function decidePermit(t, id, decision, hours, note) {
  need_(t, ['admin']);
  if (decision === 'مرفوض' && !String(note || '').trim()) throw new Error('اكتب سبب الرفض.');
  update_('permits', id, { 'قرار الإدارة': decision === 'مرفوض' ? 'مرفوض' : 'موافق', 'تاريخ ووقت القرار': new Date(), 'الصلاحية': Number(hours) || 48, 'ملاحظة الإدارة': note || '' });
  return true;
}
function checkRelease_(itemId, permitId, purpose, role) {
  const item = itemsWithStatus_().filter(function (i) { return i.id === itemId; })[0];
  if (!item) throw new Error('رقم المفقود مش موجود.');
  if (['في المخزن', 'متأخر عن مدة الحفظ'].indexOf(item.status) < 0) throw new Error('المفقود ده مش في المخزن (' + item.status + ').');
  if (!permitId && role === 'admin') return 'إذن مباشر من الإدارة';
  const p = permitsWithStatus_().filter(function (x) { return x.id === permitId; })[0];
  if (!p) throw new Error('لازم إذن خروج ساري قبل التسليم.');
  if (p.item !== itemId) throw new Error('الإذن ده لمفقود تاني.');
  if (p.purpose !== purpose) throw new Error('الإذن ده لغرض «' + p.purpose + '».');
  if (p.status !== 'ساري') throw new Error('الإذن حالته «' + p.status + '».');
  return p.id;
}
function deliver(t, o) {
  const m = need_(t, ['admin', 'keeper']);
  o = clean_(o);
  const req = [['name', 'اسم المستلم'], ['idNo', 'رقم الهوية'], ['phone', 'الموبايل'], ['blind', 'وصفه قبل العرض'], ['match', 'نتيجة المطابقة'], ['releasedBy', 'اللي خرّج المفقود']];
  req.forEach(function (x) { if (!o[x[0]]) throw new Error('اكتب ' + x[1] + '.'); });
  if (o.match === 'غير مطابق') throw new Error('الوصف غير مطابق: ما يتمّش التسليم.');
  if (['مطابق تماماً', 'مطابق جزئياً'].indexOf(o.match) < 0) throw new Error('اختار نتيجة المطابقة.');
  const proofs = String(o.proofs || '').split('،').map(function (s) { return s.trim(); }).filter(String);
  if (!proofs.length) throw new Error('اختار طريقة إثبات ملكية واحدة على الأقل.');
  if (!o.sig) throw new Error('توقيع المستلم مطلوب.');
  return withLock_(function () {
    const item = itemsWithStatus_().filter(function (i) { return i.id === o.item; })[0];
    if (!item) throw new Error('رقم المفقود مش موجود.');
    const need = (item.valuable || o.match === 'مطابق جزئياً') ? 2 : 1;
    if (proofs.length < need) throw new Error('مطلوب ' + (need === 2 ? 'طريقتين إثبات' : 'طريقة إثبات') + ' على الأقل.');
    if (item.valuable && !o.witness) throw new Error('مفقود ثمين: اكتب اسم الشاهد.');
    if (item.valuable && !o.idPhoto) throw new Error('مفقود ثمين: صورة الهوية مطلوبة.');
    const permit = checkRelease_(o.item, o.permit, 'تسليم لصاحبه', m.role);
    const id = 'DL-' + Utilities.formatDate(new Date(), TZ, 'yyMMddHHmmss') + Math.floor(Math.random() * 90 + 10);
    append_('deliv', {
      'رقم التسليم': id, 'رقم المفقود': o.item, 'رقم الإذن': permit, 'تاريخ ووقت التسليم': new Date(), 'اسم المستلم': o.name, 'علاقته بالمفقود': o.rel || '',
      'صفته في مراسي': o.res || '', 'الوحدة': o.unit || '', 'نوع الهوية': o.idType || '', 'رقم الهوية': o.idNo, 'الموبايل': o.phone, 'وصفه قبل العرض': o.blind,
      'نتيجة المطابقة': o.match, 'إثبات الملكية': o.proofs, 'خرّجه من المخزن': o.releasedBy, 'الشاهد': o.witness || '', 'حالة المفقود': o.cond || '',
      'صورة الهوية': saveFile_(o.idPhoto, id + '-هوية'), 'صورة الأبلكيشن أو QR': saveFile_(o.appPhoto, id + '-أبلكيشن'), 'توقيع المستلم': saveFile_(o.sig, id + '-توقيع'),
      'ملاحظات': o.notes || '', 'سجّله': m.user
    });
    return { id: id };
  });
}
function dispose(t, o) {
  const m = need_(t, ['admin', 'keeper']);
  o = clean_(o);
  if (!o.party || !o.releasedBy) throw new Error('اختار الجهة واسم اللي خرّج المفقود.');
  return withLock_(function () {
    const permit = checkRelease_(o.item, o.permit, 'تسليم لجهة أخرى', m.role);
    const id = 'OT-' + Utilities.formatDate(new Date(), TZ, 'yyMMddHHmmss') + Math.floor(Math.random() * 90 + 10);
    append_('other', { 'رقم العملية': id, 'رقم المفقود': o.item, 'رقم الإذن': permit, 'التاريخ': new Date(), 'الجهة': o.party, 'رقم المحضر': o.ref || '',
      'جهة الاستلام': o.to || '', 'خرّجه من المخزن': o.releasedBy, 'ملاحظات': o.notes || '', 'سجّله': m.user });
    return { id: id };
  });
}

function saveReport(t, o) {
  const m = need_(t, ['admin', 'keeper']);
  o = clean_(o);
  if (!o.name || !o.phone || !o.desc || !o.cat) throw new Error('اكتب الاسم والموبايل والنوع والوصف.');
  if (o.id) {
    update_('reports', o.id, { 'الاسم': o.name, 'الموبايل': o.phone, 'صفته في مراسي': o.res || '', 'الوحدة': o.unit || '', 'النوع': o.cat, 'الوصف': o.desc,
      'اللون والماركة': o.brand || '', 'ضاعت من': toDate_(o.from), 'لحد': toDate_(o.to || o.from), 'المكان المتوقع': o.loc || '', 'تفاصيل المكان': o.locd || '' });
    return { id: o.id };
  }
  return withLock_(function () {
    const id = nextId_('reports', 'BL-', 4);
    append_('reports', { 'رقم البلاغ': id, 'تاريخ البلاغ': new Date(), 'الاسم': o.name, 'الموبايل': o.phone, 'صفته في مراسي': o.res || '', 'الوحدة': o.unit || '',
      'النوع': o.cat, 'الوصف': o.desc, 'اللون والماركة': o.brand || '', 'ضاعت من': toDate_(o.from), 'لحد': toDate_(o.to || o.from), 'المكان المتوقع': o.loc || '',
      'تفاصيل المكان': o.locd || '', 'رقم المفقود المطابق': '', 'حالة البلاغ': 'مفتوح', 'متابعات': '', 'سجّله': m.user });
    return { id: id };
  });
}
function reportAction(t, id, action, value) {
  need_(t, ['admin', 'keeper']);
  const r = rows_('reports').filter(function (x) { return String(x['رقم البلاغ']) === id; })[0];
  if (!r) throw new Error('البلاغ مش موجود.');
  const stamp = Utilities.formatDate(new Date(), TZ, 'dd/MM HH:mm');
  if (action === 'link') {
    const it = itemsWithStatus_().filter(function (i) { return i.id === String(value); })[0];
    if (!it) throw new Error('رقم المفقود مش موجود: ' + value);
    if (['في المخزن', 'متأخر عن مدة الحفظ'].indexOf(it.status) < 0) throw new Error('المفقود ده مش في المخزن (' + it.status + ').');
  }
  if (action === 'link') update_('reports', id, { 'رقم المفقود المطابق': value, 'حالة البلاغ': 'اتطابق' });
  else if (action === 'unlink') update_('reports', id, { 'رقم المفقود المطابق': '', 'حالة البلاغ': 'مفتوح' });
  else if (action === 'close') update_('reports', id, { 'حالة البلاغ': 'مقفول', 'متابعات': (r['متابعات'] ? r['متابعات'] + '\n' : '') + stamp + ' قفل: ' + value });
  else if (action === 'note') update_('reports', id, { 'متابعات': (r['متابعات'] ? r['متابعات'] + '\n' : '') + stamp + ' ' + value });
  return true;
}

function sendToCR(t, o) {
  need_(t, ['admin']);
  o = clean_(o);
  if (!o.req) throw new Error('اكتب المطلوب من الكنترول روم.');
  if (rows_('cr').some(function (r) { return String(r['رقم البلاغ']) === o.id; })) throw new Error('البلاغ ده اتبعت للكنترول روم قبل كده.');
  const r = rows_('reports').filter(function (x) { return String(x['رقم البلاغ']) === o.id; })[0];
  if (!r) throw new Error('البلاغ مش موجود.');
  const summary = [r['النوع'], r['الوصف'], r['اللون والماركة'], 'ضاعت ' + iso_(r['ضاعت من']) + (r['لحد'] ? ' لحد ' + iso_(r['لحد']) : ''),
    'المكان: ' + (r['المكان المتوقع'] || 'غير معروف') + (r['تفاصيل المكان'] ? ' - ' + r['تفاصيل المكان'] : ''),
    'صاحب البلاغ: ' + r['الاسم'] + ' (' + (r['صفته في مراسي'] || '') + (r['الوحدة'] ? ' ' + r['الوحدة'] : '') + ')'].filter(String).join(' · ');
  append_('cr', { 'رقم البلاغ': o.id, 'تاريخ ووقت الإرسال': new Date(), 'الأولوية': o.prio || 'عادي', 'المطلوب': o.req, 'الأماكن والكاميرات': o.places || '',
    'فترة المراجعة': o.window || '', 'خط سير صاحب البلاغ': o.route || '', 'أوصاف صاحب البلاغ': o.look || '', 'ملخص البلاغ': summary });
  return true;
}
function crReply(t, o) {
  const m = need_(t, ['admin', 'cr']);
  o = clean_(o);
  if (!o.result) throw new Error('اختار نتيجة المراجعة.');
  const photos = (o.photos || []).slice(0, 4).map(function (p, i) { return saveFile_(p, o.id + '-كنترول-' + (i + 1)); }).filter(String).join('\n');
  const patch = { 'نتيجة المراجعة': o.result, 'الكاميرات اللي اتراجعت': o.cams || '', 'الإفادة': o.text || '', 'تاريخ ووقت الرد': new Date(), 'رد بواسطة': m.user };
  if (photos) patch['الصور'] = photos;
  update_('cr', o.id, patch);
  return true;
}

/* ===================== نقل بيانات البرنامج القديم ===================== */
/**
 * شغّلها من المحرر بعد ما تحط OLD_SHEET_ID فوق. بتنقل:
 *  - الأنواع والأماكن (اللي مش موجودة)،
 *  - الحسابات بكلمات السر بتاعتها (الإدارة ← الإدارة، مشرف أمن ومدير ← أمين مخزن؛ عدّل الأدوار بعدها من تاب الحسابات)،
 *  - المفقودات بنفس الكود القديم (علشان الملصقات اللي على الأكياس تفضل زي ما هي) مع صورها،
 *  - التسليمات مع صور البطايق.
 * المفقودات اللي اتمسحت في القديم مش بتتنقل. لو شغلتها تاني مش بتكرر حاجة اتنقلت قبل كده،
 * فلو ظهر في السجل «لسه فاضل»، شغّلها تاني وهتكمّل من مكان ما وقفت.
 */
function importOld() {
  if (!OLD_SHEET_ID) throw new Error('حط ID الشيت القديم في OLD_SHEET_ID في أول الملف.');
  setup_();
  const folder = ensureFolder_();
  const old = SpreadsheetApp.openById(OLD_SHEET_ID);
  const t0 = Date.now(), LIMIT = 4.5 * 60 * 1000;
  const out = { cats: 0, locs: 0, users: 0, items: 0, deliv: 0, photos: 0, photoFail: 0 };
  const read = function (name) {
    const sh = old.getSheetByName(name);
    if (!sh || sh.getLastRow() < 2) return [];
    const v = sh.getDataRange().getValues(), h = v[0].map(function (x) { return String(x).trim(); });
    return v.slice(1).map(function (r) { const o = {}; h.forEach(function (k, i) { o[k] = r[i]; }); return o; });
  };
  const str = function (v) { return v instanceof Date ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd HH:mm') : String(v == null ? '' : v).replace(/^'/, '').trim(); };
  const when = function (v) {
    if (v instanceof Date) return v;
    const t = str(v); if (!t) return '';
    let m = t.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2}))?/);
    if (m) return toDate_(m[1] + '-' + pad_(m[2], 2) + '-' + pad_(m[3], 2) + (m[4] ? 'T' + pad_(m[4], 2) + ':' + m[5] : ''));
    m = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (m) return toDate_(m[3] + '-' + pad_(m[2], 2) + '-' + pad_(m[1], 2));
    return '';
  };
  const copyPhoto = function (id, name) {
    id = str(id); if (!id) return '';
    try { const f = DriveApp.getFileById(id).makeCopy(name, folder); out.photos++; return f.getUrl(); }
    catch (e) { out.photoFail++; return ''; }
  };
  const VALUABLE = ['موبايل', 'محفظة', 'ساعة', 'مجوهرات', 'لابتوب', 'تابلت', 'كاميرا', 'كارنيهات / بطاقات', 'كروت بنك', 'فلوس'];
  const cats = {}; list_('cats', 'النوع').forEach(function (c) { cats[c] = 1; });
  const locs = {}; list_('locs', 'المكان').forEach(function (c) { locs[c] = 1; });
  const addCat = function (c) { if (!c || cats[c]) return; const v = VALUABLE.indexOf(c) >= 0; sheet_('cats').appendRow([c, v ? 90 : 30, v ? 'نعم' : 'لا']); cats[c] = 1; out.cats++; };
  const addLoc = function (c) { if (!c || locs[c]) return; sheet_('locs').appendRow([c]); locs[c] = 1; out.locs++; };

  // 1) القوائم
  read('القوائم').forEach(function (r) {
    if (str(r['نشط']) === 'لا') return;
    const v = str(r['القيمة']);
    if (str(r['القائمة']) === 'نوع') addCat(v);
    if (str(r['القائمة']) === 'مكان') addLoc(v);
  });

  // 2) الحسابات
  const accs = {}; accounts_().forEach(function (a) { accs[a.user] = 1; });
  const ROLE_MAP = { 'الإدارة': 'admin', 'admin': 'admin', 'مشرف أمن': 'keeper', 'security': 'keeper', 'مدير': 'keeper', 'manager': 'keeper' };
  read('المستخدمين').forEach(function (r) {
    const user = normUser_(r['اسم المستخدم']);
    if (!user || accs[user] || !str(r['salt']) || !str(r['hash'])) return;
    append_('accounts', { 'اسم المستخدم': user, 'الاسم': str(r['الاسم']) || user, 'الصلاحية': ROLES[ROLE_MAP[str(r['الدور'])] || 'keeper'],
      'الحالة': str(r['نشط']) === 'لا' ? 'موقوف' : 'نشط', 'كلمة السر': 'old$' + str(r['salt']) + '$' + str(r['hash']),
      'أضافه': 'البرنامج القديم', 'تاريخ الإضافة': when(r['تاريخ الإنشاء']) || new Date() });
    accs[user] = 1; out.users++;
  });

  // 3) المفقودات
  const have = {}; rows_('items').forEach(function (r) { have[String(r['رقم المفقود'])] = 1; });
  const items = read('المفقودات');
  let left = 0;
  items.forEach(function (r) {
    const code = str(r['الكود']);
    if (!code || str(r['محذوف']) || have[code]) return;
    if (Date.now() - t0 > LIMIT) { left++; return; }
    const cat = str(r['النوع']) || 'أخرى', loc = str(r['المكان']);
    addCat(cat); addLoc(loc);
    let photo = copyPhoto(r['صورة (Drive ID)'], code);
    if (!photo && /^data:image\//.test(str(r['مصغّرة']))) { try { photo = saveFile_(str(r['مصغّرة']), code); } catch (e) {} }
    const notes = ['منقول من البرنامج القديم · الحالة: ' + (str(r['الحالة']) || '—'),
      str(r['موبايل اللي لقاها']) ? 'موبايل اللي لقاها: ' + str(r['موبايل اللي لقاها']) : '',
      str(r['تبعيته']) ? 'تبعيته: ' + str(r['تبعيته']) : '',
      str(r['وقت التسليم للأمن']) ? 'اتسلّمت للأمن: ' + str(r['وقت التسليم للأمن']) : ''].filter(String).join(' · ');
    const found = when(r['تاريخ العثور']) || when(r['وقت التسجيل']) || new Date();
    const tm = str(r['وقت العثور']).match(/(\d{1,2}):(\d{2})/);
    append_('items', { 'رقم المفقود': code, 'تاريخ العثور': found, 'الوقت': tm ? "'" + pad_(tm[1], 2) + ':' + tm[2] : '', 'النوع': cat,
      'الوصف': str(r['الوصف']), 'اللون والماركة': '', 'مكان العثور': loc, 'تفاصيل المكان': '', 'المشرف': str(r['مشرف الأمن المستلم']),
      'القطاع': '', 'اللي لقاه': str(r['اسم اللي لقاها']), 'مكان التخزين': '', 'الصورة': photo, 'تأكيد الإدارة': 'مؤكد',
      'ملاحظات': notes, 'سجّله': str(r['سجّلها']), 'أكّده': 'البرنامج القديم' });
    have[code] = 1; out.items++;
  });

  // 4) التسليمات (لمفقودات اتنقلت بس)
  const done = {}; rows_('deliv').forEach(function (r) { done[String(r['رقم التسليم'])] = 1; });
  read('التسليمات').forEach(function (r) {
    const code = str(r['الكود']), id = 'DL-OLD-' + code;
    if (!code || !have[code] || done[id]) return;
    if (Date.now() - t0 > LIMIT) { left++; return; }
    append_('deliv', { 'رقم التسليم': id, 'رقم المفقود': code, 'رقم الإذن': 'البرنامج القديم', 'تاريخ ووقت التسليم': when(r['وقت التسليم']) || '',
      'اسم المستلم': str(r['اسم صاحبها']), 'علاقته بالمفقود': '', 'صفته في مراسي': str(r['الصفة']), 'الوحدة': str(r['رقم الوحدة']),
      'الموبايل': str(r['الموبايل']) ? "'" + str(r['الموبايل']) : '', 'خرّجه من المخزن': str(r['سلّمها']),
      'صورة الهوية': copyPhoto(r['صورة البطاقة (Drive ID)'], id + '-هوية'), 'ملاحظات': str(r['ملاحظات']), 'سجّله': 'البرنامج القديم' });
    done[id] = 1; out.deliv++;
  });
  // مفقودات كانت «اتسلّمت» في القديم من غير بيانات تسليم: تتسجّل مسلّمة علشان ما تظهرش في المخزن
  items.forEach(function (r) {
    const code = str(r['الكود']), id = 'DL-OLD-' + code;
    if (!code || !have[code] || done[id] || str(r['الحالة']).indexOf('اتسلّمت') < 0) return;
    append_('deliv', { 'رقم التسليم': id, 'رقم المفقود': code, 'رقم الإذن': 'البرنامج القديم', 'تاريخ ووقت التسليم': when(r['آخر تعديل']) || '',
      'ملاحظات': 'اتسلّمت في البرنامج القديم من غير بيانات تسليم', 'سجّله': 'البرنامج القديم' });
    done[id] = 1; out.deliv++;
  });

  Logger.log('اتنقل: ' + out.items + ' مفقود، ' + out.deliv + ' تسليم، ' + out.users + ' حساب، ' + out.cats + ' نوع، ' + out.locs + ' مكان، ' +
    out.photos + ' صورة' + (out.photoFail ? ' (' + out.photoFail + ' صورة ما اتنقلتش)' : '') +
    (left ? '. ⚠️ لسه فاضل ' + left + ' سجل: شغّل importOld تاني وهتكمّل.' : '. خلص النقل كله.'));
}
