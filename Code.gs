/**
 * מערכת אילוצים וסידור עבודה - מיון ילדים (גרסה 3)
 * קובץ שרת (Google Apps Script). מודבק בתוך: הרחבות > Apps Script של הגיליון.
 *
 * התקנה ראשונה או שדרוג מגרסה קודמת: בחרי בפונקציה setup בתפריט למעלה ולחצי "הפעלה".
 * ההרצה לא מוחקת נתונים קיימים.
 */
var INITIAL_ADMIN_PASSWORD = '1234';

var SH = {
  settings: 'הגדרות',
  employees: 'עובדים',
  responses: 'אילוצים_נתונים',
  schedules: 'סידורים_נתונים',
  requests: 'בקשות',
  periods: 'תקופות',
  versions: 'גרסאות_סידור'
};
var EMP_HEADERS = ['שם', 'מייל', 'פעיל', 'מינימום משמרות', 'מקסימום משמרות', 'צבע', 'ותק', 'מקסימום חסימות', 'מתחיל מתאריך', 'הגשה פתוחה עד', 'סיסמה (מוצפנת)', 'בדיקת ותק', 'רצף ברוצה'];
var RESP_HEADERS = ['תקופה', 'שם', 'מייל', 'עודכן', 'מספר חסימות', 'נתונים', 'הערה'];
var SCHED_HEADERS = ['תקופה', 'נתונים', 'פורסם', 'עודכן'];
var REQ_HEADERS = ['מזהה', 'סוג', 'סטטוס', 'מבקש', 'מול', 'פרטים', 'הערה', 'נוצר', 'טופל'];
var SETTING_KEYS = ['periodId', 'start', 'end', 'maxBlocks', 'holidays', 'open', 'title', 'emailNotify', 'deadline',
  'shiftStart', 'shiftEnd', 'wkStart', 'wkEnd', 'reminders', 'adminEmail', 'shownBlocks', 'siteUrl', 'wkBlocks'];
var PERIOD_HEADERS = ['תקופה', 'מתאריך', 'עד תאריך', 'חגים'];

/* הצבעים מהסידור הקיים */
var KNOWN_COLORS = {
  'נוה': '#FF9900', 'עמית': '#9FC5E8', 'רון': '#76A5AF', 'גיא': '#3D85C6', 'מטי': '#D0E0E3',
  'בר': '#D9D2E9', 'שירה': '#BF9000', 'עדין': '#F9CB9C', 'דוד': '#A64D79', 'סיוון': '#674EA7',
  'רונה': '#E06666', 'אור': '#6AA84F', 'Ziv': '#FFD966'
};
var PALETTE = ['#FFD966', '#93C47D', '#C27BA0', '#B4A7D6', '#F6B26B', '#8E7CC3', '#45818E', '#CC4125',
  '#B6D7A8', '#EA9999', '#A2C4C9', '#FFE599'];

/* ================= Web app ================= */

function doGet(e) {
  var page = (e && e.parameter && e.parameter.page) === 'admin' ? 'Admin' : 'Index';
  return HtmlService.createHtmlOutputFromFile(page)
    .setTitle(page === 'Admin' ? 'ניהול סידור עבודה' : 'אילוצים וסידור - מיון ילדים')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/* ================= API (לאתר שמאוחסן ב-GitHub) ================= */

/** רק הפונקציות האלה נגישות מבחוץ */
var API_ = {
  getPublicData: getPublicData, getPublishedSchedule: getPublishedSchedule, getMyData: getMyData,
  submitResponse: submitResponse, requestSwap: requestSwap, requestGive: requestGive, respondSwap: respondSwap,
  cancelRequest: cancelRequest, requestVacation: requestVacation, getLoginState: getLoginState, login: login,
  sendSetupCode: sendSetupCode, setPasswordWithCode: setPasswordWithCode,
  adminLoad: adminLoad, adminSaveSettings: adminSaveSettings, adminSaveEmployees: adminSaveEmployees,
  adminSaveResponse: adminSaveResponse, adminSaveSchedule: adminSaveSchedule, adminResolveVacation: adminResolveVacation,
  adminSetRequestStatus: adminSetRequestStatus, adminResolveSwap: adminResolveSwap, adminSeniority: adminSeniority, adminChangePassword: adminChangePassword,
  adminExport: adminExport, adminSetExtension: adminSetExtension, adminSetCurrentPeriod: adminSetCurrentPeriod,
  adminDeletePeriod: adminDeletePeriod, adminResetPassword: adminResetPassword, adminListVersions: adminListVersions
};
/** פעולות שרק קוראות נתונים (אפשר לקרוא כל לשונית פעם אחת) */
var READONLY_ = { getPublicData: 1, getPublishedSchedule: 1, getMyData: 1, getLoginState: 1, adminLoad: 1, adminListVersions: 1 };
function doPost(e) {
  var out;
  try {
    var req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var fn = API_[req.fn];
    if (!fn) throw new Error('פעולה לא מוכרת');
    MEMO_ = READONLY_[req.fn] ? {} : null;
    out = { ok: true, result: fn.apply(null, req.args || []) };
  } catch (err) {
    out = { ok: false, error: String(err && err.message ? err.message : err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/* ================= Setup / upgrade ================= */

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties().setProperty('SSID', ss.getId());
  ss.setSpreadsheetLocale('iw_IL');

  var st = ensureSheet_(SH.settings, ['מפתח', 'ערך']);
  var emp = ensureSheet_(SH.employees, EMP_HEADERS);
  var rsh = ensureSheet_(SH.responses, RESP_HEADERS);
  rsh.getRange(1, 1, 1, RESP_HEADERS.length).setValues([RESP_HEADERS]).setFontWeight('bold').setBackground('#eeeeee');
  ensureSheet_(SH.schedules, SCHED_HEADERS);
  ensureSheet_(SH.requests, REQ_HEADERS);
  ensureSheet_(SH.periods, PERIOD_HEADERS);
  ensureSheet_(SH.versions, ['תקופה', 'זמן', 'סיבה', 'נתונים']);
  installTriggers_();

  if (st.getLastRow() < 2) {
    writeSettings_({
      periodId: '2026-12-06', start: '2026-12-06', end: '2027-01-02', maxBlocks: '8',
      holidays: '', open: 'TRUE', title: 'מיון ילדים - הגשת אילוצים', emailNotify: 'TRUE', deadline: ''
    });
  } else {
    var cur = readSettings_();
    writeSettings_(settingsToRow_(cur)); // מוסיף מפתחות חדשים בשדרוג
  }
  if (!PropertiesService.getScriptProperties().getProperty('ADMIN_PW')) {
    PropertiesService.getScriptProperties().setProperty('ADMIN_PW', hash_(INITIAL_ADMIN_PASSWORD));
  }

  if (emp.getLastRow() < 2) {
    var seed = SEED_();
    var rows = seed.map(function (s, i) {
      return [s[0], s[1], 'כן', '3', '6', colorFor_(s[0], i), 'ותיק', '', '', '', '', '', ''];
    });
    emp.getRange(2, 1, rows.length, EMP_HEADERS.length).setValues(rows);
    PropertiesService.getScriptProperties().setProperty('COLORS_V2', '1');
    var days = listDays_('2026-12-06', '2027-01-02');
    seed.forEach(function (s) {
      var vals = {};
      days.forEach(function (d, i) { vals[d] = s[2].charAt(i) || 'C'; });
      saveResponseRow_('2026-12-06', s[0], s[1], vals);
    });
  } else {
    // שדרוג: כותרות חדשות + צבעים מהסידור הקיים
    emp.getRange(1, 1, 1, EMP_HEADERS.length).setValues([EMP_HEADERS]).setFontWeight('bold').setBackground('#eeeeee');
    var props = PropertiesService.getScriptProperties();
    var firstUpgrade = !props.getProperty('COLORS_V2');
    var data = vals_(emp);
    for (var i = 1; i < data.length; i++) {
      if (!norm_(data[i][0])) continue;
      if (firstUpgrade && !data[i][5]) emp.getRange(i + 1, 6).setValue(colorFor_(norm_(data[i][0]), i - 1));
      if (!data[i][6]) emp.getRange(i + 1, 7).setValue('ותיק');
    }
    props.setProperty('COLORS_V2', '1');
  }
  var cs = readSettings_();
  savePeriodMeta_(cs.periodId, cs.start, cs.end, cs.holidays);
  if (!cs.shiftStart && !cs.shiftEnd) { // שעות ברירת מחדל למשמרת חול
    var row0 = settingsToRow_(cs); row0.shiftStart = '15:00'; row0.shiftEnd = '23:00'; writeSettings_(row0);
  }
  var url = '';
  try { url = ScriptApp.getService().getUrl(); } catch (err) {}
  Logger.log('ההתקנה הסתיימה. ' + (url ? 'כתובת: ' + url : 'עכשיו צריך לפרסם: פריסה > פריסה חדשה > אפליקציית אינטרנט.'));
}

function colorFor_(name, i) {
  var first = norm_(name).split(' ')[0];
  if (KNOWN_COLORS[first]) return KNOWN_COLORS[first];
  return PALETTE[i % PALETTE.length];
}

function SEED_() {
  return [
    ['עדין פלו', 'adin.pellow@mail.huji.ac.il', 'BBWCBCCWCCCBWWWCCCBWWWCCBBCB'],
    ['שירה וולפסון', 'shira.wolfson@mail.huji.ac.il', 'BCWCBBBCCWWCCWCCWWCCBWWCCBBB'],
    ['מטי', 'mathew.mesenyashin@mail.huji.ac.il', 'BBCCCCCBBCCWWWBBCCCCCBBCCBBB'],
    ['דוד צדוק', 'davidl.zadok@gmail.com', 'WBCWBCCWBCCBCBWBCBCCCWBCCCCC'],
    ['בר', 'bar.sabah@mail.huji.ac.il', 'BBBBBCCBBBWCCCCWWWCWWCWWWCCC'],
    ['גיא רון', 'guy.ron3@mail.huji.ac.il', 'BCCWCCCCWBCWBBCCWCCWBCCCWCBB'],
    ['עמית אייל עשהאל', 'amit.eyalassael@mail.huji.ac.il', 'CCCCCBCWBBBBBCCCCCBCCCCCCBCC'],
    ['רונה טנא', 'ronnie.tene@mail.huji.ac.il', 'BCBCWWWBCBCWWWBCBCWWWBCBCWWW'],
    ['סיוון חוכמן', 'sivan.hochman@mail.huji.ac.il', 'CWWWCCCCCCWWCCCCCCCBBWWWCBCC'],
    ['Ziv Yosef', 'ziv.yosef@mail.huji.ac.il', 'BCCCCCCBCCBCCCBCCBCCCBCCBCCC'],
    ['אור זילברברג', 'or.zilberberg@mail.huji.ac.il', 'WCWBCCCCCCCCWWWBBCWCCWCWCBBC'],
    ['נוה', '', 'CBBBBCCCCCCCCCCCCCCCCCCCCCCC'],
    ['רון', '', 'WCWCBBBBCWCBBCCBWCWBCWCWCBBC']
  ];
}

/* ================= Public API (employees) ================= */

function getPublicData() {
  var s = readSettings_();
  var emps = readEmployees_().filter(function (e) { return e.active; });
  var pubs = listPublished_();
  var pick = pickByDate_(pubs, todayIso_());
  var first = pick ? readSchedule_(pick.periodId) : null;
  return {
    settings: publicSettings_(s),
    submitOpen: submitOpenFor_(s, null),
    employees: emps.map(function (e) { return e.name; }),
    colors: colorMap_(),
    published: pubs,
    schedule: first ? schedPublic_(pick.periodId, first) : null
  };
}

/** סידור מפורסם של תקופה מסוימת (לבחירת חודש) */
function getPublishedSchedule(periodId) {
  var sc = readSchedule_(periodId);
  if (!sc || !sc.published) throw new Error('הסידור לא מפורסם');
  return schedPublic_(periodId, sc);
}

/** כל מה שעובד מחובר רואה. email = אסימון הכניסה (token) */
function getMyData(name, email, withPublic) {
  var emp = verifyEmployee_(name, email);
  if (withPublic) { var res = getMyData(name, email, false); res.pub = getPublicData(); return res; }
  var s = readSettings_();
  var r = readResponses_(s.periodId)[emp.name];
  var reqs = readRequests_().filter(function (q) {
    return q.from === emp.name || q.to === emp.name || (q.type === 'give' && q.to === '*' && q.status === 'pending');
  });
  return {
    name: emp.name,
    values: r ? r.values : null,
    note: r ? r.note : '',
    updated: r ? r.updated : null,
    maxBlocks: emp.maxBlocks !== '' ? parseInt(emp.maxBlocks, 10) : s.maxBlocks,
    shownBlocks: Math.min(s.shownBlocks, emp.maxBlocks !== '' ? parseInt(emp.maxBlocks, 10) : s.maxBlocks),
    start: emp.start || '',
    canSubmit: submitOpenFor_(s, emp),
    ext: emp.ext || '',
    requests: reqs.slice(-80),
    published: listPublished_(),
    colors: colorMap_()
  };
}
/* תאימות לגרסה הקודמת */
function getMyResponse(name, email) { return getMyData(name, email); }

function submitResponse(name, email, values, note) {
  return withLock_(function () {
    var s = readSettings_();
    var emp = verifyEmployee_(name, email);
    if (!submitOpenFor_(s, emp)) throw new Error('הגשת האילוצים לתקופה הזו סגורה. אם פספסת, פנה/י לאחראית שתפתח לך.');
    var limit = emp.maxBlocks !== '' ? parseInt(emp.maxBlocks, 10) : s.maxBlocks;
    var days = listDays_(s.start, s.end);
    var clean = {}, blocks = 0, wk = 0;
    days.forEach(function (d) {
      var v = values && values[d];
      v = (v === 'B' || v === 'W') ? v : 'C';
      if (v === 'B') { blocks++; if (dow_(d) === 5 || dow_(d) === 6 || s.holidays.indexOf(d) >= 0) wk++; }
      clean[d] = v;
    });
    if (blocks > limit) throw new Error('סימנת יותר מדי חסימות. צריך עוד? פנו לאחראית. בהיעדרות ממושכת יש לעדכן את האחראית מוקדם ככל האפשר.');
    if (wk > s.wkBlocks) throw new Error('אי אפשר לחסום כל כך הרבה ימי סופ"ש.');
    var updated = saveResponseRow_(s.periodId, emp.name, emp.email, clean, note == null ? '' : note);
    return { ok: true, updated: updated, blocks: blocks };
  });
}

/* ---------- החלפות ראש בראש ---------- */

function requestSwap(name, email, periodId, giveDate, takeDate, withName, note) {
  return withLock_(function () {
    var me = verifyEmployee_(name, email);
    var other = findEmployee_(withName);
    if (!other || !other.active) throw new Error('העובד/ת לא נמצא/ה');
    if (other.name === me.name) throw new Error('אי אפשר להחליף עם עצמך');
    var sched = readSchedule_(periodId);
    if (!sched || !sched.published) throw new Error('אין סידור מפורסם');
    checkSwapValid_(sched, me.name, giveDate, other.name, takeDate);
    var dup = readRequests_().some(function (q) {
      return q.type === 'swap' && q.status === 'pending' && q.from === me.name && q.to === other.name &&
        q.details.give === giveDate && q.details.take === takeDate;
    });
    if (dup) throw new Error('כבר שלחת בקשה כזו');
    var id = Utilities.getUuid().slice(0, 8);
    var det = { give: giveDate, take: takeDate, period: periodId };
    var needsAdmin = newPairDays_(sched, { type: 'swap', from: me.name, to: other.name, details: det }).length > 0;
    if (needsAdmin) det.adminReq = true;
    appendRequest_({ id: id, type: 'swap', status: 'pending', from: me.name, to: other.name, details: det, note: note || '' });
    notify_(other, 'בקשת החלפת משמרת מ' + me.name,
      me.name + ' מבקש/ת להחליף איתך:\n' +
      'את/ה תיקח/י את ' + dayName_(giveDate) + ' ' + fmtD_(giveDate) + '\n' +
      me.name + ' ייקח/תיקח את ' + dayName_(takeDate) + ' ' + fmtD_(takeDate) + '\n' +
      (note ? 'הערה: ' + note + '\n' : '') + (needsAdmin ? '\n' + NEWPAIR_MSG_ + '\n' : '') + '\nלאישור או דחייה: ' + appUrl_());
    return { ok: true, id: id, needsAdmin: needsAdmin };
  });
}

/* ---------- מסירת משמרת (מישהו לוקח לך משמרת, בלי להחזיר) ---------- */

/** מסירת משמרת לעובד/ת מסוים/ת, שצריך/ה לאשר */
function requestGive(name, email, periodId, date, toName, note) {
  return withLock_(function () {
    var me = verifyEmployee_(name, email);
    var sched = readSchedule_(periodId);
    if (!sched || !sched.published) throw new Error('אין סידור מפורסם');
    var open = false;
    if (toName === '*' || !toName) throw new Error('בחר/י למי למסור את המשמרת');
    var other = open ? null : findEmployee_(toName);
    if (!open && (!other || !other.active)) throw new Error('העובד/ת לא נמצא/ה');
    if (!open && other.name === me.name) throw new Error('אי אפשר למסור לעצמך');
    checkGiveValid_(sched, me.name, date, open ? null : other.name);
    var dup = readRequests_().some(function (q) {
      return q.type === 'give' && q.status === 'pending' && q.from === me.name && q.details.give === date;
    });
    if (dup) throw new Error('כבר יש בקשה פתוחה למסירת המשמרת הזו');
    var id = Utilities.getUuid().slice(0, 8);
    var det = { give: date, period: periodId };
    var needsAdmin = newPairDays_(sched, { type: 'give', from: me.name, to: other.name, details: det }).length > 0;
    if (needsAdmin) det.adminReq = true;
    appendRequest_({ id: id, type: 'give', status: 'pending', from: me.name, to: open ? '*' : other.name, details: det, note: note || '' });
    var body = me.name + ' מבקש/ת שמישהו ייקח את המשמרת שלו/ה ב' + dayName_(date) + ' ' + fmtD_(date) + ' (בלי להחזיר משמרת).\n' +
      (note ? 'הערה: ' + note + '\n' : '') + (needsAdmin ? '\n' + NEWPAIR_MSG_.replace('ההחלפה', 'המסירה') + '\n' : '') + '\nכדי לקחת אותה: ' + appUrl_();
    if (open) {
      var onDay = sched.a[date] || [];
      readEmployees_().forEach(function (e) {
        if (e.active && e.name !== me.name && onDay.indexOf(e.name) < 0) notify_(e, 'משמרת פנויה ב-' + fmtD_(date), body);
      });
    } else notify_(other, 'בקשה לקחת משמרת מ' + me.name, body);
    return { ok: true, id: id, needsAdmin: needsAdmin };
  });
}

function respondSwap(name, email, id, accept) {
  return withLock_(function () {
    var me = verifyEmployee_(name, email);
    var q = findRequest_(id);
    if (!q || (q.type !== 'swap' && q.type !== 'give')) throw new Error('הבקשה לא נמצאה');
    if (q.status !== 'pending') throw new Error('הבקשה כבר טופלה');
    var open = q.type === 'give' && q.to === '*';
    if (open) {
      if (q.from === me.name) throw new Error('זו הבקשה שלך');
      if (!accept) return { ok: true, status: 'ignored' };
      var sc0 = readSchedule_(q.details.period);
      if (sc0 && (sc0.a[q.details.give] || []).indexOf(me.name) >= 0) throw new Error('את/ה כבר משובץ/ת ביום הזה');
      q.to = me.name;
    } else if (q.to !== me.name) throw new Error('הבקשה הזו לא מיועדת אליך');
    var requester = findEmployee_(q.from);
    if (!accept) {
      updateRequest_(id, 'declined');
      notify_(requester, 'הבקשה נדחתה', me.name + ' דחה/תה את הבקשה (' + reqText_(q) + ').');
      return { ok: true, status: 'declined' };
    }
    var sc = readSchedule_(q.details.period);
    if (sc && sc.published) {
      try { validReq_(sc, q); } catch (err) { updateRequest_(id, 'failed'); throw new Error('אי אפשר לבצע: ' + err.message); }
      var bad = newPairDays_(sc, q);
      if (bad.length) {
        updateRequest_(id, 'admin', q.to);
        var what = (q.type === 'give' ? 'מסירה' : 'החלפה');
        adminNotify_('ממתין לאישורך: ' + what + ' עם שני חדשים',
          q.from + ' ו' + q.to + ' הסכימו על ' + what + ' (' + reqText_(q) + ').\nאחרי הביצוע יהיו שני חדשים ביחד ב: ' +
          bad.map(function (d) { return dayName_(d) + ' ' + fmtD_(d); }).join(', ') + '.\nאפשר לאשר או לדחות בלשונית "בקשות" באתר הניהול.');
        notify_(requester, 'הבקשה ממתינה לאישור האחראית', me.name + ' הסכים/ה (' + reqText_(q) + '). ' +
          NEWPAIR_MSG_.replace('ההחלפה', q.type === 'give' ? 'המסירה' : 'ההחלפה') + ' תקבל/י מייל כשהיא תטפל בזה.');
        return { ok: true, status: 'admin' };
      }
    }
    executeSwap_(q, me.name + (q.type === 'give' ? ' לקח/ה את המשמרת' : ' אישר/ה'));
    return { ok: true, status: 'approved' };
  });
}

/** ביצוע החלפה או מסירה מאושרת: בודק שהיא עדיין אפשרית, מעדכן את הסידור ושולח מיילים */
function executeSwap_(q, byText) {
    var id = q.id, requester = findEmployee_(q.from);
    var sched = readSchedule_(q.details.period);
    try {
      if (!sched || !sched.published) throw new Error('הסידור כבר לא מפורסם');
      validReq_(sched, q);
    } catch (err) {
      updateRequest_(id, 'failed');
      throw new Error('אי אפשר לבצע: ' + err.message);
    }
    snapshot_(q.details.period, sched, 'לפני ' + (q.type === 'give' ? 'מסירה' : 'החלפה') + ': ' + q.from + ' / ' + q.to);
    applySwapSlot_(sched, q.details.give, q.from, q.to);
    if (q.type === 'swap') applySwapSlot_(sched, q.details.take, q.to, q.from);
    writeSchedule_(q.details.period, sched.a, sched.o, true, sched.m);
    updateRequest_(id, 'approved', q.to);
    // בקשות אחרות על אותן משמרות כבר לא רלוונטיות
    readRequests_().forEach(function (o) {
      if (o.id === id || (o.type !== 'swap' && o.type !== 'give') || (o.status !== 'pending' && o.status !== 'admin') || o.details.period !== q.details.period) return;
      try { validReq_(sched, o); } catch (e) { updateRequest_(o.id, 'failed'); }
    });
    var body = byText + '. הסידור עודכן:\n' +
      (q.type === 'swap' ? q.from + ': ' + dayName_(q.details.take) + ' ' + fmtD_(q.details.take) + '\n' : '') +
      q.to + ': ' + dayName_(q.details.give) + ' ' + fmtD_(q.details.give) + '\n' + appUrl_();
    notify_(requester, q.type === 'give' ? 'המשמרת נלקחה' : 'ההחלפה אושרה', body);
    notify_(findEmployee_(q.to), q.type === 'give' ? 'לקחת את המשמרת' : 'ההחלפה אושרה', body);
}

/** האחראית מאשרת או דוחה בקשת החלפה/מסירה ממתינה, גם בלי אישור העובד השני */
function adminResolveSwap(pw, id, approve) {
  checkPw_(pw);
  return withLock_(function () {
    var q = findRequest_(id);
    if (!q || (q.type !== 'swap' && q.type !== 'give')) throw new Error('הבקשה לא נמצאה');
    if (q.status !== 'pending' && q.status !== 'admin') throw new Error('הבקשה כבר טופלה');
    if (approve && q.to === '*') throw new Error('המשמרת הוצעה לכולם ועוד אף אחד לא לקח אותה. אפשר לשבץ מישהו ידנית בסידור.');
    if (approve) executeSwap_(q, 'האחראית אישרה את ' + (q.type === 'give' ? 'המסירה' : 'ההחלפה'));
    else {
      updateRequest_(id, 'rejected');
      notify_(findEmployee_(q.from), 'הבקשה לא אושרה', 'האחראית לא אישרה את הבקשה (' + reqText_(q) + ').');
      if (q.status === 'admin') notify_(findEmployee_(q.to), 'הבקשה לא אושרה', 'האחראית לא אישרה את הבקשה (' + q.from + ', ' + reqText_(q) + ').');
    }
    return readRequests_();
  });
}

function cancelRequest(name, email, id) {
  return withLock_(function () {
    var me = verifyEmployee_(name, email);
    var q = findRequest_(id);
    if (!q || q.from !== me.name) throw new Error('הבקשה לא נמצאה');
    if (q.status !== 'pending' && q.status !== 'admin') throw new Error('הבקשה כבר טופלה');
    updateRequest_(id, 'cancelled');
    return { ok: true };
  });
}

/* ---------- חופשות ---------- */

function requestVacation(name, email, start, end, note) {
  return withLock_(function () {
    var me = verifyEmployee_(name, email);
    if (!isIso_(start) || !isIso_(end) || start > end) throw new Error('טווח תאריכים לא תקין');
    if (listDays_(start, end).length > 60) throw new Error('טווח ארוך מדי');
    var id = Utilities.getUuid().slice(0, 8);
    appendRequest_({ id: id, type: 'vacation', status: 'pending', from: me.name, to: '',
      details: { start: start, end: end }, note: note || '' });
    return { ok: true, id: id };
  });
}

/* ================= Admin API ================= */

function adminLoad(pw, periodId) {
  checkPw_(pw);
  var sen = seniorityCheck_();
  var s = readSettings_();
  var pid = periodId || s.periodId;
  var sched = readSchedule_(pid);
  var view = pid === s.periodId ? { periodId: pid, start: s.start, end: s.end, holidays: s.holidays, current: true }
    : periodMeta_(pid, sched);
  return {
    settings: publicSettings_(s),
    view: view,
    periods: listPeriods_(s),
    prev: prevPeriodInfo_(view),
    employees: readEmployees_(),
    responses: readResponses_(pid),
    schedule: sched ? sched.a : {},
    orig: sched ? sched.o : {},
    published: sched ? sched.published : false,
    scheduleUpdated: sched ? sched.updated : null,
    requests: readRequests_(),
    seniority: sen,
    pairHist: pairHistory_(pid, view.start),
    url: appUrl_(),
    adminUrl: adminUrl_(),
    scriptUrl: (function () { try { return ScriptApp.getService().getUrl(); } catch (e) { return ''; } })(),
    sheetUrl: getSS_().getUrl()
  };
}

function adminSaveSettings(pw, ns, newPeriod) {
  checkPw_(pw);
  var s = readSettings_();
  if (!isIso_(ns.start) || !isIso_(ns.end) || ns.start > ns.end) throw new Error('טווח תאריכים לא תקין');
  writeSettings_({
    periodId: newPeriod ? ns.start + '_' + Utilities.formatDate(new Date(), 'Asia/Jerusalem', 'HHmmss') : s.periodId,
    start: ns.start, end: ns.end,
    maxBlocks: String(parseInt(ns.maxBlocks, 10) || 8),
    holidays: (ns.holidays || []).filter(isIso_).sort().join(','),
    open: ns.open ? 'TRUE' : 'FALSE',
    title: ns.title || s.title,
    emailNotify: ns.emailNotify ? 'TRUE' : 'FALSE',
    deadline: isIso_(ns.deadline) ? ns.deadline : '',
    shiftStart: time_(ns.shiftStart), shiftEnd: time_(ns.shiftEnd), wkStart: time_(ns.wkStart), wkEnd: time_(ns.wkEnd),
    reminders: ns.reminders === false ? 'FALSE' : 'TRUE',
    shownBlocks: String(parseInt(ns.shownBlocks, 10) || 8),
    wkBlocks: String(isNaN(parseInt(ns.wkBlocks, 10)) ? 5 : parseInt(ns.wkBlocks, 10)),
    siteUrl: /^https:\/\/\S+$/.test(String(ns.siteUrl || '').trim()) ? String(ns.siteUrl).trim().replace(/\/?$/, '/') : '',
    adminEmail: /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(String(ns.adminEmail || '').trim()) ? String(ns.adminEmail).trim() : ''
  });
  var ns2 = readSettings_();
  savePeriodMeta_(ns2.periodId, ns2.start, ns2.end, ns2.holidays);
  var sc = readSchedule_(s.periodId);
  if (sc && !newPeriod) { // מעדכן את פרטי התקופה ששמורים עם הסידור
    writeSchedule_(s.periodId, sc.a, sc.o, sc.published, { start: ns.start, end: ns.end, holidays: (ns.holidays || []).filter(isIso_).sort() }, true);
  } else if (sc && newPeriod && !(sc.m && sc.m.start)) { // התקופה הקודמת נשמרת עם התאריכים שלה
    writeSchedule_(s.periodId, sc.a, sc.o, sc.published, { start: s.start, end: s.end, holidays: s.holidays }, true);
  }
  return adminLoad(pw);
}

function adminSaveEmployees(pw, list) {
  checkPw_(pw);
  var sh = getSheet_(SH.employees);
  var seen = {};
  var oldPw = {}, oldSen = {};
  vals_(sh).slice(1).forEach(function (r0) { if (r0[0]) { oldPw[norm_(r0[0])] = r0[10] || ''; oldSen[norm_(r0[0])] = { start: r0[8] || '', asked: r0[11] || '' }; } });
  var rows = (list || []).filter(function (e) { return norm_(e.name); }).map(function (e, i) {
    var n = norm_(e.name);
    if (seen[n]) throw new Error('השם "' + n + '" מופיע פעמיים');
    seen[n] = 1;
    return [n, String(e.email || '').trim().toLowerCase(), e.active ? 'כן' : 'לא', str_(e.min), str_(e.max),
      e.color || colorFor_(n, i), e.isNew ? 'חדש' : 'ותיק', str_(e.maxBlocks), isIso_(e.start) ? e.start : '', isIso_(e.ext) ? e.ext : '',
      oldPw[n] || oldPw[norm_(e.origName || '')] || '',
      (function () { var o = oldSen[n] || oldSen[norm_(e.origName || '')]; return o && o.start === (isIso_(e.start) ? e.start : '') ? o.asked : ''; })(),
      e.consec ? 'כן' : ''];
  });
  sh.getRange(1, 1, 1, EMP_HEADERS.length).setValues([EMP_HEADERS]).setFontWeight('bold').setBackground('#eeeeee');
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, EMP_HEADERS.length).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, EMP_HEADERS.length).setNumberFormat('@').setValues(rows);
  return readEmployees_();
}

function adminSaveResponse(pw, periodId, name, values) {
  checkPw_(pw);
  var pid = periodId || readSettings_().periodId;
  var emp = findEmployee_(name);
  if (!emp) throw new Error('עובד לא נמצא');
  saveResponseRow_(pid, emp.name, emp.email, values || {});
  return readResponses_(pid);
}

/** פתיחת הגשה לעובד ספציפי עד תאריך (גם אחרי שההגשה נסגרה). date ריק = ביטול */
function adminSetExtension(pw, name, date) {
  checkPw_(pw);
  var sh = getSheet_(SH.employees);
  var rows = vals_(sh);
  for (var i = 1; i < rows.length; i++) {
    if (norm_(rows[i][0]) === norm_(name)) { sh.getRange(i + 1, 10).setNumberFormat('@').setValue(isIso_(date) ? date : ''); break; }
  }
  return readEmployees_();
}

/** baseVersion: מתי נטען הסידור אצל האחראית. אם בינתיים עובדים ביצעו החלפה, השמירה נעצרת. */
function adminSaveSchedule(pw, periodId, data, orig, published, baseVersion, force) {
  checkPw_(pw);
  return withLock_(function () {
    var s = readSettings_();
    var pid = periodId || s.periodId;
    var cur = readSchedule_(pid);
    if (!force && cur && baseVersion !== undefined && (cur.updated || '') !== (baseVersion || '')) {
      throw new Error('CONFLICT: הסידור עודכן בינתיים (כנראה החלפה בין עובדים). רענני את הדף לפני שמירה.');
    }
    var meta = pid === s.periodId ? { start: s.start, end: s.end, holidays: s.holidays } : (cur ? cur.m : null);
    if (cur && (cur.published || published)) snapshot_(pid, cur, published && !cur.published ? 'לפני פרסום' : 'לפני שמירה של האחראית');
    var updated = writeSchedule_(pid, data || {}, orig || {}, !!published, meta);
    return { ok: true, updated: updated };
  });
}

function adminResolveVacation(pw, id, approve) {
  checkPw_(pw);
  return withLock_(function () {
    var q = findRequest_(id);
    if (!q || q.type !== 'vacation') throw new Error('הבקשה לא נמצאה');
    updateRequest_(id, approve ? 'approved' : 'rejected');
    var emp = findEmployee_(q.from);
    notify_(emp, approve ? 'החופשה אושרה' : 'בקשת החופשה לא אושרה',
      'בקשת החופשה שלך ל-' + fmtD_(q.details.start) + ' עד ' + fmtD_(q.details.end) + (approve ? ' אושרה.' : ' לא אושרה. לפרטים פנה/י לאחראית.'));
    return readRequests_();
  });
}

/** אחראית יכולה לבטל החלפה או חופשה שאושרה בטעות (רק משנה סטטוס; את הסידור מתקנים ידנית) */
function adminSetRequestStatus(pw, id, status) {
  checkPw_(pw);
  updateRequest_(id, status);
  return readRequests_();
}

/** הופך תקופה קיימת לתקופה הנוכחית (ההגדרות, ההגשה והחגים יחולו עליה) */
function adminSetCurrentPeriod(pw, periodId) {
  checkPw_(pw);
  var s = readSettings_();
  savePeriodMeta_(s.periodId, s.start, s.end, s.holidays);
  var m = periodMeta_(periodId);
  if (!m.start) throw new Error('לא נמצאו תאריכים לתקופה הזו');
  var row = settingsToRow_(s);
  row.periodId = periodId; row.start = m.start; row.end = m.end; row.holidays = (m.holidays || []).join(',');
  writeSettings_(row);
  return adminLoad(pw);
}

/** מחיקת תקופה שאינה הנוכחית: האילוצים והסידור שלה */
function adminDeletePeriod(pw, periodId) {
  checkPw_(pw);
  var s = readSettings_();
  if (periodId === s.periodId) throw new Error('אי אפשר למחוק את התקופה הנוכחית. קודם הפכי תקופה אחרת לנוכחית.');
  return withLock_(function () {
    [SH.responses, SH.schedules, SH.periods].forEach(function (name) {
      var sh = getSS_().getSheetByName(name);
      if (!sh) return;
      var rows = vals_(sh);
      for (var i = rows.length - 1; i >= 1; i--) if (rows[i][0] === periodId) sh.deleteRow(i + 1);
    });
    return adminLoad(pw);
  });
}

function adminChangePassword(pw, newPw) {
  checkPw_(pw);
  if (!newPw || String(newPw).length < 4) throw new Error('הסיסמה צריכה להיות לפחות 4 תווים');
  PropertiesService.getScriptProperties().setProperty('ADMIN_PW', hash_(newPw));
  return true;
}

function adminExport(pw, periodId) {
  checkPw_(pw);
  var cs = readSettings_();
  var pid = periodId || cs.periodId;
  var sched = readSchedule_(pid) || { a: {}, o: {} };
  var s = pid === cs.periodId ? cs : (function () { var m = metaOf_(pid, sched); return { periodId: pid, start: m.start, end: m.end, holidays: m.holidays }; })();
  var color = colorMap_();
  var days = listDays_(s.start, s.end);
  var ss = getSS_();
  var label = fmtFull_(s.start) + '-' + fmtFull_(s.end);
  var heads = ['ראשון', 'שני', 'שלישי', 'רביעי', 'חמישי', 'שישי', 'שבת'];

  var name1 = 'סידור ' + label;
  var sh = ss.getSheetByName(name1) || ss.insertSheet(name1);
  sh.clear();
  sh.setRightToLeft(true);
  sh.getRange(1, 1, 1, 7).setValues([heads]).setFontWeight('bold').setFontSize(13).setHorizontalAlignment('center');
  var first = days.length ? dow_(days[0]) : 0;
  var cells = [];
  for (var p = 0; p < first; p++) cells.push(null);
  days.forEach(function (d) { cells.push(d); });
  var r = 2;
  for (var w = 0; w < cells.length; w += 7) {
    var dateRow = [], s1 = [], s2 = [], bg1 = [], bg2 = [], fc1 = [], fc2 = [];
    for (var c = 0; c < 7; c++) {
      var d = cells[w + c];
      var a = d ? (sched.a[d] || []) : [], o = d ? (sched.o[d] || []) : [];
      dateRow.push(d ? d.split('-').reverse().join('/') + (s.holidays.indexOf(d) >= 0 ? ' (חג)' : '') : '');
      s1.push(slotText_(a[0], o[0])); s2.push(slotText_(a[1], o[1]));
      bg1.push(a[0] && color[a[0]] ? color[a[0]] : '#ffffff');
      bg2.push(a[1] && color[a[1]] ? color[a[1]] : '#ffffff');
      fc1.push(textOn_(bg1[c])); fc2.push(textOn_(bg2[c]));
    }
    sh.getRange(r, 1, 1, 7).setValues([dateRow]).setFontSize(13).setHorizontalAlignment('center');
    sh.getRange(r + 1, 1, 1, 7).setValues([s1]).setBackgrounds([bg1]).setFontColors([fc1]).setHorizontalAlignment('center').setFontWeight('bold');
    sh.getRange(r + 2, 1, 1, 7).setValues([s2]).setBackgrounds([bg2]).setFontColors([fc2]).setHorizontalAlignment('center').setFontWeight('bold');
    r += 3;
  }
  sh.setColumnWidths(1, 7, 140);
  sh.getRange(1, 1, r - 1, 7).setBorder(true, true, true, true, true, true, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);

  var name2 = 'אילוצים ' + label;
  var sh2 = ss.getSheetByName(name2) || ss.insertSheet(name2);
  sh2.clear();
  sh2.setRightToLeft(true);
  var resp = readResponses_(s.periodId);
  var vac = vacationMap_();
  var L = ['א', 'ב', 'ג', 'ד', 'ה', 'ו', 'ש'];
  var head = ['שם', 'חסימות'].concat(days.map(function (d) { return L[dow_(d)] + ' ' + d.slice(8) + '/' + d.slice(5, 7); }));
  var out = [head], bgs = [head.map(function () { return '#eeeeee'; })];
  var lab = { B: 'חוסם', C: 'יכול', W: 'רוצה', V: 'חופשה' }, col = { B: '#f4c7c3', C: '#ffffff', W: '#b7e1cd', V: '#c9daf8' };
  readEmployees_().filter(function (e) { return e.active; }).forEach(function (e) {
    var rr = resp[e.name];
    var row = [e.name, rr ? rr.blocks : 'לא הגיש'], bg = ['#ffffff', '#ffffff'];
    days.forEach(function (d) {
      var v = vac[e.name] && vac[e.name][d] ? 'V' : (rr ? (rr.values[d] || 'C') : '');
      row.push(v ? lab[v] : ''); bg.push(v ? col[v] : '#f3f3f3');
    });
    out.push(row); bgs.push(bg);
  });
  sh2.getRange(1, 1, out.length, head.length).setValues(out).setBackgrounds(bgs).setHorizontalAlignment('center');
  sh2.setFrozenRows(1); sh2.setFrozenColumns(2);
  return { ok: true, sheets: [name1, name2], url: ss.getUrl() };
}

/* ================= גרסאות סידור ================= */

function snapshot_(pid, sc, reason) {
  var sh = getSS_().getSheetByName(SH.versions);
  if (!sh || !sc) return;
  var row = [pid, Utilities.formatDate(new Date(), 'Asia/Jerusalem', 'dd/MM/yyyy HH:mm:ss'), reason, JSON.stringify({ a: sc.a, o: sc.o || {} })];
  sh.getRange(sh.getLastRow() + 1, 1, 1, 4).setNumberFormat('@').setValues([row]);
  // שומרים עד 40 גרסאות לכל תקופה
  var rows = vals_(sh), idx = [];
  for (var i = 1; i < rows.length; i++) if (rows[i][0] === pid) idx.push(i + 1);
  for (var j = idx.length - 41; j >= 0; j--) sh.deleteRow(idx[j]);
}
function adminListVersions(pw, periodId) {
  checkPw_(pw);
  var sh = getSS_().getSheetByName(SH.versions);
  if (!sh) return [];
  var rows = vals_(sh), out = [];
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] !== periodId) continue;
    var d = {};
    try { d = JSON.parse(rows[i][3] || '{}'); } catch (e) {}
    out.push({ time: rows[i][1], reason: rows[i][2], a: d.a || {}, o: d.o || {} });
  }
  return out.reverse();
}

/* ================= התראות ================= */

function adminEmail_() {
  var s = readSettings_();
  if (s.adminEmail) return s.adminEmail;
  try { return Session.getEffectiveUser().getEmail(); } catch (e) { return ''; }
}
/** אחרי החלפה: אם נוצר יום עם שני חדשים, רק האחראית מקבלת מייל */
function alertNewPair_(sched, dates, q) {
  try {
    var isNew = {};
    readEmployees_().forEach(function (e) { if (e.isNew) isNew[e.name] = 1; });
    var bad = dates.filter(function (d) { var a = sched.a[d] || []; return a.length > 1 && a[0] && a[1] && isNew[a[0]] && isNew[a[1]]; });
    if (!bad.length) return;
    var to = adminEmail_();
    if (!to) return;
    MailApp.sendEmail(to, 'שימי לב: שני חדשים באותה משמרת',
      'בעקבות ' + (q.type === 'give' ? 'מסירה' : 'החלפה') + ' בין ' + q.from + ' ל' + q.to + ' נוצר מצב של שני עובדים חדשים ביחד:\n' +
      bad.map(function (d) { return dayName_(d) + ' ' + fmtD_(d) + ': ' + sched.a[d].join(', '); }).join('\n') +
      '\n\nהעובדים לא קיבלו על זה הודעה. אפשר לתקן בסידור באתר הניהול.\n' + adminUrl_());
  } catch (e) { Logger.log('newpair alert failed: ' + e); }
}

/* ================= תזכורות אוטומטיות ================= */

function installTriggers_() {
  try {
    ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'dailyJob') ScriptApp.deleteTrigger(t); });
    ScriptApp.newTrigger('dailyJob').timeBased().everyDays(1).atHour(9).inTimezone('Asia/Jerusalem').create();
  } catch (e) { Logger.log('trigger install failed: ' + e); }
}
/** רץ כל בוקר: יום לפני הסגירה וביום הסגירה, מייל למי שעוד לא הגיש */
function dailyJob() {
  seniorityCheck_(true);
  var s = readSettings_();
  if (!s.reminders || !s.open || !s.deadline) return;
  var today = todayIso_();
  var isEve = addDaysIso_(s.deadline, -1) === today, isDay = s.deadline === today;
  if (!isEve && !isDay) return;
  var props = PropertiesService.getScriptProperties(), key = 'REMIND_' + s.periodId + '_' + today;
  if (props.getProperty(key)) return;
  props.setProperty(key, '1');
  var resp = readResponses_(s.periodId);
  var p = s.deadline.split('-'), due = dayName_(s.deadline) + ' ' + (+p[2]) + '/' + (+p[1]);
  readEmployees_().forEach(function (e) {
    if (!e.active || !e.email || resp[e.name]) return;
    if (e.start && e.start > s.end) return;
    try {
      MailApp.sendEmail(e.email, isDay ? 'היום נסגרת הגשת האילוצים' : 'תזכורת: מחר נסגרת הגשת האילוצים',
        'היי ' + e.name + ',\n\n' + (isDay ? 'היום (' + due + ') ' : 'מחר (' + due + ') ') +
        'נסגרת הגשת האילוצים לתקופה ' + fmtD_(s.start) + ' עד ' + fmtD_(s.end) + ', ועוד לא הגשת.\n\nלהגשה: ' + appUrl_() +
        '\n\nמערכת הסידור - מיון ילדים');
    } catch (err) { Logger.log('reminder failed ' + e.name + ': ' + err); }
  });
}

/* ================= Helpers: periods ================= */

function todayIso_() { return Utilities.formatDate(new Date(), 'Asia/Jerusalem', 'yyyy-MM-dd'); }
/** האם עובד (או סתם מבקר כש-emp ריק) יכול להגיש כרגע */
function submitOpenFor_(s, emp) {
  var today = todayIso_();
  if (emp && emp.ext && today <= emp.ext) return true;
  if (!s.open) return false;
  if (s.deadline && today > s.deadline) return false;
  return true;
}
function metaOf_(pid, sc) {
  var reg = readPeriodRegistry_()[pid];
  if (reg && reg.start) return reg;
  if (sc && sc.m && sc.m.start) return { periodId: pid, start: sc.m.start, end: sc.m.end, holidays: sc.m.holidays || [] };
  var keys = sc ? Object.keys(sc.a || {}).filter(isIso_).sort() : [];
  return { periodId: pid, start: keys[0] || '', end: keys[keys.length - 1] || '', holidays: [] };
}
function schedPublic_(pid, sc) {
  var m = metaOf_(pid, sc);
  return { periodId: pid, start: m.start, end: m.end, holidays: m.holidays, a: sc.a, o: sc.o };
}
/** הסידור שמוצג כברירת מחדל: זה שהיום בתוכו; אחרת הבא בתור; אחרת האחרון שהסתיים */
function pickByDate_(pubs, today) {
  var cur = pubs.filter(function (p) { return p.start <= today && today <= p.end; })[0];
  if (cur) return cur;
  var next = pubs.filter(function (p) { return p.start > today; }).sort(function (a, b) { return a.start < b.start ? -1 : 1; })[0];
  return next || pubs[0] || null;
}

/** כל הסידורים שפורסמו, מהחדש לישן */
function listPublished_() {
  var rows = vals_(getSheet_(SH.schedules));
  var out = [];
  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0] || rows[i][2] !== 'TRUE') continue;
    var sc = readSchedule_(rows[i][0]);
    var m = metaOf_(rows[i][0], sc);
    if (m.start) out.push({ periodId: rows[i][0], start: m.start, end: m.end });
  }
  out.sort(function (x, y) { return x.start < y.start ? 1 : -1; });
  return out;
}
/** התקופה שלפני התקופה המוצגת: כמה סופ"שים כל אחד עשה, ומי עבד בימים האחרונים שלה */
function prevPeriodInfo_(view) {
  if (!view || !view.start) return null;
  var cands = listPeriods_(readSettings_()).filter(function (p) { return p.hasSchedule && p.end && p.end < view.start; });
  if (!cands.length) return null;
  cands.sort(function (x, y) { return x.end < y.end ? 1 : -1; });
  var p = cands[0], sc = readSchedule_(p.periodId);
  if (!sc) return null;
  // סופ"שים: סכום של שתי התקופות האחרונות (בערך חודשיים אחורה)
  var weekend = {}, total = {}, used = [];
  cands.slice(0, 2).forEach(function (pp) {
    var sc2 = readSchedule_(pp.periodId);
    if (!sc2) return;
    used.push({ start: pp.start, end: pp.end });
    var hol2 = periodMeta_(pp.periodId, sc2).holidays || [];
    listDays_(pp.start, pp.end).forEach(function (d) {
      var wk = dow_(d) === 5 || dow_(d) === 6 || hol2.indexOf(d) >= 0;
      (sc2.a[d] || []).forEach(function (n) { if (!n) return; total[n] = (total[n] || 0) + 1; if (wk) weekend[n] = (weekend[n] || 0) + 1; });
    });
  });
  var tail = {};
  [2, 1].forEach(function (k) { var d = addDaysIso_(view.start, -k); tail[d] = (sc.a[d] || []).filter(Boolean); });
  return { periodId: p.periodId, start: p.start, end: p.end, weekend: weekend, total: total, tail: tail,
    holidays: periodMeta_(p.periodId, sc).holidays || [], periods: used };
}

/** תאריכי תקופה: מהסידור אם יש, אחרת מהאילוצים שהוגשו בה */
function periodMeta_(pid, sc) {
  var reg = readPeriodRegistry_()[pid];
  if (reg && reg.start) return reg;
  var m = metaOf_(pid, sc || readSchedule_(pid));
  if (m.start) return m;
  var keys = {};
  var resp = readResponses_(pid);
  Object.keys(resp).forEach(function (n) { Object.keys(resp[n].values || {}).forEach(function (d) { if (isIso_(d)) keys[d] = 1; }); });
  var k = Object.keys(keys).sort();
  return { periodId: pid, start: k[0] || '', end: k[k.length - 1] || '', holidays: [] };
}
/** כל התקופות שיש להן סידור או אילוצים (לאחראית) + התקופה הנוכחית */
function listPeriods_(s) {
  var out = [{ periodId: s.periodId, start: s.start, end: s.end, published: false, current: true, hasSchedule: false }];
  var seen = {}; seen[s.periodId] = out[0];
  var reg = readPeriodRegistry_();
  Object.keys(reg).forEach(function (pid) {
    if (seen[pid] || !reg[pid].start) return;
    seen[pid] = { periodId: pid, start: reg[pid].start, end: reg[pid].end, published: false, current: false, hasSchedule: false };
    out.push(seen[pid]);
  });
  var rows = vals_(getSheet_(SH.schedules));
  for (var i = 1; i < rows.length; i++) {
    var pid = rows[i][0];
    if (!pid) continue;
    if (seen[pid]) { seen[pid].published = rows[i][2] === 'TRUE'; seen[pid].hasSchedule = true; continue; }
    var m = periodMeta_(pid);
    if (m.start) { seen[pid] = { periodId: pid, start: m.start, end: m.end, published: rows[i][2] === 'TRUE', current: false, hasSchedule: true }; out.push(seen[pid]); }
  }
  var rr = vals_(getSheet_(SH.responses));
  for (var j = 1; j < rr.length; j++) {
    var p2 = rr[j][0];
    if (!p2 || seen[p2]) continue;
    var m2 = periodMeta_(p2);
    seen[p2] = { periodId: p2, start: m2.start, end: m2.end, published: false, current: false, hasSchedule: false };
    if (m2.start) out.push(seen[p2]);
  }
  out.sort(function (x, y) { return x.start < y.start ? 1 : -1; });
  return out;
}

function readPeriodRegistry_() {
  var sh = getSS_().getSheetByName(SH.periods);
  var out = {};
  if (!sh) return out;
  var rows = vals_(sh);
  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    out[rows[i][0]] = { periodId: rows[i][0], start: rows[i][1], end: rows[i][2],
      holidays: rows[i][3] ? rows[i][3].split(',').filter(isIso_) : [] };
  }
  return out;
}
function savePeriodMeta_(pid, start, end, holidays) {
  var sh = getSS_().getSheetByName(SH.periods);
  if (!sh || !pid || !isIso_(start)) return;
  var row = [pid, start, end, (holidays || []).join(',')];
  var rows = vals_(sh);
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === pid) { sh.getRange(i + 1, 1, 1, 4).setNumberFormat('@').setValues([row]); return; }
  }
  sh.getRange(sh.getLastRow() + 1, 1, 1, 4).setNumberFormat('@').setValues([row]);
}

/* ================= Helpers: calendar ================= */

function time_(t) { return /^\d{1,2}:\d{2}$/.test(String(t || '')) ? String(t) : ''; }
function addDaysIso_(iso, n) { var p = iso.split('-'); var d = new Date(Date.UTC(+p[0], +p[1] - 1, +p[2])); d.setUTCDate(d.getUTCDate() + n); return d.toISOString().slice(0, 10); }

/* ================= Helpers: swaps ================= */

function checkSwapValid_(sched, a, giveDate, b, takeDate) {
  if (!isIso_(giveDate) || !isIso_(takeDate)) throw new Error('תאריך לא תקין');
  if (giveDate === takeDate) throw new Error('שתי המשמרות באותו יום');
  var g = sched.a[giveDate] || [], t = sched.a[takeDate] || [];
  if (g.indexOf(a) < 0) throw new Error(a + ' כבר לא משובץ/ת ב-' + fmtD_(giveDate));
  if (t.indexOf(b) < 0) throw new Error(b + ' כבר לא משובץ/ת ב-' + fmtD_(takeDate));
  if (g.indexOf(b) >= 0) throw new Error(b + ' כבר משובץ/ת ב-' + fmtD_(giveDate));
  if (t.indexOf(a) >= 0) throw new Error(a + ' כבר משובץ/ת ב-' + fmtD_(takeDate));
}
/* ================= ותק: אחרי 4 ו-5 חודשים שואלים, אחרי 6 הופך לותיק ================= */

/** כמה חודשים מלאים עברו מתאריך (yyyy-mm-dd) עד היום */
function monthsSince_(iso, today) {
  var a = iso.split('-').map(Number), b = today.split('-').map(Number);
  var m = (b[0] - a[0]) * 12 + (b[1] - a[1]);
  if (b[2] < a[2]) m--;
  return m;
}
function senStage_(months) { return months >= 6 ? 6 : months >= 5 ? 5 : months >= 4 ? 4 : 0; }
/** מחזיר למי צריך לשאול, ומעביר לותיק אוטומטית את מי שעבר 6 חודשים */
function seniorityCheck_(mailAdmin) {
  var out = { ask: [], auto: [] };
  try {
    var sh = getSheet_(SH.employees), rows = vals_(sh), today = todayIso_(), wrote = false;
    for (var i = 1; i < rows.length; i++) {
      var n = norm_(rows[i][0]);
      if (!n || rows[i][2] === 'לא' || rows[i][6] !== 'חדש' || !isIso_(rows[i][8])) continue;
      var months = monthsSince_(rows[i][8], today), stage = senStage_(months);
      if (stage === 6) {
        sh.getRange(i + 1, 7).setValue('ותיק'); wrote = true;
        out.auto.push({ name: n, start: rows[i][8] });
      } else if (stage && (parseInt(rows[i][11], 10) || 0) < stage) {
        out.ask.push({ name: n, start: rows[i][8], months: months, stage: stage });
      }
    }
    if (wrote && MEMO_) delete MEMO_[sh.getName()];
    if (mailAdmin && out.auto.length) adminNotify_('עדכון ותק', out.auto.map(function (x) { return x.name + ' עבר/ה 6 חודשים מתחילת העבודה (' + fmtD_(x.start) + ') ומעכשיו נחשב/ת ותיק/ה.'; }).join('\n'));
  } catch (e) { Logger.log('seniority check failed: ' + e); }
  return out;
}
/** תשובת האחראית: promote=true הופך לותיק, false שואל שוב בעוד חודש */
function adminSeniority(pw, name, promote) {
  checkPw_(pw);
  return withLock_(function () {
    var sh = getSheet_(SH.employees), rows = vals_(sh), n = norm_(name);
    for (var i = 1; i < rows.length; i++) {
      if (norm_(rows[i][0]) !== n) continue;
      if (promote) sh.getRange(i + 1, 7).setValue('ותיק');
      else {
        var st = isIso_(rows[i][8]) ? senStage_(monthsSince_(rows[i][8], todayIso_())) : 0;
        sh.getRange(i + 1, 12).setNumberFormat('@').setValue(String(Math.max(st, 4)));
      }
      return { employees: readEmployees_(), seniority: seniorityCheck_() };
    }
    throw new Error('העובד/ת לא נמצא/ה');
  });
}

/** לכל עובד חדש: עם מי עבד ובכמה משמרות, בסידורים שפורסמו לפני התקופה המוצגת (מתאריך תחילת העבודה) */
function pairHistory_(pid, before) {
  var out = {};
  try {
    var news = readEmployees_().filter(function (e) { return e.isNew; });
    if (!news.length || !before) return out;
    news.forEach(function (e) { out[e.name] = {}; });
    var rows = vals_(getSheet_(SH.schedules));
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] === pid || rows[i][2] !== 'TRUE') continue;
      var d = {};
      try { d = JSON.parse(rows[i][1] || '{}'); } catch (err) { continue; }
      var a = d && d.a ? d.a : d;
      Object.keys(a || {}).forEach(function (date) {
        if (!isIso_(date) || date >= before) return;
        var arr = (a[date] || []).filter(Boolean);
        news.forEach(function (e) {
          if (arr.indexOf(e.name) < 0 || (e.start && date < e.start)) return;
          arr.forEach(function (o) { if (o !== e.name) out[e.name][o] = (out[e.name][o] || 0) + 1; });
        });
      });
    }
  } catch (e) { Logger.log('pair history failed: ' + e); }
  return out;
}

/** הימים שבהם יהיו שני חדשים ביחד אם הבקשה תבוצע */
function newPairDays_(sched, q) {
  var isNew = {};
  readEmployees_().forEach(function (e) { if (e.isNew) isNew[e.name] = 1; });
  var dates = [q.details.give].concat(q.type === 'swap' ? [q.details.take] : []);
  var copy = { a: {}, o: {} };
  dates.forEach(function (d) { copy.a[d] = (sched.a[d] || []).slice(); copy.o[d] = (sched.o[d] || []).slice(); });
  applySwapSlot_(copy, q.details.give, q.from, q.to);
  if (q.type === 'swap') applySwapSlot_(copy, q.details.take, q.to, q.from);
  return dates.filter(function (d) { var a = copy.a[d]; return a.length > 1 && a[0] && a[1] && isNew[a[0]] && isNew[a[1]]; });
}
function adminNotify_(subject, body) {
  try { var to = adminEmail_(); if (to) MailApp.sendEmail(to, subject, body + '\n\n' + adminUrl_()); } catch (e) { Logger.log('admin mail failed: ' + e); }
}
var NEWPAIR_MSG_ = 'ההחלפה הזו תשבץ שני עובדים חדשים באותה משמרת, ולכן היא עוברת לאישור האחראית.';

/** ותיק בשורה הראשונה, חדש בשנייה */
function orderDay_(sched, date) {
  var a = sched.a[date] || [], o = sched.o[date] || [];
  if (a.length < 2 || !a[0] || !a[1]) return;
  var isNew = {};
  readEmployees_().forEach(function (e) { if (e.isNew) isNew[e.name] = 1; });
  if (isNew[a[0]] && !isNew[a[1]]) {
    sched.a[date] = [a[1], a[0]].concat(a.slice(2));
    sched.o[date] = [o[1] || null, o[0] || null];
  }
}
function checkGiveValid_(sched, a, date, b) {
  if (!isIso_(date)) throw new Error('תאריך לא תקין');
  var g = sched.a[date] || [];
  if (g.indexOf(a) < 0) throw new Error(a + ' כבר לא משובץ/ת ב-' + fmtD_(date));
  if (b && g.indexOf(b) >= 0) throw new Error(b + ' כבר משובץ/ת ב-' + fmtD_(date));
}
function validReq_(sched, q) {
  if (q.type === 'give') checkGiveValid_(sched, q.from, q.details.give, q.to === '*' ? null : q.to);
  else checkSwapValid_(sched, q.from, q.details.give, q.to, q.details.take);
}
function reqText_(q) {
  return q.type === 'give' ? 'מסירת ' + fmtD_(q.details.give) : fmtD_(q.details.give) + ' מול ' + fmtD_(q.details.take);
}
function applySwapSlot_(sched, date, oldName, newName) {
  var a = sched.a[date] = (sched.a[date] || []).slice();
  var o = sched.o[date] = (sched.o[date] || []).slice();
  var k = a.indexOf(oldName);
  if (k < 0) return;
  var original = o[k] || oldName;
  a[k] = newName;
  o[k] = original === newName ? null : original; // חזרה לבעלים המקורי מנקה את הסימון
}
function slotText_(name, orig) { return name ? (orig ? name + ' (במקום ' + orig + ')' : name) : ''; }

/* ================= Helpers: data ================= */

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}
/* מהירות: שומרים בזיכרון את הגיליון והלשוניות בתוך אותה בקשה */
var SS_ = null, SHEETS_ = {}, MEMO_ = null;
function getSS_() {
  if (SS_) return SS_;
  var id = PropertiesService.getScriptProperties().getProperty('SSID');
  SS_ = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
  return SS_;
}
function getSheet_(name) {
  var sh = SHEETS_[name] || getSS_().getSheetByName(name);
  if (!sh) throw new Error('חסרה לשונית "' + name + '". הריצי את setup.');
  SHEETS_[name] = sh;
  return sh;
}
/** קריאת כל הנתונים של לשונית. בבקשות קריאה בלבד, כל לשונית נקראת פעם אחת */
function vals_(sh) {
  if (!MEMO_) return sh.getDataRange().getDisplayValues();
  var k = sh.getName();
  if (!MEMO_[k]) MEMO_[k] = sh.getDataRange().getDisplayValues();
  return MEMO_[k];
}
function ensureSheet_(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.getRange(1, 1, sh.getMaxRows(), Math.max(headers.length, 2)).setNumberFormat('@');
  if (sh.getLastRow() < 1) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#eeeeee');
    sh.setFrozenRows(1);
  }
  sh.setRightToLeft(true);
  return sh;
}
function readSettings_() {
  var rows = vals_(getSheet_(SH.settings));
  var m = {};
  for (var i = 1; i < rows.length; i++) if (rows[i][0]) m[rows[i][0]] = rows[i][1];
  return {
    periodId: m.periodId || m.start || 'p1',
    start: m.start, end: m.end,
    maxBlocks: parseInt(m.maxBlocks, 10) || 8,
    holidays: m.holidays ? m.holidays.split(',').map(function (x) { return x.trim(); }).filter(String) : [],
    open: String(m.open).toUpperCase() === 'TRUE',
    title: m.title || 'הגשת אילוצים',
    emailNotify: m.emailNotify === undefined ? true : String(m.emailNotify).toUpperCase() === 'TRUE',
    deadline: isIso_(m.deadline) ? m.deadline : '',
    shiftStart: m.shiftStart || '', shiftEnd: m.shiftEnd || '', wkStart: m.wkStart || '', wkEnd: m.wkEnd || '',
    reminders: m.reminders === undefined || m.reminders === '' ? true : String(m.reminders).toUpperCase() === 'TRUE',
    adminEmail: m.adminEmail || '',
    shownBlocks: parseInt(m.shownBlocks, 10) || 8,
    siteUrl: m.siteUrl || '',
    wkBlocks: m.wkBlocks === undefined || m.wkBlocks === '' ? 5 : (parseInt(m.wkBlocks, 10) || 0)
  };
}
function settingsToRow_(s) {
  return { periodId: s.periodId, start: s.start, end: s.end, maxBlocks: String(s.maxBlocks), holidays: s.holidays.join(','),
    open: s.open ? 'TRUE' : 'FALSE', title: s.title, emailNotify: s.emailNotify ? 'TRUE' : 'FALSE', deadline: s.deadline || '',
    shiftStart: s.shiftStart || '', shiftEnd: s.shiftEnd || '', wkStart: s.wkStart || '', wkEnd: s.wkEnd || '',
    reminders: s.reminders ? 'TRUE' : 'FALSE', adminEmail: s.adminEmail || '', shownBlocks: String(s.shownBlocks || 8), siteUrl: s.siteUrl || '', wkBlocks: String(s.wkBlocks == null ? 5 : s.wkBlocks) };
}
function writeSettings_(o) {
  var sh = getSheet_(SH.settings);
  var rows = SETTING_KEYS.map(function (k) { return [k, o[k] == null ? '' : String(o[k])]; });
  sh.getRange(2, 1, rows.length, 2).setNumberFormat('@').setValues(rows);
}
function publicSettings_(s) {
  return { periodId: s.periodId, start: s.start, end: s.end, maxBlocks: s.maxBlocks, holidays: s.holidays, wkBlocks: s.wkBlocks,
    open: s.open, title: s.title, emailNotify: s.emailNotify, deadline: s.deadline, today: todayIso_(),
    shiftStart: s.shiftStart, shiftEnd: s.shiftEnd, wkStart: s.wkStart, wkEnd: s.wkEnd,
    reminders: s.reminders, adminEmail: s.adminEmail, shownBlocks: s.shownBlocks, siteUrl: s.siteUrl };
}
function readEmployees_() {
  var rows = vals_(getSheet_(SH.employees));
  var out = [];
  for (var i = 1; i < rows.length; i++) {
    var n = norm_(rows[i][0]);
    if (!n) continue;
    out.push({
      name: n,
      email: String(rows[i][1] || '').trim().toLowerCase(),
      active: rows[i][2] !== 'לא',
      min: rows[i][3] || '', max: rows[i][4] || '',
      color: rows[i][5] || colorFor_(n, i - 1),
      isNew: rows[i][6] === 'חדש',
      maxBlocks: String(rows[i][7] || '').trim(),
      start: isIso_(rows[i][8]) ? rows[i][8] : '',
      ext: isIso_(rows[i][9]) ? rows[i][9] : '',
      hasPw: !!rows[i][10],
      senAsked: parseInt(rows[i][11], 10) || 0,
      consec: rows[i][12] === 'כן'
    });
  }
  return out;
}
function colorMap_() {
  var m = {};
  readEmployees_().forEach(function (e) { m[e.name] = e.color; });
  return m;
}
function findEmployee_(name) {
  var n = norm_(name);
  var list = readEmployees_();
  for (var i = 0; i < list.length; i++) if (list[i].name === n) return list[i];
  return null;
}
/** בודק את אסימון הכניסה (token) שהעובד קיבל אחרי סיסמה נכונה */
function verifyEmployee_(name, token) {
  var emp = findEmployee_(name);
  if (!emp || !emp.active) throw new Error('השם לא נמצא ברשימת העובדים');
  var t = String(token || ''), dot = t.lastIndexOf('.');
  if (dot < 0) throw new Error('AUTH: נא להתחבר מחדש');
  var payload = decodeURIComponent(t.slice(0, dot)), sig = t.slice(dot + 1);
  var parts = payload.split('|');
  if (sig !== sign_(payload) || parts[0] !== emp.name || +parts[1] < Date.now() || parts[2] !== pwTag_(emp.name)) {
    throw new Error('AUTH: נא להתחבר מחדש');
  }
  return emp;
}

/* ================= כניסה עם סיסמה ================= */

var SESSION_DAYS = 60;
function secret_() {
  var p = PropertiesService.getScriptProperties(), k = p.getProperty('TOKEN_SECRET');
  if (!k) { k = Utilities.getUuid() + Utilities.getUuid(); p.setProperty('TOKEN_SECRET', k); }
  return k;
}
function sign_(payload) { return hash_(secret_() + '|' + payload).slice(0, 40); }
function pwRow_(name) {
  var sh = getSheet_(SH.employees), rows = vals_(sh), n = norm_(name);
  for (var i = 1; i < rows.length; i++) if (norm_(rows[i][0]) === n) return { sh: sh, row: i + 1, val: rows[i][10] || '' };
  return null;
}
/** חלק מהסיסמה המוצפנת: אם הסיסמה משתנה, כל החיבורים הישנים מתבטלים */
function pwTag_(name) { var r0 = pwRow_(name); return r0 && r0.val ? r0.val.slice(-8) : 'none'; }
function makeToken_(name) {
  var payload = name + '|' + (Date.now() + SESSION_DAYS * 864e5) + '|' + pwTag_(name);
  return encodeURIComponent(payload) + '.' + sign_(payload);
}
function hashPw_(name, pw, salt) { return salt + '$' + hash_('pw|' + salt + '|' + name + '|' + pw); }
function checkPassword_(name, pw) {
  var r0 = pwRow_(name);
  if (!r0 || !r0.val) return false;
  var salt = r0.val.split('$')[0];
  return hashPw_(name, String(pw), salt) === r0.val;
}
function validPw_(pw) {
  pw = String(pw || '');
  if (pw.length < 4) throw new Error('הסיסמה צריכה להיות לפחות 4 תווים (למשל 4 ספרות)');
  if (pw.length > 50) throw new Error('הסיסמה ארוכה מדי');
  return pw;
}
function tooMany_(key, max, seconds) {
  var c = CacheService.getScriptCache(), n = parseInt(c.get(key) || '0', 10) + 1;
  c.put(key, String(n), seconds);
  return n > max;
}
function maskEmail_(em) {
  if (!em) return '';
  var p = em.split('@');
  return p[0].slice(0, 2) + '***@' + (p[1] || '');
}

/** מה מציגים בכניסה: יש סיסמה? יש מייל לשליחת קוד? */
function getLoginState(name) {
  var emp = findEmployee_(name);
  if (!emp || !emp.active) throw new Error('השם לא נמצא ברשימת העובדים');
  return { hasPassword: emp.hasPw, hasEmail: !!emp.email, emailHint: maskEmail_(emp.email) };
}

function login(name, pw) {
  var emp = findEmployee_(name);
  if (!emp || !emp.active) throw new Error('השם לא נמצא ברשימת העובדים');
  if (!emp.hasPw) throw new Error('עוד אין לך סיסמה. לחצו "כניסה ראשונה / שכחתי סיסמה".');
  var cache = CacheService.getScriptCache(), key = 'fail_' + hash_(emp.name).slice(0, 16);
  if (parseInt(cache.get(key) || '0', 10) >= 5) throw new Error('יותר מדי ניסיונות. נסו שוב בעוד 15 דקות, או אפסו סיסמה דרך המייל.');
  if (!checkPassword_(emp.name, pw)) {
    tooMany_(key, 5, 900);
    throw new Error('סיסמה שגויה');
  }
  cache.remove(key);
  return { token: makeToken_(emp.name), name: emp.name };
}

/** שולח קוד אימות למייל הרשום, לכניסה ראשונה או לאיפוס סיסמה */
function sendSetupCode(name, newEmail) {
  var emp = findEmployee_(name);
  if (!emp || !emp.active) throw new Error('השם לא נמצא ברשימת העובדים');
  var to = emp.email;
  var h = hash_(emp.name).slice(0, 16);
  if (!to) {
    // עובד בלי מייל ברשימה: המייל שהוא מזין יישמר אחרי שיאמת את הקוד
    to = String(newEmail || '').trim().toLowerCase();
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(to)) throw new Error('הזינו כתובת מייל תקינה');
    var taken = readEmployees_().some(function (e) { return e.name !== emp.name && e.email === to; });
    if (taken) throw new Error('המייל הזה כבר רשום לעובד/ת אחר/ת');
  }
  if (tooMany_('codesend_' + h, 3, 600)) throw new Error('נשלחו כבר כמה קודים. נסו שוב בעוד 10 דקות.');
  var code = String(Math.floor(100000 + Math.random() * 900000));
  CacheService.getScriptCache().put('code_' + h, code, 900);
  CacheService.getScriptCache().put('codemail_' + h, to, 900);
  CacheService.getScriptCache().remove('codetry_' + h);
  MailApp.sendEmail(to, 'קוד אימות: ' + code, 'קוד האימות שלך למערכת הסידור הוא: ' + code + '\n\nהקוד בתוקף 15 דקות. אם לא ביקשת קוד, אפשר להתעלם מהמייל.\n\nמערכת הסידור - מיון ילדים');
  return { ok: true, emailHint: maskEmail_(to) };
}

/** קוד מהמייל + סיסמה חדשה = כניסה */
function setPasswordWithCode(name, code, newPw) {
  var emp = findEmployee_(name);
  if (!emp || !emp.active) throw new Error('השם לא נמצא ברשימת העובדים');
  var h = hash_(emp.name).slice(0, 16), cache = CacheService.getScriptCache();
  var real = cache.get('code_' + h);
  if (!real) throw new Error('הקוד פג תוקף. בקשו קוד חדש.');
  if (tooMany_('codetry_' + h, 5, 900)) { cache.remove('code_' + h); throw new Error('יותר מדי ניסיונות. בקשו קוד חדש.'); }
  if (String(code).trim() !== real) throw new Error('הקוד שגוי');
  newPw = validPw_(newPw);
  var r0 = pwRow_(emp.name);
  if (!emp.email) { // שומרים את המייל שאומת
    var em = cache.get('codemail_' + h);
    if (em) r0.sh.getRange(r0.row, 2).setValue(em);
  }
  r0.sh.getRange(r0.row, 11).setNumberFormat('@').setValue(hashPw_(emp.name, newPw, Utilities.getUuid().slice(0, 8)));
  cache.remove('code_' + h);
  cache.remove('fail_' + h);
  return { token: makeToken_(emp.name), name: emp.name };
}

/** האחראית מאפסת סיסמה (העובד יגדיר חדשה דרך קוד במייל) */
function adminResetPassword(pw, name) {
  checkPw_(pw);
  var r0 = pwRow_(name);
  if (r0) r0.sh.getRange(r0.row, 11).setValue('');
  return readEmployees_();
}
function readResponses_(periodId) {
  var rows = vals_(getSheet_(SH.responses));
  var out = {};
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] !== periodId) continue;
    var v = {};
    try { v = JSON.parse(rows[i][5] || '{}'); } catch (e) {}
    out[norm_(rows[i][1])] = { values: v, updated: rows[i][3], blocks: parseInt(rows[i][4], 10) || 0, note: rows[i][6] || '' };
  }
  return out;
}
/** note: undefined = להשאיר את ההערה הקיימת */
function saveResponseRow_(periodId, name, email, values, note) {
  var sh = getSheet_(SH.responses);
  var rows = vals_(sh);
  var blocks = 0;
  Object.keys(values).forEach(function (k) { if (values[k] === 'B') blocks++; });
  var now = nowStr_();
  var row = [periodId, name, email || '', now, String(blocks), JSON.stringify(values), note === undefined ? '' : String(note || '').slice(0, 1000)];
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === periodId && norm_(rows[i][1]) === name) {
      if (note === undefined) row[6] = rows[i][6] || '';
      sh.getRange(i + 1, 1, 1, row.length).setNumberFormat('@').setValues([row]);
      return now;
    }
  }
  sh.getRange(sh.getLastRow() + 1, 1, 1, row.length).setNumberFormat('@').setValues([row]);
  return now;
}
function readSchedule_(periodId) {
  var rows = vals_(getSheet_(SH.schedules));
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === periodId) {
      var d = {};
      try { d = JSON.parse(rows[i][1] || '{}'); } catch (e) {}
      var a = d && d.a ? d.a : d, o = d && d.a ? (d.o || {}) : {};
      return { a: a, o: o, m: (d && d.m) || null, published: rows[i][2] === 'TRUE', updated: rows[i][3] };
    }
  }
  return null;
}
function writeSchedule_(periodId, a, o, published, meta, keepTime) {
  var sh = getSheet_(SH.schedules);
  var rows = vals_(sh);
  var now = Utilities.formatDate(new Date(), 'Asia/Jerusalem', 'dd/MM/yyyy HH:mm:ss');
  var row = [periodId, JSON.stringify({ a: a, o: o, m: meta || null }), published ? 'TRUE' : 'FALSE', now];
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === periodId) {
      if (keepTime) row[3] = rows[i][3];
      sh.getRange(i + 1, 1, 1, 4).setNumberFormat('@').setValues([row]); return row[3];
    }
  }
  sh.getRange(sh.getLastRow() + 1, 1, 1, 4).setNumberFormat('@').setValues([row]);
  return now;
}
function readRequests_() {
  var sh = getSS_().getSheetByName(SH.requests);
  if (!sh) return [];
  var rows = vals_(sh);
  var out = [];
  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    var det = {};
    try { det = JSON.parse(rows[i][5] || '{}'); } catch (e) {}
    out.push({ id: rows[i][0], type: rows[i][1], status: rows[i][2], from: rows[i][3], to: rows[i][4],
      details: det, note: rows[i][6], created: rows[i][7], resolved: rows[i][8] });
  }
  return out;
}
function findRequest_(id) {
  var all = readRequests_();
  for (var i = 0; i < all.length; i++) if (all[i].id === id) return all[i];
  return null;
}
function appendRequest_(q) {
  var sh = getSheet_(SH.requests);
  var row = [q.id, q.type, q.status, q.from, q.to || '', JSON.stringify(q.details), q.note || '', nowStr_(), ''];
  sh.getRange(sh.getLastRow() + 1, 1, 1, row.length).setNumberFormat('@').setValues([row]);
}
function updateRequest_(id, status, to) {
  var sh = getSheet_(SH.requests);
  var rows = vals_(sh);
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === id) {
      sh.getRange(i + 1, 3).setValue(status); sh.getRange(i + 1, 9).setValue(nowStr_());
      if (to) sh.getRange(i + 1, 5).setValue(to);
      return;
    }
  }
}
/** {name: {date: true}} לחופשות שאושרו */
function vacationMap_() {
  var m = {};
  readRequests_().forEach(function (q) {
    if (q.type !== 'vacation' || q.status !== 'approved') return;
    m[q.from] = m[q.from] || {};
    listDays_(q.details.start, q.details.end).forEach(function (d) { m[q.from][d] = true; });
  });
  return m;
}
function notify_(emp, subject, body) {
  try {
    if (!emp || !emp.email || !readSettings_().emailNotify) return;
    MailApp.sendEmail(emp.email, subject, body + '\n\nמערכת הסידור - מיון ילדים');
  } catch (e) { Logger.log('mail failed: ' + e); }
}
/** הקישור לאתר העובדים: אתר GitHub אם הוגדר, אחרת הקישור של Apps Script */
function appUrl_() {
  try { var s = readSettings_(); if (s.siteUrl) return s.siteUrl; } catch (e) {}
  try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; }
}
function adminUrl_() {
  try { var s = readSettings_(); if (s.siteUrl) return s.siteUrl + 'admin.html'; } catch (e) {}
  return appUrl_() + '?page=admin';
}
function checkPw_(pw) {
  var stored = PropertiesService.getScriptProperties().getProperty('ADMIN_PW');
  if (!stored) throw new Error('המערכת לא הותקנה. הריצי setup.');
  var cache = CacheService.getScriptCache();
  if (parseInt(cache.get('adminfail') || '0', 10) >= 10) throw new Error('יותר מדי ניסיונות כניסה. נסי שוב בעוד 15 דקות.');
  if (hash_(pw || '') !== stored) { tooMany_('adminfail', 10, 900); throw new Error('סיסמה שגויה'); }
}
function hash_(s) {
  var b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'shift-salt:' + s, Utilities.Charset.UTF_8);
  return b.map(function (x) { return ('0' + (x & 0xff).toString(16)).slice(-2); }).join('');
}
function textOn_(hex) {
  var h = String(hex || '#ffffff').replace('#', '');
  var r = parseInt(h.substr(0, 2), 16), g = parseInt(h.substr(2, 2), 16), b = parseInt(h.substr(4, 2), 16);
  return (r * 299 + g * 587 + b * 114) / 1000 > 150 ? '#1d2b2a' : '#ffffff';
}
function str_(v) { return v === '' || v == null ? '' : String(v); }
function isIso_(s) { return /^\d{4}-\d{2}-\d{2}$/.test(String(s || '')); }
function norm_(s) { return String(s || '').replace(/[‎‏‪-‮]/g, '').replace(/\s+/g, ' ').trim(); }
function nowStr_() { return Utilities.formatDate(new Date(), 'Asia/Jerusalem', 'dd/MM/yyyy HH:mm'); }
function fmtD_(iso) { var p = String(iso).split('-'); return +p[2] + '/' + (+p[1]); }
function fmtFull_(iso) { return iso.split('-').reverse().join('.'); }
function dayName_(iso) { return 'יום ' + ['ראשון', 'שני', 'שלישי', 'רביעי', 'חמישי', 'שישי', 'שבת'][dow_(iso)]; }
function listDays_(start, end) {
  var out = [];
  if (!isIso_(start) || !isIso_(end)) return out;
  var p = start.split('-'), q = end.split('-');
  var d = new Date(Date.UTC(+p[0], +p[1] - 1, +p[2])), e = new Date(Date.UTC(+q[0], +q[1] - 1, +q[2]));
  while (d <= e && out.length < 120) { out.push(d.toISOString().slice(0, 10)); d.setUTCDate(d.getUTCDate() + 1); }
  return out;
}
function dow_(iso) { var p = iso.split('-'); return new Date(Date.UTC(+p[0], +p[1] - 1, +p[2])).getUTCDay(); }
