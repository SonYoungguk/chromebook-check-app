/**
 * 크롬북 충전함 점검 — 서버 코드 (구글 앱스 스크립트)
 *
 * 구글 시트에 연결된 스크립트로 씁니다. 처음 한 번 setup()을 실행하세요.
 * 로그인은 학교 구글 계정, 코인은 서버에서 뽑습니다(학생 화면에서 조작 불가).
 */

var CONFIG = {
  CART_NAME: '1번 보관함',
  SLOTS: 36,
  // AI 스튜디오에서 지금 쓸 수 있는 모델 이름을 확인해 바꾸세요.
  // 스크립트 속성 GEMINI_MODEL 을 넣으면 그 값이 우선합니다.
  GEMINI_MODEL: 'gemini-2.5-flash',
  REQUIRE_AI_FOR_COIN: true,   // true: 사진 분석이 성공한 제출만 코인 지급
  MAX_UNCERTAIN: 12,           // AI가 확신 못 한 칸이 이보다 많으면 다시 찍게 함
  MAX_ANALYZE_PER_HOUR: 10,    // 한 사람이 한 시간에 사진 분석을 요청할 수 있는 횟수
  // [코인 수, 가중치] — 1~5개, 많이 나올수록 확률이 낮습니다. (35·30·20·10·5 = 100%)
  COIN_TABLE: [[1, 35], [2, 30], [3, 20], [4, 10], [5, 5]],
  REASONS: ['수업 중 사용 중', '수리 맡김', '분실·행방 확인 중', '충전기 고장', '집에 가져감', '기타']
};

var SHEET = { students: '학생', records: '점검기록', ledger: '코인내역', teachers: '교사' };
var HEAD = {
  students: ['이메일', '이름', '당번(O)'],
  records: ['시각', '날짜', '시간대', '보관함', '제출자 이메일', '제출자 이름',
            '36칸 상태 (R충전 중 G완료 O비어 있음)', '비어 있는 칸과 사유', 'AI 판정', '학생이 고친 칸 수', '방식', '코인 지급'],
  ledger: ['시각', '이메일', '이름', '코인', '구분', '메모'],
  teachers: ['이메일']
};
var PERIOD_KO = { morning: '아침', afternoon: '방과 후' };
var KEY2CH = { red: 'R', green: 'G', off: 'O' };

/* ---------- 웹앱 진입 ---------- */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('충전함 점검')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ---------- 처음 한 번 실행 ---------- */
function setup() {
  var ss = SpreadsheetApp.getActive();
  Object.keys(SHEET).forEach(function (k) {
    var sh = ss.getSheetByName(SHEET[k]) || ss.insertSheet(SHEET[k]);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, HEAD[k].length).setValues([HEAD[k]]);
      sh.setFrozenRows(1);
    }
  });
  var first = ss.getSheets()[0];
  if (first.getName() !== SHEET.students && first.getLastRow() === 0 && ss.getSheets().length > 4) ss.deleteSheet(first);
  SpreadsheetApp.getUi().alert('시트 4개를 만들었습니다. 학생 시트에 이메일·이름·당번을 입력하고, 교사 시트에 선생님 이메일을 넣어 주세요.');
}

/* ---------- 도우미 ---------- */
function sheet_(key) {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET[key]);
  if (!sh) throw new Error('먼저 setup 함수를 실행해 주세요.');
  return sh;
}
function rows_(key) { return sheet_(key).getDataRange().getValues().slice(1); }
function tz_() { return Session.getScriptTimeZone() || 'Asia/Seoul'; }
function today_() { return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd'); }
function fmtDate_(v) { return v instanceof Date ? Utilities.formatDate(v, tz_(), 'yyyy-MM-dd') : String(v); }
function fmtWhen_(v) { return v instanceof Date ? Utilities.formatDate(v, tz_(), 'M/d HH:mm') : String(v); }
function safe_(s) { s = String(s == null ? '' : s); return /^[=+\-@]/.test(s) ? "'" + s : s; }
function isOn_(v) { return ['o', 'y', 'yes', 'true', '1', 'v', 'ㅇ', '당번', '✓'].indexOf(String(v).trim().toLowerCase()) >= 0; }

function email_() {
  var e = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  if (!e) throw new Error('로그인 정보를 확인할 수 없습니다. 학교 구글 계정으로 다시 열어 주세요.');
  return e;
}

function role_() {
  var email = email_();
  var student = null;
  rows_('students').some(function (r) {
    if (String(r[0]).trim().toLowerCase() === email) { student = { name: String(r[1] || ''), duty: isOn_(r[2]) }; return true; }
    return false;
  });
  var teacher = rows_('teachers').some(function (r) { return String(r[0]).trim().toLowerCase() === email; });
  return { email: email, student: student, teacher: teacher };
}

function balance_(email) {
  var sum = 0;
  rows_('ledger').forEach(function (r) { if (String(r[1]).trim().toLowerCase() === email) sum += Number(r[3]) || 0; });
  return sum;
}

function history_(email, n) {
  var out = [];
  rows_('ledger').forEach(function (r) {
    if (String(r[1]).trim().toLowerCase() === email) out.push({ when: fmtWhen_(r[0]), coins: Number(r[3]) || 0, kind: String(r[4]), memo: String(r[5] || '') });
  });
  return out.reverse().slice(0, n);
}

function todayStatus_() {
  var today = today_(), st = { morning: null, afternoon: null };
  rows_('records').forEach(function (r) {
    if (fmtDate_(r[1]) !== today || r[3] !== CONFIG.CART_NAME) return;
    var p = r[2] === PERIOD_KO.morning ? 'morning' : 'afternoon';
    if (!st[p]) st[p] = { by: String(r[5] || r[4]), coin: Number(r[11]) > 0 };
  });
  return st;
}

function drawCoins_() {
  var total = 0, t = CONFIG.COIN_TABLE, i;
  for (i = 0; i < t.length; i++) total += t[i][1];
  var x = Math.random() * total;
  for (i = 0; i < t.length; i++) { x -= t[i][1]; if (x < 0) return t[i][0]; }
  return t[0][0];
}

/* ---------- 화면에서 부르는 함수 ---------- */
function getInit() {
  var r = role_();
  var out = {
    email: r.email, student: !!r.student, teacher: r.teacher,
    name: r.student ? r.student.name : '', duty: !!(r.student && r.student.duty),
    cart: CONFIG.CART_NAME, reasons: CONFIG.REASONS, requireAi: CONFIG.REQUIRE_AI_FOR_COIN,
    status: todayStatus_(), coins: 0, history: []
  };
  if (r.student) { out.coins = balance_(r.email); out.history = history_(r.email, 15); }
  return out;
}

function analyze(b64) {
  var r = role_();
  if (!r.student || !r.student.duty) throw new Error('오늘 당번이 아닙니다.');
  if (typeof b64 !== 'string' || b64.length < 1000 || b64.length > 6000000) throw new Error('사진 크기가 올바르지 않습니다.');

  var cache = CacheService.getScriptCache();
  var rlKey = 'rl:' + r.email;
  var used = Number(cache.get(rlKey) || 0);
  if (used >= CONFIG.MAX_ANALYZE_PER_HOUR) return { ok: false, manual: true, message: '분석 요청이 너무 많습니다. 잠시 뒤 다시 시도해 주세요.' };
  cache.put(rlKey, String(used + 1), 3600);

  var props = PropertiesService.getScriptProperties();
  var key = props.getProperty('GEMINI_API_KEY');
  if (!key) return { ok: false, manual: true, message: '사진 분석이 아직 설정되지 않았습니다.' };
  var model = props.getProperty('GEMINI_MODEL') || CONFIG.GEMINI_MODEL;

  var prompt = [
    '이 사진은 학교 크롬북 충전 보관함을 찍은 것입니다. 보관함에는 12칸씩 3줄, 모두 36칸이 있습니다.',
    '각 칸에는 충전 표시등이 있습니다. 빨간불은 충전 중(red), 초록불은 충전 완료(green), 불이 꺼져 있으면 크롬북이 연결되지 않은 빈 칸(off)입니다.',
    '칸 번호는 맨 윗줄 왼쪽에서 오른쪽으로 1~12, 가운데 줄 13~24, 아랫줄 25~36입니다. 사진에 번호가 보여도 위치를 기준으로 세어 주세요.',
    '36칸 각각의 표시등 상태를 판정하세요. 흐리거나 가려져서 확신이 없는 칸은 uncertain에 번호(1~36)로 적으세요.',
    '다른 말 없이 JSON 객체 하나만 답하세요. 형식: {"slots": ["red","green","off", ...정확히 36개], "uncertain": [7, 20], "note": "한 문장 메모"}'
  ].join('\n');

  var body = {
    contents: [{ parts: [{ text: prompt }, { inlineData: { mimeType: 'image/jpeg', data: b64 } }] }],
    generationConfig: { responseMimeType: 'application/json', temperature: 0 }
  };
  var resp;
  try {
    resp = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(model) + ':generateContent', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'x-goog-api-key': key }, payload: JSON.stringify(body)
    });
  } catch (e) {
    console.error('Gemini fetch failed: ' + e);
    return { ok: false, manual: true, message: 'AI 분석 서버에 연결하지 못했습니다.' };
  }
  if (resp.getResponseCode() !== 200) {
    console.error('Gemini HTTP ' + resp.getResponseCode() + ': ' + resp.getContentText().slice(0, 500));
    return { ok: false, manual: true, message: 'AI 분석에 실패했습니다.' };
  }

  var parsed;
  try {
    var data = JSON.parse(resp.getContentText());
    var text = data.candidates[0].content.parts.map(function (p) { return p.text || ''; }).join('');
    var m = text.match(/\{[\s\S]*\}/);
    parsed = JSON.parse(m ? m[0] : text);
  } catch (e) {
    console.error('Gemini parse failed: ' + e);
    return { ok: false, manual: true, message: 'AI 응답을 읽지 못했습니다.' };
  }
  var slots = parsed && parsed.slots;
  var valid = Array.isArray(slots) && slots.length === CONFIG.SLOTS &&
    slots.every(function (v) { return v === 'red' || v === 'green' || v === 'off'; });
  if (!valid) return { ok: false, manual: true, message: 'AI가 36칸을 모두 읽지 못했습니다.' };

  var unc = (Array.isArray(parsed.uncertain) ? parsed.uncertain : [])
    .map(function (n) { return Number(n) - 1; })
    .filter(function (i, idx, a) { return i >= 0 && i < CONFIG.SLOTS && a.indexOf(i) === idx; });
  if (unc.length > CONFIG.MAX_UNCERTAIN) {
    return { ok: false, retake: true, message: '사진이 흐리거나 보관함 전체가 보이지 않습니다. 다시 찍어 주세요.' };
  }

  var id = Utilities.getUuid();
  cache.put('an:' + r.email + ':' + id, JSON.stringify({ slots: slots }), 21600);
  return { ok: true, analysisId: id, slots: slots, uncertain: unc, note: String(parsed.note || '').slice(0, 120) };
}

function submit(p) {
  var r = role_();
  if (!r.student || !r.student.duty) throw new Error('오늘 당번이 아닙니다.');
  if (!p || !PERIOD_KO[p.period]) throw new Error('점검 시간이 올바르지 않습니다.');
  var slots = p.slots;
  if (!Array.isArray(slots) || slots.length !== CONFIG.SLOTS ||
      !slots.every(function (v) { return v === 'red' || v === 'green' || v === 'off'; })) throw new Error('36칸 상태가 올바르지 않습니다.');

  var reasons = p.reasons || {}, detail = [];
  slots.forEach(function (s, i) {
    if (s !== 'off') return;
    var why = String(reasons[String(i + 1)] || '').trim();
    var ok = CONFIG.REASONS.indexOf(why) >= 0 || (why.indexOf('기타: ') === 0 && why.length > 4 && why.length <= 64);
    if (!ok) throw new Error((i + 1) + '번 칸의 사유가 필요합니다.');
    detail.push((i + 1) + '번: ' + why);
  });

  var cache = CacheService.getScriptCache();
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var anKey = 'an:' + r.email + ':' + String(p.analysisId || '');
    var cached = p.analysisId ? cache.get(anKey) : null;
    var ai = cached ? JSON.parse(cached).slots : null;
    var periodKo = PERIOD_KO[p.period], today = today_();

    var already = rows_('records').some(function (row) {
      return fmtDate_(row[1]) === today && row[2] === periodKo && row[3] === CONFIG.CART_NAME && Number(row[11]) > 0;
    });
    var coins = 0, why = '';
    if (already) why = 'already';
    else if (!ai && CONFIG.REQUIRE_AI_FOR_COIN) why = 'manual';
    else coins = drawCoins_();

    var edited = ai ? slots.filter(function (s, i) { return s !== ai[i]; }).length : '';
    var now = new Date();
    sheet_('records').appendRow([
      now, today, periodKo, CONFIG.CART_NAME, r.email, safe_(r.student.name),
      slots.map(function (s) { return KEY2CH[s]; }).join(''), safe_(detail.join('; ')),
      ai ? ai.map(function (s) { return KEY2CH[s]; }).join('') : '', edited, ai ? 'AI+확인' : '직접 표시', coins
    ]);
    if (coins > 0) sheet_('ledger').appendRow([now, r.email, safe_(r.student.name), coins, '획득', today + ' ' + periodKo + ' 점검']);
    if (cached) cache.remove(anKey);
    return { ok: true, coins: coins, why: why, total: balance_(r.email) };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 교사용 ---------- */
function requireTeacher_() {
  var r = role_();
  if (!r.teacher) throw new Error('선생님 계정만 쓸 수 있습니다.');
  return r;
}

// "7번: 수리 맡김; 12번: 기타: 충전기 분실" → { "7": "수리 맡김", "12": "기타: 충전기 분실" }
function parseEmpties_(s) {
  var out = {};
  String(s || '').split(/;\s+(?=\d+번: )/).forEach(function (part) {
    var m = part.match(/^(\d+)번: ([\s\S]*)$/);
    if (m) out[m[1]] = m[2];
  });
  return out;
}

function teacherData() {
  requireTeacher_();
  // 오래된 순서로, 최근 600건까지 (대시보드가 화면에서 기간별로 걸러 씁니다)
  var recs = rows_('records').slice(-600).map(function (r) {
    return { when: fmtWhen_(r[0]), date: fmtDate_(r[1]), period: r[2] === PERIOD_KO.morning ? 'morning' : 'afternoon',
             cart: String(r[3]), email: String(r[4]).trim().toLowerCase(), by: String(r[5] || r[4]),
             slots: String(r[6] || ''), reasons: parseEmpties_(r[7]),
             edited: r[9] === '' ? null : Number(r[9]), mode: String(r[10]), coin: Number(r[11]) || 0 };
  });
  var totals = {};
  rows_('ledger').forEach(function (r) {
    var e = String(r[1]).trim().toLowerCase();
    totals[e] = (totals[e] || 0) + (Number(r[3]) || 0);
  });
  var students = rows_('students').filter(function (r) { return String(r[0]).trim(); }).map(function (r) {
    var e = String(r[0]).trim().toLowerCase();
    return { email: e, name: String(r[1] || e), duty: isOn_(r[2]), total: totals[e] || 0 };
  });
  return { today: today_(), status: todayStatus_(), recent: recs, students: students };
}

function redeem(email, amount, memo) {
  requireTeacher_();
  email = String(email || '').trim().toLowerCase();
  amount = Math.floor(Number(amount));
  if (!(amount > 0)) throw new Error('사용할 코인 수를 1 이상으로 입력해 주세요.');
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var name = null;
    rows_('students').some(function (r) { if (String(r[0]).trim().toLowerCase() === email) { name = String(r[1] || email); return true; } return false; });
    if (name === null) throw new Error('명단에 없는 학생입니다.');
    var bal = balance_(email);
    if (bal < amount) throw new Error('코인이 부족합니다. 현재 ' + bal + '개입니다.');
    sheet_('ledger').appendRow([new Date(), email, safe_(name), -amount, '사용', safe_(String(memo || '').slice(0, 80))]);
    return { ok: true, total: bal - amount };
  } finally {
    lock.releaseLock();
  }
}
