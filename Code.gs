/**
 * 크롬북 충전함 점검 — 서버 코드 (구글 앱스 스크립트)
 *
 * 구글 시트에 연결된 스크립트로 씁니다. 처음 한 번 setup()을 실행하세요.
 * 로그인은 학교 구글 계정, 코인은 서버에서 뽑습니다(학생 화면에서 조작 불가).
 */

var CONFIG = {
  CART_NAME: '1번 보관함',   // 보관함 시트가 비어 있을 때 쓰는 기본 이름
  // 보관함 구조: 3단 선반, 한 단에 11칸, 모두 33칸 (칸 번호는 보관함 표시등 아래에 인쇄된 번호와 같음)
  SLOTS: 33,
  COLS: 11,
  // AI 스튜디오에서 지금 쓸 수 있는 모델 이름을 확인해 바꾸세요.
  // 스크립트 속성 GEMINI_MODEL 을 넣으면 그 값이 우선합니다.
  GEMINI_MODEL: 'gemini-2.5-flash',
  REQUIRE_AI_FOR_COIN: true,   // true: 사진 분석이 성공한 제출만 코인 지급
  MAX_UNCERTAIN: 12,           // AI가 확신 못 한 칸이 이보다 많으면 다시 찍게 함
  MAX_ANALYZE_PER_HOUR: 20,    // 한 사람이 한 시간에 사진 분석에 성공할 수 있는 횟수 (실패한 요청은 세지 않음, 교사는 제한 없음)
  // [코인 수, 가중치] — 1~5개, 많이 나올수록 확률이 낮습니다. (35·30·20·10·5 = 100%)
  COIN_TABLE: [[1, 35], [2, 30], [3, 20], [4, 10], [5, 5]],
  REASONS: ['수업 중 사용 중', '수리 맡김', '분실·행방 확인 중', '충전기 고장', '집에 가져감', '기타']
};

var SHEET = { students: '학생', carts: '보관함', records: '점검기록', ledger: '코인내역', teachers: '교사' };
var HEAD = {
  students: ['이메일', '이름', '당번(O)', '반(보관함 이름)', '번호'],
  carts: ['보관함(반)', '사용하는 칸 수', '사용하지 않는 칸 (선택, 쉼표로 구분)'],
  records: ['시각', '날짜', '시간대', '보관함', '제출자 이메일', '제출자 이름',
            '칸 상태 (R충전 중 G완료 O비어 있음 X미사용)', '비어 있는 칸과 사유', 'AI 판정', '학생이 고친 칸 수', '방식', '코인 지급'],
  ledger: ['시각', '이메일', '이름', '코인', '구분', '메모'],
  teachers: ['이메일']
};
var PERIOD_KO = { morning: '아침', afternoon: '방과 후' };
var KEY2CH = { red: 'R', green: 'G', off: 'O', unused: 'X' };

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
    var n = HEAD[k].length;
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, n).setValues([HEAD[k]]);
      sh.setFrozenRows(1);
    } else {
      // 예전에 만든 시트에 새로 생긴 머리글(예: 학생 시트의 반, 번호)이 비어 있으면 채웁니다.
      var cur = sh.getRange(1, 1, 1, n).getValues()[0];
      sh.getRange(1, 1, 1, n).setValues([cur.map(function (v, i) { return String(v).trim() ? v : HEAD[k][i]; })]);
    }
  });
  var carts = ss.getSheetByName(SHEET.carts);
  if (carts.getLastRow() < 2) carts.appendRow([CONFIG.CART_NAME, CONFIG.SLOTS, '']);
  var first = ss.getSheets()[0];
  if (first.getName() !== SHEET.students && first.getLastRow() === 0 && ss.getSheets().length > Object.keys(SHEET).length) ss.deleteSheet(first);
  SpreadsheetApp.getUi().alert('시트를 준비했습니다.\n1) 보관함 시트에 반 이름과 사용하는 칸 수를 입력하세요.\n2) 학생 시트에 이메일, 이름, 당번(O), 반을 입력하세요.\n3) 교사 시트에 선생님 이메일을 넣으세요.');
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

// 보관함 시트의 한 줄을 { name, size, used, unused }로 바꿉니다. unused는 0부터 시작하는 칸 위치입니다.
// 사용하는 칸 수가 N이면 N+1번 이후가 모두 미사용이고, '사용하지 않는 칸'에 적은 번호도 미사용입니다.
function cartInfo_(name, size, unusedText) {
  var n = Math.floor(Number(size));
  if (!(n >= 1 && n <= CONFIG.SLOTS)) n = CONFIG.SLOTS;
  var unused = [], i;
  for (i = n; i < CONFIG.SLOTS; i++) unused.push(i);
  String(unusedText == null ? '' : unusedText).split(/[,\s;]+/).forEach(function (t) {
    var k = Math.floor(Number(t));
    if (k >= 1 && k <= CONFIG.SLOTS && unused.indexOf(k - 1) < 0) unused.push(k - 1);
  });
  unused.sort(function (a, b) { return a - b; });
  return { name: name, size: n, used: CONFIG.SLOTS - unused.length, unused: unused };
}

function carts_() {
  var list = [];
  // 보관함 시트가 아직 없으면(setup을 다시 실행하기 전) 기본 보관함 하나로 동작합니다.
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET.carts);
  (sh ? sh.getDataRange().getValues().slice(1) : []).forEach(function (r) {
    var name = String(r[0]).trim();
    if (name) list.push(cartInfo_(name, r[1], r[2]));
  });
  if (!list.length) list.push(cartInfo_(CONFIG.CART_NAME, CONFIG.SLOTS, ''));
  return list;
}

// 학생이 점검할 보관함. 학생 시트의 '반'이 비어 있으면 보관함이 하나뿐일 때만 그 보관함을 씁니다.
function cartFor_(student) {
  var list = carts_(), i;
  if (student.cartName) {
    for (i = 0; i < list.length; i++) if (list[i].name === student.cartName) return list[i];
    return null;
  }
  return list.length === 1 ? list[0] : null;
}

function role_() {
  var email = email_();
  var student = null;
  rows_('students').some(function (r) {
    if (String(r[0]).trim().toLowerCase() === email) {
      student = { name: String(r[1] || ''), duty: isOn_(r[2]), cartName: String(r[3] == null ? '' : r[3]).trim(), no: String(r[4] == null ? '' : r[4]).trim() };
      return true;
    }
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

function todayStatus_(cartName) {
  var today = today_(), st = { morning: null, afternoon: null };
  rows_('records').forEach(function (r) {
    if (fmtDate_(r[1]) !== today || String(r[3]) !== cartName) return;
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
  var cart = r.student ? cartFor_(r.student) : null, notice = '';
  if (r.student && r.student.duty && !cart) {
    notice = r.student.cartName
      ? '보관함 시트에 "' + r.student.cartName + '" 반이 없습니다. 선생님께 알려 주세요.'
      : '학생 시트의 반 칸이 비어 있습니다. 선생님께 알려 주세요.';
  }
  var out = {
    email: r.email, student: !!r.student, teacher: r.teacher,
    name: r.student ? r.student.name : '', duty: !!(r.student && r.student.duty && cart),
    cart: cart ? cart.name : '', unused: cart ? cart.unused : [], notice: notice, slots: CONFIG.SLOTS, cols: CONFIG.COLS,
    reasons: CONFIG.REASONS, requireAi: CONFIG.REQUIRE_AI_FOR_COIN,
    status: cart ? todayStatus_(cart.name) : { morning: null, afternoon: null }, coins: 0, history: []
  };
  if (r.student) { out.coins = balance_(r.email); out.history = history_(r.email, 15); }
  return out;
}

function analyze(b64) {
  var r = role_();
  if (!r.student || !r.student.duty) throw new Error('오늘 당번이 아닙니다.');
  var cart = cartFor_(r.student);
  if (!cart) throw new Error('보관함(반) 설정을 찾을 수 없습니다. 선생님께 알려 주세요.');
  if (typeof b64 !== 'string' || b64.length < 1000 || b64.length > 6000000) throw new Error('사진 크기가 올바르지 않습니다.');

  var cache = CacheService.getScriptCache();
  var rlKey = 'rl:' + r.email;
  var used = Number(cache.get(rlKey) || 0);
  if (!r.teacher && used >= CONFIG.MAX_ANALYZE_PER_HOUR) {
    return { ok: false, manual: true, message: '한 시간에 사진 분석은 ' + CONFIG.MAX_ANALYZE_PER_HOUR + '번까지 할 수 있습니다. 잠시 뒤 다시 시도해 주세요.' };
  }

  var props = PropertiesService.getScriptProperties();
  var key = props.getProperty('GEMINI_API_KEY');
  if (!key) return { ok: false, manual: true, message: '사진 분석이 아직 설정되지 않았습니다.' };
  var model = props.getProperty('GEMINI_MODEL') || CONFIG.GEMINI_MODEL;

  // 실제 보관함 사진 기준: 각 단 위쪽 띠에 작은 둥근 표시등이 한 줄로 있고, 표시등 아래에 칸 번호가 인쇄되어 있습니다.
  var SL = CONFIG.SLOTS, CO = CONFIG.COLS, RO = Math.ceil(SL / CO), spans = [], ri;
  for (ri = 0; ri < RO; ri++) spans.push((ri * CO + 1) + '~' + Math.min(SL, (ri + 1) * CO));
  var prompt = [
    '이 사진은 학교 크롬북 충전 보관함을 찍은 것입니다. 보관함은 ' + RO + '단 선반이고 한 단에 ' + CO + '칸씩, 모두 ' + SL + '칸입니다.',
    '각 단의 위쪽 띠에는 칸마다 작은 둥근 LED 표시등이 한 줄로 있고, 표시등 바로 아래에 칸 번호가 인쇄되어 있습니다. 맨 위 단부터 ' + spans.join(', ') + '번입니다.',
    '판정은 표시등을 기준으로 합니다.',
    '- green: 표시등이 초록색으로 켜져 있음 (충전 완료)',
    '- red: 표시등이 빨간색이나 주황색으로 켜져 있음 (충전 중)',
    '- off: 표시등이 꺼져 있어 어두운 검은색 구멍으로 보임. 이런 칸에는 보통 크롬북이 없고 케이블만 늘어져 있습니다.',
    '표시등은 작고 안쪽으로 들어가 있어서, 켜진 표시등은 구멍 아래쪽에 밝은 초록 점으로 보일 수 있습니다. 빛이 반사되어 하얗게 보여 켜졌는지 꺼졌는지 알 수 없으면 그 칸은 uncertain에 적으세요.',
    '표시등과 그 아래 번호를 정확히 짝지어 한 단씩, 왼쪽 칸부터 차례로 하나도 건너뛰지 말고 판정하세요. 크롬북 본체는 선반 아래에 있고 표시등과 위치가 어긋나 보일 수 있으니, 크롬북 위치가 아니라 표시등과 번호로 칸을 구분하세요.',
    '확신이 없는 칸은 uncertain에 번호(1~' + SL + ')로 적으세요.',
    '다른 말 없이 JSON 객체 하나만 답하세요. 형식: {"rows": [' + spans.map(function () { return '[' + CO + '개]'; }).join(', ') + '], "uncertain": [7, 20], "note": "한 문장 메모"}',
    '각 줄은 green, red, off 중 하나인 문자열 정확히 ' + CO + '개입니다.'
  ].join('\n');

  var body = {
    contents: [{ parts: [{ text: prompt }, { inlineData: { mimeType: 'image/jpeg', data: b64 } }] }],
    // thinkingBudget 0: 답하기 전에 오래 생각하지 않게 해서 속도를 높입니다. 모델이 이 설정을 거부하면(400) 빼고 다시 시도합니다.
    generationConfig: { responseMimeType: 'application/json', temperature: 0, thinkingConfig: { thinkingBudget: 0 } }
  };
  // 일시적인 오류(429 사용량 초과, 5xx 서버 혼잡)는 잠시 기다렸다가 최대 3번까지 다시 시도합니다.
  // 404·400 같은 설정 문제는 다시 해도 소용없으므로 바로 다음 모델로 넘어갑니다.
  // 스크립트 속성 GEMINI_MODEL_FALLBACK 에 다른 모델 이름을 넣으면 그 모델로도 시도합니다.
  var models = [model];
  var fb = props.getProperty('GEMINI_MODEL_FALLBACK');
  if (fb && fb !== model) models.push(fb);
  var resp = null, code = 0, t0 = Date.now();
  outer:
  for (var mi = 0; mi < models.length; mi++) {
    model = models[mi];
    for (var attempt = 0; attempt < 3; attempt++) {
      try {
        resp = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(model) + ':generateContent', {
          method: 'post', contentType: 'application/json', muteHttpExceptions: true,
          headers: { 'x-goog-api-key': key }, payload: JSON.stringify(body)
        });
        code = resp.getResponseCode();
      } catch (e) {
        console.error('Gemini fetch failed (모델 ' + model + '): ' + e);
        resp = null; code = 0;
      }
      if (code === 200) break outer;
      if (code === 400 && body.generationConfig.thinkingConfig) {
        console.error('모델 ' + model + ' 이(가) thinking 설정을 거부해 빼고 다시 시도합니다: ' + (resp ? resp.getContentText().slice(0, 200) : ''));
        delete body.generationConfig.thinkingConfig;
        attempt--;
        continue;
      }
      // 실행 기록(Apps Script 왼쪽 '실행')에서 원인을 볼 수 있습니다. 404는 모델 이름, 400/403은 키나 권한 문제인 경우가 많습니다.
      console.error('Gemini HTTP ' + code + ' (모델 ' + model + ', 시도 ' + (attempt + 1) + '): ' + (resp ? resp.getContentText().slice(0, 300) : '연결 실패'));
      var transient = code === 0 || code === 429 || code >= 500;
      if (!transient) break;
      if (attempt < 2) Utilities.sleep(1500 * (attempt + 1));
    }
  }
  if (code !== 200) {
    var busy = code === 0 || code === 429 || code >= 500;
    return { ok: false, manual: true, message: busy
      ? 'AI 서버가 지금 바쁩니다. 잠시 뒤 다시 올려 주세요. (오류 ' + (code || '연결 실패') + ')'
      : 'AI 분석에 실패했습니다. (오류 ' + code + ')' };
  }

  cache.put(rlKey, String(used + 1), 3600);
  console.log('Gemini 소요 시간: ' + (Date.now() - t0) + 'ms (모델 ' + model + ')');
  var parsed, text = '';
  try {
    var data = JSON.parse(resp.getContentText());
    text = data.candidates[0].content.parts.map(function (p) { return p.text || ''; }).join('');
    var m = text.match(/\{[\s\S]*\}/);
    parsed = JSON.parse(m ? m[0] : text);
  } catch (e) {
    console.error('Gemini parse failed: ' + e + ' / 응답: ' + resp.getContentText().slice(0, 500));
    return { ok: false, manual: true, message: 'AI 응답을 읽지 못했습니다.' };
  }
  console.log('Gemini 응답 (모델 ' + model + '): ' + text.slice(0, 800));
  var slots = parsed && parsed.slots;
  if (!slots && parsed && Array.isArray(parsed.rows)) slots = [].concat.apply([], parsed.rows);
  var valid = Array.isArray(slots) && slots.length === CONFIG.SLOTS &&
    slots.every(function (v) { return v === 'red' || v === 'green' || v === 'off'; });
  if (!valid) return { ok: false, manual: true, message: 'AI가 ' + CONFIG.SLOTS + '칸을 모두 읽지 못했습니다. (읽은 칸 ' + (Array.isArray(slots) ? slots.length : 0) + '개)' };

  // 사용하지 않는 칸은 AI가 어떻게 읽었든 '미사용'으로 덮어씁니다.
  slots = slots.slice();
  cart.unused.forEach(function (i) { slots[i] = 'unused'; });
  var unc = (Array.isArray(parsed.uncertain) ? parsed.uncertain : [])
    .map(function (n) { return Number(n) - 1; })
    .filter(function (i, idx, a) { return i >= 0 && i < CONFIG.SLOTS && a.indexOf(i) === idx && cart.unused.indexOf(i) < 0; });
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
  var cart = cartFor_(r.student);
  if (!cart) throw new Error('보관함(반) 설정을 찾을 수 없습니다. 선생님께 알려 주세요.');
  var slots = p.slots;
  // 사용하는 칸은 red·green·off, 사용하지 않는 칸은 unused 여야 합니다.
  if (!Array.isArray(slots) || slots.length !== CONFIG.SLOTS || !slots.every(function (v, i) {
    return cart.unused.indexOf(i) >= 0 ? v === 'unused' : (v === 'red' || v === 'green' || v === 'off');
  })) throw new Error('칸 상태가 올바르지 않습니다. 화면을 새로 고친 뒤 다시 제출해 주세요.');

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
      return fmtDate_(row[1]) === today && row[2] === periodKo && String(row[3]) === cart.name && Number(row[11]) > 0;
    });
    var coins = 0, why = '';
    if (already) why = 'already';
    else if (!ai && CONFIG.REQUIRE_AI_FOR_COIN) why = 'manual';
    else coins = drawCoins_();

    var edited = ai ? slots.filter(function (s, i) { return s !== ai[i]; }).length : '';
    var now = new Date();
    sheet_('records').appendRow([
      now, today, periodKo, cart.name, r.email, safe_(r.student.name),
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
  var carts = carts_();
  var students = rows_('students').filter(function (r) { return String(r[0]).trim(); }).map(function (r) {
    var e = String(r[0]).trim().toLowerCase();
    var cartName = String(r[3] == null ? '' : r[3]).trim() || (carts.length === 1 ? carts[0].name : '');
    return { email: e, name: String(r[1] || e), duty: isOn_(r[2]), total: totals[e] || 0, cart: cartName, no: String(r[4] == null ? '' : r[4]).trim() };
  });
  var todayByCart = {};
  carts.forEach(function (c) { todayByCart[c.name] = todayStatus_(c.name); });
  return { today: today_(), slots: CONFIG.SLOTS, cols: CONFIG.COLS, carts: carts, todayByCart: todayByCart, recent: recs, students: students };
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

/* ---------- 진단용: Gemini 연결 확인 ---------- */
// Apps Script 편집기에서 함수 목록에 checkGemini 를 고르고 실행한 뒤, 아래 '실행 로그'를 보세요.
// 키가 맞는지, 지금 설정된 모델 이름을 쓸 수 있는지, 쓸 수 있는 모델 이름이 무엇인지 알려 줍니다.
function checkGemini() {
  var props = PropertiesService.getScriptProperties();
  var key = props.getProperty('GEMINI_API_KEY');
  var model = props.getProperty('GEMINI_MODEL') || CONFIG.GEMINI_MODEL;
  if (!key) { Logger.log('GEMINI_API_KEY 스크립트 속성이 없습니다. 프로젝트 설정 → 스크립트 속성에서 추가하세요.'); return; }
  Logger.log('현재 모델 설정: ' + model);

  var r = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models?pageSize=200', {
    headers: { 'x-goog-api-key': key }, muteHttpExceptions: true
  });
  Logger.log('① 모델 목록 요청: HTTP ' + r.getResponseCode());
  if (r.getResponseCode() !== 200) {
    Logger.log('키가 잘못되었거나 이 계정에서 Gemini API를 쓸 수 없습니다. 응답: ' + r.getContentText().slice(0, 400));
    return;
  }
  var names = (JSON.parse(r.getContentText()).models || [])
    .filter(function (m) { return (m.supportedGenerationMethods || []).indexOf('generateContent') >= 0; })
    .map(function (m) { return String(m.name).replace(/^models\//, ''); });
  Logger.log('사진 분석에 쓸 수 있는 모델 후보(generateContent 지원): ' + names.join(', '));
  Logger.log(names.indexOf(model) >= 0 ? '→ 현재 모델 ' + model + ' 은(는) 목록에 있습니다.'
    : '→ 현재 모델 ' + model + ' 은(는) 목록에 없습니다. 위 목록의 flash 계열 이름을 스크립트 속성 GEMINI_MODEL 에 넣어 보세요.');

  var t = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(model) + ':generateContent', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true, headers: { 'x-goog-api-key': key },
    payload: JSON.stringify({ contents: [{ parts: [{ text: '숫자 7만 답하세요.' }] }] })
  });
  Logger.log('② 현재 모델로 짧은 질문: HTTP ' + t.getResponseCode() + ' / ' + t.getContentText().slice(0, 300));
}
