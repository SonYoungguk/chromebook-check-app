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
  GEMINI_MODEL: 'gemini-3.5-flash',   // 2.5 계열은 새 사용자에게 닫혀 404가 납니다. 다른 모델이 필요하면 checkGemini 로 확인하세요.
  REQUIRE_AI_FOR_COIN: true,   // true: 사진 분석이 성공한 제출만 코인 지급
  MAX_UNCERTAIN: 12,           // AI가 확신 못 한 칸이 이보다 많으면 다시 찍게 함
  // Gemini 요청 횟수 상한. 무료 사용량을 넘거나 비용이 생기는 일을 막기 위한 안전장치입니다. (실패한 요청도 셉니다)
  MAX_GEMINI_CALLS_PER_DAY: 150,   // 앱 전체에서 하루에 Gemini에 보낼 수 있는 요청 수
  MAX_CALLS_PER_PHOTO: 6,          // 사진 한 장을 분석하려고 다시 시도하거나 다른 모델을 써 볼 수 있는 최대 횟수
  MAX_SECONDS_PER_PHOTO: 60,       // 사진 한 장의 분석에 쓸 최대 시간(초). 넘으면 더 시도하지 않고 직접 표시로 넘김
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
            '칸 상태 (R충전 중 G완료 N충전 안 됨 O비어 있음 X미사용)', '비어 있는 칸과 사유', 'AI 판정', '학생이 고친 칸 수', '방식', '코인 지급'],
  ledger: ['시각', '이메일', '이름', '코인', '구분', '메모'],
  teachers: ['이메일']
};
var PERIOD_KO = { morning: '아침', afternoon: '방과 후' };
// 칸 상태: red 충전 중, green 충전 완료, idle 크롬북은 있는데 충전 안 됨(LED 꺼짐), off 크롬북 없음(빈 칸), unused 사용하지 않는 칸
var KEY2CH = { red: 'R', green: 'G', idle: 'N', off: 'O', unused: 'X' };

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

  // 실제 보관함 사진 기준: 각 단 위쪽 띠에 작은 둥근 LED가 한 줄로 있고, LED 아래에 칸 번호가 인쇄되어 있습니다.
  // LED는 '충전기가 연결됐는지'만 알려 주므로, 크롬북이 있는지는 따로 판정합니다.
  var SL = CONFIG.SLOTS, CO = CONFIG.COLS, RO = Math.ceil(SL / CO), spans = [], ri;
  for (ri = 0; ri < RO; ri++) spans.push((ri * CO + 1) + '~' + Math.min(SL, (ri + 1) * CO));
  var grid = function (what) { return '[' + spans.map(function () { return '[' + CO + '개: ' + what + ']'; }).join(', ') + ']'; };
  var prompt = [
    '이 사진은 학교 크롬북 충전 보관함을 찍은 것입니다. 보관함은 ' + RO + '단 선반이고 한 단에 ' + CO + '칸씩, 모두 ' + SL + '칸입니다. 칸 번호는 맨 위 단부터 ' + spans.join(', ') + '번이고, 한 단 안에서는 왼쪽에서 오른쪽 순서입니다.',
    '각 단 위쪽 띠에는 칸마다 작은 둥근 LED 표시등이 한 줄로 있고, LED 바로 아래에 칸 번호가 인쇄되어 있습니다. 이 LED 띠는 선반보다 좁게 가운데에 모여 있을 수 있으니 번호로 짝을 맞추세요.',
    '각 단 아래쪽에는 크롬북을 세로로 꽂는 칸이 ' + CO + '개 있고, 칸마다 위에서 코일 케이블이 늘어져 있습니다. 한 단의 칸을 왼쪽부터 세어 1번째, 2번째, ... 칸으로 삼으세요. 크롬북이 꽂힌 칸에는 얇은 흰색 또는 회색 크롬북 본체가 보이고, 빈 칸에는 회색 거치대와 늘어진 케이블만 보입니다.',
    '칸마다 두 가지를 따로 판정하세요.',
    '1) laptop: 그 칸에 크롬북 본체가 꽂혀 있으면 "yes", 없으면 "no". 케이블이 크롬북에 연결되지 않고 늘어져 있어도 크롬북은 꽂혀 있을 수 있습니다.',
    '2) led: 그 칸 번호의 LED가 초록색으로 켜져 있으면 "green"(충전 완료), 빨간색으로 켜져 있으면 "red"(충전 중), 꺼져서 어두운 구멍으로 보이면 "off".',
    'LED는 작고 안쪽으로 들어가 있어서, 켜진 LED는 구멍 아래쪽에 밝은 초록 점으로 보일 수 있습니다. 빛이 반사되어 하얗게 보여 켜졌는지 꺼졌는지 알 수 없으면 그 칸은 uncertain에 적으세요.',
    '크롬북이 있어도 충전기가 연결되지 않으면 LED가 꺼져 있습니다. LED가 꺼진 것만 보고 크롬북이 없다고 판단하지 마세요.',
    '한 단씩, 왼쪽 칸부터 차례로 하나도 건너뛰지 말고 판정하세요. 확신이 없는 칸은 uncertain에 번호(1~' + SL + ')로 적으세요.',
    '다른 말 없이 JSON 객체 하나만 답하세요. 형식: {"laptop": ' + grid('"yes" 또는 "no"') + ', "led": ' + grid('"green", "red", "off" 중 하나') + ', "uncertain": [7, 20], "note": "한 문장 메모"}',
    'laptop의 각 줄은 yes 또는 no, led의 각 줄은 green, red, off 중 하나인 문자열 정확히 ' + CO + '개입니다.'
  ].join('\n');

  var body = {
    contents: [{ parts: [{ text: prompt }, { inlineData: { mimeType: 'image/jpeg', data: b64 } }] }],
    // thinkingBudget 0: 답하기 전에 오래 생각하지 않게 해서 속도를 높입니다. 모델이 이 설정을 거부하면(400) 빼고 다시 시도합니다.
    generationConfig: { responseMimeType: 'application/json', temperature: 0, thinkingConfig: { thinkingBudget: 0 } }
  };
  // 모델 이름 순서: 쓰던 모델 -> 스크립트 속성 GEMINI_MODEL_FALLBACK -> (그래도 안 되면) 모델 목록에서 자동으로 고른 flash 계열
  var models = [model];
  var fb = props.getProperty('GEMINI_MODEL_FALLBACK');
  if (fb && fb !== model) models.push(fb);
  // 최근 6시간 안에 성공했던 모델이 있으면 그 모델을 먼저 씁니다. (설정한 모델이 계속 바쁠 때 매번 기다리지 않도록)
  var good = null;
  try { good = cache.get('gm:good'); } catch (e) { good = null; }
  if (good) models = [good].concat(models.filter(function (m) { return m !== good; }));
  var t0 = Date.now();
  var budget = { left: CONFIG.MAX_CALLS_PER_PHOTO, until: Date.now() + CONFIG.MAX_SECONDS_PER_PHOTO * 1000 };
  // 최근 30분 안에 응답이 없거나 너무 느렸던 모델은 건너뜁니다. (그 모델을 기다리느라 매번 1분씩 걸리는 일을 막기 위해)
  var ready = models.filter(function (m) { return !isSlow_(m); });
  if (ready.length < models.length) console.log('최근 응답이 느렸던 모델은 건너뜁니다: ' + models.filter(isSlow_).join(', '));
  var call = ready.length ? callGemini_(ready, key, body, budget) : { resp: null, code: 0, model: model };
  var stop = function (c) {
    if (c.limited === 'day') return { ok: false, manual: true, message: '오늘 사진 분석 횟수를 모두 썼습니다. 내일 다시 쓸 수 있습니다. (직접 표시로 제출할 수 있습니다)' };
    if (c.limited === 'time') return { ok: false, manual: true, message: 'AI 분석이 너무 오래 걸려 중단했습니다. 잠시 뒤 다시 올리거나 직접 표시해 주세요.' };
    return null;
  };
  if (stop(call)) return stop(call);
  if (call.code !== 200 && !call.limited && (call.code === 0 || call.code === 404 || call.code === 429 || call.code >= 500)) {
    var more = autoModels_(key, models);
    if (more.length) {
      console.log('쓰던 모델이 계속 실패해 자동으로 다른 모델을 시도합니다: ' + more.join(', '));
      // 쓰던 모델을 기다리느라 시간을 다 썼어도, 대체 모델은 최소한 한두 번은 시도할 수 있게 합니다.
      budget.until = Math.max(budget.until, Date.now() + 40000);
      budget.left = Math.max(budget.left, 2);
      call = callGemini_(more, key, body, budget);
    }
  }
  if (stop(call)) return stop(call);
  var resp = call.resp, code = call.code;
  model = call.model;
  if (code === 200) { try { cache.put('gm:good', model, 21600); } catch (e) { /* 기억하지 못해도 괜찮습니다 */ } }
  if (code !== 200) {
    var busy = code === 0 || code >= 500;
    return { ok: false, manual: true, message: code === 429
      ? 'AI 무료 사용량 한도에 도달했습니다. 잠시 뒤 다시 올려 주세요. (오류 429)'
      : busy
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
  var flat = function (a) { return Array.isArray(a) ? [].concat.apply([], a) : []; };
  var slots = null, conflict = [], readCount = 0;
  if (parsed && Array.isArray(parsed.laptop) && Array.isArray(parsed.led)) {
    // 크롬북 유무(laptop)와 LED 색(led)을 따로 받아 칸 상태로 합칩니다.
    var lp = flat(parsed.laptop), ld = flat(parsed.led);
    readCount = Math.min(lp.length, ld.length);
    if (lp.length === CONFIG.SLOTS && ld.length === CONFIG.SLOTS &&
        lp.every(function (v) { return v === 'yes' || v === 'no'; }) && ld.every(function (v) { return v === 'green' || v === 'red' || v === 'off'; })) {
      slots = ld.map(function (led, i) {
        var has = lp[i] === 'yes';
        if (led === 'off') return has ? 'idle' : 'off';   // 크롬북은 있는데 LED가 꺼짐 = 충전 안 됨
        if (!has) conflict.push(i);                       // LED는 켜졌는데 크롬북이 안 보임: 가려졌을 가능성이라 확인하게 함
        return led;
      });
    }
  } else if (parsed) {
    // 예전 형식(칸 상태를 바로 적은 slots 또는 rows)도 받습니다.
    var old = parsed.slots || (Array.isArray(parsed.rows) ? flat(parsed.rows) : null);
    if (Array.isArray(old)) { readCount = old.length; slots = old; }
  }
  var valid = Array.isArray(slots) && slots.length === CONFIG.SLOTS &&
    slots.every(function (v) { return v === 'red' || v === 'green' || v === 'idle' || v === 'off'; });
  if (!valid) return { ok: false, manual: true, message: 'AI가 ' + CONFIG.SLOTS + '칸을 올바르게 읽지 못했습니다. (읽은 칸 ' + readCount + '개)' };

  // 사용하지 않는 칸은 AI가 어떻게 읽었든 '미사용'으로 덮어씁니다.
  slots = slots.slice();
  cart.unused.forEach(function (i) { slots[i] = 'unused'; });
  var unc = (Array.isArray(parsed.uncertain) ? parsed.uncertain : [])
    .map(function (n) { return Number(n) - 1; })
    .concat(conflict)
    .filter(function (i, idx, a) { return i >= 0 && i < CONFIG.SLOTS && a.indexOf(i) === idx && cart.unused.indexOf(i) < 0; });
  if (unc.length > CONFIG.MAX_UNCERTAIN) {
    return { ok: false, retake: true, message: '사진이 흐리거나 보관함 전체가 보이지 않습니다. 다시 찍어 주세요.' };
  }

  var id = Utilities.getUuid();
  cache.put('an:' + r.email + ':' + id, JSON.stringify({ slots: slots }), 21600);
  return { ok: true, analysisId: id, slots: slots, uncertain: unc, note: String(parsed.note || '').slice(0, 120) };
}

// 사진은 학생 기기 안에서 컴퓨터가 직접 읽고(LED 색), 이 함수는 그 결과만 받아 검증·저장합니다. (사진이 서버로 오지 않음, AI 호출 없음)
// LED가 꺼진 칸은 크롬북이 있는지 알 수 없으므로 '비어 있음'으로 두고 반드시 확인하게 합니다.
function scan(leds) {
  var r = role_();
  if (!r.student || !r.student.duty) throw new Error('오늘 당번이 아닙니다.');
  var cart = cartFor_(r.student);
  if (!cart) throw new Error('보관함(반) 설정을 찾을 수 없습니다. 선생님께 알려 주세요.');
  if (!Array.isArray(leds) || leds.length !== CONFIG.SLOTS || !leds.every(function (v) { return v === 'green' || v === 'red' || v === 'lit' || v === 'off'; })) {
    return { ok: false, manual: true, message: 'LED 읽기 결과가 올바르지 않습니다.' };
  }
  var cache = CacheService.getScriptCache();
  var rlKey = 'rs:' + r.email, used = Number(cache.get(rlKey) || 0);
  if (!r.teacher && used >= 60) return { ok: false, manual: true, message: '한 시간에 사진 읽기는 60번까지 할 수 있습니다. 잠시 뒤 다시 시도해 주세요.' };
  cache.put(rlKey, String(used + 1), 3600);

  var slots = leds.map(function (v) { return v === 'lit' ? 'green' : v; });
  var unc = [];
  leds.forEach(function (v, i) { if (cart.unused.indexOf(i) < 0 && (v === 'off' || v === 'lit')) unc.push(i); });
  cart.unused.forEach(function (i) { slots[i] = 'unused'; });
  var id = Utilities.getUuid();
  cache.put('an:' + r.email + ':' + id, JSON.stringify({ slots: slots }), 21600);
  return { ok: true, analysisId: id, slots: slots, uncertain: unc, note: 'LED 색은 컴퓨터가 읽었습니다. 꺼진 칸은 크롬북이 있는지 확인해 주세요.' };
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
    return cart.unused.indexOf(i) >= 0 ? v === 'unused' : (v === 'red' || v === 'green' || v === 'idle' || v === 'off');
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

/* ---------- Gemini 호출 도우미 ---------- */
// 모델 이름 목록을 차례로 시도합니다.
// 일시적인 오류(429 사용량 초과, 5xx 서버 혼잡, 연결 실패)는 잠시 기다렸다가 모델마다 최대 3번,
// 404·400 같은 설정 문제는 다시 해도 소용없으므로 바로 다음 모델로 넘어갑니다.
function callGemini_(models, key, body, budget) {
  var resp = null, code = 0, model = models[0];
  for (var mi = 0; mi < models.length; mi++) {
    model = models[mi];
    for (var attempt = 0; attempt < 3; attempt++) {
      if (budget && budget.left <= 0) return { resp: resp, code: code, model: model, limited: 'photo' };   // 사진 한 장당 최대 요청 수
      if (budget && budget.until && Date.now() > budget.until) return { resp: resp, code: code, model: model, limited: 'time' };   // 사진 한 장당 최대 시간
      if (!bumpDaily_()) return { resp: resp, code: code, model: model, limited: 'day' };                  // 하루 상한
      if (budget) budget.left--;
      var t1 = Date.now();
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
      if (code === 200) { console.log('Gemini 한 번 요청에 ' + (Date.now() - t1) + 'ms (모델 ' + model + ')'); return { resp: resp, code: 200, model: model }; }
      if (code === 400 && body.generationConfig.thinkingConfig) {
        // thinking 설정을 거부하는 모델: 예산(0) -> 수준(low) -> 설정 없음 순서로 바꿔 다시 시도합니다.
        // (thinking을 아예 빼면 답이 오래 걸릴 수 있어 마지막 수단입니다)
        console.error('모델 ' + model + ' 이(가) thinking 설정을 거부했습니다: ' + (resp ? resp.getContentText().slice(0, 200) : ''));
        if (body.generationConfig.thinkingConfig.thinkingBudget === 0) body.generationConfig.thinkingConfig = { thinkingLevel: 'low' };
        else delete body.generationConfig.thinkingConfig;
        attempt--;
        continue;
      }
      // 실행 기록(Apps Script 왼쪽 '실행')에서 원인을 볼 수 있습니다. 404는 모델 이름, 400/403은 키나 권한 문제인 경우가 많습니다.
      console.error('Gemini HTTP ' + code + ' (모델 ' + model + ', 시도 ' + (attempt + 1) + '): ' + (resp ? resp.getContentText().slice(0, 300) : '연결 실패'));
      if (code === 0 || Date.now() - t1 > 25000) markSlow_(model);   // 응답이 없거나 25초 넘게 걸린 모델은 한동안 건너뜀
      if (code === 429) markSlow_(model, 600);                       // 사용량 초과: 같은 모델로 다시 해도 소용없으니 10분 동안 건너뜀
      if (code === 404) {
        // 이 계정에서는 쓸 수 없는 모델(예: '새 사용자에게는 더 이상 제공되지 않음'): 하루 동안 건너뛰고,
        // 구글이 오류 문구에서 알려 준 대체 모델('use models/...')이 있으면 바로 이어서 시도합니다.
        markSlow_(model, 86400);
        var sug = String(resp ? resp.getContentText() : '').match(/use models\/(gemini-[A-Za-z0-9._-]+)/);
        if (sug && models.indexOf(sug[1]) < 0 && !isSlow_(sug[1])) {
          console.log('구글이 알려 준 대체 모델을 시도합니다: ' + sug[1] + ' (' + model + ' 대신)');
          models.push(sug[1]);
        }
      }
      var transient = code === 0 || code >= 500;                     // 429는 재시도하지 않음
      if (!transient) break;
      if (Date.now() - t1 > 25000) break;          // 응답이 느린 모델은 같은 모델로 다시 시도하지 않고 다음 모델로 넘어갑니다.
      if (attempt < 2) Utilities.sleep(1500 * (attempt + 1));
    }
    if (code >= 500) markSlow_(model, 120);   // 서버가 바쁜 모델은 2분 동안 건너뛰어 다음 사진이 헛되이 기다리지 않게 함
  }
  return { resp: resp, code: code, model: model };
}

// 응답이 없거나 너무 느렸던 모델은 30분 동안 건너뛰도록 기억합니다.
function isSlow_(model) { try { return !!CacheService.getScriptCache().get('gm:slow:' + model); } catch (e) { return false; } }
function markSlow_(model, seconds) { try { CacheService.getScriptCache().put('gm:slow:' + model, '1', seconds || 1800); } catch (e) { /* 기억하지 못해도 분석은 계속합니다 */ } }

// 오늘 보낸 Gemini 요청 수를 세어, 상한(MAX_GEMINI_CALLS_PER_DAY)을 넘으면 false 를 돌려줍니다.
function bumpDaily_() {
  try {
    var props = PropertiesService.getScriptProperties();
    var key = 'calls:' + today_();
    var n = Number(props.getProperty(key) || 0);
    if (n >= CONFIG.MAX_GEMINI_CALLS_PER_DAY) return false;
    if (n === 0) {   // 하루의 첫 요청: 지난 날짜의 기록은 지웁니다.
      props.getKeys().forEach(function (k) { if (k.indexOf('calls:') === 0 && k !== key) props.deleteProperty(k); });
    }
    props.setProperty(key, String(n + 1));
  } catch (e) { /* 세지 못해도 분석은 막지 않습니다 */ }
  return true;
}

// 쓰던 모델이 계속 바쁘거나 없을 때를 위한 자동 대체 모델.
// 모델 목록에서 사진을 읽을 수 있는 flash 계열을 골라, 안정 버전과 새 버전을 먼저 시도합니다. (목록은 6시간 동안 기억)
function autoModels_(key, tried) {
  var cache = CacheService.getScriptCache(), names = null;
  var cached = cache.get('gm:list');
  if (cached) { try { names = JSON.parse(cached); } catch (e) { names = null; } }
  if (!names) {
    try {
      var r = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models?pageSize=200', {
        headers: { 'x-goog-api-key': key }, muteHttpExceptions: true
      });
      if (r.getResponseCode() !== 200) return [];
      names = (JSON.parse(r.getContentText()).models || [])
        .filter(function (m) { return (m.supportedGenerationMethods || []).indexOf('generateContent') >= 0; })
        .map(function (m) { return String(m.name).replace(/^models\//, ''); });
      cache.put('gm:list', JSON.stringify(names), 21600);
    } catch (e) { return []; }
  }
  // 사진을 읽는 일반 모델만 고릅니다: gemini-2.5-flash, gemini-3-flash 처럼 '숫자 + flash'(또는 flash-lite) 이름.
  // omni, 별칭(latest), preview, 이미지·음성·임베딩 전용 같은 낯선 모델은 무료 한도가 막혀 있거나 용도가 달라 제외합니다.
  var ok = names.filter(function (n) {
    return /^gemini-\d+(\.\d+)?-flash(-lite)?$/.test(n) && tried.indexOf(n) < 0 && !isSlow_(n);
  });
  // 정렬: flash-lite 는 뒤로, 그 안에서는 버전이 낮은(오래되어 덜 붐비는) 모델부터. 새 모델은 한꺼번에 몰려 503이 잦습니다.
  var ver = function (n) { return parseFloat((n.match(/^gemini-(\d+(?:\.\d+)?)-/) || [0, 0])[1]); };
  var rank = function (n) { return (/lite/.test(n) ? 100 : 0) + ver(n); };
  ok.sort(function (x, y) { return rank(x) - rank(y); });
  return ok.slice(0, 2);
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

  try {
    var t = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(model) + ':generateContent', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true, headers: { 'x-goog-api-key': key },
      payload: JSON.stringify({ contents: [{ parts: [{ text: '숫자 7만 답하세요.' }] }] })
    });
    Logger.log('② 현재 모델로 짧은 질문: HTTP ' + t.getResponseCode() + ' / ' + t.getContentText().slice(0, 300));
  } catch (e) {
    Logger.log('② 현재 모델이 응답하지 않습니다 (연결 시간 초과). 아래 ③에서 다른 모델과 속도를 비교하세요. ' + e);
  }

  // ③ 속도 시험: 아주 작은 사진 한 장을 보내고 모델마다 응답 시간을 잽니다.
  // 새 모델은 붐비는 경우가 많아서, 오래된 안정 모델부터 폭넓게(최대 8개) 시험합니다. 몇 분 걸릴 수 있습니다.
  var png = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
  var strict = names.filter(function (n) { return /^gemini-\d+(\.\d+)?-flash(-lite)?$/.test(n); });
  var ver = function (n) { return parseFloat((n.match(/^gemini-(\d+(?:\.\d+)?)-/) || [0, 0])[1]); };
  strict.sort(function (x, y) { return ((/lite/.test(x) ? 100 : 0) + ver(x)) - ((/lite/.test(y) ? 100 : 0) + ver(y)); });
  var cands = [model].concat(strict.filter(function (n) { return n !== model; })).slice(0, 8);
  Logger.log('③ 속도 시험 대상 (' + cands.length + '개): ' + cands.join(', '));
  var results = [], startAll = Date.now();
  cands.forEach(function (m) {
    if (Date.now() - startAll > 240000) { Logger.log('③ ' + m + ': 시간이 부족해 시험하지 않음'); return; }
    var t0 = Date.now();
    try {
      var rr = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(m) + ':generateContent', {
        method: 'post', contentType: 'application/json', muteHttpExceptions: true, headers: { 'x-goog-api-key': key },
        payload: JSON.stringify({
          contents: [{ parts: [{ text: '이 사진의 색을 한 단어로만 답하세요.' }, { inlineData: { mimeType: 'image/png', data: png } }] }],
          generationConfig: { temperature: 0, thinkingConfig: { thinkingBudget: 0 } }
        })
      });
      var c = rr.getResponseCode(), ms = Date.now() - t0;
      Logger.log('③ ' + m + ': HTTP ' + c + ', ' + ms + 'ms' + (c === 200 ? '' : ' / ' + String(rr.getContentText()).replace(/\s+/g, ' ').slice(0, 230)));
      if (c === 200) results.push({ m: m, ms: ms });
    } catch (e) {
      Logger.log('③ ' + m + ': 응답 없음 (' + (Date.now() - t0) + 'ms 뒤 ' + e + ')');
    }
  });
  results.sort(function (x, y) { return x.ms - y.ms; });
  if (results.length) {
    Logger.log('→ 쓸 수 있는 모델(빠른 순서): ' + results.map(function (r) { return r.m + ' ' + r.ms + 'ms'; }).join(', '));
    Logger.log('→ 스크립트 속성 GEMINI_MODEL 에 ' + results[0].m + (results[1] ? ', GEMINI_MODEL_FALLBACK 에 ' + results[1].m : '') + ' 를 넣어 보세요.');
  } else {
    Logger.log('→ 시험한 모델이 모두 실패했습니다. 구글 서버가 붐비는 시간일 수 있으니 10~20분 뒤에 다시 실행해 보세요.');
  }
}
