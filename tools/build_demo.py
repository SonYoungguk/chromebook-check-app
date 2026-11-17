"""시연용 페이지를 만듭니다: Index.html 에 가짜 서버(google.script.run 흉내)를 붙여 demo/index.html 로 저장.

사용법:  python3 tools/build_demo.py
- 구글 로그인, 사진 분석(Gemini), 시트 저장 없이 화면과 흐름만 보여 줍니다.
- 모든 데이터는 가짜이고 브라우저 안에서만 움직입니다. 새로 고치면 처음 상태로 돌아갑니다.
"""
import os

root = os.path.join(os.path.dirname(__file__), "..")
src = open(os.path.join(root, "Index.html"), encoding="utf8").read()

BANNER = """<div id="demo-banner" style="background:#fff4d6;color:#4a3500;border-bottom:1px solid #e8cf8a;padding:8px 16px;font:500 13px/1.5 system-ui,sans-serif;text-align:center">
시연용 화면입니다. 가짜 데이터이고, 올린 사진과 관계없이 예시 분석 결과가 나오며, 아무것도 저장되지 않습니다.
</div>"""

MOCK = r"""<script>
/* ===== 시연용 가짜 서버 (실제 앱에서는 구글 앱스 스크립트가 이 역할을 합니다) ===== */
(function () {
  var REASONS = ['수업 중 사용 중', '수리 맡김', '분실·행방 확인 중', '충전기 고장', '집에 가져감', '기타'];
  var CH = { red: 'R', green: 'G', off: 'O', unused: 'X' };
  function range(a, b) { var r = []; for (var i = a; i < b; i++) r.push(i); return r; }
  // 반마다 보관함이 있고, 반 인원보다 많은 칸은 사용하지 않는 칸입니다. (실제 앱에서는 '보관함' 시트에 적습니다)
  var CARTS = [
    { name: '3학년 1반', size: 28, used: 28, unused: range(28, 36) },
    { name: '3학년 2반', size: 30, used: 29, unused: [14].concat(range(30, 36)) }
  ];
  var MY = CARTS[0];
  // 예시 분석 결과: 시험용 사진(cart-sample-01)의 정답과 같음. 빈 칸 5, 8, 9, 17, 32번 (32번은 1반의 미사용 칸이라 무시됨)
  var SAMPLE = ['red','green','green','green','off','green','red','off','off','green','red','red',
                'red','green','red','red','off','green','green','green','red','green','red','red',
                'red','red','red','red','red','red','green','off','red','red','red','red'];
  function p2(n) { return String(n).padStart(2, '0'); }
  function ymd(d) { return d.getFullYear() + '-' + p2(d.getMonth() + 1) + '-' + p2(d.getDate()); }
  var today = new Date(), recent = [];
  var rs = ['수업 중 사용 중', '수리 맡김', '충전기 고장', '집에 가져감', '기타: 분실 의심'];
  var names = ['학생 A', '학생 B', '학생 C'];
  CARTS.forEach(function (cart, ci) {
    for (var k = 13; k >= 1; k--) {
      var d = new Date(today); d.setDate(d.getDate() - k);
      if (d.getDay() === 0 || d.getDay() === 6) continue;
      ['morning', 'afternoon'].forEach(function (per) {
        if (ci === 1 && k === 3 && per === 'morning') return;
        var slots = '', reasons = {};
        for (var i = 0; i < 36; i++) {
          if (cart.unused.indexOf(i) >= 0) { slots += 'X'; continue; }
          var off = (i === 6 && k % 2 === 0) || (i === 19 && k % 3 === 0) || i === 26 || (i === 7 && k < 5 && ci === 0);
          slots += off ? 'O' : (i % 2 ? 'G' : 'R');
          if (off) reasons[String(i + 1)] = rs[(i + k) % 5];
        }
        recent.push({ when: (d.getMonth() + 1) + '/' + d.getDate() + ' ' + (per === 'morning' ? '08:10' : '15:40'), date: ymd(d), period: per,
          cart: cart.name, email: 'a' + ((k + ci) % 3) + '@demo', by: names[(k + ci) % 3], slots: slots, reasons: reasons,
          edited: k % 4 === 0 ? 2 : 0, mode: 'AI+확인', coin: 1 + (k % 5) });
      });
    }
  });
  recent.sort(function (x, y) { return x.date < y.date ? -1 : x.date > y.date ? 1 : 0; });
  var students = [
    { email: 'a0@demo', name: '학생 A', duty: false, total: 14, cart: '3학년 1반', no: '1' },
    { email: 'a1@demo', name: '학생 B', duty: true, total: 9, cart: '3학년 2반', no: '4' },
    { email: 'a2@demo', name: '학생 C', duty: false, total: 11, cart: '3학년 2반', no: '7' },
    { email: 'me@demo', name: '시연 사용자', duty: true, total: 4, cart: '3학년 1반', no: '12' }
  ];
  var me = students[3], history = [{ when: '어제 15:40', coins: 4, kind: '획득', memo: '방과 후 점검' }];
  var done = {};   // 오늘 이미 코인이 나간 시간대

  function todayBy(cartName) {
    var st = { morning: null, afternoon: null }, t = ymd(today);
    recent.forEach(function (r) { if (r.cart === cartName && r.date === t && !st[r.period]) st[r.period] = { by: r.by, coin: r.coin > 0 }; });
    return st;
  }
  function drawCoins() {
    var t = [[1, 35], [2, 30], [3, 20], [4, 10], [5, 5]], x = Math.random() * 100;
    for (var i = 0; i < t.length; i++) { x -= t[i][1]; if (x < 0) return t[i][0]; }
    return 1;
  }
  var impl = {
    getInit: function () {
      return { email: 'me@demo', name: me.name, student: true, teacher: true, duty: true, cart: MY.name, unused: MY.unused, notice: '', reasons: REASONS,
        requireAi: true, status: todayBy(MY.name), coins: me.total, history: history.slice().reverse() };
    },
    analyze: function () {
      var slots = SAMPLE.slice();
      MY.unused.forEach(function (i) { slots[i] = 'unused'; });
      return { ok: true, analysisId: 'demo', slots: slots, uncertain: [8], note: '시연용 예시 결과입니다 (올린 사진과 관계없음)' };
    },
    submit: function (p) {
      var reasons = p.reasons || {}, empties = {};
      Object.keys(reasons).forEach(function (k) { empties[k] = reasons[k]; });
      var coins = done[p.period] ? 0 : drawCoins();
      done[p.period] = true;
      var now = new Date();
      recent.push({ when: (now.getMonth() + 1) + '/' + now.getDate() + ' ' + p2(now.getHours()) + ':' + p2(now.getMinutes()), date: ymd(now), period: p.period,
        cart: MY.name, email: 'me@demo', by: me.name, slots: p.slots.map(function (s) { return CH[s]; }).join(''), reasons: empties, edited: 0, mode: 'AI+확인', coin: coins });
      if (coins) { me.total += coins; history.push({ when: '방금', coins: coins, kind: '획득', memo: ymd(now) + ' 점검' }); }
      return { ok: true, coins: coins, why: coins ? '' : 'already', total: me.total };
    },
    teacherData: function () {
      var by = {}; CARTS.forEach(function (c) { by[c.name] = todayBy(c.name); });
      return { today: ymd(today), carts: CARTS, todayByCart: by, recent: JSON.parse(JSON.stringify(recent.slice(-600))),
        students: students.map(function (s) { return { email: s.email, name: s.name, duty: s.duty, total: s.total, cart: s.cart, no: s.no }; }) };
    },
    redeem: function (email, amount) {
      var s = students.filter(function (x) { return x.email === email; })[0], n = Math.floor(Number(amount));
      if (!s) throw new Error('명단에 없는 학생입니다.');
      if (!(n > 0)) throw new Error('사용할 코인 수를 1 이상으로 입력해 주세요.');
      if (s.total < n) throw new Error('코인이 부족합니다. 현재 ' + s.total + '개입니다.');
      s.total -= n; return { ok: true, total: s.total };
    }
  };
  window.google = { script: { run: new Proxy({}, { get: function (_, k) {
    if (k === 'withSuccessHandler') return function (ok) { return { withFailureHandler: function (fail) {
      return new Proxy({}, { get: function (_, fn) { return function () {
        var a = arguments;
        setTimeout(function () { try { ok(impl[fn].apply(null, a)); } catch (e) { fail(e); } }, fn === 'analyze' ? 1800 : 120);
      }; } });
    } }; };
  } }) } };
})();
</script>
"""

out = src.replace("<body>\n<main>", "<body>\n" + BANNER + "\n<main>", 1)
assert BANNER in out, "body/main 위치를 찾지 못했습니다"
marker = "<script>\n(function () {\n  var N = 36"
assert marker in out, "본 스크립트 위치를 찾지 못했습니다"
out = out.replace(marker, MOCK + marker, 1)
out = out.replace("<title>충전함 점검</title>", "<title>충전함 점검 (시연)</title>\n<meta name=\"robots\" content=\"noindex, nofollow\">", 1)

os.makedirs(os.path.join(root, "demo"), exist_ok=True)
with open(os.path.join(root, "demo", "index.html"), "w", encoding="utf8") as f:
    f.write(out)
print("demo/index.html", len(out), "bytes")
