// 코드를 앱스 스크립트에 올리고, 기존 웹 앱 배포를 새 버전으로 갱신합니다.
// 웹 앱 주소(QR코드)는 그대로 유지됩니다.
//   사용법: npm run release
const { spawnSync } = require('child_process');
const fs = require('fs');
const path = require('path');

const root = path.join(__dirname, '..');
const isWin = process.platform === 'win32';

function clasp(args) {
  // Windows에서는 shell을 거치므로 공백이 든 값은 따옴표로 감쌉니다.
  const a = isWin ? args.map((x) => (/\s/.test(x) ? '"' + x + '"' : x)) : args;
  const r = spawnSync('npx', ['--no-install', 'clasp'].concat(a), { cwd: root, stdio: 'inherit', shell: isWin });
  if (r.status !== 0) {
    console.error('\nclasp 명령이 실패했습니다: clasp ' + args.join(' ') +
      '\n(npm install 을 먼저 했는지, npm run login 으로 로그인했는지 확인해 주세요.)');
    process.exit(r.status || 1);
  }
}

if (!fs.existsSync(path.join(root, '.clasp.json'))) {
  console.error('.clasp.json 파일이 없습니다.\n' +
    '.clasp.json.example 을 복사해 .clasp.json 으로 만들고 scriptId 를 넣어 주세요.\n' +
    '(스크립트 ID: Apps Script 왼쪽 톱니바퀴 프로젝트 설정 → ID)');
  process.exit(1);
}

let id = (process.env.DEPLOYMENT_ID || '').trim();
const idFile = path.join(root, '.deployment-id');
if (!id && fs.existsSync(idFile)) id = fs.readFileSync(idFile, 'utf8').trim();

if (!id) {
  console.error('배포 ID가 없습니다. 아래 목록에서 웹 앱 배포의 ID를 .deployment-id 파일에 한 줄로 저장해 주세요.\n' +
    '(Apps Script의 배포 → 배포 관리 화면에서도 볼 수 있습니다. @HEAD 는 테스트용이라 쓰지 않습니다.)\n');
  clasp(['deployments']);
  process.exit(1);
}

const d = new Date();
const p2 = (n) => String(n).padStart(2, '0');
const desc = d.getFullYear() + '-' + p2(d.getMonth() + 1) + '-' + p2(d.getDate()) + ' ' + p2(d.getHours()) + ':' + p2(d.getMinutes()) + ' 배포';

console.log('1/2 코드를 올립니다.');
clasp(['push', '-f']);
console.log('\n2/2 기존 배포를 새 버전으로 갱신합니다.');
clasp(['deploy', '-i', id, '-d', desc]);
console.log('\n완료. 웹 앱 주소는 그대로입니다.');
