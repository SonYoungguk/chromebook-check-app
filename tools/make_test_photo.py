"""크롬북 충전함 시험용 사진을 만듭니다. 실제 보관함 구조를 흉내 냅니다.

구조: 3단 선반, 한 단에 11칸, 모두 33칸.
      각 단 위쪽 띠에 작은 둥근 LED가 한 줄로 있고, LED 아래에 칸 번호가 인쇄되어 있습니다.
      LED 아래쪽에 크롬북이 세로로 꽂혀 있고, 빈 칸은 케이블만 늘어져 있습니다.
LED:  초록 = 충전 완료, 빨강 = 충전 중, 꺼짐 = 빈 칸

사용법:  python3 tools/make_test_photo.py [시드] [빈 칸 수]
결과:    test-photos/cart-sample-<시드>.jpg  (사진)
         test-photos/cart-sample-<시드>.answer.json  (정답표)

실제 사진보다 훨씬 단순합니다. 정확도는 실제 보관함 사진으로 확인하세요.
"""
import json, os, random, sys
from PIL import Image, ImageDraw, ImageFilter, ImageFont, ImageChops

seed = int(sys.argv[1]) if len(sys.argv) > 1 else 1
n_off = int(sys.argv[2]) if len(sys.argv) > 2 else 6
rnd = random.Random(seed)

COLS, ROWS = 11, 3
N = COLS * ROWS
W, H = 900, 1560
left, right = 70, W - 70
cw = (right - left) / COLS
shelf_h = 500

off = set(rnd.sample(range(N), n_off))
state = ["off" if i in off else ("green" if rnd.random() < 0.8 else "red") for i in range(N)]

img = Image.new("RGB", (W, H), (214, 214, 208))
d = ImageDraw.Draw(img)
# 보관함 안쪽 어두운 면
d.rectangle([30, 30, W - 30, H - 30], fill=(190, 190, 184))

try:
    font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 20)
except Exception:
    font = ImageFont.load_default()

COL = {"red": (235, 70, 55), "green": (90, 235, 70), "off": (14, 14, 16)}
glow = Image.new("RGB", (W, H), (0, 0, 0))
gd = ImageDraw.Draw(glow)

for r in range(ROWS):
    y0 = 50 + r * shelf_h
    # LED 띠
    d.rounded_rectangle([60, y0, W - 60, y0 + 100], 6, fill=(232, 232, 226), outline=(170, 170, 164))
    d.text((W // 2 - 14, y0 + 8), "dlw", fill=(120, 120, 116), font=font)
    for c in range(COLS):
        i = r * COLS + c
        cx = left + (c + 0.5) * cw
        cy = y0 + 52
        d.ellipse([cx - 11, cy - 11, cx + 11, cy + 11], fill=(30, 30, 32))              # 안쪽으로 들어간 구멍
        if state[i] != "off":
            gd.ellipse([cx - 16, cy - 10, cx + 16, cy + 22], fill=tuple(int(v * 0.5) for v in COL[state[i]]))
            d.ellipse([cx - 6, cy - 2, cx + 6, cy + 9], fill=COL[state[i]])             # 구멍 아래쪽에 밝은 점
        else:
            d.ellipse([cx - 9, cy - 9, cx + 9, cy + 9], fill=COL["off"])
        label = str(i + 1)
        tw = d.textlength(label, font=font)
        d.text((cx - tw / 2, y0 + 70), label, fill=(30, 30, 30), font=font)
    # 칸 구분판과 크롬북
    top = y0 + 112
    for c in range(COLS):
        i = r * COLS + c
        x = left + c * cw
        d.rectangle([x + cw * 0.78, top + 40, x + cw * 0.94, top + 330], fill=(120, 124, 128))   # 회색 거치대
        if state[i] != "off":
            lx = x + cw * 0.22
            d.polygon([(lx, top + 60), (lx + cw * 0.5, top + 52), (lx + cw * 0.5, top + 330), (lx, top + 336)], fill=(228, 228, 226), outline=(150, 150, 150))
            d.line([(lx + 4, top + 60), (lx + 4, top + 336)], fill=(60, 60, 64), width=3)
            # 코일 케이블
            for k in range(7):
                d.arc([x + cw * 0.1, top + 4 + k * 5, x + cw * 0.5, top + 34 + k * 5], 20, 200, fill=(245, 245, 242), width=4)
        else:
            d.line([(x + cw * 0.5, top), (x + cw * 0.46, top + 70)], fill=(245, 245, 242), width=4)   # 늘어진 케이블
            d.line([(x + cw * 0.46, top + 70), (x + cw * 0.3, top + 100)], fill=(245, 245, 242), width=3)

glow = glow.filter(ImageFilter.GaussianBlur(7))
img = ImageChops.add(img, glow)
img = img.filter(ImageFilter.GaussianBlur(0.7))
noise = Image.effect_noise((W, H), 14).convert("RGB")
img = Image.blend(img, ImageChops.add(img, noise, scale=2, offset=-64), 0.3)

out_dir = os.path.join(os.path.dirname(__file__), "..", "test-photos")
os.makedirs(out_dir, exist_ok=True)
base = os.path.join(out_dir, "cart-sample-%02d" % seed)
img.save(base + ".jpg", quality=90)

KO = {"red": "충전 중", "green": "충전 완료", "off": "비어 있음"}
ans = {
    "설명": "3단 선반, 한 단에 11칸. 윗단 1~11, 가운데 단 12~22, 아랫단 23~33 (LED 아래 인쇄된 번호 기준)",
    "slots": state,
    "빈 칸 번호": [i + 1 for i, s in enumerate(state) if s == "off"],
    "요약": {KO[k]: state.count(k) for k in ("red", "green", "off")},
}
with open(base + ".answer.json", "w", encoding="utf8") as f:
    json.dump(ans, f, ensure_ascii=False, indent=2)
print(base + ".jpg", ans["요약"], ans["빈 칸 번호"])
