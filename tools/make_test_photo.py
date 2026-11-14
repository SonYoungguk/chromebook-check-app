"""크롬북 충전함 시험용 사진을 만듭니다 (12칸 × 3줄 = 36칸).

사용법:  python3 tools/make_test_photo.py [시드] [빈 칸 수]
결과:    test-photos/cart-sample-<시드>.jpg  (사진)
         test-photos/cart-sample-<시드>.answer.json  (정답표)

표시등: 빨간불 = 충전 중, 초록불 = 충전 완료, 꺼짐 = 비어 있음
칸 번호: 윗줄 왼쪽→오른쪽 1~12, 가운데 줄 13~24, 아랫줄 25~36
실제 사진보다 단순합니다. 정확도는 실제 보관함 사진으로 따로 확인해야 합니다.
"""
import json, os, random, sys
from PIL import Image, ImageDraw, ImageFilter, ImageFont

seed = int(sys.argv[1]) if len(sys.argv) > 1 else 1
n_off = int(sys.argv[2]) if len(sys.argv) > 2 else 5
rnd = random.Random(seed)

W, H = 1600, 760
COLS, ROWS = 12, 3
pad_x, pad_top, pad_bot = 70, 70, 50
cw = (W - pad_x * 2) / COLS
rh = (H - pad_top - pad_bot) / ROWS

off = set(rnd.sample(range(36), n_off))
state = []
for i in range(36):
    state.append("off" if i in off else ("green" if rnd.random() < 0.5 else "red"))

# 몸체: 위에서 아래로 살짝 어두워지는 금속판
img = Image.new("RGB", (W, H))
px = ImageDraw.Draw(img)
for y in range(H):
    v = int(58 - 14 * y / H)
    px.line([(0, y), (W, y)], fill=(v, v + 2, v + 6))

try:
    font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 17)
except Exception:
    font = ImageFont.load_default()

glow = Image.new("RGB", (W, H), (0, 0, 0))
gd = ImageDraw.Draw(glow)
d = ImageDraw.Draw(img)
COL = {"red": (255, 59, 48), "green": (52, 209, 123), "off": (62, 66, 74)}

for i, st in enumerate(state):
    r, c = divmod(i, COLS)
    x0 = pad_x + c * cw
    y0 = pad_top + r * rh
    # 칸(어두운 홈)
    bx0, by0, bx1, by1 = x0 + 6, y0 + 4, x0 + cw - 6, y0 + rh - 34
    d.rounded_rectangle([bx0, by0, bx1, by1], 6, fill=(14, 16, 20))
    if st != "off":
        # 꽂힌 크롬북의 옆면: 은회색 판 + 윗면 하이라이트
        lx0, lx1 = bx0 + 9, bx1 - 9
        ly0, ly1 = by0 + (by1 - by0) * 0.22, by1 - 6
        for k in range(int(ly1 - ly0)):
            t = k / max(1, ly1 - ly0)
            g = int(150 - 40 * t) + rnd.randint(-3, 3)
            d.line([(lx0, ly0 + k), (lx1, ly0 + k)], fill=(g, g + 3, g + 8))
        d.line([(lx0, ly0), (lx1, ly0)], fill=(205, 210, 218), width=2)
    # 칸 번호
    d.text((x0 + 12, y0 + 8), str(i + 1), fill=(150, 156, 166), font=font)
    # 표시등
    cx, cy = x0 + cw / 2, y0 + rh - 16
    rad = 8
    if st != "off":
        gd.ellipse([cx - 22, cy - 22, cx + 22, cy + 22], fill=tuple(int(v * 0.55) for v in COL[st]))
    d.ellipse([cx - rad - 2, cy - rad - 2, cx + rad + 2, cy + rad + 2], fill=(20, 22, 26))
    d.ellipse([cx - rad, cy - rad, cx + rad, cy + rad], fill=COL[st])
    if st != "off":
        d.ellipse([cx - 3, cy - 4, cx + 1, cy], fill=tuple(min(255, v + 90) for v in COL[st]))

# 번짐(빛) 합성, 약한 흐림과 노이즈로 사진처럼
glow = glow.filter(ImageFilter.GaussianBlur(14))
img = Image.composite(img, img, Image.new("L", (W, H), 255))
from PIL import ImageChops
img = ImageChops.add(img, glow)
img = img.filter(ImageFilter.GaussianBlur(0.8))
noise = Image.effect_noise((W, H), 10).convert("RGB")
img = Image.blend(img, ImageChops.add(img, noise, scale=2, offset=-64), 0.35)

# 가장자리 어둡게
vig = Image.new("L", (W, H), 0)
vd = ImageDraw.Draw(vig)
for k in range(60):
    vd.rectangle([k, k, W - k, H - k], outline=int(70 * (1 - k / 60)))
img = ImageChops.subtract(img, Image.merge("RGB", (vig, vig, vig)).filter(ImageFilter.GaussianBlur(20)))

out_dir = os.path.join(os.path.dirname(__file__), "..", "test-photos")
os.makedirs(out_dir, exist_ok=True)
base = os.path.join(out_dir, "cart-sample-%02d" % seed)
img.save(base + ".jpg", quality=88)

KO = {"red": "충전 중", "green": "충전 완료", "off": "비어 있음"}
ans = {
    "설명": "윗줄 왼쪽에서 오른쪽으로 1~12, 가운데 줄 13~24, 아랫줄 25~36",
    "slots": state,
    "빈 칸 번호": [i + 1 for i, s in enumerate(state) if s == "off"],
    "요약": {KO[k]: state.count(k) for k in ("red", "green", "off")},
}
with open(base + ".answer.json", "w", encoding="utf8") as f:
    json.dump(ans, f, ensure_ascii=False, indent=2)
print(base + ".jpg", ans["요약"], ans["빈 칸 번호"])
