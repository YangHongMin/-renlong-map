import io, re, json, hashlib, shutil, pathlib
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from PIL import Image, ImageOps

ROOT = pathlib.Path(".")
CITY = "高雄市"


def walk(shapes):
    for s in shapes:
        if s.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from walk(s.shapes)
        else:
            yield s


def slide_lines(slide):
    out = []
    for s in walk(slide.shapes):
        if s.has_text_frame:
            for p in s.text_frame.paragraphs:
                t = "".join(r.text for r in p.runs).strip()
                if t:
                    out.append(t)
    return out


def is_boiler(t, minlen=4):
    return "富住通" in t or "關懷" in t or len(t) < minlen


def photos_of(slide, seen, min_w=2.5, min_h=2.0):
    """只取尺寸夠大的圖（排除 logo、圖示、橫幅）"""
    res = []
    for s in walk(slide.shapes):
        try:
            if s.width / 914400 < min_w or s.height / 914400 < min_h:
                continue
            blob = s.image.blob
        except Exception:
            continue
        h = hashlib.md5(blob).hexdigest()
        if h in seen:
            continue
        seen.add(h)
        res.append(blob)
    return res


def parse(path):
    prs = Presentation(path)
    slides = list(prs.slides)
    texts = [slide_lines(s) for s in slides]
    alltext = "\n".join("\n".join(t) for t in texts)

    def find(pat):
        m = re.search(pat, alltext)
        return m.group(1).strip() if m else ""

    d = {}
    # 標題：封面非公司字樣的文字
    d["title"] = " ".join(t for t in texts[0] if not is_boiler(t, 2)) if texts else path.stem
    # 售 / 租與價格
    if "售價" in alltext:
        d["type"] = "售"
        d["price"] = find(r"售價[：:]\s*([\d,\.]+\s*萬)")
        u = find(r"售價[：:][^（\n]*（每坪約\s*([\d\.]+\s*萬)")
        if u:
            d["price"] += f"（每坪約 {u}）"
    else:
        d["type"] = "租"
        d["price"] = "月租 " + find(r"月租金[：:]\s*([\d,\.]+\s*萬)")
        u = find(r"月租金[：:][^（\n]*（每坪約\s*([\d\.]+\s*元)")
        if u:
            d["price"] += f"（每坪約 {u}）"
    d["land_ping"] = find(r"基地面積[：:]\s*約?\s*([\d\.]+)")
    d["build_ping"] = find(r"建物面積[：:]\s*約?\s*([\d\.]+)")
    d["zoning"] = find(r"使用分區[：:]\s*(\S+)")
    # 地址只保留到路/街/巷，不含門牌
    m = re.search(r"([\u4e00-\u9fa5]{1,4}[區鄉鎮市][\u4e00-\u9fa5]{1,8}?[路街巷道段])[\d0-9一二三四五六七八九十]", alltext)
    d["area"] = CITY + m.group(1) if m else CITY
    # 備註與尺寸資訊
    notes = []
    n = find(r"備註[：:]([^\n]*)")
    if n:
        notes.append(n)
    road = False
    for l in sum(texts, []):
        if re.match(r"^(面寬|深度)\s*\d", l):
            notes.append(l)
        elif "路寬" in l and not road:
            road, _ = True, notes.append(re.sub(r"\s+", "", l))
    d["note"] = "、".join(notes)
    # 投資亮點
    hl = []
    for t in texts:
        if any("投資亮點" in l for l in t):
            hl = [l for l in t if not is_boiler(l) and "投資亮點" not in l]
    d["desc"] = "；".join(re.sub(r"^\d+[\.、]\s*", "", l) for l in hl)

    # 照片：封面 + 標題含「現況拍攝」的投影片
    seen, blobs = set(), []
    for i, (s, t) in enumerate(zip(slides, texts)):
        if i == 0 or any("現況拍攝" in l for l in t):
            blobs += photos_of(s, seen)
    return d, blobs


def save_photos(pid, blobs):
    out = ROOT / "img" / pid
    if out.exists():
        shutil.rmtree(out)
    out.mkdir(parents=True)
    paths = []
    for i, b in enumerate(blobs, 1):
        im = ImageOps.exif_transpose(Image.open(io.BytesIO(b))).convert("RGB")
        im.thumbnail((1600, 1600))
        im.save(out / f"{i}.jpg", "JPEG", quality=80, optimize=True)
        paths.append(f"img/{pid}/{i}.jpg")
    return paths


f = ROOT / "listings.json"
data = json.loads(f.read_text("utf-8")) if f.exists() else []
data = [l for l in data if not l["title"].startswith("【範例】")]

for p in sorted((ROOT / "pptx").glob("*")):
    if p.suffix.lower() != ".pptx":
        continue
    d, blobs = parse(p)
    old = next((l for l in data if l.get("src") == p.name), None)
    if old:
        pid = old["id"]
    else:
        nums = [int(l["id"][4:]) for l in data if l["id"][4:].isdigit()]
        pid = "IND-%03d" % (max(nums, default=0) + 1)
    rec = {"id": pid, "src": p.name, **d,
           "status": old["status"] if old else "上架",
           "photos": save_photos(pid, blobs)}
    if old:
        data[data.index(old)] = rec
    else:
        data.append(rec)
    print(pid, p.name, len(blobs), "photos")
    p.unlink()  # 處理完刪除，避免 PPTX 被公開下載

f.write_text(json.dumps(data, ensure_ascii=False, indent=2), "utf-8")


# 每個物件產生一個分享頁（讓 FB / LINE 貼連結時顯示該物件的照片與標題）
import html
SITE = "https://fulllife.blog"
pdir = ROOT / "p"
pdir.mkdir(exist_ok=True)
for l in data:
    img = f"{SITE}/{l['photos'][0]}" if l.get("photos") else ""
    t = html.escape(l["title"] + "｜富住通大型工業地產")
    d = html.escape(f"{l['price']}｜{l['area']}｜{l['zoning']}｜土地{l['land_ping']}坪 建坪{l['build_ping']}坪｜洽楊紘珉 0905-858-141")
    u = f"{SITE}/p/{l['id']}.html"
    (pdir / f"{l['id']}.html").write_text(f"""<!DOCTYPE html>
<html lang="zh-Hant"><head><meta charset="utf-8">
<title>{t}</title>
<meta property="og:type" content="website">
<meta property="og:title" content="{t}">
<meta property="og:description" content="{d}">
<meta property="og:image" content="{img}">
<meta property="og:url" content="{u}">
<meta name="twitter:card" content="summary_large_image">
<script>location.replace("../?id={l['id']}")</script>
</head><body><a href="../?id={l['id']}">查看物件：{t}</a></body></html>
""", "utf-8")
