import io, os, re, sys, json, base64, hashlib, shutil, pathlib, html
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from PIL import Image, ImageOps

ROOT = pathlib.Path(".")
SITE = "https://fulllife.blog"
MODEL = os.environ.get("OPENAI_MODEL", "gpt-4o-mini")
# 含這些字的投影片（謄本、使用執照等）完全不處理：文字不送 AI、圖片不公開
SENSITIVE = ["使用執照", "謄本", "測量成果", "登記簿", "所有權狀", "身分證", "契約"]
MAX_CAND = 30

PROMPT = """你是工業不動產物件資料整理員。以下是一份物件簡報的文字（已排除謄本、使用執照等頁面）與候選圖片。
請只輸出 JSON，欄位如下：
title：物件標題，例如「仁武區 工業廠房 小坪數好用廠房出售」，不含公司名
type："售" 或 "租"
price：價格，照原文（如 "5,300萬（每坪約 32萬）"、"月租 45.35萬"、"160萬/坪"），不要自行計算
land_ping：土地坪數，只填數字字串；沒有則 ""
build_ping：建物坪數，只填數字字串；沒有則 ""
zoning：使用分區，如 "乙種工業區"、"丁種建築用地"、"住四(50%/300%)"
area：公開地址，只到「縣市＋區/鄉/鎮＋路/街/巷」，縣市必須依原文地址填寫（可能是屏東縣、台南市等，不要預設高雄），不得含門牌號碼、巷弄號、地號、段號
note：公開備註（面寬、深度、路寬、樓高、電力、天車、載重、完工日、結構、是否帶租約等客觀資料，用「、」分隔），沒有則 ""
desc：投資亮點／訴求重點，每點一句，用「；」分隔
photo_ids：候選圖片中屬於「實景照片」的編號（整數陣列），依適合展示的順序排列（外觀、空拍優先），最多 12 個
規則：
- 實景照片＝建物外觀、內部、空拍、周邊街景實拍。排除：地圖、地籍圖、平面圖、配置圖、證件文件、謄本、截圖、logo、橫幅、人像。
- 絕對不得寫入任何欄位：屋主姓名、同行出價、議價空間、「後台內容」等內部備註、電話、公司名稱。
- 原文沒有的欄位填空字串，不要猜測或編造。"""


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


def is_sensitive(lines):
    """標題式的短行含關鍵字（如「使用執照」「測量成果圖」）才視為證件頁；
    長句內文提到（如「依謄本登記為主」）不算。"""
    return any(len(t) <= 14 and any(k in t for k in SENSITIVE) for t in lines)


def read_pptx(path):
    prs = Presentation(path)
    texts, cands, seen = [], [], set()
    for i, s in enumerate(prs.slides, 1):
        lines = slide_lines(s)
        if is_sensitive(lines):
            continue
        texts.append(f"[第{i}頁]\n" + "\n".join(lines))
        for shp in walk(s.shapes):
            try:
                if shp.width / 914400 < 2.5 or shp.height / 914400 < 2.0:
                    continue  # 太小：logo、圖示
                blob = shp.image.blob
                h = hashlib.md5(blob).hexdigest()
                if h in seen:
                    continue
                seen.add(h)
                im = Image.open(io.BytesIO(blob))
                if im.width * im.height > 60_000_000:
                    continue  # 超大掃描檔，不是現場照片
                im.draft("RGB", (1024, 1024))
                im = ImageOps.exif_transpose(im).convert("RGB")
                im.thumbnail((512, 512))
                buf = io.BytesIO()
                im.save(buf, "JPEG", quality=70)
                cands.append({"page": i, "blob": blob,
                              "b64": base64.b64encode(buf.getvalue()).decode()})
            except Exception:
                continue
    return "\n\n".join(texts), cands[:MAX_CAND]


def ask_gpt(text, cands):
    from openai import OpenAI
    content = [{"type": "text", "text": PROMPT + "\n\n簡報文字：\n" + text}]
    for k, c in enumerate(cands):
        content.append({"type": "text", "text": f"圖片編號 {k}（第{c['page']}頁）"})
        content.append({"type": "image_url",
                        "image_url": {"url": "data:image/jpeg;base64," + c["b64"], "detail": "low"}})
    r = OpenAI().chat.completions.create(
        model=MODEL, temperature=0,
        response_format={"type": "json_object"},
        messages=[{"role": "user", "content": content}])
    return json.loads(r.choices[0].message.content)


def clean(d, cands):
    g = lambda k: str(d.get(k, "") or "").strip()
    if g("type") not in ("售", "租"):
        raise ValueError("無法判斷售/租：" + g("type"))
    area = re.sub(r"[\d０-９].*$", "", g("area"))  # 再保險：砍掉門牌與地號
    ids = [k for k in d.get("photo_ids", []) if isinstance(k, int) and 0 <= k < len(cands)]
    ids = list(dict.fromkeys(ids))[:12]
    return {"title": g("title"), "type": g("type"), "price": g("price"),
            "land_ping": g("land_ping"), "build_ping": g("build_ping"),
            "zoning": g("zoning"), "area": area, "note": g("note"), "desc": g("desc")}, ids


def save_photos(pid, cands, ids):
    out = ROOT / "img" / pid
    if out.exists():
        shutil.rmtree(out)
    out.mkdir(parents=True)
    paths = []
    for n, k in enumerate(ids, 1):
        im = ImageOps.exif_transpose(Image.open(io.BytesIO(cands[k]["blob"]))).convert("RGB")
        im.thumbnail((1600, 1600))
        im.save(out / f"{n}.jpg", "JPEG", quality=80, optimize=True)
        paths.append(f"img/{pid}/{n}.jpg")
    return paths


def write_share_pages(data):
    pdir = ROOT / "p"
    pdir.mkdir(exist_ok=True)
    for l in data:
        img = f"{SITE}/{l['photos'][0]}" if l.get("photos") else ""
        t = html.escape(l["title"] + "｜富住通大型工業地產")
        meta = [l["price"], l["area"], l["zoning"]]
        if l.get("land_ping"): meta.append(f"土地{l['land_ping']}坪")
        if l.get("build_ping"): meta.append(f"建坪{l['build_ping']}坪")
        d = html.escape("｜".join(x for x in meta if x) + "｜洽楊紘珉 0905-858-141")
        (pdir / f"{l['id']}.html").write_text(f"""<!DOCTYPE html>
<html lang="zh-Hant"><head><meta charset="utf-8">
<title>{t}</title>
<meta property="og:type" content="website">
<meta property="og:title" content="{t}">
<meta property="og:description" content="{d}">
<meta property="og:image" content="{img}">
<meta property="og:url" content="{SITE}/p/{l['id']}.html">
<meta name="twitter:card" content="summary_large_image">
<script>location.replace("../?id={l['id']}")</script>
</head><body><a href="../?id={l['id']}">查看物件：{t}</a></body></html>
""", "utf-8")


def main():
    f = ROOT / "listings.json"
    data = json.loads(f.read_text("utf-8")) if f.exists() else []
    data = [l for l in data if not l["title"].startswith("【範例】")]
    failed = 0
    for p in sorted((ROOT / "pptx").glob("*")):
        if p.suffix.lower() != ".pptx":
            continue
        try:
            src = re.sub(r"^\d{14}__", "", p.name)  # n8n 上傳時會加時間前綴避免同名衝突
            text, cands = read_pptx(p)
            d, ids = clean(ask_gpt(text, cands), cands)
            old = next((l for l in data if l.get("src") == src), None)
            if old:
                pid = old["id"]
            else:
                nums = [int(l["id"][4:]) for l in data if l["id"][4:].isdigit()]
                pid = "IND-%03d" % (max(nums, default=0) + 1)
            rec = {"id": pid, "src": src, **d,
                   "status": old["status"] if old else "待確認",
                   "photos": save_photos(pid, cands, ids)}
            if old:
                data[data.index(old)] = rec
            else:
                data.append(rec)
            print("OK", pid, p.name, len(ids), "photos")
            p.unlink()  # 成功才刪除；失敗會保留檔案方便重試
        except Exception as e:
            failed += 1
            print("FAIL", p.name, repr(e))
    f.write_text(json.dumps(data, ensure_ascii=False, indent=2), "utf-8")
    write_share_pages(data)
    if failed:
        sys.exit(1)


if __name__ == "__main__":
    main()
