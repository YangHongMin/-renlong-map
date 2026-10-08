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
CATS = ["大型廠房", "小型廠房", "土地", "其他"]
DATA = ROOT / "data" / "listings.json"

PROMPT = """你是工業不動產物件資料整理員。以下是一份物件簡報的文字（已排除謄本、使用執照等頁面）與候選圖片。
請只輸出 JSON，欄位如下：
title：物件標題，例如「仁武區 工業廠房 小坪數好用廠房出售」，不含公司名
type："售" 或 "租"
price：價格，照原文（如 "5,300萬（每坪約 32萬）"、"月租 45.35萬"、"160萬/坪"），不要自行計算
land_ping：土地坪數，只填數字字串；沒有則 ""
build_ping：建物坪數，只填數字字串；沒有則 ""
zoning：使用分區，如 "乙種工業區"、"丁種建築用地"、"住四(50%/300%)"
area：公開地址，只到「縣市＋區/鄉/鎮＋路/街/巷」，縣市必須依原文地址填寫（可能是屏東縣、台南市等，不要預設高雄），不得含門牌號碼、巷弄號、地號、段號
category："大型廠房"（廠房/廠辦/倉庫，建坪約500坪以上）、"小型廠房"（建坪500坪以下的廠房/倉庫）、"土地"（無建物的土地）、"其他"（住宅、店面、辦公、車位等非廠房）擇一
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


_D = {
    "高雄市": "楠梓區 左營區 鼓山區 三民區 鹽埕區 前金區 新興區 苓雅區 前鎮區 旗津區 小港區 鳳山區 大寮區 鳥松區 林園區 仁武區 大樹區 大社區 岡山區 路竹區 橋頭區 梓官區 彌陀區 永安區 燕巢區 田寮區 阿蓮區 茄萣區 湖內區 旗山區 美濃區 內門區 杉林區 甲仙區 六龜區 茂林區 桃源區 那瑪夏區",
    "屏東縣": "屏東市 潮州鎮 東港鎮 恆春鎮 萬丹鄉 長治鄉 麟洛鄉 九如鄉 里港鄉 鹽埔鄉 高樹鄉 萬巒鄉 內埔鄉 竹田鄉 新埤鄉 枋寮鄉 新園鄉 崁頂鄉 林邊鄉 南州鄉 佳冬鄉 琉球鄉 車城鄉 滿州鄉 枋山鄉 三地門鄉 霧臺鄉 瑪家鄉 泰武鄉 來義鄉 春日鄉 獅子鄉 牡丹鄉",
    "台南市": "新營區 鹽水區 白河區 柳營區 後壁區 東山區 麻豆區 下營區 六甲區 官田區 大內區 佳里區 學甲區 西港區 七股區 將軍區 北門區 新化區 善化區 新市區 安定區 山上區 玉井區 楠西區 南化區 左鎮區 仁德區 歸仁區 關廟區 龍崎區 永康區 安平區 安南區 中西區",
}
DIST = {d: c for c, v in _D.items() for d in v.split()}


def norm_area(a):
    """統一成「縣市＋區/鄉/鎮＋路名」。縣市以區名對照表為準（修正 AI 填錯縣市）；
    找不到區名就維持原樣，不亂改。"""
    a = re.sub(r"\s+", "", str(a or "")).replace("臺", "台")
    a = re.sub(r"[\d０-９].*$", "", a)  # 門牌、地號一律砍掉
    county = dist = rest = None
    for d in sorted(DIST, key=len, reverse=True):
        if d in a:
            county, dist, rest = DIST[d], d, a[a.rindex(d) + len(d):]
            break
    if not dist:
        m = re.match(r"(.{1,3}[縣市])(.{1,4}?[區鄉鎮市])(.*)$", a)
        if not m:
            return a
        county, dist, rest = m.groups()
    rest = re.sub(r"^[^路街巷]{1,4}?[村里]", "", rest)  # 村里不顯示
    m = re.match(r"(.*?(?:大道|路|街))((?:[一二三四五六七八九十]{1,3}段)?)", rest)
    road = (m.group(1) + m.group(2)) if m else ""
    if not road:
        m = re.match(r"(.{1,6}巷)", rest)  # 只有巷名時保留；地段（如大同段）不顯示
        road = m.group(1) if m else ""
    if len(road) > 14:
        road = ""
    return county + dist + road


def clean(d, cands):
    g = lambda k: str(d.get(k, "") or "").strip()
    if g("type") not in ("售", "租"):
        raise ValueError("無法判斷售/租：" + g("type"))
    area = norm_area(g("area"))  # 統一成 縣市＋區鄉鎮＋路名
    ids = [k for k in d.get("photo_ids", []) if isinstance(k, int) and 0 <= k < len(cands)]
    ids = list(dict.fromkeys(ids))[:12]
    return {"title": g("title"), "type": g("type"), "price": g("price"),
            "land_ping": g("land_ping"), "build_ping": g("build_ping"),
            "zoning": g("zoning"), "area": area,
            "category": g("category") if g("category") in CATS else "", "note": g("note"), "desc": g("desc")}, ids


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


def cat_of(l):
    if l.get("category") in CATS:
        return l["category"]
    num = lambda k: float(re.sub(r"[^\d.]", "", str(l.get(k, ""))) or 0)
    b, d = num("build_ping"), num("land_ping")
    if not b:
        return "土地" if d else "其他"
    if not re.search("廠|倉|工業|丁種|乙種|甲種", l.get("title", "") + l.get("zoning", "")):
        return "其他"
    return "大型廠房" if b >= 500 else "小型廠房"


def jdump(o):
    return json.dumps(o, ensure_ascii=False).replace("</", "<\\/")


def page_html(l):
    E = html.escape
    ph = [p for p in (l.get("photos") or []) if re.fullmatch(r"img/[\w\-/.]+", str(p)) and ".." not in p]
    t = l["title"] + "｜富住通大型工業地產"
    meta = [l.get("price", ""), l.get("area", ""), l.get("zoning", "")]
    if l.get("land_ping"): meta.append(f"土地{l['land_ping']}坪")
    if l.get("build_ping"): meta.append(f"建坪{l['build_ping']}坪")
    desc = "｜".join(x for x in meta if x) + "｜洽楊紘珉 0905-858-141"
    url = f"{SITE}/p/{l['id']}.html"
    m = re.search(r"[市縣](.+?[區鄉鎮市])", l.get("area", ""))
    ld = {"@context": "https://schema.org", "@type": "RealEstateListing", "name": t, "url": url,
          "description": desc + ("。" + l["desc"].replace("；", "，") if l.get("desc") else ""),
          "image": [f"{SITE}/{p}" for p in ph[:6]],
          "about": {"@type": "Place", "name": l.get("area", ""),
                    "address": {"@type": "PostalAddress", "addressCountry": "TW",
                                "addressLocality": m.group(1) if m else "", "streetAddress": l.get("area", "")}},
          "category": cat_of(l) + ("出租" if l.get("type") == "租" else "出售"),
          "provider": {"@type": "RealEstateAgent", "name": "富住通商用不動產 新興店 楊紘珉",
                       "telephone": "+886-905-858-141", "url": SITE}}
    rows = [("類別", cat_of(l)), ("編號", l["id"]), ("區域", l.get("area", "")), ("價格", l.get("price", "")),
            ("使用分區", l.get("zoning", "")), ("土地坪數", l.get("land_ping", "")),
            ("建物坪數", l.get("build_ping", "")), ("備註", l.get("note", ""))]
    tr = "".join(f"<tr><th>{E(k)}</th><td>{E(str(v))}</td></tr>" for k, v in rows if v)
    pts = "".join(f"<li>{E(x)}</li>" for x in (l.get("desc") or "").split("；") if x)
    imgs = "".join(f'<img src="../{E(p)}" alt="{E(l["title"])}" width="400" loading="lazy">' for p in ph[:4])
    og = E(f"{SITE}/{ph[0]}") if ph else ""
    return f"""<!DOCTYPE html>
<html lang="zh-Hant"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{E(t)}</title>
<meta name="description" content="{E(desc)}">
<link rel="canonical" href="{url}">
<meta property="og:type" content="website">
<meta property="og:title" content="{E(t)}">
<meta property="og:description" content="{E(desc)}">
<meta property="og:image" content="{og}">
<meta property="og:url" content="{url}">
<meta name="twitter:card" content="summary_large_image">
<script type="application/ld+json">{jdump(ld)}</script>
<script>location.replace("../?id={l['id']}")</script>
</head><body>
<h1>{E(l['title'])}</h1>
<p>{imgs}</p>
<table>{tr}</table>
<ul>{pts}</ul>
<p>洽詢：楊紘珉（富住通商用不動產 新興店）0905-858-141｜<a href="https://lin.ee/S6hfHqge">LINE 諮詢</a></p>
<p><a href="../?id={l['id']}">查看完整物件頁</a>｜<a href="../all.html">全部物件清單</a></p>
</body></html>
"""


def write_share_pages(data):
    pdir = ROOT / "p"
    pdir.mkdir(exist_ok=True)
    pub = [l for l in data
           if l.get("status") != "待確認" and re.fullmatch(r"[A-Za-z0-9_-]{1,40}", str(l.get("id", "")))]
    live = set()
    for l in pub:
        live.add(f"{l['id']}.html")
        (pdir / f"{l['id']}.html").write_text(page_html(l), "utf-8")
    # 清掉不再公開的舊分享頁；資料為空時不動（防止 listings.json 被清空時誤刪）
    if data:
        for f in pdir.glob("*.html"):
            if f.name not in live:
                f.unlink()
        write_site_files(pub)


def write_site_files(pub):
    E = html.escape
    live = [l for l in pub if l.get("status") == "上架"]
    sec = ""
    for c in CATS:
        items = [l for l in live if cat_of(l) == c]
        if items:
            sec += f"<h2>{c}</h2><ul>" + "".join(
                f'<li><a href="p/{l["id"]}.html">{E(l["title"])}</a>｜{E(l.get("area",""))}｜{E(l.get("price",""))}</li>'
                for l in items) + "</ul>"
    (ROOT / "all.html").write_text(f"""<!DOCTYPE html>
<html lang="zh-Hant"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>高雄工業廠房・工業用地出售出租物件總覽｜富住通</title>
<meta name="description" content="富住通商用不動產 新興店 楊紘珉，高雄、仁武、大寮、岡山等地工業廠房與工業用地出售、出租物件清單。">
<link rel="canonical" href="{SITE}/all.html">
</head><body><h1>工業廠房・工業用地物件總覽</h1>{sec or "<p>目前沒有上架物件</p>"}
<p><a href="./">回首頁</a>｜洽詢：楊紘珉 0905-858-141｜<a href="https://lin.ee/S6hfHqge">LINE 諮詢</a></p></body></html>
""", "utf-8")
    urls = [f"{SITE}/", f"{SITE}/all.html"] + [f"{SITE}/p/{l['id']}.html" for l in live]
    (ROOT / "sitemap.xml").write_text('<?xml version="1.0" encoding="UTF-8"?>\n<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n'
        + "".join(f"<url><loc>{E(u)}</loc></url>\n" for u in urls) + "</urlset>\n", "utf-8")
    (ROOT / "robots.txt").write_text(f"User-agent: *\nAllow: /\nDisallow: /admin.html\n\nSitemap: {SITE}/sitemap.xml\n", "utf-8")


def main():
    old_f = ROOT / "listings.json"  # 舊位置：搬遷前相容
    f = DATA
    f.parent.mkdir(exist_ok=True)
    src_f = f if f.exists() else old_f
    data = json.loads(src_f.read_text("utf-8")) if src_f.exists() else []
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
    for l in data:  # 統一地址格式（已符合的不會變動）
        n = norm_area(l.get("area", ""))
        if n and n != l.get("area"):
            print("AREA", l["id"], l.get("area"), "->", n)
            l["area"] = n
    f.write_text(json.dumps(data, ensure_ascii=False, indent=2), "utf-8")
    if not f.exists() or old_f.exists():
        old_f.unlink(missing_ok=True)
    write_share_pages(data)
    if failed:
        sys.exit(1)


if __name__ == "__main__":
    main()
