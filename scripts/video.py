"""物件短片：用物件照片合成 9:16 直式短影音（每張照片緩慢推近＋字幕＋AI 旁白＋片尾 LINE）。

用法：
  python scripts/video.py            # 產生所有「上架」但還沒有影片（或資料有變）的物件
  python scripts/video.py IND-013    # 只做指定物件（會重做）
輸出：video/IND-xxx.mp4，紀錄在 data/video.json（內容指紋，資料沒變就不重做）
旁白：有 OPENAI_API_KEY 才會產生；沒有就做無聲版本。
"""
import hashlib, json, os, re, subprocess, sys, tempfile, pathlib, urllib.request

sys.path.insert(0, str(pathlib.Path(__file__).parent))
import build as B  # 共用價格換算、地區、類別等規則
from PIL import Image, ImageDraw, ImageFilter, ImageOps

ROOT = pathlib.Path(".")
OUT = ROOT / "video"
MANIFEST = ROOT / "data" / "video.json"
W, H, FPS = 1080, 1920, 30
MAX_LEN = 29.5          # 影片總長上限（秒）
TEMPO, MAX_TEMPO = 1.2, 1.35  # 旁白加速倍率（不變音調）；超過 30 秒會再加快或少一個亮點
NAVY, RED, ORANGE, WHITE = (30, 58, 138), (200, 48, 42), (255, 196, 120), (255, 255, 255)
VOICE_MODEL, VOICE = "gpt-4o-mini-tts", "onyx"
VOICE_STYLE = "用台灣口音的華語，像專業的工業不動產顧問在介紹物件，語氣沉穩、清楚、略帶親切，速度適中。"
VERSION = "v2"  # v2：語速加快、總長控制在 30 秒內  # 改版面時加一，會重做全部影片


def sh(*args):
    subprocess.run(args, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.PIPE)


def dur_of(path):
    r = subprocess.run(["ffprobe", "-v", "error", "-show_entries", "format=duration", "-of", "csv=p=0", str(path)],
                       capture_output=True, text=True)
    try:
        return float(r.stdout.strip())
    except ValueError:
        return 0.0


def tts(text, path):
    key = os.environ.get("OPENAI_API_KEY")
    if not key:
        return False
    for model in (VOICE_MODEL, "tts-1"):
        body = {"model": model, "voice": VOICE, "input": text, "response_format": "mp3"}
        if model != "tts-1":
            body["instructions"] = VOICE_STYLE
        req = urllib.request.Request("https://api.openai.com/v1/audio/speech", data=json.dumps(body).encode(),
                                     headers={"Authorization": "Bearer " + key, "Content-Type": "application/json"})
        try:
            with urllib.request.urlopen(req, timeout=90) as r:
                path.write_bytes(r.read())
            return True
        except Exception as e:
            print("TTS", model, "失敗", repr(e))
    return False


# ---------- 文案 ----------
def short_price(l):
    p = B.price_disp(l)
    p = re.split(r"[；;]", p)[0]
    p = re.sub(r"\s*[（(].*$", "", p).strip()
    if l.get("type") == "租" and "租" not in p:
        p = "月租 " + p
    elif l.get("type") != "租" and not p.startswith("總價"):
        p = "總價 " + p
    return p


def zone_short(z):
    z = re.split(r"[／/、,，(（]", z or "")[0].strip()
    return z if z and z not in ("工業區",) else ""


def lines_of(l):
    c, d = B.loc(l)
    dist = d or c
    land = B._num(l.get("land_ping"))
    bld = B._num(l.get("build_ping"))
    kind = B.kind_of(l)
    deal = "出租" if l.get("type") == "租" else "出售"
    big = f"{land:,.0f}坪" if land else (f"建物{bld:,.0f}坪" if bld else "")
    z0 = zone_short(l.get("zoning"))
    tag = {"丁種建築用地": "丁建", "乙種工業區": "乙工", "甲種工業區": "甲工", "甲種建築用地": "甲建"}.get(z0, "")
    hook = f"{dist}\n{big}{tag}{kind}{deal}".strip()
    size = "、".join(x for x in [f"土地 {land:,.0f} 坪" if land else "", f"建物 {bld:,.0f} 坪" if bld else ""] if x)
    z = zone_short(l.get("zoning"))
    pts = []
    for p in re.split(r"[；;]", l.get("desc") or ""):
        p = re.sub(r"\s+", "", p).strip("，,。 ")
        if not p or re.search(r"\d+號|地號", p):
            continue
        pts.append(p[:30])
    segs = [("hook", hook, f"{dist}，{big}{tag and z0}{kind}{deal}。"),
            ("price", short_price(l), short_price(l).replace("約", "大約") + "。")]
    if size:
        segs.append(("size", size + (f"\n{z}" if z else ""), size.replace("、", "，") + (f"，{z}" if z else "") + "。"))
    for p in pts[:3]:
        segs.append(("point", p, p + "。"))
    segs.append(("outro", "", "想看更多照片和詳細資料，歡迎加 LINE 官方帳號，或直接來電楊紘珉。"))
    return segs


# ---------- 畫面 ----------
def cover(im, w, h):
    return ImageOps.fit(im.convert("RGB"), (w, h), Image.LANCZOS, centering=(0.5, 0.5))


def wrap(draw, text, font, width):
    out = []
    for para in text.split("\n"):
        cur = ""
        for ch in para:
            if draw.textlength(cur + ch, font=font) > width and cur:
                out.append(cur)
                cur = ch
            else:
                cur += ch
        if cur:
            out.append(cur)
    return out


def overlay_png(kind, text, l, path):
    im = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    if kind == "hook":
        d.rectangle([0, 0, W, 520], fill=(14, 28, 70, 215))
        d.rectangle([0, 0, W, 16], fill=RED)
        d.text((70, 90), "富住通・大型工業地產", font=B._font(44), fill=ORANGE)
        y = 170
        for ln in wrap(d, text, B._font(104), W - 140)[:3]:
            d.text((70, y), ln, font=B._font(104), fill=WHITE)
            y += 130
    else:
        f = B._font(64)
        lines = wrap(d, text, f, W - 160)[:3]
        bh = 70 + 88 * len(lines)
        top = H - 420 - bh
        d.rounded_rectangle([50, top, W - 50, top + bh], radius=28, fill=(14, 28, 70, 205))
        if kind == "price":
            d.rounded_rectangle([50, top, 66, top + bh], radius=8, fill=ORANGE)
        y = top + 36
        for ln in lines:
            d.text((90, y), ln, font=f, fill=WHITE)
            y += 88
        # 角落：物件編號
        d.text((60, 60), l["id"], font=B._font(36, bold=False), fill=(255, 255, 255, 200))
    im.save(path)


def outro_png(l, path):
    im = Image.new("RGB", (W, H), NAVY)
    d = ImageDraw.Draw(im)
    d.rectangle([0, 0, W, 18], fill=RED)
    try:
        logo = Image.open(ROOT / "logo.png").convert("RGBA")
        logo.thumbnail((900, 110))
        im.paste(logo, ((W - logo.width) // 2, 260), logo)
    except Exception:
        pass
    y = 520
    for t, f, c in [("想看更多照片與詳細資料", B._font(66), WHITE), ("歡迎加 LINE 官方帳號", B._font(56, bold=False), (214, 222, 240)),
                    ("@447lrpzt", B._font(96), ORANGE), ("或來電 0905-858-141", B._font(60), WHITE)]:
        w = d.textlength(t, font=f)
        d.text(((W - w) / 2, y), t, font=f, fill=c)
        y += f.size + 60
    try:
        qr = Image.open(ROOT / "line_qr.png").convert("RGB")
        qr = qr.resize((360, 360))
        im.paste(qr, ((W - 360) // 2, y + 20))
        y += 420
    except Exception:
        pass
    t = "富住通商用不動產 新興店｜楊紘珉"
    f = B._font(40, bold=False)
    d.text(((W - d.textlength(t, font=f)) / 2, H - 160), t, font=f, fill=(214, 222, 240))
    t = "物件編號 " + l["id"] + "｜fulllife.blog"
    d.text(((W - d.textlength(t, font=f)) / 2, H - 100), t, font=f, fill=(214, 222, 240))
    im.save(path)


def photo_frame(src, path):
    """照片鋪滿 1080x1920：背景放大模糊，前景完整照片置中（橫式照片不會被裁太多）"""
    im = ImageOps.exif_transpose(Image.open(src)).convert("RGB")
    bg = cover(im, W, H).filter(ImageFilter.GaussianBlur(40))
    bg = Image.eval(bg, lambda v: int(v * 0.55))
    fg = im.copy()
    fg.thumbnail((W, int(H * 0.62)), Image.LANCZOS)
    if fg.width < W:  # 橫式照片放大到滿寬
        r = W / fg.width
        fg = im.resize((W, int(im.height * W / im.width)), Image.LANCZOS)
        if fg.height > int(H * 0.62):
            fg = cover(fg, W, int(H * 0.62))
    bg.paste(fg, (0, (H - fg.height) // 2 - 120))
    bg.save(path, quality=92)


def segment(still, over, dur, out, zoom=True):
    frames = max(int(dur * FPS), 1)
    z = f"zoompan=z='min(1+0.0009*on,1.12)':x='iw/2-(iw/zoom/2)':y='ih/2-(ih/zoom/2)':d={frames}:s={W}x{H}:fps={FPS}" if zoom \
        else f"zoompan=z=1:d={frames}:s={W}x{H}:fps={FPS}"
    args = ["ffmpeg", "-y", "-loop", "1", "-i", str(still)]
    if over:
        args += ["-i", str(over), "-filter_complex", f"[0:v]scale={W*2}:{H*2},{z}[b];[b][1:v]overlay=0:0,format=yuv420p[v]", "-map", "[v]"]
    else:
        args += ["-vf", f"scale={W*2}:{H*2},{z},format=yuv420p"]
    args += ["-t", f"{dur:.3f}", "-r", str(FPS), "-c:v", "libx264", "-preset", "medium", "-crf", "24", "-an", str(out)]
    sh(*args)


def make_video(l, out_path):
    photos = [ROOT / p for p in (l.get("photos") or []) if (ROOT / p).is_file()]
    if not photos:
        print("SKIP", l["id"], "沒有照片")
        return False
    segs = lines_of(l)
    with tempfile.TemporaryDirectory() as td:
        td = pathlib.Path(td)
        raw = {}
        for i, (kind, text, speech) in enumerate(segs):
            a = td / f"a{i}.mp3"
            raw[i] = dur_of(a) if tts(speech, a) else 0.0
        voiced = any(raw.values())

        def plan(idx, tempo):
            return {i: max(raw[i] / tempo + 0.2 if raw[i] else 0, 3.0 if segs[i][0] == "outro" else 2.2) for i in idx}

        idx = list(range(len(segs)))
        tempo = TEMPO
        while True:
            d = plan(idx, tempo)
            if sum(d.values()) <= MAX_LEN:
                break
            pts = [i for i in idx if segs[i][0] == "point"]
            if tempo < MAX_TEMPO:
                tempo = min(MAX_TEMPO, tempo + 0.05)
            elif len(pts) > 1:
                idx.remove(pts[-1])  # 拿掉最後一個亮點
                tempo = TEMPO
            else:
                break
        parts, auds = [], []
        for n, i in enumerate(idx):
            kind, text, speech = segs[i]
            dur = d[i]
            if kind == "outro":
                still = td / f"s{i}.png"
                outro_png(l, still)
                over = None
            else:
                still = td / f"s{i}.jpg"
                photo_frame(photos[n % len(photos)] if kind != "hook" else photos[0], still)
                over = td / f"o{i}.png"
                overlay_png(kind, text, l, over)
            v = td / f"v{i}.mp4"
            segment(still, over, dur, v, zoom=kind != "outro")
            parts.append(v)
            w = td / f"w{i}.wav"
            if raw[i]:
                sh("ffmpeg", "-y", "-i", str(td / f"a{i}.mp3"), "-af", f"atempo={tempo:.2f},apad", "-t", f"{dur:.3f}",
                   "-ar", "44100", "-ac", "2", str(w))
            else:
                sh("ffmpeg", "-y", "-f", "lavfi", "-i", "anullsrc=r=44100:cl=stereo", "-t", f"{dur:.3f}", str(w))
            auds.append(w)
        (td / "v.txt").write_text("".join(f"file '{p}'\n" for p in parts))
        (td / "a.txt").write_text("".join(f"file '{p}'\n" for p in auds))
        sh("ffmpeg", "-y", "-f", "concat", "-safe", "0", "-i", str(td / "v.txt"), "-c", "copy", str(td / "v.mp4"))
        sh("ffmpeg", "-y", "-f", "concat", "-safe", "0", "-i", str(td / "a.txt"), "-c:a", "pcm_s16le", str(td / "a.wav"))
        out_path.parent.mkdir(exist_ok=True)
        sh("ffmpeg", "-y", "-i", str(td / "v.mp4"), "-i", str(td / "a.wav"), "-c:v", "copy", "-c:a", "aac", "-b:a", "128k",
           "-shortest", "-movflags", "+faststart", str(out_path))
    print("VIDEO", l["id"], f"{dur_of(out_path):.1f}s", f"{out_path.stat().st_size / 1e6:.1f}MB",
          f"旁白 x{tempo:.2f}" if voiced else "無聲", f"{sum(1 for i in idx if segs[i][0] == 'point')} 個亮點")
    return voiced


def fingerprint(l):
    keys = ["title", "type", "price", "land_ping", "build_ping", "zoning", "area", "desc", "photos", "category"]
    return hashlib.md5((VERSION + json.dumps({k: l.get(k) for k in keys}, ensure_ascii=False, sort_keys=True)).encode()).hexdigest()[:12]


def main():
    data = json.loads(B.DATA.read_text("utf-8"))
    man = json.loads(MANIFEST.read_text("utf-8")) if MANIFEST.exists() else {}
    only = sys.argv[1:]
    live = [l for l in data if l.get("status") == "上架" and (not only or l["id"] in only)]
    for l in live:
        fp = fingerprint(l)
        old = man.get(l["id"], {})
        if not only and old.get("fp") == fp and (OUT / f"{l['id']}.mp4").exists() and (old.get("voiced") or not os.environ.get("OPENAI_API_KEY")):
            continue
        try:
            voiced = make_video(l, OUT / f"{l['id']}.mp4")
            man[l["id"]] = {"fp": fp, "voiced": bool(voiced), "path": f"video/{l['id']}.mp4"}
        except subprocess.CalledProcessError as e:
            print("FAIL", l["id"], e.stderr.decode(errors="replace")[-400:])
    # 下架的物件：刪影片
    keep = {l["id"] for l in data if l.get("status") == "上架"}
    if not only:
        for k in list(man):
            if k not in keep:
                (OUT / f"{k}.mp4").unlink(missing_ok=True)
                man.pop(k)
    MANIFEST.write_text(json.dumps(man, ensure_ascii=False, indent=1), "utf-8")


if __name__ == "__main__":
    main()
