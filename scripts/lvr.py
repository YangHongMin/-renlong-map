"""下載內政部實價登錄開放資料（買賣），篩出高雄市工業相關成交，存成 data/lvr/kaohsiung_industrial.json。

資料來源：內政部不動產成交案件實際資訊（政府資料開放授權條款第 1 版）。
GitHub Actions 每 10 天執行一次（實價登錄每月 1、11、21 日發布）。
"""
import csv, io, json, re, sys, time, zipfile, datetime, pathlib, urllib.request

ROOT = pathlib.Path(".")
OUT = ROOT / "data" / "lvr" / "kaohsiung_industrial.json"
CITY = "e"  # 高雄市
BASE = "https://plvr.land.moi.gov.tw"
UA = {"User-Agent": "Mozilla/5.0 (fulllife.blog open-data fetcher)"}
QUARTERS = 8  # 往回抓 8 季（約 2 年）


def get(url, tries=3):
    for i in range(tries):
        try:
            req = urllib.request.Request(url, headers=UA)
            with urllib.request.urlopen(req, timeout=120) as r:
                return r.read()
        except Exception as e:
            print("retry", url, repr(e))
            time.sleep(5 * (i + 1))
    return None


def seasons(n):
    today = datetime.date.today()
    y, q = today.year - 1911, (today.month - 1) // 3 + 1
    out = []
    for _ in range(n):
        q -= 1
        if q == 0:
            y, q = y - 1, 4
        out.append(f"{y}S{q}")
    return out


def rows_from_zip(blob):
    z = zipfile.ZipFile(io.BytesIO(blob))
    name = next((n for n in z.namelist() if n.lower() == f"{CITY}_lvr_land_a.csv"), None)
    if not name:
        return []
    text = z.read(name).decode("utf-8-sig", errors="replace")
    rd = list(csv.reader(io.StringIO(text)))
    if len(rd) < 3:
        return []
    head = rd[0]
    return [dict(zip(head, r)) for r in rd[2:]]  # 第 2 列是英文欄名


IND_ZONE = re.compile(r"工")  # 都市土地：工業區、乙種工業區、甲種工業區、零星工業區…
IND_TYPE = re.compile(r"工廠|廠辦|倉庫")


RESI = re.compile(r"住宅大樓|華廈|公寓|套房")


def is_industrial(r):
    """工業區（都市計畫）、丁種建築用地（非都市），或建物型態是工廠／廠辦／倉庫；排除住宅大樓類。"""
    if RESI.search(r.get("建物型態", "")):
        return False
    return bool(IND_ZONE.search(r.get("都市土地使用分區", "")) and "住" not in r.get("都市土地使用分區", "")
                or "丁種" in r.get("非都市土地使用編定", "")
                or IND_TYPE.search(r.get("建物型態", "")))


def roc_date(s):
    s = re.sub(r"\D", "", s or "")
    if len(s) < 7:
        return ""
    y, m, d = int(s[:-4]) + 1911, int(s[-4:-2]), int(s[-2:])
    try:
        return datetime.date(y, m, d).isoformat()
    except ValueError:
        return ""


def road_of(addr, dist):
    """公開位置只到「區＋路」：門牌取到路/街/大道(含段)；地號只顯示區。"""
    a = re.sub(r"\s+", "", addr or "").replace("臺", "台")
    a = a.split(dist, 1)[-1] if dist and dist in a else a
    if "地號" in a or re.search(r"段\d", a):
        return ""
    m = re.match(r"(.*?(?:大道|路|街))((?:[一二三四五六七八九十]{1,3}段)?)", a)
    return (m.group(1) + m.group(2)) if m and len(m.group(1)) <= 10 else ""


def num(s):
    try:
        return float(str(s).replace(",", ""))
    except ValueError:
        return 0.0


def slim(r):
    dist = r.get("鄉鎮市區", "")
    land = num(r.get("土地移轉總面積平方公尺")) * 0.3025
    build = num(r.get("建物移轉總面積平方公尺")) * 0.3025
    total = num(r.get("總價元"))
    note = r.get("備註", "")
    kind = r.get("交易標的", "")
    zone = r.get("都市土地使用分區", "") or r.get("非都市土地使用編定", "") or r.get("非都市土地使用分區", "")
    base_ping = land if land > 0 else build  # 工業地產習慣看「地坪單價」（含建物價值）
    return {
        "id": r.get("編號", ""), "date": roc_date(r.get("交易年月日")), "dist": dist,
        "road": road_of(r.get("土地位置建物門牌", ""), dist), "kind": kind,
        "zone": zone, "btype": r.get("建物型態", ""), "use": r.get("主要用途", ""),
        "land_ping": round(land, 1), "build_ping": round(build, 1),
        "total_wan": round(total / 10000), "unit_wan_ping": round(total / 10000 / base_ping, 1) if base_ping else 0,
        "unit_base": "地坪" if land > 0 else "建坪",
        "built": roc_date((r.get("建築完成年月", "") + "01")[:7]) if r.get("建築完成年月") else "",
        "special": bool(re.search(r"親友|特殊關係|員工|二親等|債權|急買急賣|瑕疵|凶宅|含增建|政府機關", note)),
        "note": note[:60],
    }


def main():
    found = {}
    old = json.loads(OUT.read_text("utf-8")) if OUT.exists() else {"deals": []}
    for d in old.get("deals", []):  # 保留舊資料（超過 8 季的仍留 3 年），套用目前的篩選
        if not RESI.search(d.get("btype", "")) and "住" not in d.get("zone", ""):
            found[d["id"]] = d
    urls = [f"{BASE}/Download?type=zip&fileName=lvr_landcsv.zip"] + [
        f"{BASE}/DownloadSeason?season={s}&type=zip&fileName=lvr_landcsv.zip" for s in seasons(QUARTERS)]
    ok = 0
    for u in urls:
        blob = get(u)
        if not blob or blob[:2] != b"PK":
            print("SKIP", u, len(blob or b""))
            continue
        rows = rows_from_zip(blob)
        ind = [slim(r) for r in rows if is_industrial(r)]
        print("OK", u.split("?")[-1][:40], "rows", len(rows), "industrial", len(ind))
        for d in ind:
            if d["id"] and d["date"]:
                found[d["id"]] = d
        ok += 1
        time.sleep(2)
    if not ok:
        print("沒有下載到任何資料，保留舊檔")
        sys.exit(0)
    cutoff = (datetime.date.today() - datetime.timedelta(days=365 * 3)).isoformat()
    deals = sorted((d for d in found.values() if d["date"] >= cutoff), key=lambda d: d["date"], reverse=True)
    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps({"updated": datetime.date.today().isoformat(),
                               "source": "內政部不動產成交案件實際資訊（實價登錄）開放資料",
                               "deals": deals}, ensure_ascii=False, indent=1), "utf-8")
    print("saved", len(deals))


if __name__ == "__main__":
    main()
