"""
VDG Daily Combined Report (GG + FB) → Telegram
- Nếu có META_ACCESS_TOKEN: KÉO FB từ Meta API (luôn tươi) → GHI vào FB 2026 (A:G,
  cột Tuần + F=Tổng chi tiêu ngày) → dùng cho báo cáo. Không phụ thuộc fb-bot chạy.
- Nếu chưa có token: fallback đọc FB từ sheet FB 2026 (như cũ).
- GG: đọc tab 'raw_daily' của hôm qua, tách MIỀN Bắc/Nam (từ tên campaign), chỉ CHI TIÊU.
- Gửi nhóm Telegram "VDG Report Bot". Dùng chung secret của vdg-bot.
"""
import os
import sys
import time
import re
import json
import urllib.request
import urllib.parse
from datetime import datetime, date, timedelta, timezone
import gspread
from google.oauth2.service_account import Credentials

sys.stdout.reconfigure(encoding="utf-8")

SHEET_ID = "15mOavpjqao7oR3a8NbGILi70Kj4gW-gAhtGJjiEABqA"
FB_TAB = "FB 2026"
GG_TAB = "raw_daily"
GSA_PATH = "gsa.json"
GRAPH_VER = "v21.0"
VN_TZ = timezone(timedelta(hours=7))


# ---------- helpers chung ----------
def digits(s):
    s = re.sub(r"[^\d]", "", str(s))
    return int(s) if s else 0


def fmt(n):
    return f"{int(n):,}đ".replace(",", ".")


def with_retry(fn, tries=4):
    """Thử lại khi Google API lỗi tạm thời (503/500/429) — tránh fail vô cớ."""
    for i in range(tries):
        try:
            return fn()
        except Exception as e:
            if i == tries - 1:
                raise
            print(f"⚠️ Google API lỗi tạm ({e}); thử lại {i + 2}/{tries} sau {3 * (i + 1)}s...")
            time.sleep(3 * (i + 1))


def dmy(iso_str):
    d = datetime.strptime(iso_str, "%Y-%m-%d").date()
    return f"{d.day}/{d.month}"


# ---------- FB: phân loại + tuần (copy từ fb-bot) ----------
def product_of(campaign_name):
    s = campaign_name.strip()
    s = re.sub(r"^\s*(mess(age|enger)?|inbox|lead|form|traffic|reach|tn|cd|qc)\b[\s\-|:/•·]*",
               "", s, flags=re.IGNORECASE)
    s = re.sub(r"\s+", " ", s).strip(" -|:/•·")
    if not s:
        s = campaign_name.strip()
    return s[0].upper() + s[1:] if s else s


def extract_messages(actions):
    if not actions:
        return 0
    return sum(int(round(float(a.get("value", 0))))
               for a in actions if "messaging_conversation_started" in a.get("action_type", ""))


def calc_thang_tuan(date_obj):
    if isinstance(date_obj, datetime):
        date_obj = date_obj.date()
    jan1 = date(date_obj.year, 1, 1)
    if jan1.weekday() != 0:
        first_sun = jan1 + timedelta(days=(6 - jan1.weekday()) % 7)
        if jan1 <= date_obj <= first_sun:
            return (1, 1)
    mon = date_obj if date_obj.weekday() == 0 else date_obj - timedelta(days=date_obj.weekday())

    def block_anchor(year, month):
        first = date(year, month, 1)
        fdw = first.weekday()
        if month == 1 and fdw != 0:
            return (first + timedelta(days=(7 - fdw) % 7), 2)
        elif fdw <= 2:
            return (first - timedelta(days=fdw), 1)
        else:
            return (first + timedelta(days=7 - fdw), 1)

    ty, tm = mon.year, mon.month
    fbm, wo = block_anchor(ty, tm)
    nm = tm + 1 if tm < 12 else 1
    ny = ty + 1 if tm == 12 else ty
    nf = date(ny, nm, 1)
    fbn, won = block_anchor(ny, nm)
    if fbn < nf and mon == fbn:
        ty, tm, fbm, wo = ny, nm, fbn, won
    elif mon < fbm:
        pm = tm - 1 if tm > 1 else 12
        py = ty - 1 if tm == 1 else ty
        ty, tm = py, pm
        fbm, wo = block_anchor(py, pm)
    tuan = (mon - fbm).days // 7 + wo
    return (tm, tuan)


def week_label(iso_str):
    d = datetime.strptime(iso_str, "%Y-%m-%d").date()
    t, w = calc_thang_tuan(d)
    return f"T{t}/W{w}"


def graph_insights(acct, token, since_str, until_str):
    params = {
        "level": "campaign", "time_increment": "1",
        "fields": "campaign_name,spend,actions",
        "time_range": json.dumps({"since": since_str, "until": until_str}),
        "limit": "500", "access_token": token,
    }
    url = f"https://graph.facebook.com/{GRAPH_VER}/act_{acct}/insights?" + urllib.parse.urlencode(params)
    rows, page = [], 0
    while url:
        page += 1
        try:
            with urllib.request.urlopen(url, timeout=60) as r:
                data = json.loads(r.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            print(f"❌ Graph API lỗi HTTP {e.code}: {e.read().decode('utf-8', 'ignore')}")
            raise
        rows.extend(data.get("data", []))
        url = data.get("paging", {}).get("next")
        if page > 20:
            break
    return rows


def fb_pull_and_write(sh):
    """Có token → kéo Meta N ngày, upsert FB 2026, trả về agg {(dm,product):[spend,msgs]}.
    Không có token → trả None (fallback)."""
    token = os.environ.get("META_ACCESS_TOKEN")
    acct = os.environ.get("META_AD_ACCOUNT_ID", "").replace("act_", "").strip()
    if not token or not acct:
        return None
    days = int(os.environ.get("FB_DAYS_BACK", "10"))
    today = datetime.now(VN_TZ).date()
    rows = graph_insights(acct, token, (today - timedelta(days=days)).isoformat(), today.isoformat())
    agg, week_of = {}, {}
    for r in rows:
        d = dmy(r["date_start"])
        week_of[d] = week_label(r["date_start"])
        p = product_of(r.get("campaign_name", ""))
        e = agg.setdefault((d, p), [0, 0])
        e[0] += int(round(float(r.get("spend", 0))))
        e[1] += extract_messages(r.get("actions"))
    agg = {k: v for k, v in agg.items() if v[0] > 0 or v[1] > 0}

    ws = sh.worksheet(FB_TAB)
    existing = with_retry(lambda: ws.get_all_values())
    idx, last = {}, 1
    for i, row in enumerate(existing[1:], start=2):
        if row and row[0].strip():
            last = i
            if len(row) >= 2:
                idx[(row[0].strip(), row[1].strip())] = i
    updates, nxt = [], last + 1
    for (d, p), (sp, ms) in sorted(agg.items(), key=lambda kv: (kv[0][0], kv[0][1])):
        cost = round(sp / ms) if ms > 0 else ""
        r_ = idx.get((d, p))
        if r_ is None:
            r_ = nxt
            nxt += 1
        updates.append({"range": f"A{r_}:G{r_}",
                        "values": [[d, p, sp, ms, cost, f"=SUMIF(A:A;A{r_};C:C)", week_of.get(d, "")]]})
    if updates:
        with_retry(lambda: ws.batch_update(updates, value_input_option="USER_ENTERED"))
    print(f"✓ FB: kéo Meta + ghi {len(updates)} dòng vào FB 2026")
    return agg


def fb_from_sheet(sh, dm):
    fb = {}
    for row in with_retry(lambda: sh.worksheet(FB_TAB).get_all_values())[1:]:
        if len(row) >= 4 and row[0].strip() == dm:
            p = row[1].strip()
            if not p:
                continue
            e = fb.setdefault(p, [0, 0])
            e[0] += digits(row[2])
            e[1] += digits(row[3])
    return fb


# ---------- GG ----------
def gg_product(ad_group, campaign):
    s = (ad_group or "").lower()
    if "màng đơn" in s or "mang don" in s:
        return "Màng đơn"
    if "bao bì" in s or "bao bi" in s:
        return "Bao bì"
    if "zipper" in s:
        return "Zipper"
    if "8 cạnh" in s or "8 canh" in s or "tám cạnh" in s:
        return "Túi 8 cạnh"
    if "hút chân không" in s or "hut chan khong" in s:
        return "Túi hút chân không"
    if "giấy" in s or "giay" in s:
        return "Giấy"
    return "Bao bì"  # PMax/campaign-level → policy


def gg_mien(campaign):
    c = (campaign or "").lower()
    if "bắc" in c or "bac" in c:
        return "Bắc"
    if "nam" in c:
        return "Nam"
    return "Khác"


# ---------- Telegram ----------
def send_telegram(msg):
    token = os.environ.get("TELEGRAM_BOT_TOKEN")
    chat_id = os.environ.get("TELEGRAM_CHAT_ID")
    if not token or not chat_id:
        print("⚠️ Thiếu TELEGRAM credentials, skip send")
        return False
    url = f"https://api.telegram.org/bot{token}/sendMessage"
    data = urllib.parse.urlencode({
        "chat_id": chat_id, "text": msg,
        "parse_mode": "Markdown", "disable_web_page_preview": "true",
    }).encode("utf-8")
    try:
        urllib.request.urlopen(urllib.request.Request(url, data=data), timeout=15)
        print("✓ Telegram đã gửi")
        return True
    except Exception as e:
        print(f"❌ Markdown lỗi ({e}), thử plain...")
        try:
            d2 = urllib.parse.urlencode({"chat_id": chat_id, "text": msg}).encode("utf-8")
            urllib.request.urlopen(urllib.request.Request(url, data=d2), timeout=15)
            print("✓ Telegram đã gửi (plain)")
            return True
        except Exception as e2:
            print(f"❌ Fallback lỗi: {e2}")
            return False


# ---------- MAIN ----------
def main():
    tgt = os.environ.get("TARGET_DATE", "").strip()
    y = (datetime.strptime(tgt, "%Y-%m-%d").date() if tgt
         else datetime.now(VN_TZ).date() - timedelta(days=1))
    iso = y.isoformat()
    dm = f"{y.day}/{y.month}"
    print(f"→ Báo cáo ngày {dm} ({iso})")

    creds = Credentials.from_service_account_file(
        GSA_PATH, scopes=["https://www.googleapis.com/auth/spreadsheets"])
    gc = gspread.authorize(creds)
    sh = with_retry(lambda: gc.open_by_key(SHEET_ID))

    # FB — ưu tiên kéo Meta + ghi sheet; không có token thì đọc sheet
    agg = fb_pull_and_write(sh)
    if agg is not None:
        fb = {p: v for (d, p), v in agg.items() if d == dm}
    else:
        print("⚠️ Chưa có META_ACCESS_TOKEN → đọc FB từ sheet (fallback)")
        fb = fb_from_sheet(sh, dm)

    # GG — raw_daily hôm qua, tách miền
    gg = {}
    for row in with_retry(lambda: sh.worksheet(GG_TAB).get_all_values())[1:]:
        if len(row) >= 4 and row[0].strip() == iso:
            key = (gg_mien(row[1]), gg_product(row[2], row[1]))
            gg[key] = gg.get(key, 0) + digits(row[3])

    fb_total = sum(v[0] for v in fb.values())
    gg_total = sum(gg.values())
    combined = fb_total + gg_total

    lines = [f"📊 *Báo cáo NGÀY {dm}* _(hôm qua)_", "", f"🔵 *FACEBOOK* — {fmt(fb_total)}"]
    if fb:
        for p, (sp, ms) in sorted(fb.items(), key=lambda kv: -kv[1][0]):
            lines.append(f"  📦 {p}: {fmt(sp)} · {ms} tin")
    else:
        lines.append("  _(không có dữ liệu)_")
    lines += ["", f"🟢 *GOOGLE* — {fmt(gg_total)}"]
    if gg:
        icons = {"Bắc": "🅱️", "Nam": "🅽", "Khác": "▪️"}
        for mien in ["Bắc", "Nam", "Khác"]:
            items = {p: sp for (m, p), sp in gg.items() if m == mien}
            if not items:
                continue
            lines.append(f"{icons.get(mien, '')} *Miền {mien}* — {fmt(sum(items.values()))}")
            for p, sp in sorted(items.items(), key=lambda kv: -kv[1]):
                lines.append(f"   📦 {p}: {fmt(sp)}")
    else:
        lines.append("  _(không có dữ liệu)_")
    lines += ["", "━━━━━━━━━━", f"🌏 *TỔNG GG + FB: {fmt(combined)}*"]

    msg = "\n".join(lines)
    print("\n" + msg + "\n")
    send_telegram(msg)
    print("✓ Done!")


if __name__ == "__main__":
    main()
