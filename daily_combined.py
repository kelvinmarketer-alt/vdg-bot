"""
VDG Daily Combined Report (GG + FB) → Telegram
- Đọc FB (tab 'FB 2026') + GG (tab 'raw_daily') của NGÀY hôm qua trong cùng Google Sheet
- Soạn báo cáo chi tiêu theo nhóm SP + tổng → gửi nhóm Telegram "VDG Report Bot"
- GG chỉ CHI TIÊU (theo yêu cầu user). FB có chi tiêu + tin nhắn.
- KHÔNG cần Meta token: đọc FB từ sheet (fb-bot đã ghi). Dùng chung secret của vdg-bot.
"""
import os
import sys
import re
import urllib.request
import urllib.parse
from datetime import datetime, timedelta, timezone
import gspread
from google.oauth2.service_account import Credentials

sys.stdout.reconfigure(encoding="utf-8")

SHEET_ID = "15mOavpjqao7oR3a8NbGILi70Kj4gW-gAhtGJjiEABqA"
FB_TAB = "FB 2026"
GG_TAB = "raw_daily"
GSA_PATH = "gsa.json"
VN_TZ = timezone(timedelta(hours=7))


def digits(s):
    s = re.sub(r"[^\d]", "", str(s))
    return int(s) if s else 0


def fmt(n):
    return f"{int(n):,}đ".replace(",", ".")


def gg_product(ad_group, campaign):
    """Ad group name -> nhóm SP GG (giống convention bot tuần). PMax/campaign-level -> Bao bì (policy)."""
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
    # (performance-max) / (campaign-level) / khác → policy gộp Bao bì
    return "Bao bì"


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
        print(f"❌ Markdown lỗi ({e}), thử plain text...")
        try:
            d2 = urllib.parse.urlencode({"chat_id": chat_id, "text": msg}).encode("utf-8")
            urllib.request.urlopen(urllib.request.Request(url, data=d2), timeout=15)
            print("✓ Telegram đã gửi (plain)")
            return True
        except Exception as e2:
            print(f"❌ Fallback lỗi: {e2}")
            return False


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
    sh = gc.open_by_key(SHEET_ID)

    # FB: tab FB 2026 — A=Ngày(d/m) B=SP C=Chi tiêu D=Tin nhắn
    fb = {}
    for row in sh.worksheet(FB_TAB).get_all_values()[1:]:
        if len(row) >= 4 and row[0].strip() == dm:
            p = row[1].strip()
            if not p:
                continue
            e = fb.setdefault(p, [0, 0])
            e[0] += digits(row[2])
            e[1] += digits(row[3])

    # GG: tab raw_daily — A=date(iso) B=campaign C=ad_group D=cost
    gg = {}
    for row in sh.worksheet(GG_TAB).get_all_values()[1:]:
        if len(row) >= 4 and row[0].strip() == iso:
            p = gg_product(row[2], row[1])
            gg[p] = gg.get(p, 0) + digits(row[3])

    fb_total = sum(v[0] for v in fb.values())
    gg_total = sum(gg.values())
    combined = fb_total + gg_total

    lines = [f"📊 *Báo cáo NGÀY {dm}* _(hôm qua)_", ""]
    lines.append(f"🔵 *FACEBOOK* — {fmt(fb_total)}")
    if fb:
        for p, (sp, ms) in sorted(fb.items(), key=lambda kv: -kv[1][0]):
            lines.append(f"  📦 {p}: {fmt(sp)} · {ms} tin")
    else:
        lines.append("  _(không có dữ liệu)_")
    lines += ["", f"🟢 *GOOGLE* — {fmt(gg_total)}"]
    if gg:
        for p, sp in sorted(gg.items(), key=lambda kv: -kv[1]):
            lines.append(f"  📦 {p}: {fmt(sp)}")
    else:
        lines.append("  _(không có dữ liệu)_")
    lines += ["", "━━━━━━━━━━", f"🌏 *TỔNG GG + FB: {fmt(combined)}*"]

    msg = "\n".join(lines)
    print("\n" + msg + "\n")
    send_telegram(msg)
    print("✓ Done!")


if __name__ == "__main__":
    main()
