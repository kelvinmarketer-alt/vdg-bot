"""
VDG Care Report Bot — báo cáo chăm sóc KH hằng ngày cho sếp (group Telegram riêng).
- Đọc sheet log chăm sóc KH (public CSV gviz), lọc NGÀY HÔM QUA.
- Tổng hợp: data chăm sóc theo loại KH, nhóm SP, doanh số, mức độ, số KH/nhân viên.
- Gửi group Telegram -5311055700 bằng bot VĐG (TELEGRAM_BOT_TOKEN).
- NV chưa điền (không có dòng ngày đó) → báo "⚠️ Chưa có dữ liệu".
- Cron 8h sáng VN, báo hôm qua. Chỉ dùng thư viện chuẩn (không cần pip).
"""
import os
import sys
import re
import csv
import io
import urllib.request
import urllib.parse
from datetime import datetime, timedelta, timezone
from collections import Counter, defaultdict

sys.stdout.reconfigure(encoding="utf-8")

SHEET_ID = "1-wGtnF2UYCkCH9sdGr68Csbv2g-rTQB7DoKILEeZRrQ"
CSV_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq?tqx=out:csv"
CHAT_ID = "-5311055700"  # group báo cáo chăm sóc KH (bot VĐG)
VN_TZ = timezone(timedelta(hours=7))
# Cột: 0 STT · 1 Ngày · 2 Nguồn · 3 NV · 4 Cách tư vấn · 5 TT KH · 6 TT SP
#      · 7 Nhóm SP · 8 Quá trình · 9 Mức độ KH · 10 Doanh số · 11 Ghi chú
ORDER = ["QC mới", "QC cũ", "TT mới", "TT cũ", "Khác"]


def digits(s):
    s = re.sub(r"[^\d]", "", str(s))
    return int(s) if s else 0


def fmt(n):
    return f"{int(n):,}đ".replace(",", ".")


def parse_dmy(s):
    m = re.match(r"\s*(\d{1,2})\s*/\s*(\d{1,2})\s*/\s*(\d{4})", str(s))
    return (int(m.group(1)), int(m.group(2)), int(m.group(3))) if m else None


def kh_cat(nguon):
    """Phân loại KH theo cột Nguồn. 'mới' (ad) ưu tiên trước 'cũ'
    → 'Hotline mới, QC cũ' = QC mới (quảng cáo mới) theo chốt của user."""
    s = (nguon or "").lower()
    if "tt mới" in s:
        return "TT mới"
    if "tt cũ" in s:
        return "TT cũ"
    if "mới" in s:
        return "QC mới"
    if "cũ" in s:
        return "QC cũ"
    return "Khác"


def marker_file(y):
    """Đường dẫn dấu 'đã gửi' theo ngày báo cáo (dùng chống gửi trùng khi
    chạy nhiều nhịp cron trong ngày). Chỉ bật khi có env STATE_DIR."""
    d = os.environ.get("STATE_DIR", "").strip()
    if not d:
        return None
    os.makedirs(d, exist_ok=True)
    return os.path.join(d, f"sent_{y.year}-{y.month:02d}-{y.day:02d}.txt")


def already_sent(y):
    p = marker_file(y)
    return bool(p) and os.path.exists(p)


def mark_sent(y):
    p = marker_file(y)
    if p:
        with open(p, "w") as f:
            f.write(datetime.now(VN_TZ).isoformat())


def send_telegram(msg):
    token = os.environ.get("TELEGRAM_BOT_TOKEN")
    if not token:
        print("⚠️ Thiếu TELEGRAM_BOT_TOKEN")
        return False
    url = f"https://api.telegram.org/bot{token}/sendMessage"
    data = urllib.parse.urlencode({
        "chat_id": CHAT_ID, "text": msg,
        "parse_mode": "Markdown", "disable_web_page_preview": "true",
    }).encode("utf-8")
    try:
        urllib.request.urlopen(urllib.request.Request(url, data=data), timeout=15)
        print("✓ Telegram đã gửi")
        return True
    except Exception as e:
        print(f"❌ Markdown lỗi ({e}), thử plain...")
        try:
            d2 = urllib.parse.urlencode({"chat_id": CHAT_ID, "text": msg}).encode("utf-8")
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
    ykey = (y.day, y.month, y.year)
    dstr = f"{y.day}/{y.month}"
    print(f"→ Báo cáo chăm sóc ngày {dstr}/{y.year}")

    if already_sent(y):
        print("→ Đã gửi hôm nay rồi (nhịp cron trước) — bỏ qua để khỏi trùng.")
        return

    raw = urllib.request.urlopen(CSV_URL, timeout=30).read().decode("utf-8")
    rows = list(csv.reader(io.StringIO(raw)))
    # pad mỗi dòng >=12 cột để không rớt dòng có ô cuối trống; lọc dòng CÓ NGÀY hợp lệ
    body = [r + [""] * (12 - len(r)) for r in rows[1:] if len(r) >= 2 and parse_dmy(r[1])]
    day = [r for r in body if parse_dmy(r[1]) == ykey]

    header = f"📋 *BÁO CÁO CHĂM SÓC KH — Ngày {dstr}* _(hôm qua)_"

    if not day:
        if send_telegram(header + "\n\n⚠️ *Chưa có dữ liệu* — nhân viên chưa cập nhật sheet cho ngày này."):
            mark_sent(y)
        print("→ Không có dữ liệu, đã báo.")
        return

    cat = Counter()
    rev = defaultdict(int)
    prod = Counter()
    staff = Counter()
    muc = Counter()
    for r in day:
        c = kh_cat(r[2])
        cat[c] += 1
        rev[c] += digits(r[10]) if len(r) > 10 else 0
        if len(r) > 7 and r[7].strip():
            prod[r[7].strip()] += 1
        if len(r) > 3 and r[3].strip():
            staff[r[3].strip()] += 1
        if len(r) > 9 and r[9].strip():
            muc[r[9].strip()] += 1
    tong_rev = sum(rev.values())

    L = [header, "", f"👥 *Data chăm sóc: {len(day)} lượt*"]
    for k in ORDER:
        if cat.get(k):
            L.append(f"   • KH {k}: {cat[k]}")
    L += ["", "📦 *Nhóm sản phẩm:*"]
    for k, v in prod.most_common():
        L.append(f"   • {k}: {v}")
    L += ["", f"💰 *Doanh số ngày: {fmt(tong_rev)}*"]
    if tong_rev:
        for k in ORDER:
            if rev.get(k):
                L.append(f"   • {k}: {fmt(rev[k])}")
    else:
        L.append("   _(chưa có doanh số)_")
    if muc:
        L += ["", "🎯 *Mức độ KH:* " + " · ".join(f"{k} {v}" for k, v in muc.most_common())]
    L += ["", "👤 *Số KH chăm sóc / nhân viên:*"]
    for k, v in staff.most_common():
        L.append(f"   • {k}: {v}")

    if send_telegram("\n".join(L)):
        mark_sent(y)
    print("✓ Done!")


if __name__ == "__main__":
    main()
