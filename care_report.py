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
EVAL_SHEET = "đánh giá"  # tab đánh giá hiệu quả NV theo ngày
EVAL_CSV_URL = (f"https://docs.google.com/spreadsheets/d/{SHEET_ID}"
                f"/gviz/tq?tqx=out:csv&sheet={urllib.parse.quote(EVAL_SHEET)}")
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


def fetch_eval(ykey):
    """Đọc tab 'đánh giá', forward-fill cột Ngày (chỉ điền ở dòng đầu mỗi khối),
    trả về các dòng NV của ngày ykey. Cột: 0 Ngày · 1 NV · 2 DS mục tiêu ·
    3 QC mới · 4 QC cũ · 5 tự tìm mới · 6 tự tìm cũ · 7 DS thực tế · 8 %DS ·
    9 KH cũ CS · 10 KH tự tìm · 11 Gặp khách · 12 Làm việc."""
    try:
        raw = urllib.request.urlopen(EVAL_CSV_URL, timeout=30).read().decode("utf-8")
    except Exception as e:
        print(f"⚠️ Không đọc được tab đánh giá: {e}")
        return []
    rows = list(csv.reader(io.StringIO(raw)))
    out, cur = [], None
    for r in rows[1:]:
        if len(r) < 2:
            continue
        r = r + [""] * (13 - len(r))
        d = parse_dmy(r[0])
        if d:
            cur = d
        if not r[1].strip():
            continue
        if cur == ykey:
            out.append(r)
    return out


def eval_meaningful(rows):
    """Có ai thực sự cập nhật chưa? (DS thực tế >0, hoặc điền Làm việc/gặp khách/
    mục tiêu CS). Chỉ có sẵn DS mục tiêu mẫu thì coi như CHƯA cập nhật."""
    for r in rows:
        if digits(r[7]) or r[11].strip() or r[12].strip() or r[9].strip() or r[10].strip():
            return True
    return False


def eval_section(rows, dstr):
    """Dựng phần đánh giá hiệu quả NV cho ngày; None nếu không có dòng nào."""
    if not rows:
        return None
    items = []
    for r in rows:
        name = r[1].strip()
        target = digits(r[2])
        actual = digits(r[7])
        off = "nghỉ" in r[12].strip().lower()
        p = (actual / target * 100) if target else 0
        items.append((name, target, actual, p, off, r))
    # đi làm xếp theo %DS giảm dần; người nghỉ xuống cuối
    items.sort(key=lambda x: (x[4], -x[3]))
    L = [f"📊 *ĐÁNH GIÁ HIỆU QUẢ NV — Ngày {dstr}*", ""]
    tot_t = tot_a = 0
    for name, target, actual, p, off, r in items:
        if off:
            L.append(f"😴 *{name}*: Nghỉ")
            continue
        tot_t += target
        tot_a += actual
        em = "🟢" if p >= 100 else "🟡" if p >= 50 else "🟠" if p > 0 else "🔴"
        L.append(f"{em} *{name}*: {fmt(actual)} / {fmt(target)} ({p:.0f}%)")
        srcs = []
        for lbl, ci in [("QC mới", 3), ("QC cũ", 4), ("Tự tìm mới", 5), ("Tự tìm cũ", 6)]:
            v = digits(r[ci])
            if v:
                srcs.append(f"{lbl} {fmt(v)}")
        if srcs:
            L.append("   • Nguồn: " + " · ".join(srcs))
        cs = []
        if r[9].strip():
            cs.append(f"KH cũ {r[9].strip()}")
        if r[10].strip():
            cs.append(f"tự tìm {r[10].strip()}")
        if r[11].strip():
            cs.append(f"gặp {r[11].strip()}")
        if cs:
            L.append("   • Mục tiêu CS: " + " · ".join(cs))
    if tot_t:
        pt = tot_a / tot_t * 100
        L += ["", f"Σ *Tổng đội: {fmt(tot_a)} / {fmt(tot_t)}* ({pt:.0f}%)"]
    return "\n".join(L)


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

    ev_rows = fetch_eval(ykey)
    ev = eval_section(ev_rows, dstr) if eval_meaningful(ev_rows) else None

    header = f"📋 *BÁO CÁO CHĂM SÓC KH — Ngày {dstr}* _(hôm qua)_"

    if not day and not ev:
        if send_telegram(header + "\n\n⚠️ *Chưa có dữ liệu* — nhân viên chưa cập nhật sheet cho ngày này."):
            mark_sent(y)
        print("→ Không có dữ liệu, đã báo.")
        return

    if not day:
        # có đánh giá nhưng chưa ghi lượt chăm sóc nào
        L = [header, "", "⚠️ _Chưa ghi lượt chăm sóc KH nào cho ngày này._"]
        L += ["", "━━━━━━━━━━━━━━━", "", ev]
        if send_telegram("\n".join(L)):
            mark_sent(y)
        print("✓ Done (chỉ có đánh giá)!")
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

    if ev:
        L += ["", "━━━━━━━━━━━━━━━", "", ev]

    if send_telegram("\n".join(L)):
        mark_sent(y)
    print("✓ Done!")


if __name__ == "__main__":
    main()
