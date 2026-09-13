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
from datetime import datetime, date, timedelta, timezone
from collections import Counter, defaultdict

sys.stdout.reconfigure(encoding="utf-8")

SHEET_ID = "1-wGtnF2UYCkCH9sdGr68Csbv2g-rTQB7DoKILEeZRrQ"
CSV_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq?tqx=out:csv"
EVAL_SHEET = "đánh giá"  # tab đánh giá hiệu quả NV theo ngày
EVAL_CSV_URL = (f"https://docs.google.com/spreadsheets/d/{SHEET_ID}"
                f"/gviz/tq?tqx=out:csv&sheet={urllib.parse.quote(EVAL_SHEET)}")
CHAT_ID = "-5311055700"  # group báo cáo chăm sóc KH (bot VĐG)
VN_TZ = timezone(timedelta(hours=7))
# Cột tab "Báo cáo chi tiết" (đã BỎ cột STT): 0 Ngày · 1 Nguồn · 2 NV
#   · 3 Định hướng tư vấn · 4 TT KH · 5 TT SP · 6 Nhóm SP · 7 Quá trình
#   · 8 Mức độ KH · 9 Doanh số · 10 Ghi chú
ORDER = ["QC mới", "QC cũ", "TT mới", "TT cũ", "Khác"]


def digits(s):
    s = re.sub(r"[^\d]", "", str(s))
    return int(s) if s else 0


def fmt(n):
    return f"{int(n):,}đ".replace(",", ".")


def parse_dmy(s):
    m = re.match(r"\s*(\d{1,2})\s*/\s*(\d{1,2})\s*/\s*(\d{4})", str(s))
    return (int(m.group(1)), int(m.group(2)), int(m.group(3))) if m else None


def to_date(s):
    """Chuỗi ngày → date, hoặc None."""
    t = parse_dmy(s)
    if not t:
        return None
    try:
        return date(t[2], t[1], t[0])
    except ValueError:
        return None


def header_label(mode, start, end):
    """Nhãn ngắn cho tiêu đề phần đánh giá."""
    if mode == "weekly":
        return f"Tuần {start.day}/{start.month}→{end.day}/{end.month}"
    if mode == "monthly":
        return f"Tháng {start.month}/{start.year}"
    return f"Ngày {start.day}/{start.month}"


def resolve_period():
    """Xác định kỳ báo cáo theo env MODE (daily/weekly/monthly) + TARGET_DATE
    (mốc; trống = hôm nay VN). Trả về (mode, start, end, title, key, note)."""
    mode = os.environ.get("MODE", "daily").strip().lower()
    tgt = os.environ.get("TARGET_DATE", "").strip()
    base = (datetime.strptime(tgt, "%Y-%m-%d").date() if tgt
            else datetime.now(VN_TZ).date())
    if mode == "weekly":
        this_mon = base - timedelta(days=base.weekday())   # T2 tuần chứa base
        start = this_mon - timedelta(days=7)               # T2 tuần trước
        end = start + timedelta(days=6)                     # CN tuần trước
        title = (f"📅 *BÁO CÁO TUẦN — {start.day}/{start.month}→{end.day}/{end.month}"
                 f"/{end.year}* _(tuần vừa qua)_")
        return mode, start, end, title, f"week-{start.isoformat()}", "cả tuần"
    if mode == "monthly":
        last_prev = base.replace(day=1) - timedelta(days=1)  # ngày cuối tháng trước
        start = last_prev.replace(day=1)
        end = last_prev
        title = f"🗓️ *BÁO CÁO THÁNG {start.month}/{start.year}* _(tháng vừa qua)_"
        return mode, start, end, title, f"month-{start.year}-{start.month:02d}", "cả tháng"
    # daily: TARGET_DATE = đúng ngày cần báo (không né); trống = hôm qua.
    # Công ty NGHỈ CHỦ NHẬT + NV nhập liệu thứ 7 vào thứ 2 → nếu 'hôm qua' rơi
    # vào Chủ nhật thì lùi về THỨ 7 (sáng thứ 2 báo cáo ngày cho thứ 7).
    if tgt:
        d = datetime.strptime(tgt, "%Y-%m-%d").date()
        sub, note = "_(hôm qua)_", "hôm qua"
    else:
        d = datetime.now(VN_TZ).date() - timedelta(days=1)
        if d.weekday() == 6:            # Chủ nhật (nghỉ)
            d -= timedelta(days=1)      # → Thứ 7
            sub, note = "_(thứ 7 tuần rồi)_", "ngày này"
        else:
            sub, note = "_(hôm qua)_", "hôm qua"
    title = f"📋 *BÁO CÁO CHĂM SÓC KH — Ngày {d.day}/{d.month}* {sub}"
    return mode, d, d, title, f"day-{d.isoformat()}", note


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


def marker_file(key):
    """Đường dẫn dấu 'đã gửi' theo KEY kỳ báo cáo (day-.../week-.../month-...)
    để chống gửi trùng khi chạy nhiều nhịp cron. Chỉ bật khi có env STATE_DIR."""
    d = os.environ.get("STATE_DIR", "").strip()
    if not d:
        return None
    os.makedirs(d, exist_ok=True)
    return os.path.join(d, f"sent_{key}.txt")


def already_sent(key):
    p = marker_file(key)
    return bool(p) and os.path.exists(p)


def mark_sent(key):
    p = marker_file(key)
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


def fetch_eval(start, end):
    """Đọc tab 'đánh giá', forward-fill cột Ngày (chỉ điền ở dòng đầu mỗi khối),
    trả về các dòng NV có ngày trong [start, end]. Cột: 0 Ngày · 1 NV ·
    2 DS mục tiêu · 3 QC mới · 4 QC cũ · 5 tự tìm mới · 6 tự tìm cũ · 7 DS thực tế ·
    8 %DS · 9 KH cũ CS · 10 KH tự tìm · 11 Gặp khách · 12 Làm việc."""
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
        d = to_date(r[0])
        if d:
            cur = d
        if not r[1].strip():
            continue
        if cur and start <= cur <= end:
            out.append(r)
    return out


def eval_meaningful(rows):
    """Có ai thực sự cập nhật chưa? (DS thực tế >0, hoặc điền Làm việc/gặp khách/
    mục tiêu CS). Chỉ có sẵn DS mục tiêu mẫu thì coi như CHƯA cập nhật."""
    for r in rows:
        if digits(r[7]) or r[11].strip() or r[12].strip() or r[9].strip() or r[10].strip():
            return True
    return False


def _ratio(a, t):
    """'a/t (p%)' — bỏ % khi target=0."""
    return f"{a}/{t}" + (f" ({a / t * 100:.0f}%)" if t else "")


def eval_section(rows, label, tt_by_nv=None, cu_by_nv=None):
    """Đánh giá + XẾP HẠNG NV cho kỳ (cộng dồn nếu nhiều ngày); None nếu rỗng.
    Mỗi NV có thể có NHIỀU dòng (mỗi ngày 1 dòng) → gộp theo tên. tt_by_nv /
    cu_by_nv = số lượt THỰC TẾ tự tìm / KH cũ theo NV (đếm từ tab chăm sóc cả kỳ).
    Mục tiêu KPI cộng dồn từ tab đánh giá (cột 9/10)."""
    if not rows:
        return None
    tt_by_nv = tt_by_nv or {}
    cu_by_nv = cu_by_nv or {}
    agg = {}
    order = []
    for r in rows:
        name = r[1].strip()
        if name not in agg:
            agg[name] = dict(name=name, target=0, actual=0, cu_tgt=0, tt_tgt=0,
                             off=0, days=0, src=Counter())
            order.append(name)
        a = agg[name]
        a["days"] += 1
        # ngày NGHỈ: không tính mục tiêu (không bắt KPI ngày nghỉ), chỉ đếm số ngày nghỉ
        if "nghỉ" in r[12].strip().lower():
            a["off"] += 1
            continue
        a["target"] += digits(r[2])
        a["actual"] += digits(r[7])
        a["cu_tgt"] += digits(r[9])
        a["tt_tgt"] += digits(r[10])
        for lbl, ci in [("QC mới", 3), ("QC cũ", 4), ("Tự tìm mới", 5), ("Tự tìm cũ", 6)]:
            a["src"][lbl] += digits(r[ci])

    items = []
    for name in order:
        a = agg[name]
        a["tt_act"] = tt_by_nv.get(name, 0)
        a["cu_act"] = cu_by_nv.get(name, 0)
        a["p"] = (a["actual"] / a["target"] * 100) if a["target"] else 0
        a["fulloff"] = a["off"] >= a["days"] and a["actual"] == 0
        items.append(a)

    work = sorted([it for it in items if not it["fulloff"]], key=lambda x: -x["p"])
    offs = [it for it in items if it["fulloff"]]

    L = [f"📊 *ĐÁNH GIÁ & XẾP HẠNG NV — {label}*",
         "_(xếp theo % hoàn thành KPI doanh số)_", ""]
    medals = ["🥇", "🥈", "🥉"]
    tt = ta = tct = tca = ttt = tta = 0
    for i, it in enumerate(work):
        rank = medals[i] if i < 3 else f"*{i + 1}.*"
        tt += it["target"]; ta += it["actual"]
        tct += it["cu_tgt"]; tca += it["cu_act"]
        ttt += it["tt_tgt"]; tta += it["tt_act"]
        em = ("🟢" if it["p"] >= 100 else "🟡" if it["p"] >= 50
              else "🟠" if it["p"] > 0 else "🔴")
        off_note = f"  _(nghỉ {it['off']} ngày)_" if it["off"] else ""
        L.append(f"{rank} *{it['name']}* — {em} KPI DS {it['p']:.0f}%{off_note}")
        srcs = [f"{lbl} {fmt(v)}" for lbl, v in it["src"].items() if v]
        dsline = f"   💰 DS: {fmt(it['actual'])} / {fmt(it['target'])}"
        if srcs:
            dsline += "  (" + " · ".join(srcs) + ")"
        L.append(dsline)
        L.append(f"   📞 Tự tìm: {_ratio(it['tt_act'], it['tt_tgt'])}"
                 f"  ·  🔁 KH cũ: {_ratio(it['cu_act'], it['cu_tgt'])}")
    for it in offs:
        suffix = "" if it["days"] <= 1 else f" ({it['days']} ngày)"
        L.append(f"😴 *{it['name']}*: Nghỉ{suffix}")
    if tt:
        L += ["", "Σ *Tổng đội:*",
              f"   💰 DS: {fmt(ta)} / {fmt(tt)} ({ta / tt * 100:.0f}%)",
              f"   📞 Tự tìm: {_ratio(tta, ttt)}  ·  🔁 KH cũ: {_ratio(tca, tct)}"]
    return "\n".join(L)


def main():
    mode, start, end, header, key, note = resolve_period()
    ds_label = "Doanh số ngày" if mode == "daily" else f"Doanh số {note}"
    print(f"→ {mode} | {start} → {end}")

    if already_sent(key) and not os.environ.get("FORCE", "").strip():
        print("→ Đã gửi kỳ này rồi (nhịp trước) — bỏ qua để khỏi trùng.")
        return

    raw = urllib.request.urlopen(CSV_URL, timeout=30).read().decode("utf-8")
    rows = list(csv.reader(io.StringIO(raw)))
    # pad mỗi dòng >=11 cột để không rớt dòng có ô cuối trống; lọc dòng CÓ NGÀY trong kỳ (cột 0)
    body = [r + [""] * (11 - len(r)) for r in rows[1:] if len(r) >= 1 and to_date(r[0])]
    day = [r for r in body if start <= to_date(r[0]) <= end]

    # đếm số lượt THỰC TẾ theo NV cho KPI: tự tìm (TT) và KH cũ (cột Nguồn)
    tt_by_nv, cu_by_nv = defaultdict(int), defaultdict(int)
    for r in day:
        nv = r[2].strip()
        c = kh_cat(r[1])
        if c in ("TT mới", "TT cũ"):
            tt_by_nv[nv] += 1
        if c in ("QC cũ", "TT cũ"):
            cu_by_nv[nv] += 1

    ev_rows = fetch_eval(start, end)
    ev = (eval_section(ev_rows, header_label(mode, start, end), tt_by_nv, cu_by_nv)
          if eval_meaningful(ev_rows) else None)

    if not day and not ev:
        if send_telegram(header + f"\n\n⚠️ *Chưa có dữ liệu* — chưa có dữ liệu chăm sóc/đánh giá cho {note}."):
            mark_sent(key)
        print("→ Không có dữ liệu, đã báo.")
        return

    if not day:
        # có đánh giá nhưng chưa ghi lượt chăm sóc nào
        L = [header, "", f"⚠️ _Chưa ghi lượt chăm sóc KH nào cho {note}._"]
        L += ["", "━━━━━━━━━━━━━━━", "", ev]
        if send_telegram("\n".join(L)):
            mark_sent(key)
        print("✓ Done (chỉ có đánh giá)!")
        return

    cat = Counter()
    rev = defaultdict(int)
    prod = Counter()
    staff = Counter()
    muc = Counter()
    for r in day:
        c = kh_cat(r[1])
        cat[c] += 1
        rev[c] += digits(r[9])
        if r[6].strip():
            prod[r[6].strip()] += 1
        if r[2].strip():
            staff[r[2].strip()] += 1
        if r[8].strip():
            muc[r[8].strip()] += 1
    tong_rev = sum(rev.values())

    L = [header, "", f"👥 *Data chăm sóc: {len(day)} lượt*"]
    for k in ORDER:
        if cat.get(k):
            L.append(f"   • KH {k}: {cat[k]}")
    L += ["", "📦 *Nhóm sản phẩm:*"]
    for k, v in prod.most_common():
        L.append(f"   • {k}: {v}")
    L += ["", f"💰 *{ds_label}: {fmt(tong_rev)}*"]
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
        mark_sent(key)
    print("✓ Done!")


if __name__ == "__main__":
    main()
