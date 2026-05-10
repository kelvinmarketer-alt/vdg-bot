"""
VDG Weekly Report Bot v2
- Đọc CSV Google Ads → phân loại bằng Claude → ghi vào Google Sheet
- Đã fix theo cấu trúc thật: 2 cặp Tháng/Tuần, ngày BĐ/KT ở hàng tổng tuần
"""
import os
import sys
import json
import re
import argparse
import pandas as pd
import gspread
from datetime import datetime, date, timedelta
from google.oauth2.service_account import Credentials
from anthropic import Anthropic

# Force UTF-8 output trên Windows
sys.stdout.reconfigure(encoding="utf-8")

# === CONFIG ===
SHEET_ID = "15mOavpjqao7oR3a8NbGILi70Kj4gW-gAhtGJjiEABqA"
TAB_BAC = "Báo cáo năm 2026 - miền bắc"
TAB_NAM = "Báo cáo năm 2026 - miền nam"
HEADER_ROW = 2
RAW_TAB = "raw_ads_data"
GSA_PATH = "gsa.json"

NHOM_SP = ["Túi 8 cạnh", "Zipper", "Bao bì", "Túi hút chân không", "Giấy"]


# === HELPERS ===
def norm(s):
    return re.sub(r"\s+", " ", str(s).strip().lower())

def find_col_nth(header, keyword, nth=1):
    """Tìm cột (1-based) chứa keyword, lần xuất hiện thứ nth."""
    target = norm(keyword)
    count = 0
    for i, h in enumerate(header, start=1):
        if target in norm(h):
            count += 1
            if count == nth:
                return i
    return None

def col_letter(idx):
    s = ""
    while idx > 0:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s


# === STEP 1: TỰ TÍNH TUẦN (hoặc nhận CLI override) ===
print("=" * 50)
print("VDG WEEKLY REPORT BOT v3")
print("=" * 50)

parser = argparse.ArgumentParser(description="VDG weekly report")
parser.add_argument("--from", dest="from_date", help="dd/mm/yyyy (override)")
parser.add_argument("--to", dest="to_date", help="dd/mm/yyyy (override)")
parser.add_argument("-y", "--yes", action="store_true", help="Bỏ qua confirm, ghi luôn")
args = parser.parse_args()

if args.from_date and args.to_date:
    ngay_bd = datetime.strptime(args.from_date, "%d/%m/%Y")
    ngay_kt = datetime.strptime(args.to_date, "%d/%m/%Y")
    print(f"→ Override: {args.from_date} - {args.to_date}")
else:
    today = date.today()
    # Tuần trước = T2 trước → CN trước
    this_monday = today - timedelta(days=today.weekday())
    last_monday = this_monday - timedelta(days=7)
    last_sunday = last_monday + timedelta(days=6)
    ngay_bd = datetime.combine(last_monday, datetime.min.time())
    ngay_kt = datetime.combine(last_sunday, datetime.min.time())
    print(f"→ Auto-detect tuần trước (chạy hôm nay: {today.strftime('%a %d/%m/%Y')})")

ngay_bd_str = ngay_bd.strftime("%d/%m/%Y")
ngay_kt_str = ngay_kt.strftime("%d/%m/%Y")
thang = ngay_bd.month
tuan_trong_thang = (ngay_bd.day - 1) // 7 + 1
print(f"  → Tháng {thang} / Tuần {tuan_trong_thang} ({ngay_bd_str} - {ngay_kt_str})")

# Format ngắn cho cell ngày BĐ/KT (vd: "4/5", "10/5")
ngay_bd_short = f"{ngay_bd.day}/{ngay_bd.month}"
ngay_kt_short = f"{ngay_kt.day}/{ngay_kt.month}"


# === STEP 2: ĐỌC DATA TỪ SHEET (do Google Ads Script đẩy vào) ===
print(f"\n→ Đọc data từ tab '{RAW_TAB}'...")
creds = Credentials.from_service_account_file(
    GSA_PATH,
    scopes=["https://www.googleapis.com/auth/spreadsheets"],
)
gc = gspread.authorize(creds)
sh = gc.open_by_key(SHEET_ID)
try:
    ws_raw = sh.worksheet(RAW_TAB)
except gspread.WorksheetNotFound:
    print(f"❌ Không thấy tab '{RAW_TAB}'. Chạy Google Ads Script trước.")
    sys.exit(1)

# Dùng get_all_values (raw string) để tránh gspread tự parse sai locale VN
all_rows = ws_raw.get_all_values()
if len(all_rows) < 2:
    print(f"❌ Tab '{RAW_TAB}' trống. Chạy Google Ads Script trước.")
    sys.exit(1)

header_raw = all_rows[0]
data_rows = all_rows[1:]
records = [dict(zip(header_raw, row)) for row in data_rows]

df = pd.DataFrame(records)
if "ad_group" not in df.columns:
    df["ad_group"] = "(campaign-level)"
# Lấy đủ cột cho phân tích
keep_cols = ["campaign", "ad_group", "cost"]
for c in ["impressions", "clicks"]:
    if c in df.columns:
        keep_cols.append(c)
df = df[keep_cols].copy()
# Sheet VN: "," có thể là dấu thập phân → đổi sang "." rồi parse
for c in ["cost", "impressions", "clicks"]:
    if c in df.columns:
        df[c] = df[c].astype(str).str.replace(",", ".", regex=False)
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
df = df[df["cost"] > 0].reset_index(drop=True)
# Tính CTR mỗi dòng (%)
if "impressions" in df.columns and "clicks" in df.columns:
    df["ctr"] = (df["clicks"] / df["impressions"].replace(0, 1) * 100).round(2)
    df["cpc"] = (df["cost"] / df["clicks"].replace(0, 1)).round(0).astype(int)
print(f"  Updated lúc: {records[0].get('updated_at', 'N/A')} | {len(df)} dòng có chi tiêu")

print(f"\nDanh sách (campaign / ad group / chi tiêu):")
for _, r in df.iterrows():
    print(f"  - [{r.campaign}] / {r.ad_group}: {r.cost:,.0f}đ")


# === STEP 3: CLAUDE PHÂN LOẠI ===
print("\n→ Đang gọi Claude phân loại...")
client = Anthropic()

items = "\n".join(
    f'{i}. campaign="{r.campaign}" | ad_group="{r.ad_group}"'
    for i, r in df.iterrows()
)
prompt = f"""Phân loại các dòng Google Ads vào (miền, nhóm SP).

NHÓM SP hợp lệ: {NHOM_SP} hoặc "Khác".
MIỀN: "Bắc" hoặc "Nam" (đọc từ campaign name).

QUY TẮC PHÂN LOẠI NHÓM SP (theo thứ tự ưu tiên):
1. Nếu campaign chứa keyword nhóm SP rõ ràng → dùng campaign name
   (vd "VDG - miền bắc bao bì" → Bao bì)
2. Nếu ad_group là "(campaign-level)" hoặc "(performance-max)" → dùng campaign name
   (vd "Performance Max-5- Miền nam" → Bao bì miền nam, theo policy của user)
3. Còn lại → đọc ad_group name để xác định nhóm SP

KEYWORD → NHÓM:
- "bao bì" → Bao bì
- "zipper" → Zipper
- "8 cạnh" / "tám cạnh" → Túi 8 cạnh
- "hút chân không" / "chan khong" → Túi hút chân không
- "giấy" → Giấy

CHỈ trả "Khác" khi cả campaign lẫn ad_group đều không có keyword nào ở trên.

Data:
{items}

Trả về CHỈ JSON array, KHÔNG markdown, KHÔNG giải thích:
[{{"campaign":"...","ad_group":"...","mien":"Bắc|Nam","nhom_sp":"..."}}]"""

resp = client.messages.create(
    model="claude-opus-4-7", max_tokens=2000,
    messages=[{"role": "user", "content": prompt}],
)
raw = resp.content[0].text.strip()
raw = re.sub(r"^```(?:json)?|```$", "", raw, flags=re.MULTILINE).strip()
mapping = json.loads(raw)

df = df.merge(pd.DataFrame(mapping), on=["campaign", "ad_group"], how="left")
print("\n→ Phân loại từng dòng:")
print(df[["campaign", "ad_group", "mien", "nhom_sp", "cost"]].to_string(index=False))

summary = df.groupby(["mien", "nhom_sp"])["cost"].sum().reset_index()
summary["cost"] = summary["cost"].astype(int)
print("\n→ Tổng hợp theo (miền, nhóm SP) — cộng dồn nếu nhiều nguồn:")
print(summary.to_string(index=False))

if args.yes:
    print("\n→ -y flag → ghi luôn không cần confirm")
else:
    if input("\nGhi vào Sheet? (y/n): ").strip().lower() != "y":
        print("Đã hủy.")
        sys.exit(0)


# === STEP 4: GHI SHEET ===
def ghi_mien(mien_label, tab_name):
    print(f"\n--- {tab_name} ---")
    ws = sh.worksheet(tab_name)
    rows = ws.get_all_values()
    header = rows[HEADER_ROW - 1]

    # Cột 1,2 = Tháng/Tuần cấp tổng (có merged); cột 3,4 = cấp detail (luôn có)
    idx_thang = find_col_nth(header, "tháng", 2)  # cột 3
    idx_tuan = find_col_nth(header, "tuần", 2)    # cột 4
    idx_nbd = find_col_nth(header, "ngày bắt đầu", 1)
    idx_nkt = find_col_nth(header, "ngày kết thúc", 1)
    idx_nhom = find_col_nth(header, "nhóm sp", 1)
    idx_chi = find_col_nth(header, "chi tiêu nhóm", 1)

    cols = {"Tháng(detail)": idx_thang, "Tuần(detail)": idx_tuan,
            "Ngày BĐ": idx_nbd, "Ngày KT": idx_nkt,
            "Nhóm SP": idx_nhom, "Chi tiêu": idx_chi}
    print(f"  Cột phát hiện: {cols}")
    if None in cols.values():
        missing = [k for k, v in cols.items() if v is None]
        print(f"  ❌ Thiếu cột: {missing}. Bỏ qua tab này.")
        return

    # Tìm tất cả hàng khớp tháng + tuần
    matching_rows = []
    for ri, row in enumerate(rows[HEADER_ROW:], start=HEADER_ROW + 1):
        if len(row) < idx_chi: continue
        v_thang = row[idx_thang - 1].strip()
        v_tuan = row[idx_tuan - 1].strip()
        v_nhom = row[idx_nhom - 1].strip() if idx_nhom <= len(row) else ""
        if v_thang == str(thang) and v_tuan == str(tuan_trong_thang):
            matching_rows.append((ri, v_nhom))

    if not matching_rows:
        print(f"  ❌ Không thấy hàng nào có Tháng={thang} Tuần={tuan_trong_thang}")
        return

    # Ngày BĐ/KT: luôn ghi vào hàng ĐẦU TIÊN của block tuần
    # (cột E/F thường merged 5 hàng, anchor = hàng đầu)
    row_for_date = matching_rows[0][0]
    updates = []
    updates.append({"range": f"{col_letter(idx_nbd)}{row_for_date}",
                    "values": [[ngay_bd_short]]})
    updates.append({"range": f"{col_letter(idx_nkt)}{row_for_date}",
                    "values": [[ngay_kt_short]]})
    print(f"  ✓ Hàng đầu tuần {row_for_date}: ngày {ngay_bd_short} → {ngay_kt_short}")

    # Ghi chi tiêu vào hàng nhóm SP tương ứng
    sub = summary[summary.mien == mien_label]
    for _, r in sub.iterrows():
        if r.nhom_sp == "Khác":
            print(f"  ⚠️ Bỏ qua 'Khác' ({r.cost:,.0f}đ)")
            continue
        target = next((ri for ri, n in matching_rows if n == r.nhom_sp), None)
        if target is None:
            print(f"  ⚠️ Không thấy hàng nhóm '{r.nhom_sp}' trong tuần này")
            continue
        # Ghi SỐ nguyên — sheet tự format thành "104.408 đ" theo cell format có sẵn
        updates.append({"range": f"{col_letter(idx_chi)}{target}",
                        "values": [[int(r.cost)]]})
        print(f"  ✓ Hàng {target}: {r.nhom_sp} = {int(r.cost):,}".replace(",", "."))

    if updates:
        ws.batch_update(updates, value_input_option="USER_ENTERED")


ghi_mien("Bắc", TAB_BAC)
ghi_mien("Nam", TAB_NAM)
print(f"\n✓ Sheet đã update: https://docs.google.com/spreadsheets/d/{SHEET_ID}")


# === STEP 5: PHÂN TÍCH + GỬI TELEGRAM ===
def send_telegram_report():
    import urllib.request
    import urllib.parse

    token = os.environ.get("TELEGRAM_BOT_TOKEN")
    chat_id = os.environ.get("TELEGRAM_CHAT_ID")
    if not token or not chat_id:
        print("\n⚠️ Thiếu TELEGRAM_BOT_TOKEN / TELEGRAM_CHAT_ID, bỏ qua send")
        return

    print("\n→ Đang phân tích + tạo báo cáo Telegram...")

    # Tổng hợp toàn cảnh để gửi cho Claude
    total_spend = int(df["cost"].sum())
    by_mien = df.groupby("mien")["cost"].sum().to_dict()
    bac_pct = round(by_mien.get("Bắc", 0) / total_spend * 100, 1) if total_spend else 0
    nam_pct = round(by_mien.get("Nam", 0) / total_spend * 100, 1) if total_spend else 0

    # Detail có CTR/CPC nếu có
    has_perf = "ctr" in df.columns
    if has_perf:
        detail_str = df[["campaign", "ad_group", "mien", "nhom_sp",
                         "cost", "impressions", "clicks", "ctr", "cpc"]].to_string(index=False)
    else:
        detail_str = df[["campaign", "ad_group", "mien", "nhom_sp", "cost"]].to_string(index=False)

    summary_str = summary.to_string(index=False)

    prompt = f"""Bạn là analyst Google Ads cho công ty Vua Đóng Gói (B2B bao bì).

DATA TUẦN: {ngay_bd_str} → {ngay_kt_str} (Tháng {thang}, Tuần {tuan_trong_thang})
Tổng chi 2 miền: {total_spend:,}đ — Bắc {bac_pct}% / Nam {nam_pct}%

Tổng hợp theo (miền, nhóm SP):
{summary_str}

Chi tiết từng dòng (có ad group, CTR, CPC):
{detail_str}

VIẾT 1 TIN NHẮN TELEGRAM TÁCH RIÊNG 2 MIỀN, định dạng đúng như khung dưới đây (tiếng Việt, <600 chữ):

📊 *Báo cáo Tuần {tuan_trong_thang}/Tháng {thang}* ({ngay_bd_str.replace('/2026', '')} - {ngay_kt_str.replace('/2026', '')})

🌏 *Tổng 2 miền*: {total_spend:,}đ — Bắc {bac_pct}% / Nam {nam_pct}%

━━━━━━━━━━━━━━━━━
🅱️ *MIỀN BẮC*
💰 Chi: ...đ
🏆 Top nhóm:
1. ... — ...đ (XX%)
2. ...
3. ...
🚨 Cảnh báo: (CTR thấp, CPC bất thường, chi nhiều ít click...) hoặc "Không có"
💡 Đề xuất: 1-2 ý cụ thể

━━━━━━━━━━━━━━━━━
🅽 *MIỀN NAM*
💰 Chi: ...đ
🏆 Top nhóm:
1. ...
2. ...
3. ...
🚨 Cảnh báo: ... hoặc "Không có"
💡 Đề xuất: 1-2 ý cụ thể

━━━━━━━━━━━━━━━━━
🔍 *Đánh giá tổng*: 1-2 câu so sánh Bắc vs Nam, điểm cần chú ý chung.

QUY TẮC FORMAT:
- Telegram Markdown: *bold*, _italic_. KHÔNG dùng ** hoặc ##
- Số tiền dấu chấm: 7.316.100đ
- Mỗi cảnh báo / đề xuất 1 dòng ngắn, có số liệu cụ thể
- KHÔNG lan man, KHÔNG dùng từ chung chung như "cần xem xét lại"
- Nếu 1 nhóm SP không có data trong miền nào, KHÔNG bịa, ghi "(không chạy)"
"""

    client = Anthropic()
    resp = client.messages.create(
        model="claude-opus-4-7",
        max_tokens=1500,
        messages=[{"role": "user", "content": prompt}],
    )
    msg = resp.content[0].text.strip()

    # Send Telegram
    url = f"https://api.telegram.org/bot{token}/sendMessage"
    data = urllib.parse.urlencode({
        "chat_id": chat_id,
        "text": msg,
        "parse_mode": "Markdown",
        "disable_web_page_preview": "true",
    }).encode("utf-8")
    try:
        req = urllib.request.Request(url, data=data)
        with urllib.request.urlopen(req, timeout=15) as r:
            print(f"✓ Telegram đã gửi (HTTP {r.status})")
    except Exception as e:
        print(f"❌ Telegram lỗi: {e}")
        # Fallback: gửi không có markdown nếu lỗi parse
        try:
            data_plain = urllib.parse.urlencode({
                "chat_id": chat_id, "text": msg
            }).encode("utf-8")
            urllib.request.urlopen(urllib.request.Request(url, data=data_plain), timeout=15)
            print(f"✓ Telegram đã gửi (plain text fallback)")
        except Exception as e2:
            print(f"❌ Fallback cũng lỗi: {e2}")


send_telegram_report()
print(f"\n✓ Done!")
