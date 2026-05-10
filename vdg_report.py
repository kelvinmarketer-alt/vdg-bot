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
import zoneinfo
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
parser.add_argument("--force", action="store_true", help="Bỏ qua state, gửi báo cáo bất kể đã gửi chưa")
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


# === STEP 5: STATE-AWARE TELEGRAM REPORTING ===

# --- Đọc Form/Hotline/Zalo/Mess + Tổng data nhóm cho 1 tuần ---
def read_conversions(tab_name, target_thang, target_tuan):
    ws = sh.worksheet(tab_name)
    rows = ws.get_all_values()
    header = rows[HEADER_ROW - 1]
    idx = {
        "thang": find_col_nth(header, "tháng", 2),
        "tuan": find_col_nth(header, "tuần", 2),
        "nhom": find_col_nth(header, "nhóm sp", 1),
        "total": find_col_nth(header, "tổng data nhóm", 1),
        "form": find_col_nth(header, "form", 1),
        "hotline": find_col_nth(header, "hotline", 1),
        "zalo": find_col_nth(header, "zalo", 1),
        "mess": find_col_nth(header, "mess", 1),
    }

    def parse_int(s):
        s = str(s).strip()
        if not s or s == "-":
            return 0
        try:
            return int(float(s.replace(",", ".")))
        except Exception:
            return 0

    result = {}
    for row in rows[HEADER_ROW:]:
        if len(row) < 8:
            continue
        if row[idx["thang"] - 1].strip() != str(target_thang):
            continue
        if row[idx["tuan"] - 1].strip() != str(target_tuan):
            continue
        nhom = row[idx["nhom"] - 1].strip() if idx["nhom"] and idx["nhom"] <= len(row) else ""
        if not nhom:
            continue
        result[nhom] = {
            k: parse_int(row[idx[k] - 1]) if idx[k] and idx[k] <= len(row) else 0
            for k in ["form", "hotline", "zalo", "mess", "total"]
        }
    return result


# --- State management trong tab _bot_state ---
STATE_TAB = "_bot_state"
STATE_HEADER = ["week_id", "thang", "tuan", "ngay_bd", "ngay_kt",
                "initial_sent_at", "conversion_sent_at", "last_check_at"]


def get_state_tab():
    try:
        ws = sh.worksheet(STATE_TAB)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(STATE_TAB, rows=200, cols=10)
        ws.update("A1", [STATE_HEADER])
        return ws
    rows = ws.get_all_values()
    if not rows or rows[0][:len(STATE_HEADER)] != STATE_HEADER:
        ws.update("A1", [STATE_HEADER])
    return ws


def get_state(ws, week_id):
    rows = ws.get_all_values()
    if len(rows) < 2:
        return {}
    for row in rows[1:]:
        if row and row[0] == week_id:
            return dict(zip(rows[0], row + [""] * (len(rows[0]) - len(row))))
    return {}


def upsert_state(ws, week_id, **fields):
    fields["week_id"] = week_id  # đảm bảo cột week_id luôn được set
    rows = ws.get_all_values()
    header = rows[0] if rows else STATE_HEADER
    target = None
    for i, row in enumerate(rows[1:], start=2):
        if row and row[0] == week_id:
            target = i
            break
    if target:
        existing = dict(zip(header, rows[target - 1] + [""] * (len(header) - len(rows[target - 1]))))
        existing.update({k: str(v) for k, v in fields.items()})
        ws.update(f"A{target}", [[existing.get(h, "") for h in header]])
    else:
        new_data = {h: "" for h in header}
        new_data.update({k: str(v) for k, v in fields.items()})
        ws.append_row([new_data.get(h, "") for h in header])


# --- Telegram + Claude helpers ---
def call_claude(prompt):
    resp = client.messages.create(
        model="claude-opus-4-7", max_tokens=2000,
        messages=[{"role": "user", "content": prompt}],
    )
    return resp.content[0].text.strip()


def send_telegram(msg):
    import urllib.request
    import urllib.parse
    token = os.environ.get("TELEGRAM_BOT_TOKEN")
    chat_id = os.environ.get("TELEGRAM_CHAT_ID")
    if not token or not chat_id:
        print("⚠️ Thiếu TELEGRAM credentials, skip send")
        return False
    url = f"https://api.telegram.org/bot{token}/sendMessage"
    data = urllib.parse.urlencode({
        "chat_id": chat_id, "text": msg, "parse_mode": "Markdown",
        "disable_web_page_preview": "true",
    }).encode("utf-8")
    try:
        urllib.request.urlopen(urllib.request.Request(url, data=data), timeout=15)
        print("✓ Telegram đã gửi")
        return True
    except Exception as e:
        print(f"❌ Markdown lỗi ({e}), thử plain text...")
        try:
            data_plain = urllib.parse.urlencode({"chat_id": chat_id, "text": msg}).encode("utf-8")
            urllib.request.urlopen(urllib.request.Request(url, data=data_plain), timeout=15)
            print("✓ Telegram đã gửi (plain)")
            return True
        except Exception as e2:
            print(f"❌ Fallback lỗi: {e2}")
            return False


# --- Aggregates dùng chung ---
total_spend = int(df["cost"].sum())
by_mien = df.groupby("mien")["cost"].sum().to_dict()
bac_pct = round(by_mien.get("Bắc", 0) / total_spend * 100, 1) if total_spend else 0
nam_pct = round(by_mien.get("Nam", 0) / total_spend * 100, 1) if total_spend else 0
summary_str = summary.to_string(index=False)
short_bd = ngay_bd_str.replace("/2026", "")
short_kt = ngay_kt_str.replace("/2026", "")


# --- Build messages ---
def build_initial_msg():
    """Báo cáo CHỈ chi tiêu, note NV chưa cập nhật."""
    return call_claude(f"""Bạn là analyst Google Ads cho VDG (B2B bao bì).

DATA CHI TIÊU TUẦN: {ngay_bd_str} → {ngay_kt_str} (Tháng {thang}, Tuần {tuan_trong_thang})
Tổng chi: {total_spend:,}đ — Bắc {bac_pct}% / Nam {nam_pct}%

Tổng hợp:
{summary_str}

CONVERSION: CHƯA CÓ DATA — nhân viên chưa cập nhật Form/Hotline/Zalo/Mess.

VIẾT TIN NHẮN TELEGRAM (<400 chữ tiếng Việt). KHÔNG được phân tích hiệu quả chuyển đổi (vì chưa có data). Format đúng khung dưới:

📊 *Báo cáo CHI TIÊU Tuần {tuan_trong_thang}/Tháng {thang}* ({short_bd} - {short_kt})

⚠️ *Nhân viên chưa cập nhật thông tin về chuyển đổi.*
👉 Báo cáo này CHỈ phản ánh chi tiêu. Bot sẽ tự gửi báo cáo hiệu quả khi NV cập nhật xong.

🌏 *Tổng 2 miền*: {total_spend:,}đ — Bắc {bac_pct}% / Nam {nam_pct}%

━━━━━━━━━━━━━━━━━
🅱️ *MIỀN BẮC* — Chi: ...đ
🏆 Top 3 nhóm:
1. ... — ...đ (XX%)
2. ...
3. ...

━━━━━━━━━━━━━━━━━
🅽 *MIỀN NAM* — Chi: ...đ
🏆 Top 3 nhóm:
1. ...
2. ...
3. ...

━━━━━━━━━━━━━━━━━
ℹ️ Đợi nhân viên cập nhật. Bot check 3 lần/ngày.

QUY TẮC:
- *bold* Markdown, KHÔNG ## hay **
- Số tiền dấu chấm: 7.316.100đ
- KHÔNG đánh giá hiệu quả/CPA/tốt-xấu
- KHÔNG dùng metric CTR/click làm tiêu chí đánh giá
""")


def build_full_msg(is_followup):
    """Báo cáo có conversion, phân tích CPA, đánh giá hiệu quả."""
    conv_lines = []
    for mien_label, conv in [("BẮC", conv_bac), ("NAM", conv_nam)]:
        for nhom, v in conv.items():
            if any(v[k] > 0 for k in ["form", "hotline", "zalo", "mess", "total"]):
                conv_lines.append(
                    f"  {mien_label}/{nhom}: form={v['form']}, hotline={v['hotline']}, "
                    f"zalo={v['zalo']}, mess={v['mess']}, TỔNG DATA={v['total']}"
                )
    conv_str = "\n".join(conv_lines) or "  (chưa có data)"

    cpa_lines = []
    for mien_label, conv in [("Bắc", conv_bac), ("Nam", conv_nam)]:
        for nhom, v in conv.items():
            cost_for = int(summary[(summary.mien == mien_label) & (summary.nhom_sp == nhom)]["cost"].sum())
            if v["total"] > 0 and cost_for > 0:
                cpa = int(cost_for / v["total"])
                cpa_lines.append(f"  {mien_label}/{nhom}: chi {cost_for:,}đ ÷ {v['total']} data = CPA {cpa:,}đ/data")
            elif cost_for > 0 and v["total"] == 0:
                cpa_lines.append(f"  {mien_label}/{nhom}: chi {cost_for:,}đ NHƯNG 0 data → ❌ KHÔNG HIỆU QUẢ")
    cpa_str = "\n".join(cpa_lines) or "  (chưa tính được)"

    title = "📊 *Báo cáo CẬP NHẬT (sau khi NV điền)*" if is_followup else "📊 *Báo cáo HIỆU QUẢ Tuần*"
    followup_note = "✅ NV đã cập nhật conversion → đây là báo cáo hiệu quả thật.\n" if is_followup else ""

    return call_claude(f"""Bạn là analyst Google Ads cho VDG (B2B bao bì).

DATA TUẦN: {ngay_bd_str} → {ngay_kt_str} (Tháng {thang}, Tuần {tuan_trong_thang})
Tổng chi: {total_spend:,}đ — Bắc {bac_pct}% / Nam {nam_pct}%

CHI TIÊU theo (miền, nhóm SP):
{summary_str}

CONVERSION (NV đã cập nhật):
{conv_str}

CPA tính sẵn:
{cpa_str}

VIẾT TIN NHẮN TELEGRAM (<700 chữ tiếng Việt). PHÂN TÍCH HIỆU QUẢ CHUYỂN ĐỔI (CPA) là chính. KHÔNG dùng CTR/click làm metric đánh giá. Format đúng khung dưới:

{title} *Tuần {tuan_trong_thang}/Tháng {thang}* ({short_bd} - {short_kt})

{followup_note}🌏 *Tổng 2 miền*: {total_spend:,}đ — Bắc {bac_pct}% / Nam {nam_pct}%

━━━━━━━━━━━━━━━━━
🅱️ *MIỀN BẮC*
💰 Chi: ...đ → Data thu về: ... → CPA TB: ...đ/data
🏆 Hiệu quả top 3 nhóm (CPA thấp = tốt):
1. NhómA — chi ...đ ÷ ...data → CPA ...đ
2. ...
3. ...
🚨 Cảnh báo: (gọi rõ tên nhóm + số liệu cụ thể)
   - Nhóm có chi nhưng 0 data → ❌
   - Nhóm CPA cao bất thường (>2x TB) → ⚠️
💡 Đề xuất: 1-2 ý hành động cụ thể, có số liệu (vd "Tăng budget Zipper Bắc 20% — CPA chỉ 50k", "Pause Bao bì Bắc — chi 7tr không có data")

━━━━━━━━━━━━━━━━━
🅽 *MIỀN NAM*
[tương tự miền Bắc]

━━━━━━━━━━━━━━━━━
🔍 *Tổng kết*: 1-2 câu so sánh Bắc vs Nam (CPA, conversion volume), action priority tuần tới.

QUY TẮC:
- *bold* Markdown
- Số tiền dấu chấm: 7.316.100đ; CPA dạng "250.000đ/data"
- Action-oriented, có số liệu cụ thể, KHÔNG lan man
- Nếu nhóm có chi nhưng 0 conversion → đánh dấu "❌" rõ ràng
""")


# --- Decision tree (state machine) ---
print(f"\n→ Đọc state tab _bot_state...")
ws_state = get_state_tab()
week_id = f"{ngay_bd.year}-W{ngay_bd.isocalendar()[1]:02d}"
state = get_state(ws_state, week_id)
print(f"  Week ID: {week_id}")
print(f"  State: initial={state.get('initial_sent_at') or '(empty)'}, "
      f"conversion={state.get('conversion_sent_at') or '(empty)'}")

print(f"\n→ Đọc conversions từ 2 sheet báo cáo...")
conv_bac = read_conversions(TAB_BAC, thang, tuan_trong_thang)
conv_nam = read_conversions(TAB_NAM, thang, tuan_trong_thang)
total_conv_bac = sum(v["total"] for v in conv_bac.values())
total_conv_nam = sum(v["total"] for v in conv_nam.values())
has_conversion = (total_conv_bac + total_conv_nam) > 0
print(f"  Bắc: {total_conv_bac} data | Nam: {total_conv_nam} data | has_conversion={has_conversion}")

now_str = datetime.now(zoneinfo.ZoneInfo("Asia/Ho_Chi_Minh")).strftime("%Y-%m-%d %H:%M")
initial_sent = (state.get("initial_sent_at") or "").strip()
conv_sent = (state.get("conversion_sent_at") or "").strip()

if args.force or not initial_sent:
    label = "FORCE" if args.force else "LẦN ĐẦU"
    print(f"\n→ {label} — gửi báo cáo cho tuần này...")
    if has_conversion:
        print("  Có conversion data → gửi báo cáo HIỆU QUẢ")
        ok = send_telegram(build_full_msg(is_followup=False))
        if ok:
            upsert_state(ws_state, week_id,
                         thang=thang, tuan=tuan_trong_thang,
                         ngay_bd=ngay_bd_str, ngay_kt=ngay_kt_str,
                         initial_sent_at=now_str, conversion_sent_at=now_str,
                         last_check_at=now_str)
    else:
        print("  Chưa có conversion → gửi báo cáo CHI TIÊU (initial only)")
        ok = send_telegram(build_initial_msg())
        if ok:
            upsert_state(ws_state, week_id,
                         thang=thang, tuan=tuan_trong_thang,
                         ngay_bd=ngay_bd_str, ngay_kt=ngay_kt_str,
                         initial_sent_at=now_str, conversion_sent_at="",
                         last_check_at=now_str)
elif not conv_sent and has_conversion:
    print(f"\n→ NV đã cập nhật → gửi báo cáo HIỆU QUẢ (follow-up)...")
    ok = send_telegram(build_full_msg(is_followup=True))
    if ok:
        upsert_state(ws_state, week_id, conversion_sent_at=now_str, last_check_at=now_str)
else:
    print(f"\n→ Đã gửi đủ. Skip Telegram, chỉ update last_check_at.")
    upsert_state(ws_state, week_id, last_check_at=now_str)

print(f"\n✓ Done!")
