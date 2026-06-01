"""
VDG Weekly Report Bot v2
- Đọc CSV Google Ads → phân loại bằng OpenAI → ghi vào Google Sheet
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
from openai import OpenAI

# Force UTF-8 output trên Windows
sys.stdout.reconfigure(encoding="utf-8")

# === CONFIG ===
SHEET_ID = "15mOavpjqao7oR3a8NbGILi70Kj4gW-gAhtGJjiEABqA"
TAB_BAC = "Báo cáo năm 2026 - miền bắc"
TAB_NAM = "Báo cáo năm 2026 - miền nam"
HEADER_ROW = 2
RAW_TAB = "raw_ads_data"
GSA_PATH = "gsa.json"
OPENAI_MODEL = os.environ.get("OPENAI_MODEL", "gpt-4o")

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


def calc_thang_tuan(date_obj):
    """Tính (tháng_block, tuần_trong_block) theo convention sheet VDG.

    Rules:
    - Jan partial week: 1/1 đến CN đầu tiên → T1W1 (vd 1/1/2026 Thu → 1-4/1 = T1W1)
    - Tháng bắt đầu Mon-Wed: W1 = Mon-week chứa ngày 1 (Mon có thể ở tháng trước)
    - Tháng bắt đầu Thu-Sun: W1 = Mon-week ĐẦU TIÊN trong tháng (bỏ leading partial)
    """
    if isinstance(date_obj, datetime):
        date_obj = date_obj.date()

    jan1 = date(date_obj.year, 1, 1)
    if jan1.weekday() != 0:  # Jan 1 không phải Mon
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

    target_year, target_month = mon.year, mon.month
    fbm, w_offset = block_anchor(target_year, target_month)

    # Mon là leading Mon của block tháng kế (ở tháng hiện tại)?
    next_month = target_month + 1 if target_month < 12 else 1
    next_year = target_year + 1 if target_month == 12 else target_year
    next_first = date(next_year, next_month, 1)
    fbm_next, w_offset_next = block_anchor(next_year, next_month)
    if fbm_next < next_first and mon == fbm_next:
        target_year, target_month = next_year, next_month
        fbm, w_offset = fbm_next, w_offset_next
    elif mon < fbm:
        prev_month = target_month - 1 if target_month > 1 else 12
        prev_year = target_year - 1 if target_month == 1 else target_year
        target_year, target_month = prev_year, prev_month
        fbm, w_offset = block_anchor(prev_year, prev_month)

    tuan = (mon - fbm).days // 7 + w_offset
    return (target_month, tuan)


def find_doanh_thu_cols(header):
    """Tìm cột phần Doanh thu.
    Lưu ý 'tổng' xuất hiện 3 lần: (1) Tổng data nhóm, (2) Tổng (doanh thu), (3) Tổng 2026.
    Dùng nth=2 để khớp đúng cột Tổng doanh thu."""
    return {
        "thang": find_col_nth(header, "tháng", 4),
        "tuan": find_col_nth(header, "tuần", 3),
        "nbd": find_col_nth(header, "ngày bắt đầu", 2),
        "nkt": find_col_nth(header, "ngày kết thúc", 2),
        "kc_old": find_col_nth(header, "doanh thu khách cũ", 1),
        "kc_new": find_col_nth(header, "doanh thu khách mới", 1),
        "total": find_col_nth(header, "tổng", 2),
    }


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
thang, tuan_trong_thang = calc_thang_tuan(ngay_bd)
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


# === STEP 3: OPENAI PHÂN LOẠI ===
print("\n→ Đang gọi OpenAI phân loại...")
client = OpenAI()

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

resp = client.chat.completions.create(
    model=OPENAI_MODEL, max_tokens=2000,
    messages=[{"role": "user", "content": prompt}],
)
raw = resp.choices[0].message.content.strip()
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


def ghi_doanh_thu_dates(tab_name):
    """Ghi Ngày BĐ/KT vào phần Doanh thu (cột U, V của hàng khớp Tháng/Tuần)."""
    ws = sh.worksheet(tab_name)
    rows = ws.get_all_values()
    header = rows[HEADER_ROW - 1]
    idx = find_doanh_thu_cols(header)
    if not all([idx["thang"], idx["tuan"], idx["nbd"], idx["nkt"]]):
        print(f"  ⚠️ Không tìm thấy đủ cột Doanh thu (thang/tuan/nbd/nkt)")
        return

    # Forward-fill Tháng (cột S merged cells)
    target_row = None
    cur_thang = None
    for ri, row in enumerate(rows[HEADER_ROW:], start=HEADER_ROW + 1):
        if len(row) < max(idx["thang"], idx["tuan"]):
            continue
        v_thang = row[idx["thang"] - 1].strip()
        v_tuan = row[idx["tuan"] - 1].strip()
        if v_thang:
            cur_thang = v_thang
        if cur_thang == str(thang) and v_tuan == str(tuan_trong_thang):
            target_row = ri
            break

    if not target_row:
        print(f"  ⚠️ Doanh thu: không tìm thấy hàng T{thang}/W{tuan_trong_thang}")
        return

    ws.batch_update([
        {"range": f"{col_letter(idx['nbd'])}{target_row}", "values": [[ngay_bd_short]]},
        {"range": f"{col_letter(idx['nkt'])}{target_row}", "values": [[ngay_kt_short]]},
    ], value_input_option="USER_ENTERED")
    print(f"  ✓ Doanh thu hàng {target_row}: ngày {ngay_bd_short} → {ngay_kt_short}")


ghi_mien("Bắc", TAB_BAC)
ghi_mien("Nam", TAB_NAM)
print("\n--- Ghi Ngày BĐ/KT vào phần Doanh thu ---")
ghi_doanh_thu_dates(TAB_BAC)
ghi_doanh_thu_dates(TAB_NAM)
print(f"\n✓ Sheet đã update: https://docs.google.com/spreadsheets/d/{SHEET_ID}")


# === STEP 5: STATE-AWARE TELEGRAM REPORTING ===

# --- Đọc Doanh thu (khách cũ/mới/tổng) cho 1 tuần ---
def read_revenue(tab_name, target_thang, target_tuan):
    ws = sh.worksheet(tab_name)
    rows = ws.get_all_values()
    header = rows[HEADER_ROW - 1]
    idx = find_doanh_thu_cols(header)
    if not idx["thang"] or not idx["total"]:
        return {"khach_cu": 0, "khach_moi": 0, "total": 0}

    def parse_money(s):
        s = re.sub(r"[^\d]", "", str(s))
        return int(s) if s else 0

    # Forward-fill Tháng (cột S merged cells, các row giữa block trống)
    cur_thang = None
    for row in rows[HEADER_ROW:]:
        if len(row) < idx["tuan"]:
            continue
        v_thang = row[idx["thang"] - 1].strip()
        v_tuan = row[idx["tuan"] - 1].strip()
        if v_thang:
            cur_thang = v_thang
        if cur_thang != str(target_thang) or v_tuan != str(target_tuan):
            continue
        return {
            "khach_cu": parse_money(row[idx["kc_old"] - 1]) if idx["kc_old"] and idx["kc_old"] <= len(row) else 0,
            "khach_moi": parse_money(row[idx["kc_new"] - 1]) if idx["kc_new"] and idx["kc_new"] <= len(row) else 0,
            "total": parse_money(row[idx["total"] - 1]) if idx["total"] <= len(row) else 0,
        }
    return {"khach_cu": 0, "khach_moi": 0, "total": 0}


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
# Schema mới: track 4 cờ độc lập per miền (Bắc/Nam) × loại (conv/rev)
# Khi bất kỳ cờ nào flip từ trống → có timestamp, bot gửi báo cáo cập nhật
STATE_HEADER = ["week_id", "thang", "tuan", "ngay_bd", "ngay_kt",
                "initial_sent_at",
                "bac_conv_at", "nam_conv_at",
                "bac_rev_at", "nam_rev_at",
                "last_check_at"]


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


# --- Telegram + OpenAI helpers ---
def call_llm(prompt):
    resp = client.chat.completions.create(
        model=OPENAI_MODEL, max_tokens=2000,
        messages=[{"role": "user", "content": prompt}],
    )
    return resp.choices[0].message.content.strip()


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


# --- Format helpers ---
def fmt_money(n):
    """7316100 → '7.316.100đ'."""
    return f"{int(n):,}đ".replace(",", ".")


def fmt_short(n):
    """20114 → '20k'; 1500000 → '1.5tr'; 500 → '500đ'."""
    n = int(n)
    if n == 0:
        return "—"
    if n < 1000:
        return f"{n}đ"
    if n < 1_000_000:
        return f"{n // 1000}k"
    return f"{n / 1_000_000:.1f}tr"


def cost_for(mien_label, nhom):
    return int(summary[(summary.mien == mien_label) & (summary.nhom_sp == nhom)]["cost"].sum())


def build_cost_block(mien_label):
    """Liệt kê 5 nhóm SP với chi tiêu. (cho initial msg)"""
    lines = []
    total = 0
    for nhom in NHOM_SP:
        c = cost_for(mien_label, nhom)
        if c > 0:
            lines.append(f"📦 {nhom}: {fmt_money(c)}")
            total += c
        else:
            lines.append(f"📦 {nhom}: _(không chạy)_")
    return total, "\n".join(lines)


def build_full_block(mien_label, conv_dict):
    """Liệt kê 5 nhóm SP với chi · data · CPA. (cho full msg)"""
    lines = []
    total_c = 0
    total_d = 0
    for nhom in NHOM_SP:
        c = cost_for(mien_label, nhom)
        d = conv_dict.get(nhom, {}).get("total", 0) if conv_dict else 0
        if c > 0:
            if d == 0:
                lines.append(f"📦 {nhom}: {fmt_money(c)} · 0 data · ❌")
            else:
                cpa = c // d
                lines.append(f"📦 {nhom}: {fmt_money(c)} · {d} data · CPA {fmt_short(cpa)}")
            total_c += c
            total_d += d
        else:
            lines.append(f"📦 {nhom}: _(không chạy)_")
    return total_c, total_d, "\n".join(lines)


# --- Build messages ---
def build_initial_msg():
    """Báo cáo CHỈ chi tiêu — Python build full layout (không dùng LLM)."""
    bac_cost, bac_block = build_cost_block("Bắc")
    nam_cost, nam_block = build_cost_block("Nam")
    return f"""📊 *Báo cáo CHI TIÊU Tuần {tuan_trong_thang}/Tháng {thang}* ({short_bd} - {short_kt})

⚠️ *NV chưa cập nhật conversion*
_Báo cáo này chỉ có chi tiêu. Bot sẽ tự gửi báo cáo hiệu quả khi NV cập nhật xong._

🌏 *Tổng 2 miền*: {fmt_money(total_spend)}
   Bắc {fmt_money(bac_cost)} ({bac_pct}%) / Nam {fmt_money(nam_cost)} ({nam_pct}%)

━━━━━━━━━━━━━━
🅱️ *MIỀN BẮC* — {fmt_money(bac_cost)}

{bac_block}

━━━━━━━━━━━━━━
🅽 *MIỀN NAM* — {fmt_money(nam_cost)}

{nam_block}

━━━━━━━━━━━━━━
ℹ️ Đợi NV cập nhật Form/Hotline/Zalo/Mess. Bot check 3 lần/ngày."""


def build_full_msg(has_conv, has_rev, new_flags=None):
    """Báo cáo có conversion và/hoặc revenue. new_flags = dict {bac_conv, nam_conv, bac_rev, nam_rev} → True nếu vừa cập nhật."""
    use_conv_bac = conv_bac if has_conv else {}
    use_conv_nam = conv_nam if has_conv else {}

    bac_cost, bac_data, bac_block = build_full_block("Bắc", use_conv_bac)
    nam_cost, nam_data, nam_block = build_full_block("Nam", use_conv_nam)
    bac_cpa = bac_cost // bac_data if bac_data > 0 else 0
    nam_cpa = nam_cost // nam_data if nam_data > 0 else 0

    # Header line per miền
    bac_hdr_p = [fmt_money(bac_cost)]
    nam_hdr_p = [fmt_money(nam_cost)]
    if has_conv:
        bac_hdr_p += [f"{bac_data} data", f"CPA TB {fmt_short(bac_cpa)}"]
        nam_hdr_p += [f"{nam_data} data", f"CPA TB {fmt_short(nam_cpa)}"]
    bac_hdr = " · ".join(bac_hdr_p)
    nam_hdr = " · ".join(nam_hdr_p)

    # Revenue blocks
    bac_rev_str = ""
    nam_rev_str = ""
    bac_roas = nam_roas = 0
    if has_rev:
        bac_roas = rev_bac["total"] / bac_cost if bac_cost > 0 else 0
        nam_roas = rev_nam["total"] / nam_cost if nam_cost > 0 else 0
        bac_rev_str = (
            f"\n\n💵 *Doanh thu*: {fmt_money(rev_bac['total'])}"
            f"\n   Khách cũ: {fmt_money(rev_bac['khach_cu'])} · Khách mới: {fmt_money(rev_bac['khach_moi'])}"
            f"\n   ROAS: {bac_roas:.1f}x (chi 1đ → ra {bac_roas:.1f}đ)"
        )
        nam_rev_str = (
            f"\n\n💵 *Doanh thu*: {fmt_money(rev_nam['total'])}"
            f"\n   Khách cũ: {fmt_money(rev_nam['khach_cu'])} · Khách mới: {fmt_money(rev_nam['khach_moi'])}"
            f"\n   ROAS: {nam_roas:.1f}x (chi 1đ → ra {nam_roas:.1f}đ)"
        )

    # Eval context cho LLM
    eval_ctx = ""
    if has_conv:
        eval_ctx += f"CPA TB: Bắc {fmt_short(bac_cpa)} · Nam {fmt_short(nam_cpa)}\n"
    if has_rev:
        eval_ctx += (f"Doanh thu: Bắc {fmt_money(rev_bac['total'])} (ROAS {bac_roas:.1f}x) · "
                     f"Nam {fmt_money(rev_nam['total'])} (ROAS {nam_roas:.1f}x)\n")
    eval_focus = ""
    if has_rev:
        eval_focus = "ƯU TIÊN phân tích ROAS (doanh thu/chi tiêu). ROAS<1 = lỗ, ROAS 2-5 = OK, >5 = tốt."
    elif has_conv:
        eval_focus = "Phân tích CPA per nhóm SP. CPA cao bất thường (>2x TB) hoặc 0 data → cảnh báo."

    eval_raw = call_llm(f"""Bạn là analyst Google Ads cho VDG (B2B bao bì).

DATA TUẦN {ngay_bd_str} → {ngay_kt_str}:
{eval_ctx}
MIỀN BẮC (chi {fmt_money(bac_cost)}):
{bac_block}

MIỀN NAM (chi {fmt_money(nam_cost)}):
{nam_block}

{eval_focus}

Trả về CHỈ JSON:
{{
  "bac_warn": "<1 dòng cảnh báo Bắc, max 30 chữ, có số. Nếu OK → 'Không có cảnh báo lớn'>",
  "bac_action": "<1 dòng đề xuất, có động từ Scale/Pause/Tăng/Giảm + tên nhóm + số>",
  "nam_warn": "<1 dòng cảnh báo Nam>",
  "nam_action": "<1 dòng đề xuất Nam>",
  "summary": "<1-2 câu: so sánh hiệu quả Bắc vs Nam, kèm 1 ưu tiên action>"
}}

QUY TẮC:
- Số tiền dạng "20k", "1.5tr"
- KHÔNG dùng "cần xem xét", "có thể tối ưu" — phải có động từ + số
""")
    raw = re.sub(r"^```(?:json)?|```$", "", eval_raw.strip(), flags=re.MULTILINE).strip()
    try:
        ev = json.loads(raw)
    except Exception as e:
        print(f"⚠️ Parse JSON eval lỗi: {e}")
        ev = {"bac_warn": "(lỗi parse)", "bac_action": "—",
              "nam_warn": "(lỗi parse)", "nam_action": "—",
              "summary": raw[:200]}

    # Tiêu đề tùy theo loại data có
    if has_rev and has_conv:
        title = "📊 *Báo cáo HIỆU QUẢ ĐẦY ĐỦ* (chi · data · doanh thu)"
    elif has_rev:
        title = "📊 *Báo cáo DOANH THU* _(NV chưa cập nhật conversion)_"
    else:  # has_conv only
        title = "📊 *Báo cáo CHUYỂN ĐỔI* _(chưa có doanh thu)_"

    # Dòng "Vừa cập nhật" — nếu là follow-up báo cho user biết miền nào vừa điền thêm
    new_flags = new_flags or {}
    new_items = []
    if new_flags.get("bac_conv"): new_items.append("Bắc conversion")
    if new_flags.get("nam_conv"): new_items.append("Nam conversion")
    if new_flags.get("bac_rev"): new_items.append("Bắc doanh thu")
    if new_flags.get("nam_rev"): new_items.append("Nam doanh thu")
    update_line = f"\n✅ _Vừa cập nhật: {', '.join(new_items)}_\n" if new_items else "\n"

    return f"""{title} *Tuần {tuan_trong_thang}/Tháng {thang}* ({short_bd} - {short_kt})
{update_line}
🌏 *Tổng 2 miền*: {fmt_money(total_spend)}
   Bắc {fmt_money(bac_cost)} ({bac_pct}%) / Nam {fmt_money(nam_cost)} ({nam_pct}%)

━━━━━━━━━━━━━━
🅱️ *MIỀN BẮC* — {bac_hdr}

{bac_block}{bac_rev_str}

⚠️ {ev['bac_warn']}
💡 {ev['bac_action']}

━━━━━━━━━━━━━━
🅽 *MIỀN NAM* — {nam_hdr}

{nam_block}{nam_rev_str}

⚠️ {ev['nam_warn']}
💡 {ev['nam_action']}

━━━━━━━━━━━━━━
🔍 {ev['summary']}"""


# --- Decision tree (state machine) ---
print(f"\n→ Đọc state tab _bot_state...")
ws_state = get_state_tab()
week_id = f"{ngay_bd.year}-W{ngay_bd.isocalendar()[1]:02d}"
state = get_state(ws_state, week_id)
print(f"  Week ID: {week_id}")
print(f"  State: initial={state.get('initial_sent_at') or '(empty)'}, "
      f"conversion={state.get('conversion_sent_at') or '(empty)'}")

print(f"\n→ Đọc conversions + revenue từ 2 sheet báo cáo...")
conv_bac = read_conversions(TAB_BAC, thang, tuan_trong_thang)
conv_nam = read_conversions(TAB_NAM, thang, tuan_trong_thang)
total_conv_bac = sum(v["total"] for v in conv_bac.values())
total_conv_nam = sum(v["total"] for v in conv_nam.values())

rev_bac = read_revenue(TAB_BAC, thang, tuan_trong_thang)
rev_nam = read_revenue(TAB_NAM, thang, tuan_trong_thang)

# 4 cờ data hiện tại (per miền × loại)
cur_bac_conv = total_conv_bac > 0
cur_nam_conv = total_conv_nam > 0
cur_bac_rev = rev_bac["total"] > 0
cur_nam_rev = rev_nam["total"] > 0
has_conv = cur_bac_conv or cur_nam_conv
has_rev = cur_bac_rev or cur_nam_rev

print(f"  Conversion: Bắc {total_conv_bac} · Nam {total_conv_nam}")
print(f"  Revenue: Bắc {fmt_money(rev_bac['total'])} · Nam {fmt_money(rev_nam['total'])}")
print(f"  Cờ data: bac_conv={cur_bac_conv}, nam_conv={cur_nam_conv}, "
      f"bac_rev={cur_bac_rev}, nam_rev={cur_nam_rev}")

# 4 cờ đã gửi (per miền × loại) — đọc từ state
state_bac_conv = bool((state.get("bac_conv_at") or "").strip())
state_nam_conv = bool((state.get("nam_conv_at") or "").strip())
state_bac_rev = bool((state.get("bac_rev_at") or "").strip())
state_nam_rev = bool((state.get("nam_rev_at") or "").strip())
state_initial = bool((state.get("initial_sent_at") or "").strip())
print(f"  Đã gửi: initial={state_initial}, bac_conv={state_bac_conv}, "
      f"nam_conv={state_nam_conv}, bac_rev={state_bac_rev}, nam_rev={state_nam_rev}")

# Phát hiện flip: data có nhưng state chưa ghi nhận → cần gửi cập nhật
new_bac_conv = cur_bac_conv and not state_bac_conv
new_nam_conv = cur_nam_conv and not state_nam_conv
new_bac_rev = cur_bac_rev and not state_bac_rev
new_nam_rev = cur_nam_rev and not state_nam_rev
any_new = new_bac_conv or new_nam_conv or new_bac_rev or new_nam_rev

now_str = datetime.now(zoneinfo.ZoneInfo("Asia/Ho_Chi_Minh")).strftime("%Y-%m-%d %H:%M")

# Quyết định gửi:
# 1. Force: luôn gửi
# 2. Initial chưa gửi + không có data: gửi initial
# 3. Có data mới (flip): gửi báo cáo cập nhật
need_send = args.force or any_new or (not state_initial and not has_conv and not has_rev)

if need_send:
    new_flags = {
        "bac_conv": new_bac_conv, "nam_conv": new_nam_conv,
        "bac_rev": new_bac_rev, "nam_rev": new_nam_rev,
    }
    print(f"\n→ {'FORCE' if args.force else 'CẬP NHẬT MỚI'} — gửi báo cáo...")
    print(f"  Cờ mới: {new_flags}")

    if not has_conv and not has_rev:
        msg = build_initial_msg()
    else:
        msg = build_full_msg(has_conv=has_conv, has_rev=has_rev, new_flags=new_flags)

    if send_telegram(msg):
        update_kwargs = {"last_check_at": now_str}
        # Lần đầu — populate metadata
        if not state:
            update_kwargs.update({
                "thang": thang, "tuan": tuan_trong_thang,
                "ngay_bd": ngay_bd_str, "ngay_kt": ngay_kt_str,
            })
        if not state_initial:
            update_kwargs["initial_sent_at"] = now_str
        if new_bac_conv: update_kwargs["bac_conv_at"] = now_str
        if new_nam_conv: update_kwargs["nam_conv_at"] = now_str
        if new_bac_rev: update_kwargs["bac_rev_at"] = now_str
        if new_nam_rev: update_kwargs["nam_rev_at"] = now_str
        upsert_state(ws_state, week_id, **update_kwargs)
else:
    print(f"\n→ Không có data mới so với state. Skip Telegram.")
    upsert_state(ws_state, week_id, last_check_at=now_str)

print(f"\n✓ Done!")
