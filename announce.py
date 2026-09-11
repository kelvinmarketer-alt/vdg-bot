"""Gửi 1 tin THÔNG BÁO lịch báo cáo tự động cho cả nhóm (bot VĐG).
Chạy tay qua workflow 'VDG Announce'. Dùng chung send_telegram của care_report."""
import os
import care_report as c

MSG = """📢 *THÔNG BÁO — LỊCH BÁO CÁO TỰ ĐỘNG*

Từ nay hệ thống bot sẽ tự tổng hợp dữ liệu từ *Sheet chăm sóc KH* và gửi báo cáo định kỳ vào nhóm này. Lịch cụ thể:

📋 *1. Báo cáo HẰNG NGÀY* — 9h30 sáng
   Tổng hợp dữ liệu của *ngày hôm trước*: số lượt chăm sóc, loại KH, nhóm sản phẩm, doanh số, *xếp hạng nhân viên* và tỉ lệ hoàn thành KPI (doanh số, KH tự tìm, KH cũ).

📅 *2. Báo cáo HẰNG TUẦN* — 9h30 sáng *thứ 2*
   Tổng kết *cả tuần vừa qua* (thứ 2 → chủ nhật), cộng dồn xếp hạng & KPI của từng nhân viên.

🗓️ *3. Báo cáo HẰNG THÁNG* — 9h30 sáng *ngày 1*
   Tổng kết *cả tháng vừa qua*.

⚠️ *Lưu ý quan trọng cho nhân viên:*
Vui lòng cập nhật ĐẦY ĐỦ và ĐÚNG NGÀY:
• Sheet *"Báo cáo chi tiết"* — ghi từng lượt chăm sóc khách (nguồn, nhóm SP, mức độ, doanh số...).
• Tab *"đánh giá"* — điền doanh số thực tế, mục tiêu, ngày làm / nghỉ.

Nếu không cập nhật, báo cáo sẽ hiển thị thiếu dữ liệu hoặc 0 — ảnh hưởng đến kết quả đánh giá của chính mình.

_Bot chạy tự động, số liệu lấy trực tiếp từ sheet nên hãy nhập liệu trung thực & kịp thời._ 🙏"""


def main():
    if not os.environ.get("TELEGRAM_BOT_TOKEN"):
        print("⚠️ Thiếu TELEGRAM_BOT_TOKEN")
        return
    c.send_telegram(MSG)
    print("✓ Đã gửi thông báo lịch báo cáo")


if __name__ == "__main__":
    main()
