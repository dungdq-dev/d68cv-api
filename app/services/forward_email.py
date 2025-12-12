import win32com.client
import re
from datetime import datetime, timedelta
import os


# Lấy thư mục hiện tại
cwd = os.getcwd()
print(f"Current working directory: {cwd}")


def forward_emails_by_subject(
    subject_pattern, forward_to, days_back=7, include_attachments=True, add_note=None
):
    """
    Tự động forward email dựa trên tiêu đề

    Parameters:
    - subject_pattern: Chuỗi hoặc regex pattern để tìm kiếm tiêu đề
    - forward_to: Danh sách email người nhận (string hoặc list)
    - days_back: Số ngày quay lại để tìm email
    - include_attachments: Có forward file đính kèm không
    - add_note: Thêm ghi chú vào đầu email forwarded
    """

    # Kết nối với Outlook
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    inbox = namespace.GetDefaultFolder(6)  # 6 = Inbox folder

    # Tính ngày bắt đầu tìm kiếm
    start_date = datetime.now() - timedelta(days=days_back)

    # Chuyển đổi sang định dạng Outlook
    filter_date = start_date.strftime("%m/%d/%Y %H:%M %p")

    # Tạo filter tìm kiếm
    filter_str = f"[ReceivedTime] >= '{filter_date}'"

    # Lấy email theo filter
    items = inbox.Items.Restrict(filter_str)
    items.Sort("[ReceivedTime]", True)  # Sắp xếp mới nhất đầu tiên

    forwarded_count = 0

    for item in items:
        try:
            # Kiểm tra tiêu đề
            subject = item.Subject

            # Sử dụng regex hoặc chuỗi con để tìm kiếm
            if isinstance(subject_pattern, str):
                # Tìm kiếm chuỗi đơn giản
                if subject_pattern.lower() in subject.lower():
                    should_forward = True
                else:
                    should_forward = False
            else:
                # Sử dụng regex pattern
                should_forward = bool(subject_pattern.search(subject))

            if should_forward:
                print(f"Đang forward email: {subject}")

                # Tạo email forward
                forward = item.Forward()

                # Thiết lập người nhận
                if isinstance(forward_to, list):
                    for recipient in forward_to:
                        forward.Recipients.Add(recipient)
                else:
                    forward.Recipients.Add(forward_to)

                # Thêm ghi chú nếu có
                if add_note:
                    forward.Body = add_note + "\n\n" + forward.Body

                # Giữ nguyên file đính kèm
                if include_attachments and item.Attachments.Count > 0:
                    for attachment in item.Attachments:
                        attachment.SaveAsFile(f"{cwd}\\temp\\{attachment.FileName}")
                        forward.Attachments.Add(f"{cwd}\\temp\\{attachment.FileName}")

                # Gửi email
                forward.Send()
                print(f"Đã forward email: {subject}")
                forwarded_count += 1

        except Exception as e:
            print(f"Lỗi khi xử lý email: {str(e)}")
            continue

    print(f"Hoàn thành! Đã forward {forwarded_count} email.")
    return forwarded_count


# Ví dụ sử dụng với chuỗi tìm kiếm đơn giản
def forward_emails_with_string():
    """Forward email chứa từ khóa cụ thể trong tiêu đề"""
    subject_keyword = "Báo cáo hàng ngày"
    recipients = ["colleague1@company.com", "colleague2@company.com"]

    forward_emails_by_subject(
        subject_pattern=subject_keyword,
        forward_to=recipients,
        days_back=30,
        add_note="Email được forward tự động từ hệ thống\n"
        + f"Thời gian: {datetime.now().strftime('%Y-%m-%d %H:%M')}",
    )


# Ví dụ sử dụng với regex pattern
def forward_emails_with_regex():
    """Forward email sử dụng regex pattern"""
    import re

    # Regex pattern: tìm email có chứa "urgent" hoặc "important"
    pattern = re.compile(r"urgent|important", re.IGNORECASE)

    forward_emails_by_subject(
        subject_pattern=pattern,
        forward_to="manager@company.com",
        days_back=14,
        add_note="[AUTO-FORWARDED] Email quan trọng",
    )


# Script tự động chạy theo lịch trình
def schedule_forwarding():
    """Chạy tự động theo lịch trình"""
    import schedule
    import time

    def job():
        print(f"Bắt đầu tự động forward email - {datetime.now()}")
        forward_emails_with_string()

    # Lên lịch chạy hàng ngày vào 9:00 sáng
    schedule.every().day.at("09:00").do(job)

    print("Đã lên lịch tự động forward email hàng ngày lúc 9:00 AM")

    while True:
        schedule.run_pending()
        time.sleep(60)


# Giao diện dòng lệnh đơn giản
def main():
    """Giao diện tương tác với người dùng"""
    print("=== Outlook Auto-Forward Tool ===")

    # Nhập thông tin từ người dùng
    subject_pattern = input("Nhập từ khóa hoặc regex pattern cho tiêu đề: ")
    recipients = input("Nhập email người nhận (cách nhau bằng dấu phẩy): ")
    days = int(input("Tìm email trong bao nhiêu ngày gần đây: "))

    # Xử lý danh sách người nhận
    recipient_list = [email.strip() for email in recipients.split(",")]

    # Chạy forward
    forward_emails_by_subject(
        subject_pattern=subject_pattern,
        forward_to=recipient_list,
        days_back=days,
        add_note=f"Auto-forwarded from Outlook\nDate: {datetime.now().strftime('%Y-%m-%d')}",
    )


if __name__ == "__main__":
    # Chọn một trong các hàm để chạy:

    # 1. Forward với từ khóa đơn giản
    # forward_emails_with_string()

    # 2. Forward với regex
    # forward_emails_with_regex()

    # 3. Chạy với giao diện CLI
    main()

    # 4. Chạy tự động theo lịch trình (chạy nền)
    # schedule_forwarding()
