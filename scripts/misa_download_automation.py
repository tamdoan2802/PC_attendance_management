import os
import sys
import time
from pathlib import Path

try:
    from playwright.sync_api import sync_playwright
except ImportError:
    sys.exit("Vui lòng cài đặt playwright: pip install playwright && playwright install chromium")

# --- CẤU HÌNH ---
# Nơi lưu trữ session (giữ trạng thái đăng nhập)
USER_DATA_DIR = Path(os.environ["LOCALAPPDATA"]) / "Playwright_MISA_Profile"
# Thư mục lưu file tải về
DOWNLOAD_DIR = Path(r"G:\My Drive\Dữ liệu nhân sự\Data")

# Tên các file cần tải (theo yêu cầu)
TARGET_FILES = {
    "leave": "Don_xin_nghi_Tất cả đơn vị.xlsx",
    "ot": "Đơn_đăng_ký_làm_thêm_Tất cả đơn vị.xlsx",
    "wfh": "Dang_ky_lam_viec_tu_xa_Tất cả đơn vị.xlsx",
    "lcec": "Đăng ký đi muộn, về sớm_Tất cả đơn vị.xlsx",
    "trip": "De_nghi_di_cong_tac_Tất cả đơn vị.xlsx",
    "shift": "De_nghi_doi_ca_Tất cả đơn vị.xlsx"
}

def download_misa_files():
    print("=" * 50)
    print("Khởi động tự động hóa tải file MISA AMIS...")
    print("=" * 50)

    with sync_playwright() as p:
        # Khởi chạy Chromium ở chế độ CÓ GIAO DIỆN để giữ session đăng nhập
        browser = p.chromium.launch_persistent_context(
            user_data_dir=USER_DATA_DIR,
            headless=False, # Đặt thành True nếu bạn muốn chạy ngầm hoàn toàn sau khi đã đăng nhập
            accept_downloads=True,
            viewport={"width": 1280, "height": 720}
        )
        
        page = browser.pages[0] if browser.pages else browser.new_page()

        print("[1] Đang mở MISA AMIS...")
        # URL đăng nhập MISA (bạn có thể đổi lại link chính xác của công ty)
        page.goto("https://amisapp.misa.vn/home")
        
        print("[2] Vui lòng kiểm tra trình duyệt. Nếu chưa đăng nhập, hãy đăng nhập tay.")
        print("    (Script sẽ chờ 10 giây. Ở các lần sau nó sẽ tự nhớ session).")
        page.wait_for_timeout(10000) 

        print("\n[3] Bắt đầu tải các file dữ liệu...")
        
        # Danh sách các phân hệ cần tải
        requests_config = [
            ("leave", "Đơn xin nghỉ", "https://amisapp.misa.vn/timesheet/management-request/attendance-watch"),
            ("ot", "Đơn OT", "https://amisapp.misa.vn/timesheet/management-request/register-overtime"),
            ("wfh", "Đơn WFH", "https://amisapp.misa.vn/timesheet/management-request/work-remote"),
            ("lcec", "Đơn Đi muộn về sớm", "https://amisapp.misa.vn/timesheet/management-request/late-in-early-out"),
            ("trip", "Đơn Công tác", "https://amisapp.misa.vn/timesheet/management-request/mission-allowance"),
            ("shift", "Đơn Đổi ca", "https://amisapp.misa.vn/timesheet/management-request/change-shift")
        ]
        
        for key, name, url in requests_config:
            try:
                print(f" -> Đang tải file {name}...")
                page.goto(url)
                
                # Chờ một chút để MISA tải xong giao diện bảng (nếu cần)
                page.wait_for_timeout(2000)
                
                with page.expect_download(timeout=90000) as download_info:
                    page.locator(".btn-sidebar > .mi-export").click()
                download = download_info.value
                
                file_path = DOWNLOAD_DIR / TARGET_FILES[key]
                if file_path.exists():
                    try:
                        file_path.unlink()
                    except PermissionError:
                        print(f"   [CẢNH BÁO] Không thể ghi đè {TARGET_FILES[key]} vì file đang mở.")
                        
                download.save_as(str(file_path))
                print(f"   [OK] Đã tải xong: {TARGET_FILES[key]}")
                
            except Exception as e:
                print(f"   [LỖI] Không thể tải {name}: {e}")

        print("\n[4] Đóng trình duyệt.")
        page.wait_for_timeout(3000)
        browser.close()

if __name__ == "__main__":
    download_misa_files()
