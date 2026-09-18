import os
import sys
import time
from pathlib import Path

# Đảm bảo in tiếng Việt chuẩn UTF-8 trên Windows console
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")

try:
    from playwright.sync_api import sync_playwright
except ImportError:
    sys.exit("Vui lòng cài đặt playwright: pip install playwright && playwright install chromium")

# --- CẤU HÌNH ---
# Nơi lưu trữ session (giữ trạng thái đăng nhập)
USER_DATA_DIR = Path(os.environ["LOCALAPPDATA"]) / "Playwright_MISA_Profile"
# Thư mục lưu file tải về
DOWNLOAD_DIR = Path(r"G:\My Drive\Dữ liệu nhân sự\Data")

# Tên các file cần tải (theo yêu cầu hệ thống)
TARGET_FILES = {
    "leave": "Don_xin_nghi_Tất cả đơn vị.xlsx",
    "ot": "Đơn_đăng_ký_làm_thêm_Tất cả đơn vị.xlsx",
    "wfh": "Dang_ky_lam_viec_tu_xa_Tất cả đơn vị.xlsx",
    "lcec": "Đăng ký đi muộn, về sớm_Tất cả đơn vị.xlsx",
    "trip": "De_nghi_di_cong_tac_Tất cả đơn vị.xlsx",
    "shift": "De_nghi_doi_ca_Tất cả đơn vị.xlsx"
}

REQUESTS_CONFIG = [
    ("leave", "Đơn xin nghỉ", "https://amisapp.misa.vn/timesheet/management-request/attendance-watch"),
    ("ot", "Đơn OT", "https://amisapp.misa.vn/timesheet/management-request/register-overtime"),
    ("wfh", "Đơn WFH", "https://amisapp.misa.vn/timesheet/management-request/work-remote"),
    ("lcec", "Đơn Đi muộn về sớm", "https://amisapp.misa.vn/timesheet/management-request/late-in-early-out"),
    ("trip", "Đơn Công tác", "https://amisapp.misa.vn/timesheet/management-request/mission-allowance"),
    ("shift", "Đơn Đổi ca", "https://amisapp.misa.vn/timesheet/management-request/change-shift")
]


def check_and_wait_for_login(page, max_wait_sec=300):
    """
    Kiểm tra xem trang hiện tại có phải trang đăng nhập không.
    Nếu có, thông báo và chờ người dùng đăng nhập thành công.
    """
    current_url = page.url
    if "login" in current_url or "id.misa.vn" in current_url:
        print("\n" + "!" * 65)
        print(" [CHÚ Ý] PHIÊN ĐĂNG NHẬP MISA ĐÃ HẾT HẠN HOẶC CHƯA ĐĂNG NHẬP!")
        print(" Vui lòng đăng nhập vào MISA trên cửa sổ trình duyệt vừa mở...")
        print(" Hệ thống sẽ tự động nhận diện và tiếp tục ngay khi bạn đăng nhập xong.")
        print(f" (Thời gian chờ tối đa: {max_wait_sec // 60} phút)")
        print("!" * 65 + "\n")

        start_time = time.time()
        last_notified = 0
        while time.time() - start_time < max_wait_sec:
            page.wait_for_timeout(2000)
            url = page.url
            if "amisapp.misa.vn" in url and "login" not in url and "id.misa.vn" not in url:
                print("\n[✓ THÀNH CÔNG] Đã phát hiện đăng nhập MISA thành công!")
                page.wait_for_timeout(3000)
                return True

            elapsed = int(time.time() - start_time)
            if elapsed - last_notified >= 15:
                print(f" ...Đang chờ bạn hoàn tất đăng nhập MISA ({elapsed}/{max_wait_sec}s)...")
                last_notified = elapsed

        print("\n[LỖI] Đã hết thời gian chờ đăng nhập (5 phút). Dừng tiến trình.")
        return False
    return True


def download_single_request(page, key, name, url):
    """Tải một file phân hệ từ MISA AMIS."""
    print(f"\n -> Đang tải dữ liệu: {name}...")
    try:
        page.goto(url, wait_until="domcontentloaded", timeout=45000)
    except Exception as e:
        print(f"   [CẢNH BÁO] Tải trang {name} mất nhiều thời gian: {e}")

    # Kiểm tra nếu bị văng ra trang đăng nhập
    if "login" in page.url or "id.misa.vn" in page.url:
        print(f"   [CHÚ Ý] Cần đăng nhập để xem phân hệ {name}...")
        if not check_and_wait_for_login(page):
            return False
        # Điều hướng lại sau khi đã login
        page.goto(url, wait_until="domcontentloaded", timeout=45000)

    # Chờ loading spinner nếu có
    try:
        page.wait_for_selector(".dx-loadpanel, .loading-modal, .loading-mask", state="detached", timeout=8000)
    except Exception:
        pass

    # Đợi bảng ổn định
    page.wait_for_timeout(2500)

    # Danh sách selector ưu tiên cho nút Xuất khẩu trên MISA
    export_selectors = [
        ".btn-sidebar > .mi-export",
        ".btn-sidebar:has(.mi-export)",
        "button:has(.mi-export)",
        "[title*='Xuất khẩu']",
        "[title*='xuất khẩu']",
        "[aria-label*='Xuất khẩu']",
        ".mi-export"
    ]

    export_btn = None
    for sel in export_selectors:
        loc = page.locator(sel).first
        try:
            if loc.is_visible(timeout=2000):
                export_btn = loc
                break
        except Exception:
            continue

    if not export_btn:
        print(f"   [CẢNH BÁO] Không tìm thấy nút Xuất khẩu trên trang {name}. URL: {page.url}")
        return False

    file_path = DOWNLOAD_DIR / TARGET_FILES[key]

    # Thực hiện click và bắt download
    try:
        with page.expect_download(timeout=60000) as download_info:
            export_btn.click()
        download = download_info.value

        # Nếu file đích đang tồn tại, xóa trước để tránh trùng lặp
        if file_path.exists():
            try:
                file_path.unlink()
            except PermissionError:
                print(f"   [CẢNH BÁO] Không thể ghi đè {TARGET_FILES[key]} vì file đang mở trong Excel.")
                backup_name = f"{file_path.stem}_new{file_path.suffix}"
                backup_path = DOWNLOAD_DIR / backup_name
                download.save_as(str(backup_path))
                print(f"   [TẠM THỜI] Đã lưu vào '{backup_name}'. Hãy đóng Excel và đổi tên lại.")
                return True

        download.save_as(str(file_path))
        print(f"   [OK] Đã tải xong: {TARGET_FILES[key]}")
        return True

    except Exception as e:
        print(f"   [LỖI] Không thể tải file {name}: {e}")
        return False


def download_misa_files():
    print("=" * 60)
    print("      TỰ ĐỘNG HÓA TẢI FILE DỮ LIỆU TỪ MISA AMIS")
    print("=" * 60)

    # Đảm bảo thư mục lưu trữ tồn tại
    DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as p:
        # Khởi chạy Chromium ở chế độ CÓ GIAO DIỆN (headless=False)
        # Giúp người dùng đăng nhập dễ dàng và quan sát tiến trình
        browser = p.chromium.launch_persistent_context(
            user_data_dir=USER_DATA_DIR,
            headless=False,
            accept_downloads=True,
            viewport={"width": 1366, "height": 768},
            args=[
                "--disable-features=DownloadBubble,DownloadBubbleV2",
                "--no-sandbox"
            ]
        )

        page = browser.pages[0] if browser.pages else browser.new_page()

        print("\n[1] Đang mở MISA AMIS và kiểm tra trạng thái đăng nhập...")
        try:
            page.goto("https://amisapp.misa.vn/home", wait_until="domcontentloaded", timeout=45000)
        except Exception:
            pass

        # Kiểm tra đăng nhập
        if not check_and_wait_for_login(page):
            print("\n[LỖI] Chưa thể đăng nhập vào MISA. Dừng tiến trình tải.")
            browser.close()
            sys.exit(1)

        print("\n[2] Bắt đầu tải các file dữ liệu...")
        success_count = 0
        total_count = len(REQUESTS_CONFIG)

        for key, name, url in REQUESTS_CONFIG:
            ok = download_single_request(page, key, name, url)
            if ok:
                success_count += 1
            # Nghỉ ngắn giữa các lần tải để hệ thống không bị dồn dập
            page.wait_for_timeout(1500)

        print("\n" + "=" * 60)
        print(f" KẾT QUẢ: Tải thành công {success_count}/{total_count} file dữ liệu.")
        print("=" * 60)

        print("\n[3] Đóng trình duyệt sau 3 giây...")
        page.wait_for_timeout(3000)
        browser.close()

        if success_count == 0:
            sys.exit(1)


if __name__ == "__main__":
    download_misa_files()

