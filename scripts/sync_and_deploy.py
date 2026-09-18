import os
import sys
import time
import subprocess
import win32com.client
from pathlib import Path

# --- CẤU HÌNH ĐƯỜNG DẪN ---
THIS_DIR = Path(__file__).parent.resolve()
BASE_DIR = THIS_DIR.parent
EXCEL_PATH = Path(r"G:\My Drive\Dữ liệu nhân sự\Data\Timesheet\HR_Fact_Attendance.xlsx")
MISA_SCRIPT = THIS_DIR / "misa_download_automation.py"
ETL_SCRIPT = THIS_DIR / "generate_dashboard_data.py"

def step_1_download_misa():
    print("\n[BƯỚC 1] Tải dữ liệu từ MISA...")
    try:
        subprocess.run([sys.executable, str(MISA_SCRIPT)], check=True)
        print("Hoàn tất tải dữ liệu MISA.")
    except subprocess.CalledProcessError:
        print("[CẢNH BÁO] Script tải MISA bị lỗi. Vẫn tiếp tục xử lý với dữ liệu cũ...")

def step_2_refresh_excel():
    print("\n[BƯỚC 2] Kiến trúc mới đã đọc trực tiếp 100% từ Raw Timesheets & Entities.")
    print("         Bỏ qua bước làm mới PowerQuery Excel trung gian.")

def step_3_generate_json():
    print("\n[BƯỚC 3] Chạy ETL script để tạo data.json...")
    try:
        subprocess.run([sys.executable, str(ETL_SCRIPT)], check=True, cwd=str(THIS_DIR))
    except subprocess.CalledProcessError:
        print("[LỖI] Không thể tạo file JSON. Dừng quá trình.")
        sys.exit(1)

def get_git_executable():
    import shutil
    # Nếu git đã có trong PATH thì dùng luôn
    if shutil.which("git"):
        return "git"
    
    # Nếu không, tự động tìm git đi kèm với GitHub Desktop
    github_desktop_dir = Path(os.environ.get("LOCALAPPDATA", "")) / "GitHubDesktop"
    if github_desktop_dir.exists():
        # Tìm các thư mục app-*
        app_dirs = list(github_desktop_dir.glob("app-*"))
        if app_dirs:
            # Lấy bản app mới nhất
            latest_app = sorted(app_dirs)[-1]
            git_path = latest_app / "resources" / "app" / "git" / "cmd" / "git.exe"
            if git_path.exists():
                return str(git_path)
    return "git"

def step_4_git_deploy():
    print("\n[BƯỚC 4] Deploy lên GitHub Pages...")
    try:
        repo_dir = str(BASE_DIR)
        git_cmd = get_git_executable()
        
        print("Git add...")
        subprocess.run([git_cmd, "add", "reports/index.html", "references/data.json", "references/data.js", "scripts/generate_dashboard_data.py"], cwd=repo_dir, check=True)
        
        print("Git commit...")
        commit_msg = f"Auto-sync from MISA at {time.strftime('%Y-%m-%d %H:%M:%S')}"
        res = subprocess.run([git_cmd, "commit", "-m", commit_msg], cwd=repo_dir, capture_output=True, text=True)
        if "nothing to commit" in res.stdout:
            print("Không có thay đổi dữ liệu nào mới để commit.")
            return

        print("Git push...")
        subprocess.run([git_cmd, "push"], cwd=repo_dir, check=True)
        print("Đã deploy lên GitHub Pages thành công!")
    except FileNotFoundError:
        print("[LỖI] Lệnh 'git' không tồn tại. Vui lòng cài đặt Git và cấu hình Github Desktop.")
    except subprocess.CalledProcessError as e:
        print(f"[CẢNH BÁO] Lỗi trong quá trình Git commit/push: {e}")

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Tự động hóa Workforce Dashboard")
    parser.add_argument("--step", type=str, default="all", help="Chọn bước để chạy (1, 2, 3, 4, hoặc all)")
    args = parser.parse_args()

    print("=" * 60)
    print("      TỰ ĐỘNG HÓA WORKFORCE DASHBOARD")
    print("=" * 60)
    
    if args.step in ["1", "all"]:
        step_1_download_misa()
    if args.step in ["2", "all"]:
        step_2_refresh_excel()
    if args.step in ["3", "all"]:
        step_3_generate_json()
    if args.step in ["4", "all"]:
        step_4_git_deploy()
    
    print("\n" + "=" * 60)
    print("                 HOÀN TẤT!")
    print("=" * 60)
    if args.step == "all":
        time.sleep(3)
