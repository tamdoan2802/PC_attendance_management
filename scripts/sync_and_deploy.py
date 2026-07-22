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
    print(f"\n[BƯỚC 2] Làm mới PowerQuery trong {EXCEL_PATH.name}...")
    if not EXCEL_PATH.exists():
        print(f"[LỖI] Không tìm thấy file Excel tại {EXCEL_PATH}")
        return

    xl = None
    wb = None
    try:
        xl = win32com.client.Dispatch("Excel.Application")
        xl.Visible = False # Chạy ngầm
        xl.DisplayAlerts = False # Tắt popup cảnh báo
        wb = xl.Workbooks.Open(str(EXCEL_PATH))
        
        print("Đang refresh các QueryTable liên quan đến Request...")
        for sheet in wb.Sheets:
            for lo in sheet.ListObjects:
                if "Req_" in lo.Name:
                    try:
                        # Ép refresh đồng bộ từng bảng
                        lo.QueryTable.Refresh(BackgroundQuery=False)
                        print(f" Đã refresh bảng {lo.Name} trên sheet {sheet.Name}")
                    except Exception:
                        pass
                    
        print("Đang refresh các Connection liên quan đến Request...")
        for conn in wb.Connections:
            if "Req_" in conn.Name:
                try:
                    conn.Refresh()
                    print(f" Đã refresh connection {conn.Name}")
                except Exception:
                    pass
                
        print("Làm mới PowerQuery xong. Đang lưu Excel...")
        wb.Save()
        print("Lưu Excel thành công.")
    except Exception as e:
        print(f"[LỖI] Lỗi khi Refresh Excel: {e}")
    finally:
        if wb:
            wb.Close(False)
        if xl:
            xl.DisplayAlerts = True
            xl.Quit()

def step_3_generate_json():
    print("\n[BƯỚC 3] Chạy ETL script để tạo data.json...")
    try:
        subprocess.run([sys.executable, str(ETL_SCRIPT)], check=True, cwd=str(THIS_DIR))
    except subprocess.CalledProcessError:
        print("[LỖI] Không thể tạo file JSON. Dừng quá trình.")
        sys.exit(1)

def step_4_git_deploy():
    print("\n[BƯỚC 4] Deploy lên GitHub Pages...")
    try:
        # cd vào thư mục gốc của repo
        repo_dir = str(BASE_DIR)
        
        print("Git add...")
        subprocess.run(["git", "add", "reports/index.html", "references/data.json"], cwd=repo_dir, check=True)
        
        print("Git commit...")
        # Lấy thời gian hiện tại làm message
        commit_msg = f"Auto-sync from MISA at {time.strftime('%Y-%m-%d %H:%M:%S')}"
        res = subprocess.run(["git", "commit", "-m", commit_msg], cwd=repo_dir, capture_output=True, text=True)
        if "nothing to commit" in res.stdout:
            print("Không có thay đổi dữ liệu nào mới để commit.")
            return

        print("Git push...")
        subprocess.run(["git", "push"], cwd=repo_dir, check=True)
        print("Đã deploy lên GitHub Pages thành công!")
    except FileNotFoundError:
        print("[LỖI] Lệnh 'git' không tồn tại. Vui lòng cài đặt Git và cấu hình Github Desktop.")
    except subprocess.CalledProcessError as e:
        print(f"[CẢNH BÁO] Lỗi trong quá trình Git commit/push: {e}")

if __name__ == "__main__":
    print("=" * 60)
    print("      TỰ ĐỘNG HÓA WORKFORCE DASHBOARD (1-CLICK)")
    print("=" * 60)
    
    step_1_download_misa()
    step_2_refresh_excel()
    step_3_generate_json()
    step_4_git_deploy()
    
    print("\n" + "=" * 60)
    print("                 HOÀN TẤT!")
    print("=" * 60)
    time.sleep(3)
