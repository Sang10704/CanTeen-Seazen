# -*- coding: utf-8 -*-
"""
Kịch bản đóng gói ứng dụng SEAZEN Canteen Kiosk thành 1 file EXE độc lập duy nhất.
Chỉ cần gửi 1 file SEAZEN_Canteen_Kiosk.exe đi là chạy được trên máy trạm.
"""

import os
import sys
import shutil
import subprocess

if hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SPEC_FILE = os.path.join(BASE_DIR, "SEAZEN_Canteen_Kiosk.spec")
DIST_DIR = os.path.join(BASE_DIR, "dist")
OUTPUT_EXE = os.path.join(DIST_DIR, "SEAZEN_Canteen_Kiosk.exe")
ROOT_EXE = os.path.join(BASE_DIR, "SEAZEN_Canteen_Kiosk.exe")
README_FILE = os.path.join(BASE_DIR, "HUONG_DAN_SU_DUNG.txt")

def main():
    print("=" * 70)
    print("   DONG GOI SEAZEN CANTEEN KIOSK - BAN STANDALONE 1 FILE EXE")
    print("=" * 70)

    # 1. Thực thi PyInstaller
    print("[*] Dang bien dich ma nguon voi PyInstaller...")
    cmd = [sys.executable, "-m", "PyInstaller", "--clean", "-y", SPEC_FILE]
    res = subprocess.run(cmd, cwd=BASE_DIR)

    if res.returncode != 0 or not os.path.exists(OUTPUT_EXE):
        print(f"[-] Bien dich that bai! Ma loi: {res.returncode}")
        sys.exit(1)

    exe_size_mb = os.path.getsize(OUTPUT_EXE) / (1024 * 1024)
    print(f"[+] Bien dich thanh cong: {OUTPUT_EXE} ({exe_size_mb:.2f} MB)")

    # 2. Sao chép file EXE ra thư mục gốc
    try:
        shutil.copy2(OUTPUT_EXE, ROOT_EXE)
        print(f"[+] Da cap nhat file EXE tai thu muc goc: {ROOT_EXE}")
    except PermissionError:
        print(f"[!] File {ROOT_EXE} dang duoc su dung boi tien trinh khac.")

    # 3. Tạo tài liệu hướng dẫn vận hành súc tích
    readme_content = """======================================================================
         HE THONG DIEM DANH CANTEEN SEAZEN (BAN DOC LAP 1 FILE EXE)
======================================================================

1. CACH SU DUNG:
- Chi can gui duy nhat tep tin: SEAZEN_Canteen_Kiosk.exe toi may tram.
- May tinh nhan KHONG CAN cai dat Python, thu vien hay bat ky tep phu tro nao.
- Nhap dup chuot vao SEAZEN_Canteen_Kiosk.exe de khoi chay he thong.
- Ung dung tu dong khoi chay may chu noi bo va mo trinh duyet tai:
  http://localhost:5000

2. XAC THUC TAI KHOAN MICROSOFT 365 (BAT BUOC):
- Khi chay lan dau, Kiosk hien thi man hinh yeu cau xac thuc tai khoan Microsoft.
- Chon "DANG NHAP BANG MICROSOFT 365" de xac thuc tai khoan @seazen.com.vn.
- Token duoc tu dong ma hoa bang Windows DPAPI va luu tru an tai:
  %LOCALAPPDATA%\\SeaZen_Canteen_Kiosk\\token_cache.bin
- He thong tu dong gia han token ngam, khong doi hoi dang nhap lai hang ngay.

3. QUY TAC NHAN DIEN MA NHAN VIEN:
- Ho tro chuan hoa ma nhan vien dang so: loai bo cac so 0 o dau.
- Dac quyen VIP Tong Giam Doc DO VAN TOAN:
  Quet voi bat ky bien the nao: 1, 01, 001, 0001, 00001 deu nhan dien chinh xac
  la 1 nguoi duy nhat: Tong Giam Doc DO VAN TOAN (Ma NV: 1).
- Tu dong phat am thanh thong bao va duyet nhan com tuc thi.
======================================================================
"""
    with open(README_FILE, "w", encoding="utf-8") as f:
        f.write(readme_content)
    print(f"[+] Da cap nhat: {README_FILE}")

    print("=" * 70)
    print(f"[OK] HOAN TAT: Ban dong goi SEAZEN_Canteen_Kiosk.exe da san sang gui di!")
    print("=" * 70)

if __name__ == '__main__':
    main()
