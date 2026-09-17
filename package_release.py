import os
import shutil
import zipfile

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PACKAGE_DIR = os.path.join(BASE_DIR, "SEAZEN_Canteen_Kiosk_Package")
ZIP_OUTPUT = os.path.join(BASE_DIR, "SEAZEN_Canteen_Kiosk_Full_Package.zip")

print("[*] Bat dau dong goi...")

# 1. Tao thu muc package
os.makedirs(PACKAGE_DIR, exist_ok=True)
try:
    for item in os.listdir(PACKAGE_DIR):
        item_path = os.path.join(PACKAGE_DIR, item)
        if item != "SEAZEN_Canteen_Kiosk.exe":
            if os.path.isdir(item_path):
                shutil.rmtree(item_path, ignore_errors=True)
            else:
                try:
                    os.remove(item_path)
                except Exception:
                    pass
except Exception:
    pass

# 2. Sao chep cac file can thiet
files_to_copy = [
    (os.path.join(BASE_DIR, "dist", "SEAZEN_Canteen_Kiosk.exe"), "SEAZEN_Canteen_Kiosk.exe"),
    (os.path.join(BASE_DIR, "token_cache.bin"), "token_cache.bin"),
    (os.path.join(BASE_DIR, "Canteen_Kiosk_Production.html"), "Canteen_Kiosk_Production.html"),
    (os.path.join(BASE_DIR, "CHAY_CANTEEN_KIOSK.bat"), "CHAY_CANTEEN_KIOSK.bat"),
    (os.path.join(BASE_DIR, "canteen_icon.ico"), "canteen_icon.ico"),
    (os.path.join(BASE_DIR, "tao_token_moi.py"), "tao_token_moi.py"),
    (os.path.join(BASE_DIR, "DANG_NHAP_TAI_KHOAN_MOI.bat"), "DANG_NHAP_TAI_KHOAN_MOI.bat"),
    (os.path.join(BASE_DIR, "MO_CONG_FIREWALL_LAN.bat"), "MO_CONG_FIREWALL_LAN.bat"),
]

for src, dst_name in files_to_copy:
    if os.path.exists(src):
        dst = os.path.join(PACKAGE_DIR, dst_name)
        try:
            shutil.copy2(src, dst)
            print(f" [+] Da copy: {dst_name}")
        except PermissionError:
            print(f" [!] Bo qua {dst_name} (file dang chay boi ung dung)")
    else:
        print(f" [!] Khong tim thay: {src}")

# Copy thu muc mp3
src_mp3 = os.path.join(BASE_DIR, "mp3")
dst_mp3 = os.path.join(PACKAGE_DIR, "mp3")
if os.path.exists(src_mp3):
    if os.path.exists(dst_mp3):
        shutil.rmtree(dst_mp3)
    shutil.copytree(src_mp3, dst_mp3)
    print(" [+] Da copy thu muc: mp3")

# 3. Tao file HUONG_DAN_SU_DUNG.txt
readme_text = """===================================================================
      HE THONG DIEM DANH CANTEEN SEAZEN (PORTABLE STANDALONE)
===================================================================

[1. HUONG DAN SU DUNG CHO MAY MOI / MAY KHAC]:
1. May tinh nhan KHONG CAN cai dat Python hay bat ky phan mem nao khac.
2. Chi can giai nen toan bo thu muc nay ra mot o dia (Vi du: C:\\Canteen_Kiosk hoac Desktop).
3. Nhap dup chuot vao file: SEAZEN_Canteen_Kiosk.exe de khoi chay.
4. He thong se tu dong bat trinh duyet hien thi man hinh Kiosk toan man hinh (hoac truy cap: http://localhost:5000).
5. Muon tat he thong: Dong cua so Terminal chay ngam hoac nhan to hop phim Ctrl + C.

===================================================================
[2. HUONG DAN DOI TAI KHOAN / LAM MOI TOKEN (KHI CO SU CO TAI KHOAN)]:
===================================================================
File 'token_cache.bin' chua ma xac thuc ngam ket noi toi SharePoint Microsoft 365.
Kiosk se tu dong lam moi token ngam nen ban khong can dang nhap lai hang ngay.

KHI NAO BAN CAN THAY DOI TAI KHOAN / CAP NHAT LAI TOKEN?
- Khi tai khoan ca nhan hien tai doi mat khau.
- Khi nguoi phu trach cu chuyen cong tac / nghi viec hoac khong muon dung tai khoan ca nhan nua.
- Khi muon chuyen quyen sang tai khoan chung cua Canteen hoac IT (Service Account).
- Khi Kiosk bao loi: 'TOKEN_HET_HAN', 'TOKEN_RONG' hoac khong tai duoc danh sach don com.

👉 CAC BUOC THUC HIEN DOI TAI KHOAN (CHI MAT 1 PHUT):
-------------------------------------------------------------------
Cach 1: Su dung cong cu tu dong "DANG_NHAP_TAI_KHOAN_MOI.bat" (Khuyen dung)
1. Tat Kiosk dang chay (neu co).
2. Nhap dup chuot vao file: DANG_NHAP_TAI_KHOAN_MOI.bat
3. Man hinh se hien ra mot doan ma xac nhan (Vi du: ABCD-EFGH) va tu dong mo trinh duyet vao trang:
   https://microsoft.com/devicelogin
4. Nhap ma hien thi tren man hinh -> Nhan "Tiep tuc" (Next).
5. Dang nhap tai khoan Microsoft 365 moi cua cong ty SeaZen (Tai khoan bat ky co quyen truy cap vao SharePoint ONESEAZEN).
6. Bam "Tiep tuc" / "Cho phep" de cap quyen.
7. Cua so dong lenh se bao: "[+] XAC THUC THANH CONG!", va file 'token_cache.bin' moi duoc tu dong sinh ra va ghi de.
8. Mo lai file "SEAZEN_Canteen_Kiosk.exe" de su dung binh thuong voi tai khoan moi!

Cach 2: Neu may Kiosk khong cai Python (May doc lap cua nha an)
1. Chay file "DANG_NHAP_TAI_KHOAN_MOI.bat" tren bat ky may tinh nao cua ban hoac IT co san Python.
2. Thuc hien dang nhap theo cac buoc o Cach 1.
3. Sau khi thanh cong, copy duy nhat file "token_cache.bin" vua tao ra va dan de vao thu muc Kiosk tren may nha an.

===================================================================
[3. KHUYEN NGHI DAI HAN VE TAI KHOAN CANTEEN]:
- De he thong hoat dong on dinh lau dai ma khong phu thuoc vao ca nhan (tranh truong hop nhan vien doi mat khau hoac nghi viec lam gian doan Kiosk), de xuat IT cong ty tao mot tai khoan dich vu rieng:
  Vi du: canteen.kiosk@seazen.com.vn hoac it.service@seazen.com.vn
- Cap quyen thanh vien (Member / Edit) cho tai khoan nay vao SharePoint ONESEAZEN (danh sach Canteen_Orders).
- Dang nhap tai khoan nay vao Kiosk mot lan duy nhat theo Huong dan o Muc 2.
===================================================================
"""
readme_path = os.path.join(PACKAGE_DIR, "HUONG_DAN_SU_DUNG.txt")
with open(readme_path, "w", encoding="utf-8") as f:
    f.write(readme_text)
print(" [+] Da tao: HUONG_DAN_SU_DUNG.txt")

# 4. Nen thanh file ZIP
if os.path.exists(ZIP_OUTPUT):
    os.remove(ZIP_OUTPUT)

with zipfile.ZipFile(ZIP_OUTPUT, "w", zipfile.ZIP_DEFLATED) as zipf:
    for root, dirs, files in os.walk(PACKAGE_DIR):
        for file in files:
            full_path = os.path.join(root, file)
            rel_path = os.path.relpath(full_path, PACKAGE_DIR)
            zipf.write(full_path, rel_path)

zip_size_mb = os.path.getsize(ZIP_OUTPUT) / (1024 * 1024)
print(f"[OK] DONG GOI HOAN TAT: {ZIP_OUTPUT} ({zip_size_mb:.2f} MB)")
