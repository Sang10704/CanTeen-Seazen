import os
import sys
import msal
import webbrowser

TENANT_ID = "29101901-3778-4339-827c-7d2b50914db0"
CLIENT_ID = "04b07795-8ddb-461a-bbee-02f9e1bf7b46" # Azure CLI Client ID
SCOPES = ["https://graph.microsoft.com/.default"]

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TOKEN_CACHE_FILE = os.path.join(BASE_DIR, "token_cache.bin")

def nhap_mat_khau_co_sao(prompt="Nhap Mat khau: "):
    """Cho phep go mat khau hien thi dau hoa thi * va co the Backspace xoa"""
    try:
        import msvcrt
        sys.stdout.write(prompt)
        sys.stdout.flush()
        pwd = []
        while True:
            ch = msvcrt.getch()
            # Enter (CR hoac LF)
            if ch in (b'\r', b'\n'):
                sys.stdout.write('\n')
                sys.stdout.flush()
                break
            # Ctrl+C
            elif ch == b'\x03':
                sys.stdout.write('\n')
                raise KeyboardInterrupt
            # Backspace (xoa 1 ky tu *)
            elif ch in (b'\x08', b'\x7f'):
                if pwd:
                    pwd.pop()
                    sys.stdout.write('\b \b')
                    sys.stdout.flush()
            # Phim dac biet / phim mui ten (2-byte code tren Windows)
            elif ch in (b'\x00', b'\xe0'):
                msvcrt.getch() # Bo qua byte thu 2
            else:
                try:
                    char = ch.decode('utf-8')
                    pwd.append(char)
                    sys.stdout.write('*')
                    sys.stdout.flush()
                except Exception:
                    pass
        return "".join(pwd)
    except Exception:
        import getpass
        return getpass.getpass(prompt)

def main():
    print("=" * 70)
    print("   CONG CU DANG NHAP / DOI TAI KHOAN CANTEEN KIOSK (SEAZEN)")
    print("=" * 70)
    print()
    print("[1] Dang nhap truc tiep bang Email & Mat khau tren CMD (Nhanh nhat)")
    print("[2] Dang nhap qua trinh duyet bang Ma xac nhan (Device Code)")
    print()
    
    chon = input("Chon phuong thuc [1 hoac 2, mac dinh 1]: ").strip() or "1"
    
    cache = msal.SerializableTokenCache()
    app = msal.PublicClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        token_cache=cache
    )
    
    if chon == "1":
        print("\n--- DANG NHAP TRUC TIEP TREN CMD ---")
        username = input("Nhap Email (@seazen.com.vn): ").strip()
        if not username:
            print("[-] Email khong duoc de trong!")
            input("\nNhan Enter de thoat...")
            return
            
        password = nhap_mat_khau_co_sao("Nhap Mat khau (se hien thi dau *): ")
        if not password:
            print("[-] Mat khau khong duoc de trong!")
            input("\nNhan Enter de thoat...")
            return
            
        print("\n[*] Dang gui xac thuc truc tiep toi Microsoft...")
        res = app.acquire_token_by_username_password(username, password, scopes=SCOPES)
        
        if res and "access_token" in res:
            with open(TOKEN_CACHE_FILE, "w", encoding="utf-8") as f:
                f.write(cache.serialize())
            print("\n" + "=" * 70)
            print(" [+] DANG NHAP TRUC TIEP THANH CONG!")
            print(f" [+] Tai khoan dang hoat dong: {username}")
            print(f" [+] Da tao va luu file token moi tai:")
            print(f"     {TOKEN_CACHE_FILE}")
            print("=" * 70)
            print("\n[OK] Bay gio he thong Canteen Kiosk se tu dong dung tai khoan nay de ket noi SharePoint!")
            input("\nNhan Enter de hoan tat...")
            return
            
        err_desc = res.get("error_description", "")
        err_code = res.get("error", "")
        
        # Kiem tra neu tai khoan co bat bao mat 2 lop MFA
        if any(k in err_desc for k in ["AADSTS50076", "AADSTS50079", "MFA", "multi-factor"]):
            print("\n" + "-" * 70)
            print("[!] THONG BAO: Tai khoan cua ban co bat Bao mat 2 lop (MFA / Authenticator).")
            print("[*] Chinh sach Microsoft yeu cau xac thuc qua trinh duyet de nhan dien OTP.")
            print("[*] Dang tu dong chuyen sang che do Device Code...")
            print("-" * 70)
        else:
            print(f"\n[-] Dang nhap tren CMD that bai: {err_desc or err_code}")
            print("[*] Ban co muon thu dang nhap qua Trinh duyet (Device Code) khong? (Y/N): ", end="")
            ans = input().strip().lower()
            if ans not in ("y", "yes", ""):
                input("\nNhan Enter de thoat...")
                return

    # Phuong thuc 2: Device Code Flow
    print("\n--- DANG NHAP QUA TRINH DUYET (DEVICE CODE) ---")
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        print("[!] Khong the khoi tao Device Code Flow:", flow)
        input("\nNhan Enter de thoat...")
        return
        
    user_code = flow['user_code']
    login_url = flow['verification_uri']
    
    print("\n" + "=" * 70)
    print(" 👉 HUONG DAN XAC NHAN:")
    print("=" * 70)
    print(f" 1. He thong da tu dong mo trinh duyet (hoac truy cap: {login_url})")
    print(f" 2. Nhap chinh xac ma sau: >>>  {user_code}  <<<")
    print(" 3. Dang nhap tai khoan Microsoft 365 cua ban va bam Tiep tuc.")
    print("=" * 70)
    print("\n[*] Dang cho ban xac nhan tren trinh duyet...")
    
    try:
        webbrowser.open(login_url)
    except Exception:
        pass
        
    res_dev = app.acquire_token_by_device_flow(flow)
    if res_dev and "access_token" in res_dev:
        acc_name = (
            res_dev.get("id_token_claims", {}).get("preferred_username") or
            res_dev.get("id_token_claims", {}).get("upn") or
            "Tai khoan moi"
        )
        with open(TOKEN_CACHE_FILE, "w", encoding="utf-8") as f:
            f.write(cache.serialize())
        print("\n" + "=" * 70)
        print(" [+] XAC THUC THANH CONG!")
        print(f" [+] Tai khoan dang hoat dong: {acc_name}")
        print(f" [+] Da tao va luu file token moi tai:")
        print(f"     {TOKEN_CACHE_FILE}")
        print("=" * 70)
        print("\n[OK] Bay gio he thong Canteen Kiosk se tu dong dung tai khoan nay de ket noi SharePoint!")
    else:
        print(f"\n[-] Dang nhap that bai: {res_dev.get('error_description')}")
        
    input("\nNhan Enter de hoan tat...")

if __name__ == "__main__":
    main()
