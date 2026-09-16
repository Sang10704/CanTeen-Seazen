import os
import sys
import json
import requests
import msal

SITE_URL = "https://seazenvn.sharepoint.com/sites/ONESEAZEN"
TENANT = "29101901-3778-4339-827c-7d2b50914db0"
# Client ID chính thức của Microsoft Azure CLI (pre-approved trong mọi tổ chức Microsoft 365)
CLIENT_ID = "04b07795-8ddb-461a-bbee-02f9e1bf7b46"
SCOPES = ["https://seazenvn.sharepoint.com/.default"]

TOKEN_CACHE_FILE = os.path.join(os.path.dirname(__file__), "token_cache.bin")

def get_token():
    cache = msal.SerializableTokenCache()
    if os.path.exists(TOKEN_CACHE_FILE):
        with open(TOKEN_CACHE_FILE, "r") as f:
            cache.deserialize(f.read())

    app = msal.PublicClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT}",
        token_cache=cache
    )

    # 1. Thử lấy token từ cache đã lưu trước đó
    accounts = app.get_accounts(username="Sang.Nguyen@seazen.com.vn")
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            print("[+] Đã lấy token tự động từ bộ nhớ đệm (Cache)!")
            return result["access_token"]

    # 2. Device Flow chính thức của Microsoft
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        print("[-] Không thể khởi tạo device flow:", flow)
        return None

    print("\n=======================================================")
    print("👉 XÁC THỰC HIỆN ĐẠI MICROSOFT (CHỈ CẦN LÀM 1 LẦN DUY NHẤT):")
    print(f"1. Mở trình duyệt vào link: {flow['verification_uri']}")
    print(f"2. Nhập mã hiển thị: {flow['user_code']}")
    print("3. Đăng nhập tài khoản Sang.Nguyen@seazen.com.vn và bấm Xác nhận.")
    print("=======================================================\n")
    print("[*] Đang chờ bạn xác nhận trên trình duyệt...")

    result = app.acquire_token_by_device_flow(flow)
    if "access_token" in result:
        print("[+] ĐĂNG NHẬP THÀNH CÔNG!")
        if cache.has_state_changed:
            with open(TOKEN_CACHE_FILE, "w") as f:
                f.write(cache.serialize())
        return result["access_token"]
    else:
        print("[-] Đăng nhập thất bại:", result.get("error_description"))
        return None

def test_sharepoint_api():
    token = get_token()
    if not token:
        return

    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json;odata=verbose"
    }

    url = f"{SITE_URL}/_api/web/lists/getbytitle('Canteen_Orders')/items"
    print(f"\n[*] Đang đọc danh sách từ SharePoint: {url}")
    res = requests.get(url, headers=headers)
    print(f"[*] HTTP Status Code: {res.status_code}")
    
    if res.status_code == 200:
        data = res.json()
        results = data.get("d", {}).get("results", [])
        print(f"\n[+] THÀNH CÔNG RỰC RỠ! Đã lấy được {len(results)} đơn đặt cơm trong Canteen_Orders:")
        for item in results:
            print(f"  - [{item.get('Title')}] {item.get('HoTen')} | Suất: {item.get('SoLuongSuat')} | Trạng thái: {item.get('TrangThai')}")
    else:
        print(f"[-] Lỗi response: {res.text}")

if __name__ == "__main__":
    test_sharepoint_api()
