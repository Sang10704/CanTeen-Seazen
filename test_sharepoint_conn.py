import sys
import os
import json
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext

CONFIG_FILE = os.path.join(os.path.dirname(__file__), "config.json")

def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {
        "site_url": "https://seazenvn.sharepoint.com/sites/ONESEAZEN",
        "username": "Sang.Nguyen@seazen.com.vn",
        "password": ""
    }

def test_connection():
    cfg = load_config()
    site_url = cfg.get("site_url", "https://seazenvn.sharepoint.com/sites/ONESEAZEN")
    username = cfg.get("username", "Sang.Nguyen@seazen.com.vn")
    password = cfg.get("password", "")

    if not password:
        if len(sys.argv) > 1:
            password = sys.argv[1]
        else:
            import getpass
            print(f"[*] Đang dùng tài khoản: {username}")
            print(f"[!] Mật khẩu trong config.json đang trống. Vui lòng nhập mật khẩu:")
            password = getpass.getpass("Password: ")

    print(f"\n[*] Đang thử kết nối tới SharePoint: {site_url}")
    print(f"[*] Tài khoản: {username}")
    
    try:
        user_cred = UserCredential(username, password)
        ctx = ClientContext(site_url).with_credentials(user_cred)
        
        web = ctx.web
        ctx.load(web)
        ctx.execute_query()
        print(f"\n=======================================================")
        print(f"[+] KẾT NỐI THÀNH CÔNG TỚI SHAREPOINT!")
        print(f"[+] Tên Trang (Web Title): {web.properties.get('Title')}")
        
        # Thử đọc list Canteen_Orders
        print(f"[*] Đang truy vấn danh sách 'Canteen_Orders'...")
        target_list = ctx.web.lists.get_by_title("Canteen_Orders")
        items = target_list.items
        ctx.load(items)
        ctx.execute_query()
        print(f"[+] ĐÃ ĐỌC ĐƯỢC DANH SÁCH Canteen_Orders! Tổng số dòng hiện tại: {len(items)}")
        
        if len(items) == 0:
            print(f"[*] Danh sách hiện đang trống (chưa có dòng nào).")
        else:
            for item in items:
                print(f"  - [{item.properties.get('Title')}] {item.properties.get('HoTen')} | Suất: {item.properties.get('SoLuongSuat')} | Trạng thái: {item.properties.get('TrangThai')}")
        print(f"=======================================================\n")
        return True
    except Exception as e:
        print(f"\n[-] Kết nối thất bại: {e}")
        return False

if __name__ == "__main__":
    test_connection()
