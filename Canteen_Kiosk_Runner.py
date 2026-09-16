import os
import sys
import json
import datetime
import threading
import time
import webbrowser
import requests
import msal
from flask import Flask, jsonify, request, send_file
from flask_cors import CORS
import logging

# Tắt log GET định kỳ của Werkzeug trên Terminal để màn hình sạch sẽ
logging.getLogger('werkzeug').setLevel(logging.ERROR)

app = Flask(__name__)
CORS(app)

TENANT_ID = "29101901-3778-4339-827c-7d2b50914db0"
CLIENT_ID = "04b07795-8ddb-461a-bbee-02f9e1bf7b46"
SCOPES = ["https://graph.microsoft.com/.default"]
SITE_ID = "seazenvn.sharepoint.com,d47e794f-870a-4d63-8173-4d0721d5d77d,69313196-298d-4be2-af46-f7893c08cccf"
LIST_NAME = "Canteen_Orders"

if getattr(sys, 'frozen', False):
    EXE_DIR = os.path.dirname(sys.executable)
    BUNDLE_DIR = getattr(sys, '_MEIPASS', EXE_DIR)
else:
    EXE_DIR = os.path.dirname(os.path.abspath(__file__))
    BUNDLE_DIR = EXE_DIR

candidate_token_paths = [
    os.path.join(EXE_DIR, "token_cache.bin"),
    os.path.join(os.getcwd(), "token_cache.bin"),
    os.path.join(BUNDLE_DIR, "token_cache.bin"),
]
TOKEN_CACHE_FILE = candidate_token_paths[0]
for p in candidate_token_paths:
    if os.path.exists(p):
        TOKEN_CACHE_FILE = p
        break

HTML_FILE = os.path.join(BUNDLE_DIR, "Canteen_Kiosk_Production.html")
if not os.path.exists(HTML_FILE):
    HTML_FILE = os.path.join(EXE_DIR, "OneSEAZEN", "Canteen_Kiosk_Production.html")
if not os.path.exists(HTML_FILE):
    HTML_FILE = os.path.join(EXE_DIR, "Canteen_Kiosk_Production.html")

VN_TZ = datetime.timezone(datetime.timedelta(hours=7))

CONNECTION_DIAGNOSTICS = {
    "token_file_found": os.path.exists(TOKEN_CACHE_FILE),
    "token_file_path": TOKEN_CACHE_FILE,
    "last_error": "",
    "user_account": "",
    "last_sync_time": ""
}

LOCAL_CHECKIN_OVERLAY = {}

def get_today_vn_str():
    return datetime.datetime.now(VN_TZ).strftime("%Y-%m-%d")

def normalize_sp_date_to_vn(raw_date_str):
    if not raw_date_str:
        return ""
    try:
        str_val = str(raw_date_str).strip()
        if len(str_val) == 10 and str_val[4] == '-' and str_val[7] == '-':
            return str_val
        clean = str_val.replace("Z", "+00:00")
        dt = datetime.datetime.fromisoformat(clean)
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=datetime.timezone.utc)
        dt_vn = dt.astimezone(VN_TZ)
        return dt_vn.strftime("%Y-%m-%d")
    except Exception:
        return str(raw_date_str)[:10]

def get_access_token():
    CONNECTION_DIAGNOSTICS["token_file_found"] = os.path.exists(TOKEN_CACHE_FILE)
    if not os.path.exists(TOKEN_CACHE_FILE):
        CONNECTION_DIAGNOSTICS["last_error"] = "THIEU_FILE_TOKEN"
        return None

    cache = msal.SerializableTokenCache()
    try:
        with open(TOKEN_CACHE_FILE, "r", encoding="utf-8") as f:
            cache.deserialize(f.read())
    except Exception as ex:
        CONNECTION_DIAGNOSTICS["last_error"] = f"LOI_DOC_TOKEN: {ex}"
        return None

    msal_app = msal.PublicClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        token_cache=cache
    )

    accounts = msal_app.get_accounts()
    if not accounts:
        CONNECTION_DIAGNOSTICS["last_error"] = "TOKEN_RONG"
        return None

    acc = accounts[0]
    CONNECTION_DIAGNOSTICS["user_account"] = acc.get("username", "")

    res = msal_app.acquire_token_silent(SCOPES, account=acc)
    if res and "access_token" in res:
        if cache.has_state_changed:
            try:
                with open(TOKEN_CACHE_FILE, "w", encoding="utf-8") as f:
                    f.write(cache.serialize())
            except Exception:
                pass
        return res["access_token"]
    
    err_desc = res.get("error_description", res.get("error", "Khong ro")) if res else "Silent token failed"
    CONNECTION_DIAGNOSTICS["last_error"] = f"TOKEN_HET_HAN: {err_desc[:60]}"
    return None

LAST_FETCH_TIME = 0
CACHED_ORDERS_TODAY = {}
CACHED_CLAIMED_LIST = []
CACHED_ALL_TODAY_LIST = []
CACHE_TTL_SECONDS = 10

def fetch_sharepoint_data(force=False):
    global LAST_FETCH_TIME, CACHED_ORDERS_TODAY, CACHED_CLAIMED_LIST, CACHED_ALL_TODAY_LIST
    now_ts = time.time()
    today_vn = get_today_vn_str()

    if not force and (now_ts - LAST_FETCH_TIME < CACHE_TTL_SECONDS) and CACHED_ORDERS_TODAY:
        for ma_nv, ov in LOCAL_CHECKIN_OVERLAY.items():
            if ov.get("date") == today_vn and ma_nv in CACHED_ORDERS_TODAY:
                CACHED_ORDERS_TODAY[ma_nv]["status"] = "Đã nhận cơm"
                CACHED_ORDERS_TODAY[ma_nv]["claimedAt"] = ov.get("claimedAt", "")
        return CACHED_ORDERS_TODAY, CACHED_CLAIMED_LIST, CACHED_ALL_TODAY_LIST

    token = get_access_token()
    if not token:
        return CACHED_ORDERS_TODAY or {}, CACHED_CLAIMED_LIST or [], CACHED_ALL_TODAY_LIST or []

    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json"
    }

    url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/lists/{LIST_NAME}/items?$expand=fields&$top=999"

    try:
        r = requests.get(url, headers=headers, timeout=10)
        if r.status_code == 200:
            CONNECTION_DIAGNOSTICS["last_error"] = ""
            CONNECTION_DIAGNOSTICS["last_sync_time"] = datetime.datetime.now(VN_TZ).strftime("%H:%M:%S")
            data = r.json()
            items = data.get("value", [])
            next_link = data.get("@odata.nextLink")
            while next_link and len(items) < 3000:
                try:
                    nr = requests.get(next_link, headers=headers, timeout=10)
                    if nr.status_code == 200:
                        ndata = nr.json()
                        items.extend(ndata.get("value", []))
                        next_link = ndata.get("@odata.nextLink")
                    else:
                        break
                except Exception:
                    break

            orders_today = {}
            claimed_list = []
            all_today_list = []

            for item in items:
                fields = item.get("fields", {})
                ma_nv = str(fields.get("Title", "")).strip()
                ho_ten = fields.get("HoTen", "")
                chuc_vu = fields.get("ChucVu", "")
                loai_mon = fields.get("LoaiMon") or "Món mặn"
                so_luong = int(fields.get("SoLuongSuat", 1) or 1)
                trang_thai = fields.get("TrangThai", "Đã đăng ký")
                raw_ngay_an = fields.get("NgayAn", "")
                gio_nhan = fields.get("GioNhan", "")

                order_date_vn = normalize_sp_date_to_vn(raw_ngay_an)
                try:
                    display_ngay_an = datetime.datetime.strptime(order_date_vn, "%Y-%m-%d").strftime("%d/%m/%Y")
                except Exception:
                    display_ngay_an = order_date_vn

                # Đồng bộ tức thì với overlay check-in cục bộ
                if ma_nv in LOCAL_CHECKIN_OVERLAY:
                    overlay = LOCAL_CHECKIN_OVERLAY[ma_nv]
                    if overlay.get("date") == today_vn:
                        trang_thai = "Đã nhận cơm"
                        if not gio_nhan:
                            gio_nhan = overlay.get("claimedAt", "")

                if order_date_vn == today_vn and trang_thai != "Đã hủy":
                    order_obj = {
                        "itemId": item.get("id"),
                        "empId": ma_nv,
                        "fullName": ho_ten,
                        "role": chuc_vu,
                        "mealType": loai_mon,
                        "quantity": so_luong,
                        "status": trang_thai,
                        "claimedAt": gio_nhan or "",
                        "orderDate": order_date_vn,
                        "mealDate": display_ngay_an
                    }
                    orders_today[ma_nv] = order_obj
                    all_today_list.append(order_obj)

                    if trang_thai == "Đã nhận cơm":
                        claimed_list.append({
                            "name": ho_ten,
                            "empId": ma_nv,
                            "meal": f"{so_luong} {loai_mon}",
                            "time": gio_nhan or "11:30"
                        })

            claimed_list.sort(key=lambda x: x["time"], reverse=True)
            CACHED_ORDERS_TODAY = orders_today
            CACHED_CLAIMED_LIST = claimed_list
            CACHED_ALL_TODAY_LIST = all_today_list
            LAST_FETCH_TIME = now_ts
            return orders_today, claimed_list, all_today_list
        else:
            CONNECTION_DIAGNOSTICS["last_error"] = f"HTTP_{r.status_code}"
    except Exception as e:
        CONNECTION_DIAGNOSTICS["last_error"] = f"LOI_MANG: {str(e)[:40]}"

    return CACHED_ORDERS_TODAY or {}, CACHED_CLAIMED_LIST or [], CACHED_ALL_TODAY_LIST or []

@app.route("/")
def index():
    if os.path.exists(HTML_FILE):
        return send_file(HTML_FILE)
    return f"<h3>Lỗi: Không tìm thấy file giao diện Kiosk tại: {HTML_FILE}</h3>"

@app.route("/api/orders/today", methods=["GET"])
def get_orders_today():
    orders, claimed_list, all_today = fetch_sharepoint_data()
    has_error = bool(CONNECTION_DIAGNOSTICS["last_error"])
    return jsonify({
        "status": "ERROR" if has_error else "SUCCESS",
        "errorMessage": CONNECTION_DIAGNOSTICS["last_error"],
        "account": CONNECTION_DIAGNOSTICS["user_account"],
        "tokenFound": CONNECTION_DIAGNOSTICS["token_file_found"],
        "date": get_today_vn_str(),
        "total": len(orders),
        "orders": orders,
        "recentClaimed": claimed_list,
        "allOrders": all_today
    })

@app.route("/api/checkin", methods=["POST"])
def checkin():
    body = request.get_json(force=True) or {}
    ma_nv = str(body.get("maNV", "")).strip()
    gio_nhan = datetime.datetime.now(VN_TZ).strftime("%H:%M:%S")

    if not ma_nv:
        return jsonify({"status": "ERROR", "message": "Thiếu mã nhân viên"}), 400

    orders_today, _, _ = fetch_sharepoint_data()

    if ma_nv not in orders_today:
        print(f"[-] [{gio_nhan}] TU CHOI: Ma NV {ma_nv} chua dang ky com trua hom nay.")
        return jsonify({
            "status": "NOT_REGISTERED",
            "message": "Chưa đăng ký cơm hôm nay",
            "empId": ma_nv
        })

    order = orders_today[ma_nv]

    if order["status"] == "Đã nhận cơm":
        print(f"[!] [{gio_nhan}] CANH BAO: Ma NV {ma_nv} ({order.get('fullName')}) da nhan com truoc do luc {order.get('claimedAt')}.")
        return jsonify({
            "status": "ALREADY_CLAIMED",
            "message": "Đã nhận cơm trước đó",
            "order": order
        })

    # Ghi nhận ngay vào bộ nhớ đệm cục bộ
    today_vn = get_today_vn_str()
    LOCAL_CHECKIN_OVERLAY[ma_nv] = {
        "date": today_vn,
        "status": "Đã nhận cơm",
        "claimedAt": gio_nhan
    }
    order["status"] = "Đã nhận cơm"
    order["claimedAt"] = gio_nhan

    # Cập nhật ngay trong danh sách tất cả các đơn CACHED_ALL_TODAY_LIST
    for item in CACHED_ALL_TODAY_LIST:
        if item.get("empId") == ma_nv:
            item["status"] = "Đã nhận cơm"
            item["claimedAt"] = gio_nhan
            break

    # Cập nhật danh sách vừa nhận CACHED_CLAIMED_LIST lên đầu
    CACHED_CLAIMED_LIST = [c for c in CACHED_CLAIMED_LIST if c.get("empId") != ma_nv]
    CACHED_CLAIMED_LIST.insert(0, {
        "name": order.get("fullName"),
        "empId": ma_nv,
        "meal": f"{order.get('quantity')} {order.get('mealType')}",
        "time": gio_nhan
    })

    # In log rõ ràng, ngắn gọn cho người vận hành Kiosk
    print(f"[+] [{gio_nhan}] CHECK-IN THANH CONG: Ma NV {ma_nv} | {order.get('fullName')} | {order.get('quantity')} suat ({order.get('mealType')})")

    token = get_access_token()
    if token:
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }
        patch_url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/lists/{LIST_NAME}/items/{order['itemId']}/fields"
        try:
            requests.patch(patch_url, headers=headers, json={
                "TrangThai": "Đã nhận cơm",
                "GioNhan": gio_nhan
            }, timeout=6)
        except Exception as ex:
            print("[-] Loi ghi SharePoint:", ex)

    return jsonify({
        "status": "SUCCESS",
        "message": "Check-in thành công",
        "order": order
    })

def open_browser():
    time.sleep(1.5)
    webbrowser.open("http://localhost:5000")

if __name__ == "__main__":
    today = get_today_vn_str()
    print("=" * 70)
    print("      🍱 SEAZEN CANTEEN CHECK-IN SYSTEM - STANDALONE KIOSK 🍱      ")
    print("=" * 70)
    print(f"[*] Ngay lam viec (VN UTC+7): {today}")
    print(f"[*] Thu muc thuc thi: {EXE_DIR}")
    print(f"[*] Kiem tra token_cache.bin: {'DA TIM THAY [OK]' if os.path.exists(TOKEN_CACHE_FILE) else 'KHONG TIM THAY [FAIL]'}")
    print("[*] Server dang chay tai: http://localhost:5000")
    print("[*] Che do Terminal sach: Da tat log GET dinh ky, chi hien thi Check-in")
    print("=" * 70)
    print("[*] San sang nhan ma quet barcode/QR...")
    
    threading.Thread(target=open_browser, daemon=True).start()
    app.run(host="0.0.0.0", port=5000, debug=False)
