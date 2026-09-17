import os
import sys
import json
import datetime
import threading
import time
import webbrowser
import requests
import msal
import socket
from flask import Flask, jsonify, request, send_file, send_from_directory
from flask_cors import CORS
import logging

if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass
if sys.stderr and hasattr(sys.stderr, 'reconfigure'):
    try:
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass

# Tắt log GET định kỳ của Werkzeug trên Terminal để màn hình sạch sẽ
logging.getLogger('werkzeug').setLevel(logging.ERROR)

app = Flask(__name__)
CORS(app)

TENANT_ID = "29101901-3778-4339-827c-7d2b50914db0"
CLIENT_ID = "04b07795-8ddb-461a-bbee-02f9e1bf7b46"
SCOPES = ["https://graph.microsoft.com/.default"]
SITE_ID = "seazenvn.sharepoint.com,d47e794f-870a-4d63-8173-4d0721d5d77d,69313196-298d-4be2-af46-f7893c08cccf"
LIST_NAME = "Canteen_Orders"
USER_LOGIN_LIST_NAME = "UserLogin"

ACTIVE_SESSIONS = {}
CACHED_USER_LOGINS = []
LAST_USER_LOGIN_FETCH = 0
USER_LOGIN_CACHE_TTL = 60  # 60s cache để giảm tải SharePoint Graph API

def get_lan_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.settimeout(0.5)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        try:
            return socket.gethostbyname(socket.gethostname())
        except Exception:
            return "127.0.0.1"

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

MP3_DIR = os.path.join(BUNDLE_DIR, "mp3")
if not os.path.exists(MP3_DIR):
    MP3_DIR = os.path.join(EXE_DIR, "mp3")

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
CACHE_TTL_SECONDS = 3

def fetch_sharepoint_data(force=False):
    global LAST_FETCH_TIME, CACHED_ORDERS_TODAY, CACHED_CLAIMED_LIST, CACHED_ALL_TODAY_LIST, LOCAL_CHECKIN_OVERLAY
    now_ts = time.time()
    today_vn = get_today_vn_str()

    if not force and (now_ts - LAST_FETCH_TIME < CACHE_TTL_SECONDS) and CACHED_ORDERS_TODAY:
        for ma_nv, ov in list(LOCAL_CHECKIN_OVERLAY.items()):
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

            # Nếu có phân trang (@odata.nextLink)
            next_link = data.get("@odata.nextLink")
            while next_link:
                try:
                    r_next = requests.get(next_link, headers=headers, timeout=10)
                    if r_next.status_code == 200:
                        d_next = r_next.json()
                        items.extend(d_next.get("value", []))
                        next_link = d_next.get("@odata.nextLink")
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
                trang_thai = str(fields.get("TrangThai", "Đã đăng ký")).strip()
                raw_ngay_an = fields.get("NgayAn", "")
                gio_nhan = str(fields.get("GioNhan", "") or "").strip()

                order_date_vn = normalize_sp_date_to_vn(raw_ngay_an)
                try:
                    display_ngay_an = datetime.datetime.strptime(order_date_vn, "%Y-%m-%d").strftime("%d/%m/%Y")
                except Exception:
                    display_ngay_an = order_date_vn

                # ƯU TIÊN TUYỆT ĐỐI SHAREPOINT:
                # Nếu trên SharePoint người quản lý sửa thủ công lại thành "Đã đăng ký" -> Xóa ngay overlay cục bộ!
                if trang_thai != "Đã nhận cơm":
                    LOCAL_CHECKIN_OVERLAY.pop(ma_nv, None)
                    gio_nhan = ""
                elif ma_nv in LOCAL_CHECKIN_OVERLAY:
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

@app.route("/mp3/<path:filename>")
def serve_mp3(filename):
    for target_dir in [MP3_DIR, os.path.join(EXE_DIR, "mp3"), os.path.join(BUNDLE_DIR, "mp3")]:
        if os.path.exists(os.path.join(target_dir, filename)):
            return send_from_directory(target_dir, filename)
    return jsonify({"error": "File not found"}), 404

@app.route("/api/orders/today", methods=["GET"])
def get_orders_today():
    force_refresh = (request.args.get("force", "0") in ("1", "true", "True"))
    if force_refresh:
        LOCAL_CHECKIN_OVERLAY.clear()
    orders, claimed_list, all_today = fetch_sharepoint_data(force=force_refresh)
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

import re

@app.route("/api/checkin", methods=["POST"])
def checkin():
    global CACHED_CLAIMED_LIST, CACHED_ALL_TODAY_LIST, LOCAL_CHECKIN_OVERLAY
    body = request.get_json(force=True) or {}
    raw_ma_nv = str(body.get("maNV", "")).strip()
    raw_scan = str(body.get("rawScan", "")).strip()
    gio_nhan = datetime.datetime.now(VN_TZ).strftime("%H:%M:%S")

    # Quy tắc cốt lõi: Nếu chuỗi có |, lấy hết text phía trước, gặp | thì dừng
    candidate = raw_ma_nv or raw_scan
    if "|" in candidate:
        candidate = candidate.split("|")[0].strip()
    ma_nv = candidate

    if not ma_nv:
        return jsonify({"status": "ERROR", "message": "Thiếu mã nhân viên"}), 400

    # Dùng trực tiếp dữ liệu danh sách hôm nay trong bộ nhớ RAM (0.001 ms)
    orders_today = CACHED_ORDERS_TODAY
    if not orders_today:
        orders_today, _, _ = fetch_sharepoint_data()

    # Nhận diện tự động nếu ma_nv bị kẹt phím rác (vd: ttttt14703) hoặc trùng khớp
    if ma_nv not in orders_today:
        matched_id = None
        for valid_id in orders_today.keys():
            if ma_nv == valid_id or ma_nv.endswith(valid_id) or valid_id in ma_nv:
                matched_id = valid_id
                break
        if not matched_id:
            num_match = re.search(r'(\d{2,})$', ma_nv)
            if num_match and num_match.group(1) in orders_today:
                matched_id = num_match.group(1)
        if matched_id:
            ma_nv = matched_id

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

    # Ghi nhận ngay vào bộ nhớ đệm cục bộ trong 0.001 giây
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
    print(f"[+] [{gio_nhan}] CHECK-IN THANH CONG (PHAN HOI TUC THI): Ma NV {ma_nv} | {order.get('fullName')} | {order.get('quantity')} suat ({order.get('mealType')})")

    # ĐỒNG BỘ NỀN (BACKGROUND SYNC): Gửi lệnh PATCH lên SharePoint trên luồng ngầm để Kiosk không bị trễ 1 giây
    def async_patch_sharepoint(item_id, gio_nhan_val, ma_nv_val):
        token = get_access_token()
        if token:
            headers = {
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json"
            }
            patch_url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/lists/{LIST_NAME}/items/{item_id}/fields"
            try:
                resp = requests.patch(patch_url, headers=headers, json={
                    "TrangThai": "Đã nhận cơm",
                    "GioNhan": gio_nhan_val
                }, timeout=8)
                if resp.status_code in (200, 204):
                    print(f"[*] [{gio_nhan_val}] DONG BO SHAREPOINT THANH CONG: Item {item_id} (Ma NV {ma_nv_val})")
                else:
                    print(f"[-] [{gio_nhan_val}] SharePoint Error Status {resp.status_code}: {resp.text[:150]}")
            except Exception as ex:
                print("[-] Loi ghi SharePoint ngam:", ex)

    threading.Thread(target=async_patch_sharepoint, args=(order['itemId'], gio_nhan, ma_nv), daemon=True).start()

    # Trả về kết quả ngay lập tức (< 0.05s) cho giao diện Kiosk
    return jsonify({
        "status": "SUCCESS",
        "message": "Check-in thành công",
        "order": order,
        "spSynced": True
    })

def fetch_user_logins(force=False):
    global CACHED_USER_LOGINS, LAST_USER_LOGIN_FETCH
    now_ts = time.time()
    if not force and CACHED_USER_LOGINS and (now_ts - LAST_USER_LOGIN_FETCH < USER_LOGIN_CACHE_TTL):
        return CACHED_USER_LOGINS

    token = get_access_token()
    if not token:
        return CACHED_USER_LOGINS

    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json"
    }

    url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/lists/{USER_LOGIN_LIST_NAME}/items?expand=fields&$top=999"
    try:
        r = requests.get(url, headers=headers, timeout=10)
        if r.status_code == 200:
            data = r.json()
            items = data.get("value", [])
            next_link = data.get("@odata.nextLink")
            while next_link:
                try:
                    r_next = requests.get(next_link, headers=headers, timeout=10)
                    if r_next.status_code == 200:
                        d_next = r_next.json()
                        items.extend(d_next.get("value", []))
                        next_link = d_next.get("@odata.nextLink")
                    else:
                        break
                except Exception:
                    break

            parsed_users = []
            for it in items:
                f = it.get("fields", {})
                u = str(f.get("UserName", "")).strip()
                p = str(f.get("Credential", "")).strip()
                roles = f.get("AdminRoles", [])
                if isinstance(roles, str):
                    roles = [roles]
                elif not isinstance(roles, list):
                    roles = []

                parsed_users.append({
                    "id": it.get("id"),
                    "username": u,
                    "password": p,
                    "fullName": f.get("FullName") or f.get("MSNV_x003a__x0020_H_x1ecc__x0020") or f.get("Title") or u,
                    "department": f.get("Department", ""),
                    "securityRole": f.get("SecurityRole", ""),
                    "adminRoles": roles
                })

            CACHED_USER_LOGINS = parsed_users
            LAST_USER_LOGIN_FETCH = now_ts
            return parsed_users
        else:
            print(f"[-] Loi lay danh sach UserLogin: HTTP {r.status_code}")
    except Exception as e:
        print("[-] Loi lay danh sach UserLogin:", e)

    return CACHED_USER_LOGINS

@app.route("/api/login", methods=["POST"])
def api_login():
    data = request.get_json(force=True, silent=True) or {}
    username = str(data.get("username", "")).strip()
    password = str(data.get("password", "")).strip()

    if not username or not password:
        return jsonify({
            "status": "ERROR",
            "message": "Vui lòng nhập đầy đủ Tên đăng nhập và Mật khẩu!"
        }), 400

    users = fetch_user_logins()
    if not users:
        users = fetch_user_logins(force=True)

    matched = None
    for u in users:
        if u["username"].lower() == username.lower() and u["password"] == password:
            matched = u
            break

    if not matched:
        return jsonify({
            "status": "ERROR",
            "message": "Tên đăng nhập hoặc mật khẩu không chính xác!"
        }), 401

    # Kiểm tra quyền: AdminRoles phải chứa "Xem lịch trực toàn bộ"
    roles = matched.get("adminRoles", [])
    has_permission = any("xem lịch trực toàn bộ" in str(r).lower() for r in roles)
    if not has_permission:
        return jsonify({
            "status": "FORBIDDEN",
            "message": "Tài khoản không có quyền truy cập hệ thống! (Yêu cầu quyền 'Xem lịch trực toàn bộ' trong AdminRoles)."
        }), 403

    # Tạo token ngẫu nhiên lưu phiên
    session_token = f"sz_{abs(hash(username + password + str(time.time()))):x}"
    session_data = {
        "username": matched["username"],
        "fullName": matched["fullName"],
        "department": matched["department"],
        "securityRole": matched["securityRole"],
        "adminRoles": matched["adminRoles"],
        "loginAt": datetime.datetime.now(VN_TZ).strftime("%H:%M:%S %d/%m/%Y")
    }
    ACTIVE_SESSIONS[session_token] = session_data

    print(f"[+] DANG NHAP THANH CONG: {matched['username']} ({matched['fullName']})")

    return jsonify({
        "status": "SUCCESS",
        "message": "Đăng nhập thành công!",
        "token": session_token,
        "user": session_data
    })

@app.route("/api/auth/verify", methods=["GET", "POST"])
def api_auth_verify():
    token = request.headers.get("X-Auth-Token") or request.args.get("token", "")
    if not token and request.is_json:
        token = (request.get_json(force=True, silent=True) or {}).get("token", "")

    if token and token in ACTIVE_SESSIONS:
        return jsonify({
            "status": "SUCCESS",
            "valid": True,
            "user": ACTIVE_SESSIONS[token]
        })

    return jsonify({
        "status": "UNAUTHORIZED",
        "valid": False,
        "message": "Phiên làm việc hết hạn hoặc không hợp lệ."
    }), 401

@app.route("/api/auth/logout", methods=["POST"])
def api_auth_logout():
    token = request.headers.get("X-Auth-Token") or (request.get_json(force=True, silent=True) or {}).get("token", "")
    if token in ACTIVE_SESSIONS:
        del ACTIVE_SESSIONS[token]
    return jsonify({"status": "SUCCESS", "message": "Đã đăng xuất"})

def open_browser():
    time.sleep(1.5)
    # Thử mở bằng Microsoft Edge trước để có giọng đọc AI Nam/Nữ tự nhiên chất lượng cao nhất
    try:
        edge_paths = [
            r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
            r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        ]
        opened = False
        for ep in edge_paths:
            if os.path.exists(ep):
                import subprocess
                subprocess.Popen([ep, "http://localhost:5000"])
                opened = True
                break
        if not opened:
            webbrowser.open("http://localhost:5000")
    except Exception:
        webbrowser.open("http://localhost:5000")

if __name__ == "__main__":
    today = get_today_vn_str()
    lan_ip = get_lan_ip()
    hostname = socket.gethostname()

    print("=" * 75)
    print("      [+] SEAZEN CANTEEN CHECK-IN SYSTEM - STANDALONE KIOSK [+]      ")
    print("=" * 75)
    print(f"[*] Ngay lam viec (VN UTC+7) : {today}")
    print(f"[*] Thu muc thuc thi         : {EXE_DIR}")
    print(f"[*] Kiem tra token_cache.bin : {'DA TIM THAY [OK]' if os.path.exists(TOKEN_CACHE_FILE) else 'KHONG TIM THAY [FAIL]'}")
    print("=" * 75)
    print("   🌐 CAC DIA CHI TRUY CAP HE THONG (TRUONG HOP A - MANG LAN / WI-FI):")
    print(f"   👉 [1] Tren may nay          : http://localhost:5000")
    print(f"   👉 [2] Tu may khac (Bang IP) : http://{lan_ip}:5000")
    print(f"   👉 [3] Tu may khac (Ten may) : http://{hostname}:5000")
    print("=" * 75)
    print("[*] San sang nhan ma quet barcode/QR & phuc vu he thong...")

    threading.Thread(target=open_browser, daemon=True).start()
    app.run(host="0.0.0.0", port=5000, debug=False)

