import os
import sys
import json
import datetime
import requests
import msal
from flask import Flask, jsonify, request, send_file
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

# ==========================================
# CẤU HÌNH SHAREPOINT
# ==========================================
TENANT_ID = "29101901-3778-4339-827c-7d2b50914db0"
CLIENT_ID = "04b07795-8ddb-461a-bbee-02f9e1bf7b46"
SCOPES = ["https://graph.microsoft.com/.default"]
SITE_ID = "seazenvn.sharepoint.com,d47e794f-870a-4d63-8173-4d0721d5d77d,69313196-298d-4be2-af46-f7893c08cccf"
LIST_NAME = "Canteen_Orders"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TOKEN_CACHE_FILE = os.path.join(BASE_DIR, "token_cache.bin")
HTML_FILE = os.path.join(BASE_DIR, "OneSEAZEN", "Canteen_Kiosk_RetroRed_Preview.html")

# Timezone Việt Nam (UTC+7)
VN_TZ = datetime.timezone(datetime.timedelta(hours=7))

def get_today_vn_str():
    return datetime.datetime.now(VN_TZ).strftime("%Y-%m-%d")

def normalize_sp_date_to_vn(raw_date_str):
    """Chuẩn hóa ngày từ SharePoint (ISO UTC) về ngày Việt Nam (YYYY-MM-DD)"""
    if not raw_date_str:
        return ""
    try:
        clean = raw_date_str.replace("Z", "+00:00")
        dt = datetime.datetime.fromisoformat(clean)
        dt_vn = dt.astimezone(VN_TZ)
        return dt_vn.strftime("%Y-%m-%d")
    except Exception:
        return str(raw_date_str)[:10]

def get_access_token():
    cache = msal.SerializableTokenCache()
    if os.path.exists(TOKEN_CACHE_FILE):
        with open(TOKEN_CACHE_FILE, "r") as f:
            cache.deserialize(f.read())

    msal_app = msal.PublicClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        token_cache=cache
    )

    accounts = msal_app.get_accounts()
    if not accounts:
        return None

    res = msal_app.acquire_token_silent(SCOPES, account=accounts[0])
    if res and "access_token" in res:
        if cache.has_state_changed:
            with open(TOKEN_CACHE_FILE, "w") as f:
                f.write(cache.serialize())
        return res["access_token"]
    return None

# ==========================================
# ĐỒNG BỘ SHAREPOINT CHÍNH XÁC
# ==========================================
def fetch_sharepoint_data():
    """Đọc dữ liệu SharePoint và lọc CHÍNH XÁC chỉ lấy ngày HÔM NAY"""
    token = get_access_token()
    if not token:
        return {}, []

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/lists/{LIST_NAME}/items?expand=fields"
    today_vn = get_today_vn_str()

    try:
        res = requests.get(url, headers=headers, timeout=8)
        if res.status_code == 200:
            raw_items = res.json().get("value", [])
            orders_today = {}
            claimed_list = []

            for item in raw_items:
                f = item.get("fields", {})
                ma_nv = str(f.get("Title", "")).strip()
                if not ma_nv:
                    continue

                raw_date = str(f.get("NgayAn", ""))
                order_date_vn = normalize_sp_date_to_vn(raw_date)

                # YÊU CẦU 1: CHỈ LẤY ĐÚNG NGÀY HÔM NAY
                if order_date_vn != today_vn:
                    continue

                status = str(f.get("TrangThai", "Đã đăng ký")).strip()
                if status == "Đã hủy":
                    continue

                gio_nhan = f.get("GioNhan", None)
                qty = int(float(f.get("SoLuongSuat", 1) or 1))

                order_obj = {
                    "itemId": item.get("id"),
                    "empId": ma_nv,
                    "fullName": f.get("HoTen", ""),
                    "role": f.get("ChucVu", "Nhân viên"),
                    "mealType": f.get("LoaiMon", "Cơm mặn"),
                    "quantity": qty,
                    "status": status,
                    "claimedAt": gio_nhan,
                    "date": order_date_vn
                }

                orders_today[ma_nv] = order_obj

                # YÊU CẦU 3: DANH SÁCH VỪA NHẬN PHẢI LẤY ĐÚNG TỪ SHAREPOINT
                if status == "Đã nhận cơm":
                    claimed_list.append({
                        "name": order_obj["fullName"] or f"NV ({ma_nv})",
                        "meal": f"{order_obj['mealType']} ({qty} suất)",
                        "time": gio_nhan or "Đã nhận",
                        "empId": ma_nv
                    })

            # Sắp xếp danh sách vừa nhận: mới nhất lên đầu
            claimed_list.sort(key=lambda x: str(x.get("time", "")), reverse=True)
            return orders_today, claimed_list
        else:
            print("[-] Lỗi Graph API:", res.status_code, res.text)
    except Exception as e:
        print("[-] Ngoại lệ:", e)

    return {}, []

# ==========================================
# CÁC API CHO GIAO DIỆN KIOSK
# ==========================================

@app.route("/")
def index():
    if os.path.exists(HTML_FILE):
        return send_file(HTML_FILE)
    return "Không tìm thấy file giao diện Kiosk HTML!"

@app.route("/api/orders/today", methods=["GET"])
def get_orders_today():
    orders, claimed_list = fetch_sharepoint_data()
    return jsonify({
        "status": "SUCCESS",
        "date": get_today_vn_str(),
        "total": len(orders),
        "orders": orders,
        "recentClaimed": claimed_list  # Dữ liệu lấy THẬT từ SharePoint
    })

@app.route("/api/checkin", methods=["POST"])
def checkin():
    body = request.get_json(force=True) or {}
    ma_nv = str(body.get("maNV", "")).strip()
    gio_nhan = datetime.datetime.now(VN_TZ).strftime("%H:%M:%S")

    if not ma_nv:
        return jsonify({"status": "ERROR", "message": "Thiếu mã nhân viên"}), 400

    # Đọc dữ liệu mới nhất từ SharePoint
    orders_today, _ = fetch_sharepoint_data()

    if ma_nv not in orders_today:
        return jsonify({
            "status": "NOT_REGISTERED",
            "message": "Chưa đăng ký cơm hôm nay",
            "empId": ma_nv
        })

    order = orders_today[ma_nv]

    # Kiểm tra nếu đã nhận cơm rồi
    if order["status"] == "Đã nhận cơm":
        return jsonify({
            "status": "ALREADY_CLAIMED",
            "message": "Đã nhận cơm trước đó",
            "order": order
        })

    # Cập nhật SharePoint trực tiếp
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
            order["status"] = "Đã nhận cơm"
            order["claimedAt"] = gio_nhan
        except Exception as ex:
            print("[-] Lỗi ghi SharePoint:", ex)

    return jsonify({
        "status": "SUCCESS",
        "message": "Check-in thành công",
        "order": order
    })

if __name__ == "__main__":
    today = get_today_vn_str()
    print("\n" + "="*60)
    print(f"🚀 CANTEEN KIOSK HOST SERVER ĐANG CHẠY (NGÀY: {today})...")
    print("👉 Mở trình duyệt tại: http://localhost:5000")
    print("="*60 + "\n")
    app.run(host="0.0.0.0", port=5000, debug=False)
