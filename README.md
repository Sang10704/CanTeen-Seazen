# 🍱 SEAZEN CANTEEN CHECK-IN SYSTEM & POWER APPS

Hệ thống điểm danh suất ăn Canteen tự động cho nhân viên SEAZEN, kết nối trực tiếp với SharePoint List `Canteen_Orders` qua Microsoft Graph API.

---

## 🌟 Tính năng chính

1. **Điểm danh tự động qua Barcode / QR 2D Scanner:**
   - Trích xuất thông minh mã nhân viên trước ký tự `|` (ví dụ `14703|Nguyễn Văn A` -> `14703`).
   - Cảnh báo âm thanh giọng nói tiếng Việt chuẩn khi nhận cơm hoặc khi đã lấy trước đó.
   - Hạn mức: mỗi nhân viên nhận đúng số suất đã đăng ký.

2. **Bảng đối soát & Danh sách chi tiết hôm nay:**
   - Bật/tắt nhanh bằng mã QR / phím `LIST` hoặc nút trên màn hình.
   - Hiển thị 100% dữ liệu thực từ SharePoint List `Canteen_Orders`:
     `Số hiệu` | `Họ tên` | `Chức vụ` | `Ngày ăn` | `Giờ nhận` | `Số suất` | `Loại món` | `Trạng thái`.
   - Phân biệt rõ ràng người **✓ ĐÃ NHẬN** và người **⏳ ĐÃ ĐĂNG KÝ (chờ nhận)**.

3. **Giao diện Web Kiosk Retro & Ứng dụng Standalone `.EXE`:**
   - Hoạt động độc lập không cần cài đặt phức tạp, tự động bật trình duyệt toàn màn hình.
   - Bộ đệm phím toàn cục (Global Key Buffer) chống mất focus khi quét mã.

4. **Tích hợp Power Apps:**
   - Màn hình đăng ký suất ăn nhân viên trên ứng dụng di động SEAZEN HR.

---

## 📁 Cấu trúc thư mục

```text
CanTeen-Seazen/
├── Canteen_Kiosk_Production.html   # Giao diện Web Kiosk chính (Retro Red & Gold)
├── Canteen_Kiosk_Runner.py        # Backend Flask server & Graph API sync
├── canteen_server.py              # Server phụ trợ API Kiosk
├── SEAZEN_Canteen_Kiosk.spec      # File cấu hình đóng gói PyInstaller sang .EXE
├── CHAY_CANTEEN_KIOSK.bat         # File phím tắt 1-click khởi động hệ thống
├── canteen_icon.ico               # Icon ứng dụng (chuẩn Windows multi-res)
├── canteen_icon.png               # Icon PNG độ phân giải cao
├── ep_mo_cua.py                   # Script điều khiển thiết bị phần cứng (relay mở cửa)
├── test_sharepoint_conn.py        # Kiểm tra kết nối SharePoint Graph API
├── test_modern_auth.py            # Tạo và làm mới token_cache.bin từ Microsoft 365
├── requirements.txt               # Danh sách thư viện Python cần thiết
├── PowerApps/
│   └── DangKyCom_Screen.yaml      # Định nghĩa màn hình Power Apps đặt cơm Canteen
└── OneSEAZEN/
    └── Canteen_Kiosk_Production.html # Bản sao đồng bộ của giao diện Kiosk
```

---

## 🚀 Hướng dẫn cài đặt & Chạy ứng dụng

### 1. Cài đặt môi trường Python
Yêu cầu Python 3.10 trở lên:
```bash
pip install -r requirements.txt
```

### 2. Cấu hình xác thực Microsoft 365 (SharePoint Token)
Hệ thống sử dụng OAuth2 MSAL. Chạy script để đăng nhập và tạo file `token_cache.bin`:
```bash
python test_modern_auth.py
```
*Lưu ý: File `token_cache.bin` chứa phiên đăng nhập nội bộ của tổ chức và đã được đưa vào `.gitignore` để bảo mật.*

### 3. Khởi chạy Kiosk
Chạy trực tiếp bằng lệnh:
```bash
python Canteen_Kiosk_Runner.py
```
Hoặc nhấp đúp vào file **`CHAY_CANTEEN_KIOSK.bat`**.

Hệ thống sẽ tự động mở giao diện tại: `http://localhost:5000`.

### 4. Đóng gói thành file chạy `.EXE` độc lập
Để tạo file `SEAZEN_Canteen_Kiosk.exe`:
```bash
python -m PyInstaller --noconfirm SEAZEN_Canteen_Kiosk.spec
```
File `.exe` thành phẩm sẽ nằm trong thư mục `dist/`.

---

## 📋 Cấu trúc danh sách SharePoint `Canteen_Orders`

| Tên cột SharePoint | Kiểu dữ liệu | Ý nghĩa |
| :--- | :--- | :--- |
| `Title` | Single line of text | Mã nhân viên (Số hiệu) |
| `HoTen` | Single line of text | Họ và tên nhân viên |
| `ChucVu` | Single line of text | Chức vụ / Bộ phận |
| `NgayAn` | Date and Time | Ngày đăng ký ăn trưa |
| `GioNhan` | Single line of text | Thời gian quét nhận cơm thực tế (HH:mm:ss) |
| `SoLuongSuat` | Number | Số lượng suất ăn đăng ký |
| `LoaiMon` | Choice / Text | Loại món ăn (Món mặn / Món chay) |
| `TrangThai` | Choice / Text | Trạng thái (`Đã đăng ký`, `Đã nhận cơm`, `Đã hủy`) |
| `GhiChu` | Multiple lines of text | Ghi chú cho nhà bếp |

---

## 🔒 Bảo mật
- Tệp `token_cache.bin` được tự động bỏ qua khi đẩy lên Git nhằm bảo vệ thông tin đăng nhập Azure AD / Microsoft 365 của doanh nghiệp.
