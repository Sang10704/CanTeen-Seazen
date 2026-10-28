# -*- coding: utf-8 -*-
"""
HỆ THỐNG KIỂM THỬ TỰ ĐỘNG CANTEEN KIOSK (UNIT TESTS)
1. Kiểm tra chuẩn hóa mã nhân viên (1, 01, 001, 0001, 00001 đều là Tổng Giám Đốc Đỗ Văn Toàn).
2. Kiểm tra bắt buộc xác thực Microsoft 365 (REQUIRE_MS_LOGIN).
3. Kiểm tra mã hóa và lưu trữ token DPAPI tại AppData.
4. Kiểm tra luồng check-in nhân viên thường và bảo vệ dữ liệu.
"""

import os
import sys
import datetime
import unittest
from unittest import mock

# Nạp module Runner từ thư mục cha
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
PARENT_DIR = os.path.dirname(CURRENT_DIR)
if PARENT_DIR not in sys.path:
    sys.path.insert(0, PARENT_DIR)

import Canteen_Kiosk_Runner as kiosk


class Noon(datetime.datetime):
    @classmethod
    def now(cls, tz=None):
        return cls(2026, 9, 30, 12, 0, tzinfo=tz)


class Morning(datetime.datetime):
    @classmethod
    def now(cls, tz=None):
        return cls(2026, 9, 30, 10, 30, tzinfo=tz)


class KioskComprehensiveTests(unittest.TestCase):
    def setUp(self):
        kiosk.CACHED_ORDERS_TODAY.clear()
        kiosk.CACHED_ALL_TODAY_LIST.clear()
        kiosk.CACHED_CLAIMED_LIST.clear()
        kiosk.LOCAL_CHECKIN_OVERLAY.clear()
        self.client = kiosk.app.test_client()
        self.clock = mock.patch.object(kiosk.datetime, 'datetime', Noon)
        self.clock.start()
        self.addCleanup(self.clock.stop)

    def post(self, emp_id, raw=None):
        return self.client.post('/api/checkin', json={
            'maNV': emp_id,
            'rawScan': raw if raw is not None else emp_id,
        })

    # =========================================================================
    # 1. KIỂM THỬ CHUẨN HÓA MÃ NHÂN VIÊN
    # =========================================================================
    def test_normalize_emp_id_variants(self):
        """Tất cả các biến thể 1, 01, 001, 0001, 00001 đều chuẩn hóa về '1'"""
        tgd_variants = ['1', '01', '001', '0001', '00001', '000001', '00001|Đỗ Văn Toàn|Tổng Giám Đốc']
        for var in tgd_variants:
            norm = kiosk.normalize_emp_id(var)
            self.assertEqual(norm, '1', f"Biến thể '{var}' không chuẩn hóa về '1' (nhận được: {norm})")

        # Kiểm tra nhân viên thông thường
        self.assertEqual(kiosk.normalize_emp_id('12345'), '12345')
        self.assertEqual(kiosk.normalize_emp_id('012345'), '12345')
        self.assertEqual(kiosk.normalize_emp_id('00123'), '123')
        self.assertEqual(kiosk.normalize_emp_id('00000'), '0')
        self.assertEqual(kiosk.normalize_emp_id('NV001'), 'NV001')

    # =========================================================================
    # 2. KIỂM THỬ ĐẶC QUYỀN VIP CHO TỔNG GIÁM ĐỐC ĐỖ VĂN TOÀN
    # =========================================================================
    def test_tgd_scanned_with_leading_zeros_all_recognized_as_same_person(self):
        """Quét với 1, 01, 001, 0001, 00001 đều được ghi nhận là ĐỖ VĂN TOÀN (Mã NV: 1)"""
        variants = ['1', '01', '001', '0001', '00001', '00001|ĐỖ VĂN TOÀN|Tổng Giám Đốc']

        for var in variants:
            # Dọn dẹp cache cho mỗi lần thử
            kiosk.CACHED_ORDERS_TODAY.clear()
            kiosk.CACHED_ALL_TODAY_LIST.clear()
            kiosk.CACHED_CLAIMED_LIST.clear()
            kiosk.LOCAL_CHECKIN_OVERLAY.clear()

            with mock.patch.object(kiosk, 'get_access_token', return_value='fake_token'), \
                 mock.patch.object(kiosk, 'save_checkin_to_sharepoint', return_value=('item_99', None)) as save_sp:
                response = self.post(var)
                self.assertEqual(response.status_code, 200)
                data = response.json
                self.assertEqual(data['status'], 'SUCCESS')
                self.assertEqual(data['order']['empId'], '1')
                self.assertEqual(data['order']['fullName'], 'ĐỖ VĂN TOÀN')
                self.assertEqual(data['order']['role'], 'Tổng Giám Đốc')
                self.assertTrue(data['spSynced'])
                save_sp.assert_called_once()

    def test_tgd_bypasses_morning_time_restriction(self):
        """Tổng Giám Đốc quét trước 11h30 (ví dụ 10h30) vẫn được phê duyệt thành công ngay lập tức"""
        with mock.patch.object(kiosk.datetime, 'datetime', Morning), \
             mock.patch.object(kiosk, 'get_access_token', return_value='fake_token'), \
             mock.patch.object(kiosk, 'save_checkin_to_sharepoint', return_value=('item_99', None)):
            # Quét biến thể 0001 trước giờ ăn
            response = self.post('0001')
            self.assertEqual(response.status_code, 200)
            self.assertEqual(response.json['status'], 'SUCCESS')
            self.assertEqual(response.json['order']['fullName'], 'ĐỖ VĂN TOÀN')

    def test_tgd_already_claimed_warning(self):
        """Tổng Giám Đốc quét lần 2 sẽ nhận cảnh báo đã nhận cơm trước đó"""
        with mock.patch.object(kiosk, 'get_access_token', return_value='fake_token'), \
             mock.patch.object(kiosk, 'save_checkin_to_sharepoint', return_value=('item_99', None)):
            res1 = self.post('00001')
            self.assertEqual(res1.json['status'], 'SUCCESS')

            # Quét lần 2 bằng mã '1'
            res2 = self.post('1')
            self.assertEqual(res2.json['status'], 'ALREADY_CLAIMED')

    # =========================================================================
    # 3. KIỂM THỬ ÉP BUỘC ĐĂNG NHẬP MICROSOFT 365
    # =========================================================================
    def test_unauthenticated_blocks_checkin(self):
        """Khi chưa đăng nhập Microsoft 365, hệ thống từ chối mọi thao tác quét mã"""
        with mock.patch.object(kiosk, 'get_access_token', return_value=None):
            response = self.post('1')
            self.assertEqual(response.status_code, 401)
            self.assertEqual(response.json['status'], 'REQUIRE_MS_LOGIN')
            self.assertFalse(response.json['authenticated'])

    def test_orders_today_requires_ms_login(self):
        """API lấy danh sách suất ăn yêu cầu đăng nhập Microsoft 365"""
        with mock.patch.object(kiosk, 'get_access_token', return_value=None):
            response = self.client.get('/api/orders/today')
            self.assertEqual(response.json['status'], 'REQUIRE_MS_LOGIN')
            self.assertFalse(response.json['authenticated'])

    # =========================================================================
    # 4. KIỂM THỬ NHÂN VIÊN THÔNG THƯỜNG
    # =========================================================================
    def test_regular_employee_normalization_match(self):
        """Nhân viên đăng ký mã 12345, quét với 012345 hoặc 12345 đều khớp đơn"""
        order = {
            'itemId': '42', 'empId': '12345', 'normEmpId': '12345',
            'fullName': 'Nguyễn Văn A', 'quantity': 1, 'mealType': 'Món mặn',
            'status': 'Đã đăng ký'
        }
        kiosk.CACHED_ORDERS_TODAY['12345'] = order

        with mock.patch.object(kiosk, 'get_access_token', return_value='fake_token'), \
             mock.patch.object(kiosk, 'save_checkin_to_sharepoint', return_value=('42', None)):
            # Quét với số 0 ở đầu
            response = self.post('012345')
            self.assertEqual(response.status_code, 200)
            self.assertEqual(response.json['status'], 'SUCCESS')
            self.assertEqual(response.json['order']['fullName'], 'Nguyễn Văn A')

    def test_regular_employee_before_time_blocked(self):
        """Nhân viên thường quét trước 11h30 bị chặn với trạng thái BEFORE_TIME"""
        order = {
            'itemId': '42', 'empId': '12345', 'normEmpId': '12345',
            'fullName': 'Nguyễn Văn A', 'quantity': 1, 'mealType': 'Món mặn',
            'status': 'Đã đăng ký'
        }
        kiosk.CACHED_ORDERS_TODAY['12345'] = order

        with mock.patch.object(kiosk.datetime, 'datetime', Morning), \
             mock.patch.object(kiosk, 'get_access_token', return_value='fake_token'):
            response = self.post('12345')
            self.assertEqual(response.json['status'], 'BEFORE_TIME')

    def test_unregistered_employee_rejected(self):
        """Mã nhân viên chưa đăng ký bị từ chối"""
        with mock.patch.object(kiosk, 'get_access_token', return_value='fake_token'):
            response = self.post('99999')
            self.assertEqual(response.json['status'], 'NOT_REGISTERED')

    # =========================================================================
    # 5. KIỂM THỬ DPAPI & LƯU TRỮ TOKEN ẨN APPDATA
    # =========================================================================
    def test_dpapi_encrypt_decrypt_cycle(self):
        """Mã hóa và giải mã DPAPI bảo toàn 100% nội dung dữ liệu"""
        test_payload = b'{"access_token": "secret_msal_token_123", "account": "toan.dovan@seazen.vn"}'
        enc = kiosk.dpapi_encrypt(test_payload)
        if os.name == 'nt':
            self.assertTrue(enc.startswith(b'DPAPI:'))
        dec = kiosk.dpapi_decrypt(enc)
        self.assertEqual(dec, test_payload)

    def test_appdata_storage_path(self):
        """Đường dẫn token cache nằm trong thư mục ẩn AppData của máy cục bộ"""
        self.assertIn("SeaZen_Canteen_Kiosk", kiosk.TOKEN_CACHE_FILE)
        self.assertTrue(os.path.isabs(kiosk.TOKEN_CACHE_FILE))


if __name__ == '__main__':
    unittest.main()
