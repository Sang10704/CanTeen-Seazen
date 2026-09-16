import sys
from zk import ZK

def mo_cua(ip_address, port=4370, delay_sec=3):
    print(f"=== ĐANG THỬ KẾT NỐI TỚI MÁY {ip_address}:{port} (UDP Mode) ===")
    zk = ZK(ip_address, port=port, timeout=4, password=0, force_udp=True, ommit_ping=True)
    conn = None
    try:
        conn = zk.connect()
        print(" -> KẾT NỐI THÀNH CÔNG!")
        
        firmware = conn.get_firmware_version()
        sn = conn.get_serialnumber()
        print(f" -> Firmware: {firmware}")
        print(f" -> Serial Number: {sn}")
        
        print(f" -> ĐANG GỬI LỆNH ÉP MỞ CỬA ({delay_sec} GIÂY)...")
        conn.unlock(delay_sec)
        print(" -> ĐÃ KÍCH HOẠT RELAY MỞ CỬA THÀNH CÔNG! (Hãy lắng nghe tiếng chốt cửa)")
        
        conn.disconnect()
        return True
    except Exception as e:
        print(f" -> Thất bại qua UDP ({e})")
        if conn:
            try:
                conn.disconnect()
            except:
                pass
        
    print(f"=== ĐANG THỬ KẾT NỐI TỚI MÁY {ip_address}:{port} (TCP Mode) ===")
    zk_tcp = ZK(ip_address, port=port, timeout=4, password=0, force_udp=False, ommit_ping=True)
    conn_tcp = None
    try:
        conn_tcp = zk_tcp.connect()
        print(" -> KẾT NỐI THÀNH CÔNG QUA TCP!")
        
        print(f" -> ĐANG GỬI LỆNH ÉP MỞ CỬA ({delay_sec} GIÂY)...")
        conn_tcp.unlock(delay_sec)
        print(" -> ĐÃ KÍCH HOẠT RELAY MỞ CỬA THÀNH CÔNG!")
        
        conn_tcp.disconnect()
        return True
    except Exception as e:
        print(f" -> Thất bại qua TCP ({e})")
        if conn_tcp:
            try:
                conn_tcp.disconnect()
            except:
                pass
        return False

if __name__ == "__main__":
    ips = ["172.16.0.211", "172.16.0.151", "172.16.56.2"]
    for target_ip in ips:
        if mo_cua(target_ip, port=4370, delay_sec=3):
            print(f"\n ĐÃ ÉP MỞ CỬA THÀNH CÔNG TẠI IP {target_ip}!")
            break
