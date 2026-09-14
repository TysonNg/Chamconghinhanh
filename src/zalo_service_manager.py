# -*- coding: utf-8 -*-
"""
Zalo Service Process Manager
Quản lý vòng đời tiến trình Node.js Zalo Service:
- Khởi động an toàn không bị chặn pipe buffer
- Tự động dọn dẹp port xung đột (3001)
- Kiểm tra trạng thái sẵn sàng (health check)
- Tự phục hồi (auto-recovery) khi service gặp sự cố
- Tự động đóng tiến trình con khi app chính thoát (Windows Job Object)
"""

import os
import sys
import time
import logging
import subprocess
import urllib.request
import urllib.error

ZALO_HOST = "127.0.0.1"
ZALO_PORT = 3001
ZALO_BASE_URL = f"http://{ZALO_HOST}:{ZALO_PORT}"

# Toàn cục lưu subprocess
_zalo_process = None
_zalo_job_handle = None
_zalo_log_file_handle = None


def get_base_dir():
    """Lấy thư mục gốc chứa data và zalo-service"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def is_zalo_service_healthy(timeout=1.5):
    """Kiểm tra Zalo Service có đang phản hồi 200 OK trên port 3001 không"""
    url = f"{ZALO_BASE_URL}/api/health"
    try:
        req = urllib.request.Request(url, method='GET')
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return resp.status == 200
    except Exception:
        return False


def kill_process_on_port(port=ZALO_PORT):
    """Tìm và cưỡng chế tắt mọi tiến trình cũ đang chiếm port (tránh EADDRINUSE)"""
    if sys.platform != 'win32':
        return
    try:
        cmd = f'netstat -ano | findstr :{port}'
        output = subprocess.check_output(cmd, shell=True, text=True, errors='replace')
        pids = set()
        for line in output.strip().split('\n'):
            parts = line.strip().split()
            # Cần dòng có trạng thái LISTENING và PID ở cột cuối cùng
            if len(parts) >= 5 and 'LISTENING' in parts:
                try:
                    pid = int(parts[-1])
                    if pid > 0 and pid != os.getpid():
                        pids.add(pid)
                except ValueError:
                    pass
        for pid in pids:
            try:
                subprocess.run(['taskkill', '/F', '/PID', str(pid)], capture_output=True, timeout=5)
                logging.info(f"[ZaloManager] Đã đóng tiến trình cũ (PID {pid}) đang chiếm port {port}")
            except Exception:
                pass
    except Exception:
        pass


def _assign_to_job_object(proc):
    """Gắn tiến trình con vào Windows Job Object để tự động tắt khi process cha thoát"""
    global _zalo_job_handle
    if sys.platform != 'win32' or not proc:
        return
    try:
        import ctypes
        from ctypes import wintypes
        kernel32 = ctypes.windll.kernel32

        if _zalo_job_handle is None:
            _zalo_job_handle = kernel32.CreateJobObjectW(None, None)
            if not _zalo_job_handle:
                return

            class IO_COUNTERS(ctypes.Structure):
                _fields_ = [('ReadOperationCount', ctypes.c_uint64),
                            ('WriteOperationCount', ctypes.c_uint64),
                            ('OtherOperationCount', ctypes.c_uint64),
                            ('ReadTransferCount', ctypes.c_uint64),
                            ('WriteTransferCount', ctypes.c_uint64),
                            ('OtherTransferCount', ctypes.c_uint64)]

            class JOBOBJECT_BASIC_LIMIT_INFORMATION(ctypes.Structure):
                _fields_ = [('PerProcessUserTimeLimit', ctypes.c_int64),
                            ('PerJobUserTimeLimit', ctypes.c_int64),
                            ('LimitFlags', wintypes.DWORD),
                            ('MinimumWorkingSetSize', ctypes.c_size_t),
                            ('MaximumWorkingSetSize', ctypes.c_size_t),
                            ('ActiveProcessLimit', wintypes.DWORD),
                            ('Affinity', ctypes.c_size_t),
                            ('PriorityClass', wintypes.DWORD),
                            ('SchedulingClass', wintypes.DWORD)]

            class JOBOBJECT_EXTENDED_LIMIT_INFORMATION(ctypes.Structure):
                _fields_ = [('BasicLimitInformation', JOBOBJECT_BASIC_LIMIT_INFORMATION),
                            ('IoInfo', IO_COUNTERS),
                            ('ProcessMemoryLimit', ctypes.c_size_t),
                            ('JobMemoryLimit', ctypes.c_size_t),
                            ('PeakProcessMemoryLimit', ctypes.c_size_t),
                            ('PeakJobMemoryLimit', ctypes.c_size_t)]

            info = JOBOBJECT_EXTENDED_LIMIT_INFORMATION()
            info.BasicLimitInformation.LimitFlags = 0x2000  # JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
            JobObjectExtendedLimitInformation = 9
            kernel32.SetInformationJobObject(
                _zalo_job_handle, JobObjectExtendedLimitInformation,
                ctypes.byref(info), ctypes.sizeof(info)
            )

        hProcess = kernel32.OpenProcess(0x1F0FFF, False, proc.pid)
        if hProcess:
            kernel32.AssignProcessToJobObject(_zalo_job_handle, hProcess)
            kernel32.CloseHandle(hProcess)
            logging.info(f"[ZaloManager] Đã gắn PID {proc.pid} vào Windows Job Object")
    except Exception as e:
        logging.warning(f"[ZaloManager] Không thể gán Job Object: {e}")


def start_zalo_service(base_dir=None, wait_ready_seconds=10):
    """
    Khởi động Zalo Service Node.js và đợi cho tới khi port 3001 phản hồi.
    Ghi log ra file để không bao giờ bị nghẽn buffer subprocess pipe.
    """
    global _zalo_process, _zalo_log_file_handle

    if base_dir is None:
        base_dir = get_base_dir()

    zalo_service_dir = os.path.join(base_dir, 'zalo-service')
    if not os.path.exists(zalo_service_dir):
        logging.warning(f"[ZaloManager] Không tìm thấy thư mục: {zalo_service_dir}")
        return False

    # 1. Kiểm tra Node.js
    try:
        node_check = subprocess.run(['node', '--version'], capture_output=True, text=True, timeout=5)
        if node_check.returncode != 0:
            logging.warning("[ZaloManager] Node.js không khả dụng")
            return False
        logging.info(f"[ZaloManager] Node.js phiên bản: {node_check.stdout.strip()}")
    except Exception as e:
        logging.warning(f"[ZaloManager] Lỗi kiểm tra Node.js: {e}")
        return False

    # 2. Kiểm tra node_modules
    if not os.path.exists(os.path.join(zalo_service_dir, 'node_modules')):
        logging.info("[ZaloManager] Đang cài đặt npm packages cho zalo-service...")
        try:
            subprocess.run(['npm', 'install'], cwd=zalo_service_dir, timeout=120, capture_output=True, shell=True)
        except Exception as e:
            logging.error(f"[ZaloManager] Lỗi npm install: {e}")

    # 3. Nếu port 3001 đang phản hồi tốt thì tái sử dụng
    if is_zalo_service_healthy(timeout=1.0):
        logging.info("[ZaloManager] Zalo Service đã đang chạy và phản hồi tốt trên port 3001!")
        return True

    # 4. Nếu port 3001 bị chiếm bởi tiến trình treo, kill nó trước
    kill_process_on_port(ZALO_PORT)
    time.sleep(0.5)

    # 5. Mở file log riêng để tránh tràn pipe buffer
    log_file_path = os.path.join(zalo_service_dir, 'zalo_service.log')
    try:
        _zalo_log_file_handle = open(log_file_path, 'a', encoding='utf-8')
        _zalo_log_file_handle.write(f"\n\n--- Zalo Service Started at {time.strftime('%Y-%m-%d %H:%M:%S')} ---\n")
        _zalo_log_file_handle.flush()
    except Exception:
        _zalo_log_file_handle = subprocess.DEVNULL

    logging.info(f"[ZaloManager] Đang khởi động 'node server.js' tại {zalo_service_dir}...")
    try:
        creationflags = subprocess.CREATE_NO_WINDOW if sys.platform == 'win32' else 0
        _zalo_process = subprocess.Popen(
            ['node', 'server.js'],
            cwd=zalo_service_dir,
            stdout=_zalo_log_file_handle,
            stderr=subprocess.STDOUT,
            creationflags=creationflags
        )

        _assign_to_job_object(_zalo_process)

        # 6. Polling kiểm tra trạng thái sẵn sàng
        start_time = time.time()
        while time.time() - start_time < wait_ready_seconds:
            if _zalo_process.poll() is not None:
                exit_code = _zalo_process.poll()
                logging.error(f"[ZaloManager] Node.js thoát sớm với mã lỗi {exit_code}. Xem chi tiết: {log_file_path}")
                return False

            if is_zalo_service_healthy(timeout=0.8):
                logging.info(f"[ZaloManager] ✓ Zalo Service đã khởi động và sẵn sàng trên port {ZALO_PORT}!")
                return True

            time.sleep(0.4)

        if is_zalo_service_healthy(timeout=1.0):
            logging.info(f"[ZaloManager] ✓ Zalo Service sẵn sàng!")
            return True
        else:
            logging.warning(f"[ZaloManager] Quá thời gian {wait_ready_seconds}s chờ Zalo Service phản hồi.")
            return False

    except Exception as e:
        logging.error(f"[ZaloManager] Lỗi khi tạo subprocess Zalo: {e}")
        return False


def ensure_zalo_service(base_dir=None):
    """Đảm bảo Zalo service đang chạy, nếu chưa hoặc crash thì khởi động lại"""
    if is_zalo_service_healthy(timeout=1.0):
        return True
    logging.info("[ZaloManager] Zalo Service không phản hồi, đang tự động khởi chạy lại...")
    return start_zalo_service(base_dir=base_dir, wait_ready_seconds=6)


def stop_zalo_service():
    """Dừng tiến trình Zalo Service sạch sẽ"""
    global _zalo_process, _zalo_log_file_handle
    if _zalo_process and _zalo_process.poll() is None:
        logging.info("[ZaloManager] Đang dừng Zalo Service...")
        try:
            _zalo_process.terminate()
            _zalo_process.wait(timeout=3)
        except Exception:
            try:
                _zalo_process.kill()
            except Exception:
                pass
        _zalo_process = None

    if _zalo_log_file_handle and hasattr(_zalo_log_file_handle, 'close'):
        try:
            _zalo_log_file_handle.close()
        except Exception:
            pass
        _zalo_log_file_handle = None


def restart_zalo_service(base_dir=None):
    """Cưỡng chế khởi động lại Zalo Service"""
    stop_zalo_service()
    kill_process_on_port(ZALO_PORT)
    time.sleep(0.5)
    return start_zalo_service(base_dir=base_dir, wait_ready_seconds=8)
