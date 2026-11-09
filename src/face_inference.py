"""Bounded TensorFlow workers shared by all scan threads.

Native inference cannot be interrupted safely inside a Python thread. A child
process gives cancellation and timeout a real boundary, without changing the
face detector, recognition model, input resolution or comparison threshold.
"""
import atexit
import multiprocessing
import os
import sys
import threading
import time
from collections import OrderedDict
from pathlib import Path
import json
import ctypes
import math


def usable_face_representations(values, image_path):
    """Reject tiny avatar/ID faces before their vectors can establish attendance."""
    from PIL import Image, ImageOps
    with Image.open(image_path) as image:
        width, height = ImageOps.exif_transpose(image).size
    accepted = []
    for value in values:
        area = value.get('facial_area') or {}
        try:
            x, y, w, h = (float(area[k]) for k in ('x', 'y', 'w', 'h'))
            confidence = float(value.get('face_confidence', 0))
        except (KeyError, TypeError, ValueError):
            continue
        if (not all(math.isfinite(v) for v in (x, y, w, h, confidence))
                or confidence < .9 or min(w, h) < 40
                or x < 0 or y < 0 or x + w > width or y + h > height
                or w * h / (width * height) < .005):
            continue
        accepted.append(value)
    return accepted

from src.scan_runtime import ScanCancelled, check_scan_cancelled, report_scan_activity


class FaceInferenceError(RuntimeError):
    pass


def _worker(connection, threads=None):
    # These must be set before importing TensorFlow/DeepFace in this process.
    threads = str(threads or max(1, min(2, (os.cpu_count() or 2) - 1)))
    os.environ['TF_NUM_INTRAOP_THREADS'] = threads
    os.environ['TF_NUM_INTEROP_THREADS'] = '1'
    os.environ['OMP_NUM_THREADS'] = threads
    os.environ.setdefault('TF_CPP_MIN_LOG_LEVEL', '2')
    if os.name == 'nt':
        # Let the UI and other desktop applications preempt these CPU-heavy jobs.
        import subprocess
        kernel = ctypes.WinDLL('kernel32', use_last_error=True)
        kernel.GetCurrentProcess.restype = ctypes.c_void_p
        kernel.SetPriorityClass.argtypes = [ctypes.c_void_p, ctypes.c_ulong]
        kernel.SetPriorityClass(kernel.GetCurrentProcess(), subprocess.BELOW_NORMAL_PRIORITY_CLASS)
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, 'reconfigure'):
            stream.reconfigure(encoding='utf-8', errors='replace')
    try:
        from deepface import DeepFace
        connection.send(('ready', None))
        while True:
            request = connection.recv()
            if request is None:
                return
            try:
                values = DeepFace.represent(**request)
                values = usable_face_representations(values, request['img_path'])
                connection.send(('ok', [{'embedding': value['embedding']} for value in values]))
            except Exception as exc:
                from deepface.modules.exceptions import FaceNotDetected
                if isinstance(exc, FaceNotDetected):
                    connection.send(('ok', []))
                else:
                    connection.send(('error', f'{type(exc).__name__}: {exc}'))
    except EOFError:
        pass
    except Exception as exc:
        try:
            connection.send(('fatal', f'{type(exc).__name__}: {exc}'))
        except (OSError, EOFError):
            pass
    finally:
        connection.close()


class _InferenceWorker:
    def __init__(self, timeout=120, worker_target=_worker, startup_timeout=300, threads=1):
        self.timeout = timeout
        self.startup_timeout = startup_timeout
        self._worker_target = worker_target
        self._lock = threading.Lock()
        self._process = None
        self._connection = None
        self._cache = OrderedDict()
        self.threads = threads
        self.warmed = False

    def _start(self):
        if self._connection is not None and self._process is not None and self._process.is_alive():
            return
        self.close()
        if self._process is not None:
            raise FaceInferenceError('Bộ nhận diện cũ đang dừng; vui lòng thử lại sau')
        context = multiprocessing.get_context('spawn')
        parent, child = context.Pipe()
        args = (child, self.threads) if self._worker_target is _worker else (child,)
        process = context.Process(target=self._worker_target, args=args, daemon=True)
        try:
            process.start()
        except BaseException:
            parent.close()
            child.close()
            raise
        child.close()
        self._process, self._connection = process, parent
        report_scan_activity('Đang nạp bộ nhận diện lần đầu')
        deadline = time.monotonic() + self.startup_timeout
        while not parent.poll(0.1):
            check_scan_cancelled()
            if not process.is_alive():
                raise FaceInferenceError('Bộ nhận diện đã thoát khi khởi tạo')
            if time.monotonic() >= deadline:
                raise FaceInferenceError(f'Khởi tạo AI vượt quá {self.startup_timeout:g} giây')
        status, value = parent.recv()
        if status != 'ready':
            raise FaceInferenceError(f'Không khởi tạo được bộ nhận diện: {value}')

    def close(self):
        if self._connection is not None:
            self._connection.close()
            self._connection = None
        if self._process is not None:
            if self._process.is_alive():
                self._process.terminate()
            self._process.join(timeout=2)
            if self._process.is_alive():
                self._process.kill()
                self._process.join(timeout=2)
            if not self._process.is_alive():
                self._process.close()
                self._process = None
        self.warmed = False

    def represent(self, **request):
        report_scan_activity('Đang chờ bộ nhận diện')
        while not self._lock.acquire(timeout=0.1):
            check_scan_cancelled()
        try:
            check_scan_cancelled()
            image_path = request.get('img_path')
            stat = os.stat(image_path) if isinstance(image_path, str) else None
            key = (tuple(sorted(request.items())), stat.st_size, stat.st_mtime_ns) if stat else None
            if key is not None and key in self._cache:
                self._cache.move_to_end(key)
                return self._cache[key]
            self._start()
            report_scan_activity('Đang nhận diện ảnh')
            deadline = time.monotonic() + self.timeout
            self._connection.send(request)
            while not self._connection.poll(0.1):
                check_scan_cancelled()
                if not self._process.is_alive():
                    raise FaceInferenceError('Bộ nhận diện đã thoát bất thường')
                if time.monotonic() >= deadline:
                    raise FaceInferenceError(f'Nhận diện một ảnh vượt quá {self.timeout:g} giây; đã dừng bộ xử lý để tránh treo lượt quét')
            check_scan_cancelled()
            status, value = self._connection.recv()
            if status == 'fatal':
                raise FaceInferenceError(f'Không khởi tạo được bộ nhận diện: {value}')
            if status == 'error':
                raise ValueError(value)
            if key is not None:
                self._cache[key] = value
                if len(self._cache) > 256:
                    self._cache.popitem(last=False)
            return value
        except (ScanCancelled, FaceInferenceError):
            self.close()
            raise
        except (OSError, EOFError) as exc:
            self.close()
            raise FaceInferenceError(f'Mất kết nối bộ nhận diện: {exc}') from exc
        finally:
            self._lock.release()


GIB = 1024 ** 3
ESTIMATED_WORKER_RAM = 2 * GIB


def detect_hardware():
    """Read host memory without importing TensorFlow or adding a dependency."""
    total = available = None
    if os.name == 'nt':
        class MemoryStatus(ctypes.Structure):
            _fields_ = [('length', ctypes.c_ulong), ('load', ctypes.c_ulong),
                        ('total', ctypes.c_ulonglong), ('available', ctypes.c_ulonglong),
                        ('page_total', ctypes.c_ulonglong), ('page_available', ctypes.c_ulonglong),
                        ('virtual_total', ctypes.c_ulonglong), ('virtual_available', ctypes.c_ulonglong),
                        ('extended', ctypes.c_ulonglong)]
        memory = MemoryStatus()
        memory.length = ctypes.sizeof(memory)
        if ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(memory)):
            total, available = int(memory.total), int(memory.available)
    else:
        try:
            page = os.sysconf('SC_PAGE_SIZE')
            total = page * os.sysconf('SC_PHYS_PAGES')
            available = page * os.sysconf('SC_AVPHYS_PAGES')
        except (AttributeError, OSError, ValueError):
            pass
    return {'cpu_count': os.cpu_count() or 1, 'total_ram': total, 'available_ram': available}


def hardware_limit(hardware):
    cpu_limit = max(1, hardware['cpu_count'] - 1)
    total = hardware.get('total_ram')
    memory_limit = max(1, int((total - max(2 * GIB, total * 0.2)) // ESTIMATED_WORKER_RAM)) if total else 1
    return min(8, cpu_limit, memory_limit)


class FaceInferenceRunner:
    """A CPU/RAM-sized pool; one native call per child and shared result cache."""
    def __init__(self, timeout=120, worker_target=_worker, startup_timeout=300,
                 base_dir=None, hardware_provider=detect_hardware):
        self.timeout, self.startup_timeout = timeout, startup_timeout
        self._worker_target = worker_target
        self._hardware_provider = hardware_provider
        self._hardware = hardware_provider()
        self.hardware_max = hardware_limit(self._hardware)
        self._condition = threading.Condition(threading.RLock())
        self._workers, self._busy, self._inflight = [], set(), set()
        self._cache = OrderedDict()
        self._waiting = self._memory_waiting = 0
        self._pending = None
        self._settings_path = Path(base_dir) / 'data' / 'ai_execution.json' if base_dir else None
        mode = 'auto'
        if self._settings_path and self._settings_path.exists():
            try:
                stored = json.loads(self._settings_path.read_text(encoding='utf-8')).get('mode', 'auto')
                mode = self._validate(stored)
            except (OSError, ValueError, TypeError):
                pass
        self.mode = mode
        self.max_workers = self.hardware_max if mode == 'auto' else mode
        self._threads = max(1, min(2, self._hardware['cpu_count'] // self.max_workers))
        atexit.register(self.close)

    @property
    def _process(self):
        # Kept for callers inspecting a one-worker runner during diagnostics.
        return self._workers[0]._process if self._workers else None

    def _validate(self, mode):
        if mode == 'auto':
            return mode
        if type(mode) is not int or not 1 <= mode <= self.hardware_max:
            raise ValueError(f'Số lượt AI phải từ 1 đến {self.hardware_max}, hoặc chọn Tự động')
        return mode

    def configure(self, mode):
        mode = self._validate(mode)
        with self._condition:
            if self._settings_path:
                self._settings_path.parent.mkdir(parents=True, exist_ok=True)
                temporary = self._settings_path.with_suffix('.tmp')
                temporary.write_text(json.dumps({'mode': mode}), encoding='utf-8')
                os.replace(temporary, self._settings_path)
            desired = self.hardware_max if mode == 'auto' else mode
            if desired == self.max_workers:
                self.mode = mode
                self._pending = None
            elif self._busy:
                # Drain current photos before changing model CPU allocations.
                # New requests wait, so continuous traffic cannot starve resizing.
                self._pending = (mode, desired)
            else:
                self._apply(mode, desired)
            self._condition.notify_all()
        return self.status()

    def _apply(self, mode, desired):
        threads = max(1, min(2, self._hardware['cpu_count'] // desired))
        for index, worker in enumerate(self._workers):
            if index >= desired or worker.threads != threads:
                worker.close()
            worker.threads = threads
        self.mode, self.max_workers, self._threads = mode, desired, threads
        self._pending = None

    def status(self):
        hardware = self._hardware_provider()
        with self._condition:
            return {'mode': self.mode, 'max_workers': self.max_workers, 'hardware_max': self.hardware_max,
                    'active': len(self._busy), 'waiting': self._waiting, 'memory_waiting': self._memory_waiting,
                    'threads_per_worker': self._threads, 'cpu_count': hardware['cpu_count'],
                    'total_ram_gb': round(hardware['total_ram'] / GIB, 1) if hardware.get('total_ram') else None,
                    'available_ram_gb': round(hardware['available_ram'] / GIB, 1) if hardware.get('available_ram') is not None else None,
                    'pending_mode': self._pending[0] if self._pending else None,
                    'pending_max_workers': self._pending[1] if self._pending else None}

    def close(self):
        for worker in self._workers:
            worker.close()

    def represent(self, **request):
        image_path = request.get('img_path')
        stat = os.stat(image_path) if isinstance(image_path, str) else None
        key = (tuple(sorted(request.items())), stat.st_size, stat.st_mtime_ns) if stat else None
        report_scan_activity('Đang chờ bộ nhận diện', ai_slot=None)
        slot, memory_wait = None, False
        last_reason = None
        with self._condition:
            self._waiting += 1
            try:
                while slot is None:
                    check_scan_cancelled()
                    if key is not None and key in self._cache:
                        self._cache.move_to_end(key)
                        return self._cache[key]
                    candidate = next((i for i in range(self.max_workers) if i not in self._busy), None)
                    reason = 'Đang chờ bộ nhận diện'
                    if self._pending:
                        candidate = None
                        reason = 'Đang chờ áp dụng số lượt AI mới'
                    elif key is not None and key in self._inflight:
                        candidate = None
                        reason = 'Đang chờ kết quả ảnh đang được xử lý'
                    elif candidate is not None:
                        worker = self._workers[candidate] if candidate < len(self._workers) else None
                        cold = worker is None or worker._process is None
                        available = self._hardware_provider().get('available_ram')
                        if available is not None:
                            available -= sum(ESTIMATED_WORKER_RAM for i in self._busy if not self._workers[i].warmed)
                        if cold and available is not None and available < ESTIMATED_WORKER_RAM + GIB and self._busy:
                            candidate = None
                            reason = 'Đang chờ đủ RAM để mở thêm lượt AI'
                    needs_memory = 'RAM' in reason
                    if needs_memory != memory_wait:
                        self._memory_waiting += 1 if needs_memory else -1
                        memory_wait = needs_memory
                    if candidate is not None:
                        slot = candidate
                        while len(self._workers) <= slot:
                            self._workers.append(_InferenceWorker(self.timeout, self._worker_target, self.startup_timeout, self._threads))
                        self._busy.add(slot)
                        if key is not None:
                            self._inflight.add(key)
                    else:
                        # Do not reset the activity timer on every polling tick.
                        if reason != last_reason:
                            report_scan_activity(reason, ai_slot=None)
                            last_reason = reason
                        self._condition.wait(timeout=0.1)
            finally:
                self._waiting -= 1
                if memory_wait:
                    self._memory_waiting -= 1
        try:
            worker = self._workers[slot]
            worker.timeout, worker.startup_timeout = self.timeout, self.startup_timeout
            report_scan_activity('Đang nhận diện ảnh', ai_slot=slot + 1)
            value = worker.represent(**request)
            worker.warmed = True
            if key is not None:
                with self._condition:
                    self._cache[key] = value
                    if len(self._cache) > 256:
                        self._cache.popitem(last=False)
            return value
        finally:
            with self._condition:
                self._busy.discard(slot)
                if key is not None:
                    self._inflight.discard(key)
                if not self._busy and self._pending:
                    self._apply(*self._pending)
                self._condition.notify_all()
