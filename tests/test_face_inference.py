"""Exercise real process boundaries without loading TensorFlow in unit tests."""
import threading
import time
from concurrent.futures import ThreadPoolExecutor

import pytest

from src.face_inference import FaceInferenceRunner, FaceInferenceError, hardware_limit, GIB
from src.scan_runtime import ScanCancelled, scan_context


def fake_worker(connection):
    connection.send(('ready', None))
    while True:
        request = connection.recv()
        if request.get('hang'):
            time.sleep(60)
        elif request.get('crash'):
            connection.close()
            return
        elif request.get('error'):
            connection.send(('error', 'bad photo'))
        else:
            connection.send(('ok', [{'embedding': [1., 0.]}]))


def counting_worker(connection):
    connection.send(('ready', None))
    count = 0
    while True:
        connection.recv()
        count += 1
        connection.send(('ok', [{'embedding': [float(count), 0.]}]))


def timed_worker(connection):
    import os
    connection.send(('ready', None))
    while True:
        connection.recv()
        start = time.time()
        time.sleep(0.4)
        connection.send(('ok', [{'embedding': [start, time.time(), os.getpid()]}]))


def plenty_of_memory():
    return {'cpu_count': 4, 'total_ram': 7.9 * GIB, 'available_ram': 6 * GIB}


def test_native_timeout_is_bounded_and_worker_can_restart():
    runner = FaceInferenceRunner(timeout=1, worker_target=fake_worker)
    try:
        start = time.monotonic()
        with pytest.raises(FaceInferenceError, match='vượt quá'):
            runner.represent(hang=True)
        assert time.monotonic() - start < 6
        assert runner._process is None
        runner.timeout = 5
        assert runner.represent()[0]['embedding'] == [1., 0.]
    finally:
        runner.close()


def test_cancel_kills_native_call_and_releases_waiting_scan():
    runner = FaceInferenceRunner(timeout=20, worker_target=fake_worker)
    cancel, started = threading.Event(), threading.Event()

    def run():
        with scan_context(lambda detail: started.set() if detail['stage'] == 'Đang nhận diện ảnh' else None, cancel.is_set):
            return runner.represent(hang=True)

    try:
        with ThreadPoolExecutor(max_workers=2) as pool:
            future = pool.submit(run)
            assert started.wait(3)
            cancel.set()
            with pytest.raises(ScanCancelled):
                future.result(timeout=5)
        assert runner.represent() == [{'embedding': [1., 0.]}]
    finally:
        runner.close()


def test_worker_crash_and_model_errors_are_not_reported_as_matches():
    runner = FaceInferenceRunner(timeout=5, worker_target=fake_worker)
    try:
        with pytest.raises(FaceInferenceError):
            runner.represent(crash=True)
        with pytest.raises(ValueError, match='bad photo'):
            runner.represent(error=True)
        assert runner.represent()[0]['embedding'] == [1., 0.]
    finally:
        runner.close()


def test_concurrent_requests_share_one_result_and_changed_image_is_recomputed(tmp_path):
    image = tmp_path / 'same.jpg'
    image.write_bytes(b'first')
    runner = FaceInferenceRunner(timeout=5, worker_target=counting_worker)
    try:
        with ThreadPoolExecutor(max_workers=3) as pool:
            values = list(pool.map(lambda _: runner.represent(img_path=str(image)), range(3)))
        assert all(value[0]['embedding'] == [1., 0.] for value in values)
        image.write_bytes(b'changed file')
        assert runner.represent(img_path=str(image))[0]['embedding'] == [2., 0.]
    finally:
        runner.close()


def test_pool_runs_two_different_photos_at_the_same_time():
    runner = FaceInferenceRunner(worker_target=timed_worker, hardware_provider=plenty_of_memory)
    try:
        assert runner.max_workers == 2
        with ThreadPoolExecutor(max_workers=2) as pool:
            results = list(pool.map(lambda _: runner.represent(), range(2)))
        first, second = (value[0]['embedding'] for value in results)
        assert first[2] != second[2]
        assert max(first[0], second[0]) < min(first[1], second[1]), 'Native calls must actually overlap'
        assert runner.status()['active'] == 0
        assert runner.status()['threads_per_worker'] * runner.max_workers <= 4
    finally:
        runner.close()


def test_low_available_ram_queues_extra_worker_instead_of_overcommitting():
    hardware = lambda: {**plenty_of_memory(), 'available_ram': GIB}
    runner = FaceInferenceRunner(worker_target=timed_worker, hardware_provider=hardware)
    activities = []
    def run(_):
        with scan_context(lambda detail: activities.append(detail['stage'])):
            return runner.represent()
    try:
        with ThreadPoolExecutor(max_workers=2) as pool:
            results = list(pool.map(run, range(2)))
        first, second = (value[0]['embedding'] for value in results)
        assert first[2] == second[2], 'Low-RAM requests must reuse the existing worker'
        assert any('RAM' in stage for stage in activities)
        assert runner.status()['memory_waiting'] == 0
    finally:
        runner.close()


def test_resize_waits_for_current_photos_and_persists_selection(tmp_path):
    runner = FaceInferenceRunner(worker_target=timed_worker, hardware_provider=plenty_of_memory, base_dir=tmp_path)
    try:
        with ThreadPoolExecutor(max_workers=2) as pool:
            futures = [pool.submit(runner.represent) for _ in range(2)]
            deadline = time.monotonic() + 3
            while runner.status()['active'] < 2 and time.monotonic() < deadline:
                time.sleep(0.01)
            status = runner.configure(1)
            assert status['pending_max_workers'] == 1
            for future in futures:
                future.result(timeout=5)
        assert runner.status()['max_workers'] == 1
        assert runner.status()['pending_mode'] is None
        restored = FaceInferenceRunner(base_dir=tmp_path, hardware_provider=plenty_of_memory)
        assert restored.mode == 1
        restored.close()
        for invalid in (True, 0, 3, '2', None):
            with pytest.raises(ValueError):
                runner.configure(invalid)
    finally:
        runner.close()


def test_cancelling_one_worker_does_not_kill_another_worker():
    runner = FaceInferenceRunner(worker_target=fake_worker, hardware_provider=plenty_of_memory)
    cancelled = threading.Event()
    def hanging():
        with scan_context(cancel_check=cancelled.is_set):
            return runner.represent(hang=True)
    try:
        with ThreadPoolExecutor(max_workers=2) as pool:
            first = pool.submit(hanging)
            healthy = pool.submit(runner.represent)
            assert healthy.result(timeout=5)[0]['embedding'] == [1., 0.]
            cancelled.set()
            with pytest.raises(ScanCancelled):
                first.result(timeout=5)
        assert runner.represent()[0]['embedding'] == [1., 0.]
    finally:
        runner.close()


def test_unknown_memory_defaults_to_one_and_large_hosts_are_bounded():
    assert hardware_limit({'cpu_count': 16, 'total_ram': None}) == 1
    assert hardware_limit({'cpu_count': 1, 'total_ram': 64 * GIB}) == 1
    assert hardware_limit({'cpu_count': 64, 'total_ram': 256 * GIB}) == 8
