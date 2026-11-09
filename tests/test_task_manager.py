# -*- coding: utf-8 -*-
"""
Tests for TaskManager: FIFO Queue, SQLite Persistence, Cancellation and Recovery
"""
import time
import threading
import pytest
from pathlib import Path

from src.task_manager import TaskManager, TaskRecord


def test_scan_snapshot_includes_all_waiting_scans_and_updates_after_cancel(tmp_path):
    tm = TaskManager(str(tmp_path))
    started, release = threading.Event(), threading.Event()

    def blocker(task, cancel_check):
        started.set()
        release.wait(5)

    first = tm.submit_task('excel_face', 'first', 'A', blocker)
    assert started.wait(2)
    try:
        waiting = [tm.submit_task('pdf_face', str(i), 'B', lambda *_: None) for i in range(5)]
        snapshot = tm.scan_snapshot(limit=1)
        assert snapshot['counts'] == {'running': 1, 'queued': 5, 'cancelling': 0}
        assert len(snapshot['active']) == 6  # History limit cannot hide active scans.
        assert tm.queue_info(waiting[-1].task_id) == {'queue_position': 5, 'ahead_count': 5}
        assert tm.cancel_task(waiting[0].task_id)
        assert tm.queue_info(waiting[1].task_id) == {'queue_position': 1, 'ahead_count': 1}
        assert tm.cancel_task(first.task_id)
        assert tm.cancel_task(first.task_id)  # Repeating a pending cancel is harmless.
        assert tm.scan_snapshot()['counts'] == {'running': 0, 'queued': 4, 'cancelling': 1}
    finally:
        release.set()
        tm._task_queue.join()
    assert first.status == 'cancelled'
    assert waiting[0].task_id not in tm._callbacks
    assert tm.scan_snapshot()['active'] == []


def test_idle_submission_is_preparing_without_claiming_another_scan(tmp_path):
    tm = TaskManager(str(tmp_path))
    # Keep worker outside the transition while inspecting the newly submitted scan.
    with tm._lock:
        task = tm.submit_task('excel_face', 'only', 'A', lambda *_: None)
        assert tm.queue_info(task.task_id) == {'queue_position': 1, 'ahead_count': 0}
        assert tm.cancel_task(task.task_id)
    tm._task_queue.join()
    assert task.status == 'cancelled'


def test_activity_distinguishes_active_person_from_last_completed(tmp_path):
    tm = TaskManager(str(tmp_path))
    release = threading.Event()
    task = tm.submit_task('excel_face', 'batch', 'A', lambda *_: release.wait(5))
    try:
        tm.update_progress(task.task_id, 19, 79, 'Last completed')
        tm.update_activity(task.task_id, 'worker-1', {'person': 'Still scanning', 'day': '2026-10-03', 'image': '002.jpg'})
        tm.update_activity(task.task_id, 'worker-1', {'stage': 'Đang nhận diện ảnh'})
        snapshot = tm.scan_snapshot()['active'][0]
        assert snapshot['current'] == 'Last completed'
        worker = snapshot['active_workers'][0]
        assert worker['person'] == 'Still scanning'
        assert worker['image'] == '002.jpg'
        assert worker['idle_seconds'] >= 0
        tm.update_activity(task.task_id, 'worker-1', None)
        assert tm.scan_snapshot()['active'][0]['active_workers'] == []
    finally:
        release.set()
        tm._task_queue.join()


def test_task_queue_fifo_execution(tmp_path):
    tm = TaskManager(str(tmp_path))
    execution_order = []

    def task1_fn(task, cancel_check):
        time.sleep(0.1)
        execution_order.append(1)

    def task2_fn(task, cancel_check):
        execution_order.append(2)

    # Submit 2 tasks
    t1 = tm.submit_task("test", "folder1", "proj", task1_fn)
    t2 = tm.submit_task("test", "folder2", "proj", task2_fn)

    # Wait for completion
    for _ in range(30):
        if t2.status == "completed":
            break
        time.sleep(0.1)

    assert t1.status == "completed"
    assert t2.status == "completed"
    assert execution_order == [1, 2]


def test_task_cancellation_running(tmp_path):
    tm = TaskManager(str(tmp_path))
    started = False
    cancelled_observed = False

    def slow_fn(task, cancel_check):
        nonlocal started, cancelled_observed
        started = True
        for _ in range(50):
            if cancel_check():
                cancelled_observed = True
                return
            time.sleep(0.05)

    t = tm.submit_task("test", "folder", "proj", slow_fn)

    # Wait until it starts running
    for _ in range(20):
        if started:
            break
        time.sleep(0.05)

    assert tm.cancel_task(t.task_id) is True

    # Wait for cancellation
    for _ in range(20):
        if t.status == "cancelled":
            break
        time.sleep(0.05)

    assert t.status == "cancelled"
    assert cancelled_observed is True


def test_task_cancellation_queued(tmp_path):
    tm = TaskManager(str(tmp_path))
    gate = False

    def blocker_fn(task, cancel_check):
        while not gate:
            time.sleep(0.05)

    def queued_fn(task, cancel_check):
        pass

    t1 = tm.submit_task("test", "folder1", "proj", blocker_fn)
    t2 = tm.submit_task("test", "folder2", "proj", queued_fn)

    assert t2.status == "queued"
    assert tm.cancel_task(t2.task_id) is True
    assert t2.status == "cancelled"

    gate = True  # unblock t1
    time.sleep(0.2)
    assert t1.status == "completed"
    assert t2.status == "cancelled"


def test_task_recovery_on_startup(tmp_path):
    # 1. Create a task manager and write a 'running' task directly
    tm1 = TaskManager(str(tmp_path))
    with tm1._connect() as conn:
        conn.execute("""
            INSERT INTO tasks (
                task_id, task_type, project_name, target_folder, status,
                progress, total, current_item, files_json, aggregate_report,
                errors_json, payload_json, start_time, end_time, created_at
            ) VALUES ('abandoned_task', 'excel_face', 'Proj', 'fold', 'running',
                      5, 10, 'Nguyen Van A', '[]', NULL, '[]', '{}', NULL, NULL, '2026-09-18')
        """)
        conn.commit()

    # 2. Simulate server restart by creating new TaskManager instance
    tm2 = TaskManager(str(tmp_path))
    rec = tm2.get_task("abandoned_task")
    assert rec is not None
    assert rec.status == "interrupted"
    assert "gián đoạn" in rec.errors[0].lower()
