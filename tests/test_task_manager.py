# -*- coding: utf-8 -*-
"""
Tests for TaskManager: FIFO Queue, SQLite Persistence, Cancellation and Recovery
"""
import time
import pytest
from pathlib import Path

from src.task_manager import TaskManager, TaskRecord


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
