# -*- coding: utf-8 -*-
"""
Task Manager Module
Quản lý hàng đợi tác vụ tuần tự (Single Worker FIFO Queue) cho các tác vụ AI nặng:
- Giới hạn tối đa 1 tác vụ AI chạy tại một thời điểm (khóa độc quyền).
- Tự động xếp hàng các tác vụ bấm sau (trạng thái 'queued').
- Lưu trữ persistent xuống SQLite (data/tasks.sqlite3).
- Hỗ trợ dừng/hủy tác vụ tức thì (Cancel/Abort) và quét lại (Retry).
- Tự động đánh dấu 'interrupted' cho các tác vụ dở dang khi khởi động lại server.
"""

import json
import logging
import os
import queue
import sqlite3
import threading
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional

logger = logging.getLogger(__name__)


class TaskRecord:
    """Đại diện cho một tác vụ trong hệ thống"""

    def __init__(
        self,
        task_id: str,
        task_type: str,
        project_name: str = "",
        target_folder: str = "",
        status: str = "queued",
        progress: int = 0,
        total: int = 0,
        current_item: str = "",
        files: Optional[List[Dict]] = None,
        aggregate_report: Optional[str] = None,
        errors: Optional[List[str]] = None,
        payload: Optional[Dict] = None,
        start_time: Optional[str] = None,
        end_time: Optional[str] = None,
        created_at: Optional[str] = None,
    ):
        self.task_id = task_id
        self.task_type = task_type
        self.project_name = project_name
        self.target_folder = target_folder
        self.status = status  # queued | running | cancelling | cancelled | completed | failed | interrupted
        self.progress = progress
        self.total = total
        self.current_item = current_item
        self.files = files or []
        self.aggregate_report = aggregate_report
        self.errors = errors or []
        self.payload = payload or {}
        self.start_time = start_time
        self.end_time = end_time
        self.created_at = created_at or datetime.now().isoformat()
        self.cancel_requested = False

    def to_dict(self) -> Dict[str, Any]:
        return {
            "task_id": self.task_id,
            "task_type": self.task_type,
            "project_name": self.project_name,
            "target_folder": self.target_folder,
            "status": self.status,
            "progress": self.progress,
            "total": self.total,
            "current": self.current_item,
            "files": self.files,
            "aggregate_report": self.aggregate_report,
            "errors": self.errors,
            "start_time": self.start_time,
            "end_time": self.end_time,
            "created_at": self.created_at,
        }


class TaskManager:
    """Quản lý hàng đợi tác vụ và lưu trữ SQLite"""

    def __init__(self, base_dir: str):
        self.base_dir = Path(base_dir)
        self.data_dir = self.base_dir / "data"
        self.data_dir.mkdir(parents=True, exist_ok=True)
        self.db_path = self.data_dir / "tasks.sqlite3"

        self._lock = threading.RLock()
        self._task_queue = queue.Queue()
        self._tasks: Dict[str, TaskRecord] = {}
        self._callbacks: Dict[str, Callable] = {}

        self._current_running_id: Optional[str] = None
        self._current_cancel_event = threading.Event()

        self._init_db()
        self._recover_interrupted_tasks()
        self._load_active_tasks()

        # Khởi động Single Worker Daemon Thread
        self._worker_thread = threading.Thread(target=self._worker_loop, daemon=True)
        self._worker_thread.start()

    def _connect(self) -> sqlite3.Connection:
        conn = sqlite3.connect(self.db_path, timeout=30)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA synchronous=NORMAL")
        return conn

    def _init_db(self):
        with self._connect() as conn:
            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS tasks (
                    task_id TEXT PRIMARY KEY,
                    task_type TEXT NOT NULL,
                    project_name TEXT NOT NULL,
                    target_folder TEXT NOT NULL,
                    status TEXT NOT NULL,
                    progress INTEGER NOT NULL DEFAULT 0,
                    total INTEGER NOT NULL DEFAULT 0,
                    current_item TEXT NOT NULL DEFAULT '',
                    files_json TEXT NOT NULL DEFAULT '[]',
                    aggregate_report TEXT,
                    errors_json TEXT NOT NULL DEFAULT '[]',
                    payload_json TEXT NOT NULL DEFAULT '{}',
                    start_time TEXT,
                    end_time TEXT,
                    created_at TEXT NOT NULL
                )
            """
            )
            conn.execute(
                """
                CREATE INDEX IF NOT EXISTS idx_tasks_type_created
                ON tasks(task_type, created_at DESC)
            """
            )
            conn.commit()

    def _recover_interrupted_tasks(self):
        """Đánh dấu interrupted cho các task đang chạy hoặc chờ khi server vừa bật lại"""
        try:
            with self._connect() as conn:
                conn.execute(
                    """
                    UPDATE tasks
                    SET status = 'interrupted',
                        end_time = datetime('now'),
                        errors_json = '["Bị gián đoạn do máy chủ khởi động lại"]'
                    WHERE status IN ('running', 'cancelling', 'queued')
                """
                )
                conn.commit()
        except Exception as e:
            logger.error(f"[TaskManager] Lỗi khôi phục tác vụ gián đoạn: {e}")

    def _load_active_tasks(self):
        """Nạp các task gần nhất vào bộ nhớ cache RAM"""
        try:
            with self._connect() as conn:
                rows = conn.execute(
                    "SELECT * FROM tasks ORDER BY created_at DESC LIMIT 50"
                ).fetchall()
                for r in rows:
                    task = TaskRecord(
                        task_id=r["task_id"],
                        task_type=r["task_type"],
                        project_name=r["project_name"],
                        target_folder=r["target_folder"],
                        status=r["status"],
                        progress=r["progress"],
                        total=r["total"],
                        current_item=r["current_item"],
                        files=json.loads(r["files_json"] or "[]"),
                        aggregate_report=r["aggregate_report"],
                        errors=json.loads(r["errors_json"] or "[]"),
                        payload=json.loads(r["payload_json"] or "{}"),
                        start_time=r["start_time"],
                        end_time=r["end_time"],
                        created_at=r["created_at"],
                    )
                    self._tasks[task.task_id] = task
        except Exception as e:
            logger.error(f"[TaskManager] Lỗi nạp task từ database: {e}")

    def _save_task_to_db(self, task: TaskRecord):
        """Lưu hoặc cập nhật task xuống SQLite"""
        try:
            with self._connect() as conn:
                conn.execute(
                    """
                    INSERT INTO tasks (
                        task_id, task_type, project_name, target_folder, status,
                        progress, total, current_item, files_json, aggregate_report,
                        errors_json, payload_json, start_time, end_time, created_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ON CONFLICT(task_id) DO UPDATE SET
                        status=excluded.status,
                        progress=excluded.progress,
                        total=excluded.total,
                        current_item=excluded.current_item,
                        files_json=excluded.files_json,
                        aggregate_report=excluded.aggregate_report,
                        errors_json=excluded.errors_json,
                        payload_json=excluded.payload_json,
                        start_time=excluded.start_time,
                        end_time=excluded.end_time
                """,
                    (
                        task.task_id,
                        task.task_type,
                        task.project_name,
                        task.target_folder,
                        task.status,
                        task.progress,
                        task.total,
                        task.current_item,
                        json.dumps(task.files, ensure_ascii=False),
                        task.aggregate_report,
                        json.dumps(task.errors, ensure_ascii=False),
                        json.dumps(task.payload, ensure_ascii=False),
                        task.start_time,
                        task.end_time,
                        task.created_at,
                    ),
                )
                conn.commit()
        except Exception as e:
            logger.error(f"[TaskManager] Lỗi lưu task xuống database: {e}")

    def submit_task(
        self,
        task_type: str,
        target_folder: str,
        project_name: str,
        execute_fn: Callable[[TaskRecord, Callable[[], bool]], None],
        payload: Optional[Dict] = None,
        task_id: Optional[str] = None,
    ) -> TaskRecord:
        """Đăng ký tác vụ mới vào hàng đợi"""
        tid = task_id or str(uuid.uuid4())[:8]
        task = TaskRecord(
            task_id=tid,
            task_type=task_type,
            project_name=project_name,
            target_folder=target_folder,
            status="queued",
            payload=payload or {},
        )

        with self._lock:
            self._tasks[tid] = task
            self._callbacks[tid] = execute_fn
            self._save_task_to_db(task)
            self._task_queue.put(tid)

        return task

    def get_task(self, task_id: str) -> Optional[TaskRecord]:
        """Lấy thông tin task theo ID (từ RAM hoặc DB)"""
        with self._lock:
            if task_id in self._tasks:
                return self._tasks[task_id]

        # Truy vấn từ DB nếu chưa có trong RAM
        try:
            with self._connect() as conn:
                r = conn.execute("SELECT * FROM tasks WHERE task_id=?", (task_id,)).fetchone()
                if r:
                    task = TaskRecord(
                        task_id=r["task_id"],
                        task_type=r["task_type"],
                        project_name=r["project_name"],
                        target_folder=r["target_folder"],
                        status=r["status"],
                        progress=r["progress"],
                        total=r["total"],
                        current_item=r["current_item"],
                        files=json.loads(r["files_json"] or "[]"),
                        aggregate_report=r["aggregate_report"],
                        errors=json.loads(r["errors_json"] or "[]"),
                        payload=json.loads(r["payload_json"] or "{}"),
                        start_time=r["start_time"],
                        end_time=r["end_time"],
                        created_at=r["created_at"],
                    )
                    with self._lock:
                        self._tasks[task_id] = task
                    return task
        except Exception as e:
            logger.error(f"[TaskManager] Lỗi đọc task {task_id}: {e}")

        return None

    def cancel_task(self, task_id: str) -> bool:
        """Yêu cầu hủy task đang chạy hoặc xóa khỏi hàng đợi chờ"""
        with self._lock:
            task = self.get_task(task_id)
            if not task:
                return False

            if task.status == "queued":
                task.status = "cancelled"
                task.cancel_requested = True
                task.end_time = datetime.now().isoformat()
                self._save_task_to_db(task)
                return True

            if task.status == "running" and self._current_running_id == task_id:
                task.cancel_requested = True
                task.status = "cancelling"
                self._current_cancel_event.set()
                self._save_task_to_db(task)
                return True

        return False

    def update_progress(
        self,
        task_id: str,
        progress: int,
        total: int,
        current_item: str = "",
        file_path: Optional[str] = None,
    ):
        """Cập nhật tiến độ xử lý của task"""
        task = self.get_task(task_id)
        if not task:
            return

        with self._lock:
            task.progress = progress
            task.total = total
            if current_item:
                task.current_item = current_item
            if file_path:
                task.files.append(
                    {
                        "name": os.path.basename(file_path),
                        "path": file_path,
                        "person": current_item,
                    }
                )
            self._save_task_to_db(task)

    def get_recent_tasks(self, limit: int = 20, task_type: Optional[str] = None) -> List[Dict]:
        """Lấy danh sách các task gần nhất"""
        try:
            with self._connect() as conn:
                if task_type:
                    rows = conn.execute(
                        "SELECT * FROM tasks WHERE task_type=? ORDER BY created_at DESC LIMIT ?",
                        (task_type, limit),
                    ).fetchall()
                else:
                    rows = conn.execute(
                        "SELECT * FROM tasks ORDER BY created_at DESC LIMIT ?",
                        (limit,),
                    ).fetchall()

                return [
                    {
                        "task_id": r["task_id"],
                        "task_type": r["task_type"],
                        "project_name": r["project_name"],
                        "target_folder": r["target_folder"],
                        "status": r["status"],
                        "progress": r["progress"],
                        "total": r["total"],
                        "current": r["current_item"],
                        "file_count": len(json.loads(r["files_json"] or "[]")),
                        "aggregate_report": r["aggregate_report"],
                        "errors": json.loads(r["errors_json"] or "[]"),
                        "payload": json.loads(r["payload_json"] or "{}"),
                        "start_time": r["start_time"],
                        "end_time": r["end_time"],
                        "created_at": r["created_at"],
                    }
                    for r in rows
                ]
        except Exception as e:
            logger.error(f"[TaskManager] Lỗi lấy danh sách task gần nhất: {e}")
            return []

    def _worker_loop(self):
        """Vòng lặp luồng worker thực thi tuần tự các task từ hàng đợi"""
        while True:
            task_id = self._task_queue.get()
            task = self.get_task(task_id)
            execute_fn = self._callbacks.get(task_id)

            if not task or not execute_fn:
                self._task_queue.task_done()
                continue

            # Nếu task đã bị hủy khi đang chờ trong queue
            if task.status == "cancelled" or task.cancel_requested:
                self._task_queue.task_done()
                continue

            # Bắt đầu chạy task
            with self._lock:
                self._current_running_id = task_id
                self._current_cancel_event.clear()
                task.status = "running"
                task.start_time = datetime.now().isoformat()
                self._save_task_to_db(task)

            logger.info(f"[TaskManager] Bắt đầu thực thi task {task_id} ({task.task_type} - {task.target_folder})")

            try:
                # Gọi hàm thực thi và truyền hàm kiểm tra hủy: self._current_cancel_event.is_set
                execute_fn(task, self._current_cancel_event.is_set)

                # Kiểm tra trạng thái kết thúc
                with self._lock:
                    if self._current_cancel_event.is_set() or task.cancel_requested:
                        task.status = "cancelled"
                    elif task.status not in ("failed", "cancelled"):
                        task.status = "completed"
                    task.end_time = datetime.now().isoformat()
                    self._save_task_to_db(task)

                logger.info(f"[TaskManager] Hoàn thành task {task_id} với trạng thái: {task.status}")

            except Exception as e:
                import traceback
                traceback.print_exc()
                with self._lock:
                    task.status = "failed"
                    task.errors.append(str(e))
                    task.end_time = datetime.now().isoformat()
                    self._save_task_to_db(task)
                logger.error(f"[TaskManager] Lỗi khi thực thi task {task_id}: {e}")

            finally:
                with self._lock:
                    self._current_running_id = None
                    self._current_cancel_event.clear()
                    self._callbacks.pop(task_id, None)
                self._task_queue.task_done()
