"""Persistent identities; payroll codes and portrait bindings are project scoped."""
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import date, datetime, timezone
from pathlib import Path
import hashlib
import secrets
import re
import shutil
import sqlite3
import uuid

@dataclass(frozen=True)
class IdentityResolution:
    status: str
    project_id: str
    employee_id: str | None = None
    reason: str = ""

def _component(value):
    if not isinstance(value, str) or not value.strip() or value.strip() in (".", "..") or re.search(r'[<>:"/\\|?*\x00-\x1f]', value):
        raise ValueError("Tên thư mục không hợp lệ")
    return value.strip()

def _day(value):
    if isinstance(value, date):
        return value.isoformat()
    if not isinstance(value, str) or not re.fullmatch(r"\d{4}-\d{2}-\d{2}", value):
        raise ValueError("Ngày hiệu lực phải là YYYY-MM-DD")
    return date.fromisoformat(value).isoformat()

class IdentityRegistry:
    def __init__(self, db_path, portrait_root):
        self.db_path = Path(db_path)
        self.portrait_root = Path(portrait_root).resolve()
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        with self._connect() as c:
            c.executescript("""
                CREATE TABLE IF NOT EXISTS projects (
                    project_id TEXT PRIMARY KEY, display_name TEXT NOT NULL,
                    storage_dir TEXT NOT NULL UNIQUE COLLATE NOCASE,
                    active INTEGER NOT NULL DEFAULT 1);
                CREATE TABLE IF NOT EXISTS employees (
                    employee_id TEXT PRIMARY KEY, display_name TEXT NOT NULL,
                    active INTEGER NOT NULL DEFAULT 1);
                CREATE TABLE IF NOT EXISTS employee_internal_codes (
                    employee_id TEXT PRIMARY KEY REFERENCES employees(employee_id),
                    code TEXT NOT NULL UNIQUE COLLATE NOCASE,
                    created_at TEXT NOT NULL);
                CREATE TABLE IF NOT EXISTS memberships (
                    membership_id TEXT PRIMARY KEY,
                    project_id TEXT NOT NULL REFERENCES projects(project_id),
                    employee_id TEXT NOT NULL REFERENCES employees(employee_id),
                    payroll_code TEXT NOT NULL,
                    valid_from TEXT NOT NULL, valid_to TEXT);
                CREATE INDEX IF NOT EXISTS memberships_lookup
                    ON memberships(project_id, payroll_code, valid_from, valid_to);
                CREATE TABLE IF NOT EXISTS portrait_bindings (
                    project_id TEXT NOT NULL REFERENCES projects(project_id),
                    employee_id TEXT NOT NULL REFERENCES employees(employee_id),
                    relative_path TEXT NOT NULL, confirmed_at TEXT NOT NULL, confirmed_by TEXT NOT NULL,
                    PRIMARY KEY(project_id, employee_id, relative_path));
                CREATE TABLE IF NOT EXISTS portrait_exclusions (
                    project_id TEXT NOT NULL REFERENCES projects(project_id),
                    employee_id TEXT NOT NULL REFERENCES employees(employee_id),
                    relative_path TEXT NOT NULL, valid_from TEXT NOT NULL,
                    PRIMARY KEY(project_id,employee_id,relative_path));
                CREATE TABLE IF NOT EXISTS legacy_sources (
                    project_id TEXT NOT NULL REFERENCES projects(project_id),
                    employee_id TEXT NOT NULL REFERENCES employees(employee_id),
                    relative_path TEXT NOT NULL,
                    PRIMARY KEY(project_id, relative_path));
            """)

    @contextmanager
    def _connect(self):
        c = sqlite3.connect(self.db_path, timeout=30)
        c.row_factory = sqlite3.Row
        c.execute("PRAGMA foreign_keys=ON")
        try:
            with c:
                yield c
        finally:
            c.close()

    def register_project(self, display_name, storage_dir=None):
        name = _component(display_name)
        folder = _component(storage_dir or name)
        with self._connect() as c:
            c.execute("BEGIN IMMEDIATE")
            found = c.execute("SELECT * FROM projects WHERE storage_dir=? COLLATE NOCASE", (folder,)).fetchone()
            if found:
                return dict(found)
            pid = uuid.uuid4().hex
            c.execute("INSERT INTO projects(project_id,display_name,storage_dir) VALUES(?,?,?)", (pid,name,folder))
        return self.get_project(pid)

    def get_project(self, reference):
        with self._connect() as c:
            rows = c.execute("SELECT * FROM projects WHERE project_id=? OR storage_dir=? COLLATE NOCASE OR display_name=?",
                             (reference,reference,reference)).fetchall()
        if len(rows) != 1:
            raise ValueError("Dự án không tồn tại hoặc tên dự án không duy nhất")
        return dict(rows[0])

    def list_projects(self, include_archived=False):
        with self._connect() as c:
            return [dict(r) for r in c.execute("SELECT * FROM projects" + ("" if include_archived else " WHERE active=1") + " ORDER BY display_name")]

    def rename_project(self, project_id, name):
        name = _component(name)
        with self._connect() as c:
            c.execute("BEGIN IMMEDIATE")
            if c.execute("SELECT 1 FROM projects WHERE display_name=? AND project_id<>?",(name,project_id)).fetchone():
                raise ValueError("Tên dự án đã tồn tại")
            if not c.execute("UPDATE projects SET display_name=? WHERE project_id=?", (name,project_id)).rowcount:
                raise ValueError("Không tìm thấy dự án")

    def archive_project(self, project_id):
        with self._connect() as c:
            c.execute("UPDATE projects SET active=0 WHERE project_id=?", (project_id,))

    def project_portrait_dir(self, project_id):
        folder = self.portrait_root / self.get_project(project_id)["storage_dir"]
        if not folder.resolve().is_relative_to(self.portrait_root):
            raise ValueError("Đường dẫn dự án không hợp lệ")
        return folder

    def create_employee(self, name):
        if not isinstance(name,str) or not name.strip():
            raise ValueError("Thiếu tên nhân viên")
        eid = uuid.uuid4().hex
        with self._connect() as c:
            c.execute("INSERT INTO employees(employee_id,display_name) VALUES(?,?)",(eid,name.strip()))
        return self.get_employee(eid)

    def generate_internal_codes(self):
        """Assign display codes once, without modifying payroll or identity approval."""
        with self._connect() as c:
            c.execute("BEGIN IMMEDIATE")
            pending=c.execute("""SELECT e.employee_id FROM employees e WHERE active=1
                AND NOT EXISTS (SELECT 1 FROM employee_internal_codes n WHERE n.employee_id=e.employee_id)
                ORDER BY e.employee_id""").fetchall()
            used={row[0].upper() for row in c.execute("SELECT code FROM employee_internal_codes")}
            used.update(row[0].upper() for row in c.execute("SELECT payroll_code FROM memberships"))
            for row in pending:
                for _ in range(100):
                    code="NV-"+secrets.token_hex(4).upper()
                    if code not in used:
                        break
                else:
                    raise ValueError("Không tạo được mã duy nhất; vui lòng thử lại")
                c.execute("INSERT INTO employee_internal_codes VALUES(?,?,?)",
                          (row["employee_id"],code,datetime.now(timezone.utc).isoformat()))
                used.add(code)
            total=c.execute("""SELECT COUNT(*) FROM employee_internal_codes n
                JOIN employees e ON e.employee_id=n.employee_id WHERE e.active=1""").fetchone()[0]
        return {"created_count":len(pending),"total_count":total,"unchanged_count":total-len(pending)}

    def get_employee(self, employee_id):
        with self._connect() as c:
            row=c.execute("SELECT e.*, (SELECT code FROM employee_internal_codes WHERE employee_id=e.employee_id) AS internal_code FROM employees e WHERE employee_id=?",(employee_id,)).fetchone()
        if row is None:
            raise ValueError("Không tìm thấy nhân viên")
        return dict(row)

    def _assign(self,c,project_id,employee_id,code,start,end):
        if not isinstance(code,str) or not code.strip():
            raise ValueError("Mã chấm công phải là chuỗi không rỗng")
        code=code.strip()
        if end and end <= start:
            raise ValueError("Ngày kết thúc phải sau ngày bắt đầu")
        if not c.execute("SELECT 1 FROM projects WHERE project_id=? AND active=1",(project_id,)).fetchone():
            raise ValueError("Dự án không tồn tại hoặc đã lưu trữ")
        if not c.execute("SELECT 1 FROM employees WHERE employee_id=? AND active=1",(employee_id,)).fetchone():
            raise ValueError("Nhân viên không tồn tại hoặc đã lưu trữ")
        existing = c.execute("""SELECT * FROM memberships WHERE project_id=?
            AND (payroll_code=? OR employee_id=?) AND valid_from < ?
            AND (valid_to IS NULL OR valid_to > ?)""",
            (project_id,code,employee_id,end or "9999-12-31",start)).fetchall()
        if existing:
            if len(existing)==1 and all(existing[0][k]==v for k,v in {
                "employee_id":employee_id,"payroll_code":code,"valid_from":start,"valid_to":end}.items()):
                return
            raise ValueError("Mã chấm công hoặc nhân viên bị trùng khoảng hiệu lực trong dự án")
        c.execute("INSERT INTO memberships VALUES(?,?,?,?,?,?)",
                  (uuid.uuid4().hex,project_id,employee_id,code,start,end))

    def assign_employee(self,project_id,employee_id,payroll_code,valid_from,valid_to=None):
        start,end=_day(valid_from),_day(valid_to) if valid_to else None
        with self._connect() as c:
            c.execute("BEGIN IMMEDIATE")
            self._assign(c,project_id,employee_id,payroll_code,start,end)

    def resolve_employee(self,project_id,payroll_code,on_date):
        if not isinstance(payroll_code,str) or not payroll_code.strip():
            return IdentityResolution("needs_confirmation",project_id,reason="Chưa có mã chấm công đã xác nhận")
        day=_day(on_date)
        with self._connect() as c:
            rows=c.execute("""SELECT DISTINCT employee_id FROM memberships WHERE project_id=? AND payroll_code=?
                AND valid_from<=? AND (valid_to IS NULL OR valid_to>?)""",
                (project_id,payroll_code.strip(),day,day)).fetchall()
        if len(rows)==1:
            return IdentityResolution("resolved",project_id,rows[0]["employee_id"])
        return IdentityResolution("ambiguous" if rows else "missing",project_id,
                                  reason="Cần xác nhận mã chấm công trong đúng dự án và ngày hiệu lực")

    def is_member(self,project_id,employee_id,on_date):
        day=_day(on_date)
        with self._connect() as c:
            return bool(c.execute("""SELECT 1 FROM memberships WHERE project_id=? AND employee_id=?
                AND valid_from<=? AND (valid_to IS NULL OR valid_to>?)""",(project_id,employee_id,day,day)).fetchone())

    def _bound_path(self,project_id,relative_path):
        root=self.project_portrait_dir(project_id)
        rel=Path(relative_path)
        if rel.is_absolute() or ".." in rel.parts or not rel.parts:
            raise ValueError("Chân dung phải nằm trong dự án")
        path=root/rel
        if not path.resolve().is_relative_to(root.resolve()) or path.is_symlink():
            raise ValueError("Chân dung phải nằm trong dự án")
        return path

    def bind_portrait(self,project_id,employee_id,relative_path,confirmed_by):
        path=self._bound_path(project_id,relative_path)
        if not path.exists() or not str(confirmed_by).strip():
            raise ValueError("Thiếu ảnh hoặc người xác nhận")
        self.get_employee(employee_id)
        with self._connect() as c:
            if not c.execute("SELECT 1 FROM memberships WHERE project_id=? AND employee_id=?",(project_id,employee_id)).fetchone():
                raise ValueError("Chưa xác nhận mã nhân viên trong dự án")
            self._bind(c,project_id,employee_id,relative_path,confirmed_by)

    @staticmethod
    def _bind(c,project_id,employee_id,relative_path,reviewer):
        c.execute("INSERT OR REPLACE INTO portrait_bindings VALUES(?,?,?,?,?)",
                  (project_id,employee_id,str(relative_path),datetime.now(timezone.utc).isoformat(),reviewer.strip()))

    def portrait_paths(self,project_id,employee_id,on_date):
        if not self.is_member(project_id,employee_id,on_date):
            return []
        with self._connect() as c:
            rows=c.execute("SELECT relative_path FROM portrait_bindings WHERE project_id=? AND employee_id=?",
                           (project_id,employee_id)).fetchall()
        result=[]
        for row in rows:
            try:
                bound=self._bound_path(project_id,row["relative_path"])
                paths=sorted(bound.rglob("*")) if bound.is_dir() else [bound]
                root=self.project_portrait_dir(project_id).resolve()
                result.extend(p for p in paths if p.is_file() and not p.is_symlink()
                              and p.resolve().is_relative_to(root) and p.suffix.lower() in {".jpg",".jpeg",".png",".bmp",".webp"})
            except ValueError:
                continue
        with self._connect() as c:
            excluded={row[0] for row in c.execute(
                "SELECT relative_path FROM portrait_exclusions WHERE project_id=? AND employee_id=? AND valid_from<=?",
                (project_id,employee_id,_day(on_date)))}
        root=self.project_portrait_dir(project_id)
        return sorted({p for p in result if p.relative_to(root).as_posix() not in excluded})

    def import_legacy(self,project_id):
        root=self.project_portrait_dir(project_id)
        if not root.exists():
            return []
        with self._connect() as c:
            c.execute("BEGIN IMMEDIATE")
            for path in sorted(root.iterdir()):
                if path.is_symlink() or not path.resolve().is_relative_to(root.resolve()):
                    continue
                if not path.is_dir() and path.suffix.lower() not in {".jpg",".jpeg",".png",".bmp",".webp"}:
                    continue
                rel=path.name
                if c.execute("SELECT 1 FROM legacy_sources WHERE project_id=? AND relative_path=?",(project_id,rel)).fetchone():
                    continue
                bindings=c.execute("SELECT relative_path FROM portrait_bindings WHERE project_id=?",(project_id,)).fetchall()
                if any(Path(binding[0]).parts[0] == rel for binding in bindings):
                    continue
                eid=uuid.uuid4().hex
                c.execute("INSERT INTO employees(employee_id,display_name) VALUES(?,?)",(eid,path.name if path.is_dir() else path.stem))
                c.execute("INSERT INTO legacy_sources VALUES(?,?,?)",(project_id,eid,rel))
        return self.list_employees(project_id)

    def confirm_source(self,project_id,employee_id,payroll_code,valid_from,reviewer):
        if not isinstance(reviewer,str) or not reviewer.strip():
            raise ValueError("Thiếu người xác nhận")
        with self._connect() as c:
            c.execute("BEGIN IMMEDIATE")
            sources=c.execute("SELECT relative_path FROM legacy_sources WHERE project_id=? AND employee_id=?",
                              (project_id,employee_id)).fetchall()
            if not sources:
                raise ValueError("Không có nguồn chân dung cần xác nhận")
            for row in sources:
                if not self._bound_path(project_id,row["relative_path"]).exists():
                    raise ValueError("Không tìm thấy ảnh nguồn")
            self._assign(c,project_id,employee_id,payroll_code,_day(valid_from),None)
            for row in sources:
                self._bind(c,project_id,employee_id,row["relative_path"],reviewer)

    def list_employees(self,project_id):
        with self._connect() as c:
            rows=c.execute("""SELECT e.*, (SELECT code FROM employee_internal_codes WHERE employee_id=e.employee_id) AS internal_code FROM employees e WHERE employee_id IN
                (SELECT employee_id FROM memberships WHERE project_id=? UNION SELECT employee_id FROM legacy_sources WHERE project_id=?)
                ORDER BY e.display_name,e.employee_id""",(project_id,project_id)).fetchall()
            items=[]
            for row in rows:
                item=dict(row)
                memberships=[dict(r) for r in c.execute("SELECT * FROM memberships WHERE project_id=? AND employee_id=? ORDER BY valid_from DESC",
                                                      (project_id,item["employee_id"]))]
                sources=[r[0] for r in c.execute("SELECT relative_path FROM legacy_sources WHERE project_id=? AND employee_id=?",
                                               (project_id,item["employee_id"]))]
                item.update(project_id=project_id,memberships=memberships,source_paths=sources,
                            status="confirmed" if memberships else "needs_confirmation")
                items.append(item)
        return items

    def transfer_employee(self,source_id,target_id,employee_id,effective_date,payroll_code,reviewer):
        if source_id==target_id or not isinstance(reviewer,str) or not reviewer.strip():
            raise ValueError("Chọn dự án đích khác nguồn và người xác nhận")
        effective=_day(effective_date)
        # Exclusive end date preserves the source assignment before the transfer.
        from datetime import timedelta
        previous=date.fromisoformat(effective)-timedelta(days=1)
        portraits=self.portrait_paths(source_id,employee_id,previous)
        target=self.project_portrait_dir(target_id)/employee_id/("transfer_"+uuid.uuid4().hex)
        with self._connect() as c:
            c.execute("BEGIN IMMEDIATE")
            current=c.execute("""SELECT * FROM memberships WHERE project_id=? AND employee_id=? AND valid_from<?
                AND (valid_to IS NULL OR valid_to>?)""",(source_id,employee_id,effective,effective)).fetchall()
            if len(current)!=1:
                raise ValueError("Không xác định được lịch sử dự án nguồn tại ngày chuyển")
            self._assign(c,target_id,employee_id,payroll_code,effective,None)
            # Source remains untouched. Unique target directory avoids name collisions.
            if portraits:
                target.mkdir(parents=True,exist_ok=False)
                for index,source in enumerate(portraits):
                    dest=target/(f"{index:03}_"+source.name)
                    shutil.copy2(source,dest)
                    if hashlib.sha256(dest.read_bytes()).digest()!=hashlib.sha256(source.read_bytes()).digest():
                        raise OSError("Bản sao chân dung không khớp nguồn")
                rel=target.relative_to(self.project_portrait_dir(target_id)).as_posix()
                self._bind(c,target_id,employee_id,rel,reviewer)
            c.execute("UPDATE memberships SET valid_to=? WHERE membership_id=?",(effective,current[0]["membership_id"]))

    def archive_employee(self,employee_id):
        with self._connect() as c:
            c.execute("UPDATE employees SET active=0 WHERE employee_id=?",(employee_id,))
