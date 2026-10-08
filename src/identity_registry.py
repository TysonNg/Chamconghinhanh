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
import unicodedata
import uuid


def _casefold_unicode(value):
    """Unicode-aware case-insensitive normalization.

    SQLite COLLATE NOCASE only handles ASCII a-z/A-Z.
    Vietnamese characters like Á/á are treated as different,
    which causes duplicate projects on case-insensitive filesystems (Windows).
    """
    return unicodedata.normalize('NFC', value).casefold()

def _is_uuid_like(value):
    """Kiểm tra chuỗi có phải dạng UUID/hex kỹ thuật (không phải họ tên người)."""
    if not isinstance(value, str):
        return False
    v = value.strip()
    return bool(re.fullmatch(r"^[0-9a-fA-F]{32}$", v) or re.fullmatch(r"^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$", v))


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
                CREATE TABLE IF NOT EXISTS project_employee_archives (
                    project_id TEXT NOT NULL REFERENCES projects(project_id),
                    employee_id TEXT NOT NULL REFERENCES employees(employee_id),
                    PRIMARY KEY(project_id, employee_id));
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
        folder_key = _casefold_unicode(folder)
        with self._connect() as c:
            c.execute("BEGIN IMMEDIATE")
            # Use Python-side Unicode casefold instead of SQLite COLLATE NOCASE
            # which only handles ASCII and fails for Vietnamese characters.
            all_projects = c.execute("SELECT * FROM projects ORDER BY active DESC").fetchall()
            for row in all_projects:
                if _casefold_unicode(row["storage_dir"]) == folder_key or _casefold_unicode(row["display_name"]) == folder_key:
                    return dict(row)
            pid = uuid.uuid4().hex
            c.execute("INSERT INTO projects(project_id,display_name,storage_dir) VALUES(?,?,?)", (pid,name,folder))
        return self.get_project(pid)

    def get_project(self, reference):
        if not reference:
            raise ValueError("Dự án không tồn tại hoặc tên dự án không duy nhất")
        ref_str = str(reference).strip()
        ref_key = _casefold_unicode(ref_str)
        with self._connect() as c:
            row = c.execute("SELECT * FROM projects WHERE project_id=?", (ref_str,)).fetchone()
            if row:
                return dict(row)
            all_p = c.execute("SELECT * FROM projects").fetchall()
            matches = [dict(r) for r in all_p if _casefold_unicode(r["storage_dir"]) == ref_key or _casefold_unicode(r["display_name"]) == ref_key]
            active_matches = [m for m in matches if m["active"]]
            if len(active_matches) == 1:
                return active_matches[0]
            if len(matches) == 1:
                return matches[0]
            if not matches:
                raise ValueError("Dự án không tồn tại hoặc tên dự án không duy nhất")
            raise ValueError("Dự án không tồn tại hoặc tên dự án không duy nhất")

    def list_projects(self, include_archived=False):
        with self._connect() as c:
            rows = [dict(r) for r in c.execute("SELECT * FROM projects" + ("" if include_archived else " WHERE active=1") + " ORDER BY display_name")]
            seen = {}
            for r in rows:
                key = _casefold_unicode(r["storage_dir"])
                if key not in seen:
                    seen[key] = r
                else:
                    if r["active"] and not seen[key]["active"]:
                        seen[key] = r
            return list(seen.values())

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

    def create_employee(self, name, auto_code=False):
        if not isinstance(name,str) or not name.strip():
            raise ValueError("Thiếu tên nhân viên")
        eid = uuid.uuid4().hex
        with self._connect() as c:
            c.execute("INSERT INTO employees(employee_id,display_name) VALUES(?,?)",(eid,name.strip()))
            if auto_code:
                self._ensure_internal_code(c, eid)
        return self.get_employee(eid)

    def _ensure_internal_code(self, c, employee_id, used_codes=None):
        """Đảm bảo nhân viên có mã nội bộ (NV-XXXXXXXX). Trả về (code, created: bool)."""
        row = c.execute("SELECT code FROM employee_internal_codes WHERE employee_id=?", (employee_id,)).fetchone()
        if row:
            return row[0], False

        if used_codes is None:
            used_codes = {r[0].upper() for r in c.execute("SELECT code FROM employee_internal_codes")}
            used_codes.update(r[0].upper() for r in c.execute("SELECT payroll_code FROM memberships"))

        for _ in range(100):
            code = "NV-" + secrets.token_hex(4).upper()
            if code not in used_codes:
                break
        else:
            raise ValueError("Không tạo được mã duy nhất; vui lòng thử lại")

        now_str = datetime.now(timezone.utc).isoformat()
        c.execute("INSERT INTO employee_internal_codes VALUES(?,?,?)", (employee_id, code, now_str))
        used_codes.add(code)
        return code, True

    def generate_internal_codes(self):
        """Assign display codes once, without modifying payroll or identity approval."""
        with self._connect() as c:
            c.execute("BEGIN IMMEDIATE")
            pending=c.execute("""SELECT e.employee_id FROM employees e WHERE active=1
                AND NOT EXISTS (SELECT 1 FROM employee_internal_codes n WHERE n.employee_id=e.employee_id)
                ORDER BY e.employee_id""").fetchall()
            used={row[0].upper() for row in c.execute("SELECT code FROM employee_internal_codes")}
            used.update(row[0].upper() for row in c.execute("SELECT payroll_code FROM memberships"))
            created = 0
            for row in pending:
                _, is_new = self._ensure_internal_code(c, row["employee_id"], used)
                if is_new:
                    created += 1
            total=c.execute("""SELECT COUNT(*) FROM employee_internal_codes n
                JOIN employees e ON e.employee_id=n.employee_id WHERE e.active=1""").fetchone()[0]
        return {"created_count":created,"total_count":total,"unchanged_count":total-created}

    def get_employee(self, employee_id):
        with self._connect() as c:
            row=c.execute("SELECT e.*, (SELECT code FROM employee_internal_codes WHERE employee_id=e.employee_id) AS internal_code FROM employees e WHERE employee_id=?",(employee_id,)).fetchone()
        if row is None:
            raise ValueError("Không tìm thấy nhân viên")
        return dict(row)

    def _assign(self,c,project_id,employee_id,code,start,end):
        if not isinstance(code,str):
            raise ValueError("Mã chấm công phải là chuỗi")
        code=code.strip()
        if end and end <= start:
            raise ValueError("Ngày kết thúc phải sau ngày bắt đầu")
        if not c.execute("SELECT 1 FROM projects WHERE project_id=? AND active=1",(project_id,)).fetchone():
            raise ValueError("Dự án không tồn tại hoặc đã lưu trữ")
        if not c.execute("SELECT 1 FROM employees WHERE employee_id=? AND active=1",(employee_id,)).fetchone():
            raise ValueError("Nhân viên không tồn tại hoặc đã lưu trữ")
        existing = c.execute("""SELECT * FROM memberships WHERE project_id=?
            AND ((? <> '' AND payroll_code=?) OR employee_id=?) AND valid_from < ?
            AND (valid_to IS NULL OR valid_to > ?)""",
            (project_id,code,code,employee_id,end or "9999-12-31",start)).fetchall()
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
        valid_exts = {".jpg", ".jpeg", ".png", ".bmp", ".webp"}
        with self._connect() as c:
            c.execute("BEGIN IMMEDIATE")
            for path in sorted(root.iterdir()):
                if path.is_symlink() or not path.resolve().is_relative_to(root.resolve()):
                    continue
                if not path.is_dir() and path.suffix.lower() not in valid_exts:
                    continue

                rel=path.name
                stem_or_name = path.name if path.is_dir() else path.stem

                # Tuyệt đối không dùng mã hex/UUID kỹ thuật làm tên nhân viên hiển thị
                if _is_uuid_like(stem_or_name):
                    emp_row = c.execute("SELECT employee_id FROM employees WHERE employee_id=?", (stem_or_name,)).fetchone()
                    if emp_row and path.is_dir():
                        target_eid = emp_row["employee_id"]
                        bound = c.execute("SELECT 1 FROM portrait_bindings WHERE project_id=? AND employee_id=?", (project_id, target_eid)).fetchone()
                        if not bound and c.execute("SELECT 1 FROM memberships WHERE project_id=? AND employee_id=?", (project_id, target_eid)).fetchone():
                            self._bind(c, project_id, target_eid, rel, "system")
                    continue

                if c.execute("SELECT 1 FROM legacy_sources WHERE project_id=? AND relative_path=?",(project_id,rel)).fetchone():
                    continue
                bindings=c.execute("SELECT relative_path FROM portrait_bindings WHERE project_id=?",(project_id,)).fetchall()
                if any(Path(binding[0]).parts[0] == rel for binding in bindings):
                    continue
                existing_emp = c.execute("""
                    SELECT e.employee_id FROM employees e
                    WHERE e.display_name=? AND e.employee_id IN (
                        SELECT employee_id FROM memberships WHERE project_id=?
                        UNION SELECT employee_id FROM portrait_bindings WHERE project_id=?
                        UNION SELECT employee_id FROM legacy_sources WHERE project_id=?
                    )
                """, (stem_or_name, project_id, project_id, project_id)).fetchone()
                if existing_emp:
                    eid = existing_emp["employee_id"]
                else:
                    eid=uuid.uuid4().hex
                    c.execute("INSERT INTO employees(employee_id,display_name) VALUES(?,?)",(eid,stem_or_name))
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
                if c.execute("SELECT 1 FROM project_employee_archives WHERE project_id=? AND employee_id=?",
                             (project_id, item["employee_id"])).fetchone():
                    item["active"] = 0
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
        if not portraits:
            # Fallback: tìm ảnh từ portrait_bindings hoặc legacy_sources của nhân viên trong dự án nguồn
            with self._connect() as c:
                rows = c.execute("""SELECT relative_path FROM portrait_bindings WHERE project_id=? AND employee_id=?
                    UNION SELECT relative_path FROM legacy_sources WHERE project_id=? AND employee_id=?""",
                    (source_id, employee_id, source_id, employee_id)).fetchall()
                excluded = {row[0] for row in c.execute(
                    "SELECT relative_path FROM portrait_exclusions WHERE project_id=? AND employee_id=? AND valid_from<=?",
                    (source_id, employee_id, effective))}
            root = self.project_portrait_dir(source_id)
            result = []
            for r in rows:
                try:
                    bound = self._bound_path(source_id, r["relative_path"])
                    candidates = sorted(bound.rglob("*")) if bound.is_dir() else [bound]
                    for p in candidates:
                        if (p.is_file() and not p.is_symlink()
                                and p.resolve().is_relative_to(root.resolve())
                                and p.suffix.lower() in {".jpg", ".jpeg", ".png", ".bmp", ".webp"}
                                and p.relative_to(root).as_posix() not in excluded):
                            result.append(p)
                except ValueError:
                    continue
            portraits = sorted(set(result))

        target=self.project_portrait_dir(target_id)/employee_id/("transfer_"+uuid.uuid4().hex)
        with self._connect() as c:
            c.execute("BEGIN IMMEDIATE")
            current=c.execute("""SELECT * FROM memberships WHERE project_id=? AND employee_id=?
                AND (valid_to IS NULL OR valid_to>=?) ORDER BY valid_from DESC""",
                (source_id,employee_id,effective)).fetchall()
            
            if current:
                cur = current[0]
                if cur["valid_from"] >= effective:
                    c.execute("UPDATE memberships SET valid_from=?, valid_to=? WHERE membership_id=?",
                              (previous.isoformat(), effective, cur["membership_id"]))
                else:
                    c.execute("UPDATE memberships SET valid_to=? WHERE membership_id=?",
                              (effective, cur["membership_id"]))
            else:
                legacy_rows = c.execute("SELECT relative_path FROM legacy_sources WHERE project_id=? AND employee_id=?",
                                        (source_id, employee_id)).fetchall()
                bound_rows = c.execute("SELECT relative_path FROM portrait_bindings WHERE project_id=? AND employee_id=?",
                                       (source_id, employee_id)).fetchall()
                any_hist = c.execute("SELECT * FROM memberships WHERE project_id=? AND employee_id=? ORDER BY valid_from DESC",
                                     (source_id, employee_id)).fetchone()

                if not legacy_rows and not bound_rows and not any_hist:
                    raise ValueError("Không xác định được lịch sử dự án nguồn tại ngày chuyển")

                if any_hist:
                    if any_hist["valid_from"] >= effective:
                        c.execute("UPDATE memberships SET valid_from=?, valid_to=? WHERE membership_id=?",
                                  (previous.isoformat(), effective, any_hist["membership_id"]))
                    else:
                        c.execute("UPDATE memberships SET valid_to=? WHERE membership_id=?",
                                  (effective, any_hist["membership_id"]))
                else:
                    start_date = "2000-01-01"
                    if start_date >= effective:
                        start_date = previous.isoformat()
                    c.execute("INSERT INTO memberships VALUES(?,?,?,?,?,?)",
                              (uuid.uuid4().hex, source_id, employee_id, payroll_code or "", start_date, effective))

                for r in legacy_rows:
                    self._bind(c, source_id, employee_id, r["relative_path"], reviewer)

            self._assign(c,target_id,employee_id,payroll_code,effective,None)
            # Confirmed legacy employees also retain legacy_sources. Remove that
            # pending-source link so ended memberships hide them from active lists;
            # portrait bindings and the original files preserve historical access.
            c.execute("DELETE FROM legacy_sources WHERE project_id=? AND employee_id=?",
                      (source_id, employee_id))
            # Source remains untouched. Unique target directory avoids name collisions.
            if portraits:
                target.mkdir(parents=True,exist_ok=True)
                for index,source in enumerate(portraits):
                    dest=target/(f"{index:03}_"+source.name)
                    shutil.copy2(source,dest)
                    if hashlib.sha256(dest.read_bytes()).digest()!=hashlib.sha256(source.read_bytes()).digest():
                        raise OSError("Bản sao chân dung không khớp nguồn")
                rel=target.relative_to(self.project_portrait_dir(target_id)).as_posix()
                self._bind(c,target_id,employee_id,rel,reviewer)

    def archive_employee(self,employee_id):
        with self._connect() as c:
            c.execute("UPDATE employees SET active=0 WHERE employee_id=?",(employee_id,))

    def archive_project_employees(self, project_id):
        self.get_project(project_id)
        self.import_legacy(project_id)
        with self._connect() as c:
            c.execute("BEGIN IMMEDIATE")
            cursor = c.execute("""INSERT OR IGNORE INTO project_employee_archives(project_id, employee_id)
                SELECT ?, employee_id FROM employees WHERE active=1 AND employee_id IN
                (SELECT employee_id FROM memberships WHERE project_id=?
                 UNION SELECT employee_id FROM legacy_sources WHERE project_id=?)""",
                (project_id, project_id, project_id))
            return cursor.rowcount
