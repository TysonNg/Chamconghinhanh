"""Use a selected attendance roster as the active project directory."""
from collections import Counter
from datetime import date
import hashlib
import json
import shutil
import uuid

from src.employee_audit import load_roster, current_assignment
from src.roster_import import matching_name
from src.identity_registry import _day


def sync_roster(r, project_id, apply=False, expected_token=None, effective_date=None, selections=None, skipped=None):
    supplied_selections = selections
    selections = selections or {}
    if not isinstance(selections, dict):
        raise ValueError('Lựa chọn hồ sơ không hợp lệ')
    p = r.get_project(project_id)
    if not p['active']:
        raise ValueError('Dự án đã lưu trữ')
    roster = load_roster(r, project_id)
    if not roster['rows']:
        raise ValueError('Chưa có bảng chấm công')
    token = hashlib.sha256(json.dumps(roster, sort_keys=True, ensure_ascii=False).encode()).hexdigest()
    with r._connect() as c:
        c.execute('CREATE TABLE IF NOT EXISTS roster_file_sync_receipts '
                  '(project_id TEXT, token TEXT, effective_date TEXT, selections_json TEXT, skipped_json TEXT, '
                  'PRIMARY KEY(project_id,token))')
        if c.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name='roster_sync_receipts'").fetchone():
            c.execute('INSERT OR IGNORE INTO roster_file_sync_receipts SELECT * FROM roster_sync_receipts')
        receipt = c.execute('SELECT * FROM roster_file_sync_receipts WHERE project_id=? AND token=?',(project_id,token)).fetchone()
    if receipt:
        effective_date = effective_date or receipt['effective_date']
        if supplied_selections is None:
            selections = json.loads(receipt['selections_json'])
        if skipped is None:
            skipped = json.loads(receipt['skipped_json'])
    skipped = skipped or []
    if not isinstance(skipped,list) or not all(isinstance(key,str) for key in skipped):
        raise ValueError('Danh sách bỏ qua không hợp lệ')
    if apply and token != expected_token:
        raise ValueError('Bảng đã thay đổi; xem trước lại')
    effective = _day(effective_date or date.today().isoformat())
    if effective > date.today().isoformat():
        raise ValueError('Ngày hiệu lực không được nằm trong tương lai')
    projects = {p['project_id']: p for p in r.list_projects() if p['active']}
    all_people = []
    for pid in projects:
        r.import_legacy(pid)
        all_people.extend(r.list_employees(pid))
    local = [e for e in all_people if e['project_id'] == project_id]
    rows = list({(row['payroll_code'], matching_name(row['name'])): row for row in roster['rows']}.values())
    codes = Counter(row['payroll_code'] for row in rows)
    keys = {f"{row['payroll_code']}|{row['name']}" for row in rows}
    if set(skipped)-keys:
        raise ValueError('Bảng đã thay đổi; xem trước lại danh sách bỏ qua')
    plan = []
    protected = set()
    for row in rows:
        code, name = row['payroll_code'], row['name']
        key = matching_name(name)
        same = [e for e in all_people if matching_name(e['display_name']) == key]
        coded = [e for e in local if any(m['payroll_code'] == code for m in e['memberships'])]
        eligible = coded or same
        chosen = selections.get(code)
        if chosen:
            if not any(e['employee_id'] == chosen for e in eligible):
                raise ValueError('Hồ sơ được chọn đã thay đổi; xem trước lại')
            eligible = [e for e in eligible if e['employee_id'] == chosen]
        target = None
        row_key = f'{code}|{name}'
        if row_key in skipped:
            action, reason = 'skip', 'Không đồng bộ; giữ nguyên hồ sơ'
            protected.update(e['employee_id'] for e in coded + same)
        elif not code:
            action, reason = 'conflict', 'Thiếu mã chấm công trong bảng'
        elif not key:
            action, reason = 'conflict', 'Không đọc được tên nhân viên trong bảng'
        elif codes[code] > 1:
            action, reason = 'conflict', 'Một mã có nhiều tên trong bảng; kiểm tra lại bảng nguồn'
        else:
            ids = {e['employee_id'] for e in eligible}
            if len(ids) > 1:
                action, reason = 'conflict', 'Có nhiều hồ sơ cùng tên; chọn đúng người trước'
            elif ids:
                eid = next(iter(ids))
                target = next(e for e in eligible if e['employee_id'] == eid)
                sources = [e for e in all_people if e['employee_id'] == eid and e['project_id'] != project_id and current_assignment(e)]
                if sources:
                    action, reason = 'transfer', 'Chuyển về dự án này, dùng mã trong bảng'
                elif target['project_id'] != project_id:
                    action, reason = 'restore', 'Khôi phục vào dự án này'
                elif any(m['payroll_code'] == code and m['valid_from'] <= effective and (not m['valid_to'] or m['valid_to'] > effective) for m in target['memberships']) and target['active']:
                    action, reason = 'existing', 'Đã có'
                else:
                    action, reason = 'assign', 'Cập nhật mã theo bảng'
            else:
                action, reason = 'create', 'Bổ sung'
        item = dict(row, action=action, reason=reason, row_key=row_key)
        if action == 'conflict' and code and codes[code] == 1:
            options = {}
            for e in eligible:
                options.setdefault(e['employee_id'], {'employee_id':e['employee_id'], 'name':e['display_name'], 'projects':[]})['projects'].append(projects[e['project_id']]['display_name'])
            item['candidates'] = list(options.values())
        if target:
            item['employee_id'] = target['employee_id']
            if target['display_name'] != name:
                item['previous_name'] = target['display_name']
                item['rename'] = True
            item['source_projects'] = [projects[e['project_id']]['display_name'] for e in all_people
                if e['employee_id'] == target['employee_id'] and e['project_id'] != project_id and current_assignment(e)]
        plan.append(item)
    keep = {row.get('employee_id') for row in plan if row['action'] != 'conflict'}
    outside = [e for e in local if current_assignment(e) and e['employee_id'] not in keep and e['employee_id'] not in protected]
    copied = []
    if apply:
        if any(row['action'] == 'conflict' for row in plan):
            raise ValueError('Chưa thể đồng bộ toàn bộ: cần xử lý các dòng trùng hồ sơ hoặc mã')
        try:
            with r._connect() as c:
                c.execute('BEGIN IMMEDIATE')
                c.execute('CREATE TABLE IF NOT EXISTS employee_name_history '
                          '(employee_id TEXT, project_id TEXT, previous_name TEXT, new_name TEXT, '
                          'source_file TEXT, changed_at TEXT)')
                for row in plan:
                    if row['action']=='skip':
                        continue
                    eid = row.get('employee_id') or uuid.uuid4().hex
                    if row['action'] == 'create':
                        c.execute('INSERT INTO employees(employee_id,display_name) VALUES (?,?)', (eid,row['name']))
                    elif row.get('rename'):
                        previous = c.execute('SELECT display_name FROM employees WHERE employee_id=?',(eid,)).fetchone()[0]
                        c.execute('INSERT INTO employee_name_history VALUES (?,?,?,?,?,datetime(\'now\'))',
                                  (eid,project_id,previous,row['name'],roster['filename']))
                        c.execute('UPDATE employees SET display_name=? WHERE employee_id=?',(row['name'],eid))
                    r._ensure_internal_code(c,eid)
                    c.execute('UPDATE employees SET active=1 WHERE employee_id=?',(eid,))
                    c.execute('DELETE FROM project_employee_archives WHERE project_id=? AND employee_id=?',(project_id,eid))
                    # Close other-project assignments; preserve codes, files and history.
                    for source in [e for e in all_people if e['employee_id']==eid and e['project_id']!=project_id and current_assignment(e)]:
                        pid = source['project_id']
                        c.execute('INSERT OR IGNORE INTO project_employee_archives VALUES (?,?)',(pid,eid))
                        c.execute('UPDATE memberships SET valid_to=CASE WHEN valid_from>? THEN valid_from ELSE ? END WHERE project_id=? AND employee_id=? AND (valid_to IS NULL OR valid_to>?)',(effective,effective,pid,eid,effective))
                        paths = c.execute('SELECT relative_path FROM portrait_bindings WHERE project_id=? AND employee_id=? UNION SELECT relative_path FROM legacy_sources WHERE project_id=? AND employee_id=?',(pid,eid,pid,eid)).fetchall()
                        excluded = {x[0] for x in c.execute('SELECT relative_path FROM portrait_exclusions WHERE project_id=? AND employee_id=?',(pid,eid))}
                        root = r.project_portrait_dir(pid)
                        for path in paths:
                            bound = r._bound_path(pid,path[0])
                            for image in (bound.rglob('*') if bound.is_dir() else [bound]):
                                if not image.is_file() or image.is_symlink() or not image.resolve().is_relative_to(root.resolve()) or image.suffix.lower() not in {'.jpg','.jpeg','.png','.webp','.bmp'} or image.relative_to(root).as_posix() in excluded:
                                    continue
                                folder = r.project_portrait_dir(project_id)/eid/'synced'
                                folder.mkdir(parents=True,exist_ok=True)
                                dest = folder/(uuid.uuid4().hex+image.suffix)
                                shutil.copy2(image,dest); copied.append(dest)
                                r._bind(c,project_id,eid,dest.relative_to(r.project_portrait_dir(project_id)).as_posix(),'roster-sync')
                    active = c.execute('SELECT * FROM memberships WHERE project_id=? AND employee_id=? AND valid_from<=? AND (valid_to IS NULL OR valid_to>?)',(project_id,eid,effective,effective)).fetchall()
                    if not any(m['payroll_code']==row['payroll_code'] for m in active):
                        c.execute('UPDATE memberships SET valid_to=CASE WHEN valid_from>? THEN valid_from ELSE ? END WHERE project_id=? AND employee_id=? AND (valid_to IS NULL OR valid_to>?)',(effective,effective,project_id,eid,effective))
                        r._assign(c,project_id,eid,row['payroll_code'],effective,None)
                    row['employee_id']=eid
                for e in outside:
                    c.execute('INSERT OR IGNORE INTO project_employee_archives VALUES (?,?)',(project_id,e['employee_id']))
                c.execute('INSERT OR REPLACE INTO roster_file_sync_receipts VALUES (?,?,?,?,?)',
                          (project_id,token,effective,json.dumps(selections),json.dumps(skipped)))
        except Exception:
            for path in copied:
                path.unlink(missing_ok=True)
            raise
    counts = Counter(row['action'] for row in plan)
    return dict(project_id=project_id, filename=roster['filename'], rows=plan, total=len(plan),
                counts={k:counts[k] for k in ('create','assign','existing','restore','transfer','conflict','skip')},
                outside_roster=[{'name':e['display_name'],'employee_id':e['employee_id']} for e in outside],
                renamed=sum(bool(row.get('rename')) for row in plan),
                roster_token=token, effective_date=effective, applied=apply,
                sync_approved=bool(receipt) or apply)
