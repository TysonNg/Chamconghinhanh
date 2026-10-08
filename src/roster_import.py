"""Preview and apply a project roster without merging ambiguous identities."""
from collections import Counter
import uuid
import hashlib
import json

from src.employee_audit import load_roster, name_key
from src.employee_importer import normalize_vietnamese_name


def matching_name(value):
    return normalize_vietnamese_name(value).replace(' ', '')


def import_roster(registry, project_id, apply=False, expected_token=None):
    project = registry.get_project(project_id)
    if not project['active']:
        raise ValueError('Dự án đã được lưu trữ')
    roster = load_roster(registry, project_id)
    if not roster['rows']:
        raise ValueError('Chọn bảng chấm công để đối chiếu trước')
    token = hashlib.sha256(json.dumps(roster, ensure_ascii=False, sort_keys=True).encode()).hexdigest()
    if apply and expected_token and token != expected_token:
        raise ValueError('Bảng đối chiếu đã thay đổi; bấm Xem trước lại')
    registry.import_legacy(project_id)
    rows = list({(r['payroll_code'], name_key(r['name'])): r for r in roster['rows']}.values())
    code_counts = Counter(r['payroll_code'] for r in rows)
    name_counts = Counter(matching_name(r['name']) for r in rows)
    others = []
    for p in registry.list_projects():
        if p['project_id'] == project_id or not p['active']:
            continue
        for e in registry.list_employees(p['project_id']):
            if e['active'] and (e['source_paths'] or any(not m['valid_to'] for m in e['memberships'])):
                others.append({**e, 'source_project': p['display_name'], 'source_project_id': p['project_id']})
    people = registry.list_employees(project_id)
    results = []
    with registry._connect() as c:
        c.execute('BEGIN IMMEDIATE')
        # Read again under the write lock to make repeated/concurrent imports idempotent.
        people = registry.list_employees(project_id)
        for row in rows:
            name, code = row['name'], row['payroll_code']
            key = matching_name(name)
            candidates = [e for e in people if any(m['payroll_code'] == code for m in e['memberships'])] if code else []
            similar = [e for e in people if matching_name(e['display_name']) == key]
            result = dict(row)
            target = None
            if not code or not key:
                action, reason = 'conflict', 'Thiếu tên hoặc mã chấm công'
            elif code_counts[code] > 1:
                action, reason = 'conflict', 'Một mã có nhiều tên trong bảng'
            elif len(candidates) > 1:
                action, reason = 'conflict', 'Mã đã gắn với nhiều hồ sơ'
            elif candidates:
                target = candidates[0]
                if matching_name(target['display_name']) != key:
                    action, reason = 'conflict', 'Mã khớp nhưng tên khác; cần kiểm tra'
                elif not target['active']:
                    action, reason = 'restore', 'Khôi phục hồ sơ đã lưu trữ'
                else:
                    action, reason = 'existing', 'Đã có đúng tên và mã'
            elif similar:
                if len(similar) != 1 or name_counts[key] > 1 or similar[0]['memberships']:
                    action, reason = 'conflict', 'Tên có hồ sơ khác hoặc mã khác; cần kiểm tra'
                else:
                    target = similar[0]
                    action, reason = 'assign', 'Gắn mã vào hồ sơ cũ cùng tên chưa có mã'
            else:
                matches = [e for e in others if matching_name(e['display_name']) == key]
                if matches:
                    action, reason = 'transfer', 'Có hồ sơ tại dự án khác; xác nhận đúng người và ngày chuyển'
                    result['transfer_candidates'] = [{'employee_id': e['employee_id'], 'name': e['display_name'],
                        'source_project_id': e['source_project_id'], 'source_project': e['source_project'],
                        'source_codes': sorted({m['payroll_code'] for m in e['memberships'] if m['payroll_code']})} for e in matches]
                else:
                    action, reason = 'create', 'Thêm hồ sơ mới'
            result.update(action=action, reason=reason)
            if apply and action not in ('conflict', 'existing', 'transfer'):
                eid = target['employee_id'] if target else uuid.uuid4().hex
                if action == 'create':
                    c.execute('INSERT INTO employees(employee_id, display_name) VALUES (?,?)', (eid, name))
                c.execute('UPDATE employees SET active=1 WHERE employee_id=?', (eid,))
                c.execute('DELETE FROM project_employee_archives WHERE project_id=? AND employee_id=?', (project_id, eid))
                registry._ensure_internal_code(c, eid)
                if action in ('create', 'assign'):
                    registry._assign(c, project_id, eid, code, '2000-01-01', None)
                # Keep legacy photos pending: a name match does not verify a face.
                result['employee_id'] = eid
            results.append(result)
    counts = Counter(r['action'] for r in results)
    return {'project_id': project_id, 'filename': roster['filename'], 'rows': results, 'roster_token': token,
            'counts': {k: counts[k] for k in ('create', 'assign', 'restore', 'existing', 'conflict', 'transfer')},
            'total': len(results), 'applied': apply}
