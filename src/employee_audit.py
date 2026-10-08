"""Read-only identity comparisons; a roster match does not verify a face."""
from collections import Counter
import json
import unicodedata
from datetime import date
from src.employee_importer import normalize_vietnamese_name


def name_key(value):
    return ' '.join(unicodedata.normalize('NFC', str(value)).casefold().split())


def current_assignment(employee):
    today = date.today().isoformat()
    return bool(employee['active'] and (employee['source_paths'] or any(
        m['valid_from'] <= today and (not m['valid_to'] or m['valid_to'] > today)
        for m in employee['memberships'])))


def load_roster(registry, project_id):
    with registry._connect() as c:
        c.execute('CREATE TABLE IF NOT EXISTS audit_rosters (project_id TEXT PRIMARY KEY, filename TEXT NOT NULL, rows_json TEXT NOT NULL)')
        row = c.execute('SELECT filename, rows_json FROM audit_rosters WHERE project_id=?', (project_id,)).fetchone()
    return {'filename': row['filename'], 'rows': json.loads(row['rows_json'])} if row else {'filename': '', 'rows': []}


def save_roster(registry, project_id, filename, rows):
    load_roster(registry, project_id)
    with registry._connect() as c:
        c.execute('INSERT OR REPLACE INTO audit_rosters VALUES (?,?,?)', (project_id, filename, json.dumps(rows, ensure_ascii=False)))


def audit_project(registry, project_id):
    roster = load_roster(registry, project_id)
    projects = registry.list_projects()
    people = {p['project_id']: registry.list_employees(p['project_id']) for p in projects}
    local = [e for e in people.get(project_id, []) if current_assignment(e)]
    counts = Counter(name_key(e['display_name']) for e in local)
    similar_counts = Counter(normalize_vietnamese_name(e['display_name']) for e in local)
    result = []
    for e in local:
        codes = {m['payroll_code'] for m in e['memberships'] if m['payroll_code']}
        rows = [r for r in roster['rows'] if r['payroll_code'] in codes]
        issues = []
        if counts[name_key(e['display_name'])] > 1:
            issues.append('Trùng tên trong dự án; cần phân biệt bằng mã và ảnh')
        elif similar_counts[normalize_vietnamese_name(e['display_name'])] > 1:
            issues.append('Tên tương tự trong dự án (khác dấu/tiền tố); có thể là hồ sơ trùng')
        if not codes:
            issues.append('Chưa có mã chấm công')
        if roster['filename']:
            if not rows:
                issues.append('Không tìm thấy mã trong bảng đã chọn')
            elif len(rows) != 1:
                issues.append('Mã xuất hiện với nhiều tên trong bảng')
            elif name_key(rows[0]['name']) != name_key(e['display_name']):
                issues.append('Tên hồ sơ khác tên trong bảng')
        else:
            issues.append('Chưa chọn bảng chấm công để đối chiếu')
        others = []
        for p in projects:
            if p['project_id'] == project_id:
                continue
            for other in people[p['project_id']]:
                if not other['active']:
                    continue
                same_id = other['employee_id'] == e['employee_id']
                similar_name = normalize_vietnamese_name(other['display_name']) == normalize_vietnamese_name(e['display_name'])
                if same_id or similar_name:
                    others.append({'project': p['display_name'], 'employee_id': other['employee_id'],
                                   'name': other['display_name'], 'same_identity': same_id,
                                   'current': current_assignment(other),
                                   'memberships': other['memberships']})
        result.append({'employee_id': e['employee_id'], 'name': e['display_name'],
                       'codes': sorted(codes), 'roster_rows': rows, 'issues': issues,
                       'other_projects': others})
    return {**roster, 'employees': result,
            'missing_profiles': [r for r in roster['rows'] if not any(r['payroll_code'] and r['payroll_code'] in e['codes'] for e in result)]}
