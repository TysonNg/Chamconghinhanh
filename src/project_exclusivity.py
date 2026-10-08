"""Explicitly retire confirmed duplicate assignments without deleting history."""
from datetime import date
from src.employee_audit import load_roster, current_assignment
from src.roster_import import matching_name


def exclusive_assignment(registry, project_id, employee_id, apply=False, confirmed=None):
    target = next((e for e in registry.list_employees(project_id)
                   if e['employee_id'] == employee_id and current_assignment(e)), None)
    if not target:
        raise ValueError('Hồ sơ không hoạt động trong dự án')
    codes = {m['payroll_code'] for m in target['memberships'] if m['payroll_code']}
    rows = [r for r in load_roster(registry, project_id)['rows'] if r['payroll_code'] in codes
            and matching_name(r['name']) == matching_name(target['display_name'])]
    if len(rows) != 1:
        raise ValueError('Chọn hồ sơ có tên và mã khớp một dòng trong bảng')
    affected = []
    for p in registry.list_projects():
        for e in registry.list_employees(p['project_id']):
            if p['project_id'] == project_id and e['employee_id'] == employee_id:
                continue
            if current_assignment(e) and (e['employee_id'] == employee_id or
                    matching_name(e['display_name']) == matching_name(target['display_name'])):
                affected.append({'project_id': p['project_id'], 'employee_id': e['employee_id'],
                                 'project': p['display_name'], 'name': e['display_name']})
    keys = sorted(f"{e['project_id']}:{e['employee_id']}" for e in affected)
    if apply:
        if sorted(confirmed or []) != keys:
            raise ValueError('Danh sách hồ sơ đã thay đổi; xem trước lại')
        today = date.today().isoformat()
        with registry._connect() as c:
            c.execute('BEGIN IMMEDIATE')
            for e in affected:
                c.execute('INSERT OR IGNORE INTO project_employee_archives VALUES (?,?)',
                          (e['project_id'], e['employee_id']))
                c.execute('UPDATE memberships SET valid_to=? WHERE project_id=? AND employee_id=? '
                          'AND valid_from<=? AND (valid_to IS NULL OR valid_to>?)',
                          (today, e['project_id'], e['employee_id'], today, today))
    return {'affected': affected, 'confirmation_keys': keys, 'name': rows[0]['name'],
            'payroll_code': rows[0]['payroll_code'], 'applied': apply}
