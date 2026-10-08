import openpyxl
from src.identity_registry import IdentityRegistry
from src.employee_audit import save_roster
from src.employee_importer import extract_employees_from_excel
from src.roster_import import import_roster


def test_all_sheets_repeated_days_and_conflicting_names(tmp_path):
    w = openpyxl.Workbook()
    for sheet, rows in [(w.active, [('001', 'Nguyen Van A'), ('001', 'Nguyen Van A')]),
                        (w.create_sheet('Other'), [('002', 'Tran Van B'), ('001', 'Different Person')])]:
        sheet.append(['Mã nhân viên', 'Tên nhân viên'])
        for row in rows:
            sheet.append(row)
        sheet.append([None, None])
    path = tmp_path / 'attendance.xlsx'
    w.save(path)
    rows = extract_employees_from_excel(str(path))
    assert len(rows) == 3
    assert {r['payroll_code'] for r in rows} == {'001', '002'}


def test_preview_apply_idempotence_and_conflicts(tmp_path):
    r = IdentityRegistry(tmp_path / 'ids.db', tmp_path / 'portraits')
    a, b = r.register_project('A'), r.register_project('B')
    save_roster(r, a['project_id'], 'month.xlsx', [
        {'name': 'Nguyen Van A', 'payroll_code': '001'},
        {'name': 'Tran Van B', 'payroll_code': '002'},
        {'name': 'Different Person', 'payroll_code': '002'},
        {'name': 'No Code', 'payroll_code': ''}])
    preview = import_roster(r, a['project_id'])
    assert preview['counts']['create'] == 1
    assert preview['counts']['conflict'] == 3
    assert r.list_employees(a['project_id']) == []
    result = import_roster(r, a['project_id'], True)
    assert result['counts'] == preview['counts']
    assert len(r.list_employees(a['project_id'])) == 1
    assert r.list_employees(b['project_id']) == []
    assert import_roster(r, a['project_id'], True)['counts']['existing'] == 1
    r.archive_project_employees(a['project_id'])
    assert import_roster(r, a['project_id'], True)['counts']['restore'] == 1
    assert r.list_employees(a['project_id'])[0]['active'] == 1


def test_legacy_name_gets_code_without_verifying_photo(tmp_path):
    r = IdentityRegistry(tmp_path / 'ids.db', tmp_path / 'portraits')
    a = r.register_project('A')
    folder = r.project_portrait_dir(a['project_id']) / 'Nguyen Van A'
    folder.mkdir(parents=True)
    (folder / 'portrait.jpg').write_bytes(b'keep original photo')
    r.import_legacy(a['project_id'])
    save_roster(r, a['project_id'], 'month.xlsx', [{'name': 'Nguyen Van A', 'payroll_code': '001'}])
    assert import_roster(r, a['project_id'], True)['counts']['assign'] == 1
    assert len(r.list_employees(a['project_id'])) == 1
    with r._connect() as c:
        assert c.execute('SELECT count(*) FROM portrait_bindings').fetchone()[0] == 0
    assert (folder / 'portrait.jpg').read_bytes() == b'keep original photo'


def test_endpoint_requires_project_and_only_applies_explicitly(tmp_path):
    from flask import Flask
    from src.identity_routes import register_identity_routes
    r = IdentityRegistry(tmp_path / 'ids.db', tmp_path / 'portraits')
    p = r.register_project('A')
    save_roster(r, p['project_id'], 'month.xlsx', [{'name': 'Nguyen Van A', 'payroll_code': '001'}])
    app = Flask(__name__)
    register_identity_routes(app, lambda: r, lambda: tmp_path / 'camera')
    c = app.test_client()
    assert c.post('/api/portraits/import-roster', json={}).status_code == 400
    assert c.post('/api/portraits/import-roster', json={'project_id': p['project_id']}).json['counts']['create'] == 1
    assert r.list_employees(p['project_id']) == []
    assert c.post('/api/portraits/import-roster', json={'project_id': p['project_id'], 'apply': True}).json['applied']
    assert len(r.list_employees(p['project_id'])) == 1


def test_cross_project_candidate_requires_transfer_confirmation(tmp_path):
    from flask import Flask
    from src.identity_routes import register_identity_routes
    r = IdentityRegistry(tmp_path / 'ids.db', tmp_path / 'portraits')
    source, target = r.register_project('Source'), r.register_project('Target')
    eid = r.create_employee('DangVanMinh')['employee_id']
    with r._connect() as db:
        r._assign(db, source['project_id'], eid, 'OLD', '2020-01-01', None)
    save_roster(r, target['project_id'], 'month.xlsx', [{'name': 'Dang Van Minh', 'payroll_code': 'NEW'}])
    preview = import_roster(r, target['project_id'], True)
    assert preview['counts']['transfer'] == 1
    assert not r.list_employees(target['project_id'])
    app = Flask(__name__)
    register_identity_routes(app, lambda: r, lambda: tmp_path / 'camera')
    c = app.test_client()
    data = dict(project_id=target['project_id'], source_project_id=source['project_id'],
                employee_id=eid, payroll_code='NEW', effective_date='2026-09-01')
    assert c.post('/api/portraits/roster-transfer', json={**data, 'payroll_code': 'WRONG'}).status_code == 400
    assert c.post('/api/portraits/roster-transfer', json=data).status_code == 200
    assert r.list_employees(target['project_id'])[0]['employee_id'] == eid
    assert r.list_employees(target['project_id'])[0]['memberships'][0]['payroll_code'] == 'NEW'
    old = r.list_employees(source['project_id'])[0]['memberships'][0]
    assert old['payroll_code'] == 'OLD' and old['valid_to'] == '2026-09-01'
    assert import_roster(r, target['project_id'])['counts']['existing'] == 1


def test_changed_roster_requires_new_preview(tmp_path):
    import pytest
    r = IdentityRegistry(tmp_path/'ids.db', tmp_path/'portraits')
    p = r.register_project('A')
    save_roster(r, p['project_id'], 'old.xlsx', [{'name':'Person One','payroll_code':'001'}])
    token = import_roster(r, p['project_id'])['roster_token']
    save_roster(r, p['project_id'], 'new.xlsx', [{'name':'Person Two','payroll_code':'002'}])
    with pytest.raises(ValueError):
        import_roster(r, p['project_id'], True, token)
    assert r.list_employees(p['project_id']) == []
