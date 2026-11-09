from datetime import date
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor

from docx import Document
from flask import Flask
import pytest

from src.aggregate_report_exporter import collect_issue_rows, export_aggregate_report
from src.identity_registry import IdentityRegistry
from src.identity_routes import register_identity_routes
from src.report_order import sort_report_persons


def setup_project(registry, name='Site'):
    project = registry.register_project(name)
    people = []
    for index, name in enumerate(['Anh', 'Bình', 'Chi']):
        employee = registry.create_employee(name)
        registry.assign_employee(project['project_id'], employee['employee_id'], str(index + 1), '2026-01-01')
        people.append(employee)
    return project, people


def test_order_persists_across_restart_rename_and_new_employees(tmp_path):
    registry = IdentityRegistry(tmp_path / 'ids.db', tmp_path / 'portraits')
    project, people = setup_project(registry)
    order = [p['employee_id'] for p in people[::-1]]
    registry.save_report_order(project['project_id'], order)
    registry.rename_project(project['project_id'], 'New name')
    registry = IdentityRegistry(registry.db_path, registry.portrait_root)
    new = registry.create_employee('A new employee')
    registry.assign_employee(project['project_id'], new['employee_id'], 'new', '2026-01-01')
    assert [p['employee_id'] for p in registry.report_order(project['project_id'])] == order + [new['employee_id']]
    assert registry.get_project(project['project_id'])['display_name'] == 'New name'


def test_report_order_api_is_project_scoped_and_invalid_save_is_atomic(tmp_path):
    registry = IdentityRegistry(tmp_path / 'ids.db', tmp_path / 'portraits')
    project, people = setup_project(registry)
    other, others = setup_project(registry, 'Other')
    app = Flask(__name__)
    register_identity_routes(app, lambda: registry, lambda: tmp_path / 'camera')
    client = app.test_client()
    order = [p['employee_id'] for p in people[::-1]]
    assert client.post('/api/projects/report-order', json={'project_id': project['project_id'], 'employee_ids': order}).status_code == 200
    for bad in [None, 'bad', [1], [order[0], order[0]], [others[0]['employee_id']]]:
        assert client.post('/api/projects/report-order', json={'project_id': project['project_id'], 'employee_ids': bad}).status_code == 400
        result = client.get('/api/projects/report-order', query_string={'project_id': project['project_id']})
        assert [e['employee_id'] for e in result.json['employees']] == order
    assert [e['employee_id'] for e in registry.report_order(other['project_id'])] == [e['employee_id'] for e in others]
    registry.archive_project(project['project_id'])
    assert client.post('/api/projects/report-order', json={'project_id': project['project_id'], 'employee_ids': order}).status_code == 400


def test_custom_order_survives_scrambled_completion_and_keeps_dates_sorted(tmp_path):
    persons = [
        {'employee_id': eid, 'name': name, 'records': [
            {'date': day, 'is_absent': True} for day in ['03/09/2026', '01/09/2026']]}
        for eid, name in [('a', 'Anh'), ('new', 'New'), ('c', 'Chi'), ('b', 'Bình')]
    ]
    order = ['c', 'b', 'a']
    rows = collect_issue_rows(persons, employee_order=order)
    assert [r['name'] for r in rows] == ['Chi', 'Chi', 'Bình', 'Bình', 'Anh', 'Anh', 'New', 'New']
    assert [r['date_value'] for r in rows[:2]] == [date(2026, 9, 1), date(2026, 9, 3)]
    path = export_aggregate_report(persons, str(tmp_path), 'Site', employee_order=order)
    doc = Document(path)
    names = [row.cells[0].text for row in doc.tables[0].rows[1:]]
    assert names == ['Chi', '', 'Bình', '', 'Anh', '', 'New', '']
    # Same-name employees are ranked by identity, and an unresolved name cannot take a saved slot.
    same_names = [{'employee_id': 'a', 'name': 'Same'}, {'employee_id': None, 'name': 'Same'}, {'employee_id': 'b', 'name': 'Same'}]
    assert [p['employee_id'] for p in sort_report_persons(same_names, ['b', 'a'])] == ['b', 'a', None]


@pytest.mark.parametrize('source', ['excel', 'pdf'])
@pytest.mark.parametrize('paused', [False, True, 'all'])
def test_both_analyzers_submit_in_saved_order_and_export_that_order(tmp_path, monkeypatch, source, paused):
    import src.excel_face_analyzer as excel
    import src.pdf_face_analyzer as pdf
    registry = IdentityRegistry(tmp_path / 'ids.db', tmp_path / 'portraits')
    project, people = setup_project(registry)
    order = [p['employee_id'] for p in people[::-1]]
    registry.save_report_order(project['project_id'], order)
    if paused:
        for person in people if paused == 'all' else [people[1]]:
            registry.set_scan_paused(project['project_id'], person['employee_id'], True)
    module = excel if source == 'excel' else pdf
    parser_class = module.ExcelPersonFileParser if source == 'excel' else module.PDFPersonFileParser
    analyzer_class = module.ExcelFaceAnalyzer if source == 'excel' else module.PDFFaceAnalyzer
    input_dir = tmp_path / 'input'
    input_dir.mkdir()
    for index in range(3):
        (input_dir / f'{index}{".xlsx" if source == "excel" else ".docx"}').touch()
    def parse_person(self):
        index = int(Path(self.path).stem)
        return {'id': str(index + 1), 'name': people[index]['display_name'],
                'records': [{'date': '01/09/2026', 'is_absent': True}]}
    monkeypatch.setattr(parser_class, '__init__', lambda self, path: setattr(self, 'path', path))
    monkeypatch.setattr(parser_class, 'parse_person', parse_person)
    submitted = []
    class RecordingExecutor(ThreadPoolExecutor):
        def submit(self, fn, person):
            submitted.append(person['employee_id'])
            return super().submit(fn, person)
    monkeypatch.setattr(module, 'ThreadPoolExecutor', RecordingExecutor)
    def export_person(self, person, *args, **kwargs):
        return str(tmp_path / f"{person['employee_id']}.docx")
    monkeypatch.setattr(module.ExcelToWordExporter, 'export_person', export_person)
    analyzer = analyzer_class(str(tmp_path / 'portraits'), str(tmp_path / 'camera'), matcher=None,
                              project_id=project['project_id'], identity_registry=registry)
    files = analyzer.analyze_folder(str(input_dir), str(tmp_path / 'out'),
                                   report_options={'project_name': 'Site', 'employee_order': order,
                                                   'paused_employee_ids': registry.paused_employee_ids(project['project_id'])})
    if paused == 'all':
        assert submitted == [] and files == []
        return
    assert submitted == ([order[0], order[2]] if paused else order)
    report = next(path for path in files if 'GIAI_TRINH_' in path)
    assert [r.cells[0].text for r in Document(report).tables[0].rows[1:]] == (['Chi', 'Anh'] if paused else ['Chi', 'Bình', 'Anh'])


def test_pause_api_persists_scoped_to_project_and_resuming_keeps_order(tmp_path):
    registry = IdentityRegistry(tmp_path / 'ids.db', tmp_path / 'portraits')
    project, people = setup_project(registry)
    other, others = setup_project(registry, 'Other')
    # One employee may be known in several projects: leave in A must not pause B.
    registry.assign_employee(other['project_id'], people[1]['employee_id'], 'shared', '2026-01-01')
    order = [e['employee_id'] for e in people[::-1]]
    registry.save_report_order(project['project_id'], order)
    app = Flask(__name__)
    register_identity_routes(app, lambda: registry, lambda: tmp_path / 'camera')
    client = app.test_client()
    payload = {'project_id': project['project_id'], 'employee_id': people[1]['employee_id'], 'paused': True}
    assert client.post('/api/portraits/employee/scan-pause', json=payload).json['scan_paused'] is True
    registry = IdentityRegistry(registry.db_path, registry.portrait_root)
    assert registry.paused_employee_ids(project['project_id']) == [people[1]['employee_id']]
    assert registry.paused_employee_ids(other['project_id']) == []
    listing = client.get('/api/portraits', query_string={'project_id': project['project_id']}).json['employees']
    assert next(e for e in listing if e['employee_id'] == people[1]['employee_id'])['scan_paused'] is True
    for bad in [None, 'false', 1]:
        assert client.post('/api/portraits/employee/scan-pause', json={**payload, 'paused': bad}).status_code == 400
    assert client.post('/api/portraits/employee/scan-pause', json={**payload, 'employee_id': others[0]['employee_id']}).status_code == 400
    assert client.post('/api/portraits/employee/scan-pause', json={**payload, 'paused': False}).json['scan_paused'] is False
    assert registry.paused_employee_ids(project['project_id']) == []
    assert [e['employee_id'] for e in registry.report_order(project['project_id'])] == order


def test_exporter_excludes_paused_even_if_supplied_in_input(tmp_path):
    people = [{'employee_id': eid, 'name': name, 'records': [{'date': '01/09/2026', 'is_absent': True}]}
              for eid, name in [('a', 'Anh'), ('b', 'Bình')]]
    path = export_aggregate_report(people, str(tmp_path), 'Site', employee_order=['b', 'a'], paused_employee_ids=['b'])
    assert [r.cells[0].text for r in Document(path).tables[0].rows[1:]] == ['Anh']
