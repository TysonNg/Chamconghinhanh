import threading
from types import SimpleNamespace

import pytest

from src.task_manager import TaskManager, TaskRecord


@pytest.fixture
def scan_client(tmp_path, monkeypatch):
    import src.app as app_module
    from src.face_inference import FaceInferenceRunner, GIB
    manager = TaskManager(str(tmp_path))
    monkeypatch.setattr(app_module, 'task_manager', manager)
    runner = FaceInferenceRunner(base_dir=tmp_path, hardware_provider=lambda: {
        'cpu_count': 4, 'total_ram': 7.9 * GIB, 'available_ram': 6 * GIB})
    monkeypatch.setattr(app_module, 'face_inference_runner', runner)
    return app_module.app.test_client(), manager


def test_ai_config_is_validated_and_separate_from_scan_queue(scan_client):
    client, _ = scan_client
    initial = client.get('/api/tasks/ai-config').get_json()['inference']
    assert initial['max_workers'] == initial['hardware_max'] == 2
    response = client.post('/api/tasks/ai-config', json={'mode': 1})
    assert response.status_code == 200
    assert response.get_json()['inference']['max_workers'] == 1
    for mode in (0, 3, True, '2', None):
        assert client.post('/api/tasks/ai-config', json={'mode': mode}).status_code == 400
    assert client.post('/api/tasks/ai-config', json=[]).status_code == 400
    assert client.post('/api/tasks/ai-config', json={'mode': 'auto'}).get_json()['inference']['max_workers'] == 2
    snapshot = client.get('/api/tasks/scans').get_json()
    assert snapshot['max_running'] == 1
    assert snapshot['inference']['max_workers'] == 2


def test_snapshot_and_cancel_api(scan_client):
    client, manager = scan_client
    release = threading.Event()
    started = threading.Event()

    def run(task, cancel_check):
        started.set()
        release.wait(5)

    first = manager.submit_task('excel_face', 'month', 'A', run)
    assert started.wait(2)
    try:
        queued = manager.submit_task('pdf_face', 'month2', 'B', lambda *_: None)
        data = client.get('/api/tasks/scans?limit=1').get_json()
        assert data['counts'] == {'running': 1, 'queued': 1, 'cancelling': 0}
        assert len(data['active']) == 2
        status = client.get(f'/api/pdf/face/status/{queued.task_id}').get_json()
        assert status['ahead_count'] == 1
        assert client.post(f'/api/task/retry/{first.task_id}').status_code == 409
        assert client.post(f'/api/task/cancel/{queued.task_id}').get_json()['status'] == 'cancelled'
        assert client.post(f'/api/task/cancel/{queued.task_id}').status_code == 409
        assert client.post('/api/task/cancel/missing').status_code == 404
        data = client.get('/api/tasks/scans').get_json()
        assert data['counts']['queued'] == 0
        assert any(t['task_id'] == queued.task_id for t in data['history'])
    finally:
        release.set()
        manager._task_queue.join()


def test_monitor_is_available_in_page(scan_client):
    client, _ = scan_client
    response = client.get('/')
    assert response.status_code == 200
    assert b'id="scan-task-summary"' in response.data
    assert b'/static/js/scan-tasks.js' in response.data


@pytest.mark.parametrize('task_type', ['excel_face', 'pdf_face'])
def test_retry_preserves_parameters_and_new_scan_can_be_cancelled(scan_client, monkeypatch, task_type):
    import src.app as app_module
    client, manager = scan_client
    monkeypatch.setattr(app_module, 'get_identity_registry', lambda: SimpleNamespace(
        get_project=lambda reference: {'display_name': 'A', 'project_id': 'project-a'},
        report_order=lambda project_id: [{'employee_id': 'employee-b'}, {'employee_id': 'employee-a'}],
        paused_employee_ids=lambda project_id: ['employee-b']))
    payload = {'folder': 'month', 'project_name': 'A', 'project_id': 'project-a',
               'threshold': 0.4, 'input_dir': 'input', 'output_dir': 'output',
               'p_dir': 'portraits', 'i_dir': 'daily', 'report_options': {'project_name': 'Stale project'}}
    old = TaskRecord('failed-scan', task_type, status='failed', payload=payload)
    manager._tasks[old.task_id] = old
    manager._save_task_to_db(old)
    release = threading.Event()
    started = threading.Event()

    def run(task, cancel_check):
        started.set()
        release.wait(5)

    manager.submit_task('excel_face', 'blocking', 'A', run)
    assert started.wait(2)
    try:
        response = client.post(f'/api/task/retry/{old.task_id}')
        assert response.status_code == 200
        task_id = response.get_json()['task_id']
        assert task_id != old.task_id
        new = manager.get_task(task_id)
        assert new.payload == {**payload, 'report_options': {'project_name': 'A', 'employee_order': ['employee-b', 'employee-a'], 'paused_employee_ids': ['employee-b']}}
        assert old.payload['report_options']['project_name'] == 'Stale project'
        assert new.task_type == task_type
        assert new.status == 'queued'
        assert old.status == 'failed'
        assert client.post(f'/api/task/cancel/{task_id}').get_json()['status'] == 'cancelled'
    finally:
        release.set()
        manager._task_queue.join()
