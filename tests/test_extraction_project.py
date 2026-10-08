from src.extraction_project import read_batch_project, write_batch_project


def test_batch_project_survives_reload_and_legacy_is_unknown(tmp_path):
    project={'project_id':'project-a','display_name':'Dự án Long An'}
    assert read_batch_project(tmp_path)['project_id']==''
    write_batch_project(tmp_path,project,'monthly.xlsx')
    assert read_batch_project(tmp_path)=={'project_id':'project-a','project_name':'Dự án Long An','source_filename':'monthly.xlsx'}


def test_extract_requires_explicit_project_before_starting(tmp_path,monkeypatch):
    import src.app as module
    monkeypatch.setattr(module,'PDF_EXTRACTOR_AVAILABLE',True)
    monkeypatch.setattr(module.pdf_extractor,'is_available',lambda:True)
    c=module.app.test_client()
    for kind in ('excel','pdf'):
        response=c.post(f'/api/{kind}/extract',json={'filename':f'monthly.{"xlsx" if kind=="excel" else "pdf"}'})
        assert response.status_code==400
        assert 'Chọn dự án' in response.json['error']


def test_assign_legacy_batch_project_is_visible_in_list(tmp_path, monkeypatch):
    import src.app as module
    from types import SimpleNamespace
    output = tmp_path / 'word'
    people = tmp_path / 'people'
    uploads = tmp_path / 'uploads'
    for root in (output, people, uploads):
        root.mkdir()
    (output / 'monthly').mkdir()
    (people / 'monthly').mkdir()
    (uploads / 'monthly.xlsx').touch()
    project = {'project_id': 'project-a', 'display_name': 'Long An', 'active': True}
    monkeypatch.setattr(module, 'get_identity_registry', lambda: SimpleNamespace(get_project=lambda _: project))
    monkeypatch.setattr(module, 'EXCEL_OUTPUT_DIR', str(output))
    monkeypatch.setattr(module, 'EXCEL_PERSON_DIR', str(people))
    monkeypatch.setattr(module, 'EXCEL_UPLOAD_DIR', str(uploads))
    client = module.app.test_client()
    response = client.post('/api/extraction/project', json={'kind': 'excel', 'folder': 'monthly', 'project_id': 'project-a'})
    assert response.status_code == 200
    assert read_batch_project(people / 'monthly')['project_name'] == 'Long An'
    batch = client.get('/api/excel/files').json['folders'][0]
    assert batch['project_name'] == 'Long An'
    assert batch['source_filename'] == 'monthly.xlsx'
    assert client.post('/api/extraction/project', json={'kind': 'excel', 'folder': '../outside', 'project_id': 'project-a'}).status_code == 400


def test_excel_extraction_saves_selected_project_before_worker(tmp_path, monkeypatch):
    import src.app as module
    from types import SimpleNamespace
    project = {'project_id': 'project-a', 'display_name': 'EHome3', 'active': True}
    monkeypatch.setattr(module, 'get_identity_registry', lambda: SimpleNamespace(get_project=lambda _: project))
    for constant in ('EXCEL_UPLOAD_DIR', 'EXCEL_OUTPUT_DIR', 'EXCEL_PERSON_DIR'):
        root = tmp_path / constant
        root.mkdir()
        monkeypatch.setattr(module, constant, str(root))
    (tmp_path / 'EXCEL_UPLOAD_DIR' / 'monthly.xlsx').touch()
    monkeypatch.setattr(module.threading, 'Thread', lambda **kwargs: SimpleNamespace(start=lambda: None))
    monkeypatch.setattr(module, 'excel_tasks', {})
    client = module.app.test_client()
    response = client.post('/api/excel/extract', json={'filename': 'monthly.xlsx', 'project_id': 'project-a'})
    assert response.status_code == 200
    batch = client.get('/api/excel/files').json['folders'][0]
    assert batch['project_id'] == 'project-a'
    assert batch['project_name'] == 'EHome3'
