import io
import json
from concurrent.futures import ThreadPoolExecutor
from datetime import date
from pathlib import Path
from unittest.mock import patch

import pytest
from flask import Flask
from PIL import Image
from src.fake_photo_service import FakePhotoService
from src.identity_registry import IdentityRegistry
from src.face_matcher import MatchResult
from src.supplement_batches import register_batches


def image_bytes(color='blue', captured='2026:09:01 08:00:00'):
    stream = io.BytesIO()
    exif = Image.Exif()
    if captured:
        exif[36867] = captured
    Image.new('RGB', (24, 24), color).save(stream, 'JPEG', exif=exif)
    return stream.getvalue()


class Matcher:
    status = 'matched'
    model_name = 'test'
    detector_backend = 'test'
    distance_metric = 'cosine'
    def match_employee_in_photo(self, **kwargs):
        return MatchResult(self.status, kwargs['image_path'] if self.status == 'matched' else None,
                           .1, kwargs['project_id'], kwargs['employee_id'], self.status)
    def _get_default_threshold(self):
        return .37


@pytest.fixture
def env(tmp_path):
    registry = IdentityRegistry(tmp_path/'identity.sqlite3', tmp_path/'portraits')
    project = registry.register_project('Site')
    employee = registry.create_employee('Employee')
    registry.assign_employee(project['project_id'], employee['employee_id'], '001', '2026-01-01')
    matcher = Matcher()
    service = FakePhotoService(str(tmp_path/'supplement_data'), registry_provider=lambda: registry,
                               matcher_provider=lambda: matcher)
    app = Flask(__name__)
    app.testing = True
    register_batches(app, service.data_dir, registry_provider=lambda: registry,
                     matcher_provider=lambda: matcher, service=service)
    return service, registry, project, employee, matcher, app.test_client()


def record(env, color='blue'):
    service, registry, project, employee, matcher, client = env
    with patch.object(service.ai_service, 'inspect_photo', return_value={'status':'unconfigured'}):
        return service.create_supplement_photo(image_bytes(color), project['storage_dir'], employee['display_name'],
                    '2026-09-01', 'a.jpg', project_id=project['project_id'], employee_id=employee['employee_id'])


@pytest.mark.parametrize('ids', [None, [], '', {}, [None], [1]])
def test_delete_requires_explicit_ids(env, ids):
    item = record(env)
    client = env[-1]
    payload = {} if ids is None else {'ids':ids}
    assert client.post('/api/supplement/delete-staging', json=payload).status_code == 400
    assert len(env[0].get_staging_photos()) == 1


def test_delete_missing_id_is_atomic(env):
    item = record(env)
    response = env[-1].post('/api/supplement/delete-staging', json={'ids':[item['id'], 'missing']})
    assert response.status_code == 409
    assert env[0].get_staging_image_path(item['id']).exists()


@pytest.mark.parametrize('value', [None, '99:99', '24:00', '12:60:00', 800])
def test_invalid_time_is_400_without_creation(env, value):
    service, _, project, employee, _, client = env
    response = client.post('/api/supplement/records', data={'project_id':project['project_id'],
            'employee_id':employee['employee_id'], 'configs':json.dumps([{'target_date':'2026-09-01','target_time':value}]),
            'photos':(io.BytesIO(image_bytes()), 'a.jpg')})
    assert response.status_code == 400
    assert service.get_staging_photos() == []


def test_name_only_and_wrong_membership_rejected(env):
    _, registry, project, employee, _, client = env
    args={'project':'Site','employee':'Employee','configs':'[{"target_date":"2026-09-01"}]'}
    assert client.post('/api/supplement/records', data={**args,'photos':(io.BytesIO(image_bytes()),'a.jpg')}).status_code == 400
    other=registry.create_employee('Employee')
    args.update(project_id=project['project_id'],employee_id=other['employee_id'])
    assert client.post('/api/supplement/records', data={**args,'photos':(io.BytesIO(image_bytes()),'a.jpg')}).status_code == 400


def test_missing_time_retained_as_null_and_no_mutation(env):
    item = record(env)
    assert item['target_time'] is None
    assert item['processing'] == {'watermark':'not_requested','exif':'not_requested'}
    assert item['face_status'] == 'matched'
    with env[0]._connect() as c:
        assert c.execute('SELECT target_time FROM staging_photos WHERE id=?',(item['id'],)).fetchone()[0] is None


def test_tamper_same_day_is_blocked(env):
    item = record(env)
    (env[0].raw_dir/item['original_file_name']).write_bytes(image_bytes('red'))
    result=env[0].apply_to_attendance([item['id']])
    assert not result['success'] and result['count']==0
    assert not list(env[0].input_images_dir.rglob('*.jpg'))


def test_no_match_never_enters_attendance(env):
    item = record(env)
    env[4].status='no_match'
    assert not env[0].apply_to_attendance([item['id']])['success']
    assert not list(env[0].input_images_dir.rglob('*.jpg'))


def test_duplicate_content_different_id_does_not_apply_twice(env):
    first, second=record(env), record(env)
    assert env[0].apply_to_attendance([first['id']])['success']
    result=env[0].apply_to_attendance([second['id']])
    assert result['success']
    assert len(list(env[0].input_images_dir.rglob('*.jpg')))==1


def test_get_keeps_missing_derived_record(env):
    item = record(env)
    env[0].get_staging_image_path(item['id']).unlink()
    items=env[0].get_staging_photos()
    assert len(items)==1 and items[0]['storage_status']=='derived_missing'
    with env[0]._connect() as c:
        assert c.execute('SELECT count(*) FROM staging_photos').fetchone()[0]==1


def test_apply_failure_second_metadata_rolls_back_all(env, monkeypatch):
    first,second=record(env),record(env,'red')
    original=Path.write_text
    def fail(path,*args,**kwargs):
        if str(path).endswith(second['id']+'.jpg.json'):
            raise OSError('injected disk failure')
        return original(path,*args,**kwargs)
    monkeypatch.setattr(Path,'write_text',fail)
    result=env[0].apply_to_attendance([first['id'],second['id']])
    assert not result['success'] and result['count']==0
    assert not list(env[0].input_images_dir.rglob('*.jpg'))
    with env[0]._connect() as c:
        assert not any(r[0] for r in c.execute('SELECT applied_path FROM staging_photos'))


def test_apply_timestamp_independent_of_delete_after_and_idempotent(env):
    item=record(env)
    assert env[0].apply_to_attendance([item['id']],delete_after=False)['success']
    with env[0]._connect() as c:
        before=c.execute('SELECT applied_at FROM staging_photos WHERE id=?',(item['id'],)).fetchone()[0]
    assert before
    assert env[0].apply_to_attendance([item['id']])['success']
    with env[0]._connect() as c:
        assert before==c.execute('SELECT applied_at FROM staging_photos WHERE id=?',(item['id'],)).fetchone()[0]
    assert env[-1].post('/api/supplement/delete-staging',json={'ids':[item['id']]}).status_code==409


def test_zip_without_approval_and_approve_retired(env):
    client=env[-1]
    response=client.post('/api/supplement/batches',data={'employee':'A','supplement_date':'2026-09-01',
        'photos':(io.BytesIO(image_bytes()),'a.jpg')})
    batch=response.json['batch']
    url='/api/supplement/batches/'+batch['id']
    assert client.get(url+'/download').status_code==200
    assert client.post(url+'/approve',json={}).status_code==410
    assert client.patch(url+'/items/'+batch['items'][0]['id'],json={'supplement_date':'2026-09-02'}).status_code==200
    assert client.get(url+'/download').status_code==200


def test_face_result_invalidates_when_model_changes(env):
    item=record(env)
    assert item['can_apply']
    env[4].model_name='changed'
    refreshed=env[0].get_staging_photos()[0]
    assert not refreshed['can_apply'] and refreshed['face_status']=='not_run'


def test_face_result_invalidates_when_portrait_changes(env):
    _,r,p,e,_,_=env
    folder=r.project_portrait_dir(p['project_id'])/e['employee_id']
    folder.mkdir(parents=True)
    portrait=folder/'p.jpg'
    portrait.write_bytes(image_bytes())
    r.bind_portrait(p['project_id'],e['employee_id'],e['employee_id'],'test')
    item=record(env)
    portrait.write_bytes(image_bytes('red'))
    assert not env[0].get_staging_photos()[0]['can_apply']


def test_membership_rechecked_before_apply(env):
    item=record(env)
    with env[1]._connect() as c:
        c.execute('UPDATE memberships SET valid_to=?',('2026-09-01',))
    result=env[0].apply_to_attendance([item['id']])
    assert not result['success'] and result['count']==0


def test_same_content_different_employee_conflicts(env):
    one=record(env)
    r=env[1]
    other=r.create_employee('Employee')
    r.assign_employee(env[2]['project_id'],other['employee_id'],'002','2026-01-01')
    two=env[0].create_supplement_photo(image_bytes(),env[2]['storage_dir'],'Employee','2026-09-01','a.jpg',
                                    project_id=env[2]['project_id'],employee_id=other['employee_id'])
    assert env[0].apply_to_attendance([one['id']])['success']
    result=env[0].apply_to_attendance([two['id']])
    assert not result['success'] and result['count']==0
    assert len(list(env[0].input_images_dir.rglob('*.jpg')))==1


def test_uncommitted_files_hidden_and_recovered(env):
    from src.supplement_evidence import evidence_is_visible
    import hashlib
    item=record(env)
    service=env[0]
    dest=service.input_images_dir/'Site'/'2026-09-01'/(item['id']+'.jpg')
    dest.parent.mkdir(parents=True)
    raw=image_bytes()
    digest=hashlib.sha256(raw).hexdigest()
    op={'id':'unfinished','kind':'apply','files':[{'dest':str(dest),'sha256':digest}]}
    with service._connect() as c:
        c.execute('INSERT INTO supplement_operations VALUES(?,?,?,?)',('unfinished','prepared',json.dumps(op),'time'))
    dest.with_suffix('.jpg.json').write_text(json.dumps({'operation_id':'unfinished','operation_db':str(service.db_path), 'sha256':digest}),encoding='utf-8')
    dest.write_bytes(raw)
    assert not evidence_is_visible(dest)
    reopened=FakePhotoService(str(service.data_dir),registry_provider=service.registry_provider,matcher_provider=service.matcher_provider)
    assert not dest.exists() and not dest.with_suffix('.jpg.json').exists()
    with reopened._connect() as c:
        assert c.execute('SELECT state FROM supplement_operations WHERE id=?',('unfinished',)).fetchone()[0]=='rolled_back'


def test_committed_files_visible(env):
    from src.supplement_evidence import evidence_is_visible
    item=record(env)
    result=env[0].apply_to_attendance([item['id']])
    assert result['success']
    dest=Path(result['applied'][0]['dest_path'])
    assert evidence_is_visible(dest)
    dest.write_bytes(image_bytes('red'))
    assert not evidence_is_visible(dest)


def test_two_applies_concurrent_do_not_publish_duplicates(env):
    one,two=record(env),record(env)
    with ThreadPoolExecutor(max_workers=2) as pool:
        results=list(pool.map(lambda i:env[0].apply_to_attendance([i['id']]),[one,two]))
    assert all(r['success'] for r in results)
    assert len(list(env[0].input_images_dir.rglob('*.jpg')))==1


def test_recovery_does_not_remove_foreign_destination(env):
    import hashlib
    service=env[0]
    dest=service.input_images_dir/'Site'/'2026-09-01'/'foreign.jpg'
    dest.parent.mkdir(parents=True)
    raw=image_bytes()
    dest.write_bytes(raw)
    op={'id':'unwritten','kind':'apply','files':[{'dest':str(dest),'sha256':hashlib.sha256(raw).hexdigest()}]}
    with service._connect() as c:
        c.execute('INSERT INTO supplement_operations VALUES(?,?,?,?)',('unwritten','prepared',json.dumps(op),'time'))
    service._recover_operations()
    assert dest.read_bytes()==raw


def test_corrupt_inspection_blocks_apply(env):
    item=record(env)
    with env[0]._connect() as c:
        c.execute('UPDATE staging_photos SET inspection_json=? WHERE id=?',('[]',item['id']))
    assert not env[0].apply_to_attendance([item['id']])['success']


def test_original_exif_inside_ifd_and_date_conflict(env):
    import piexif
    service=env[0]
    buffer=io.BytesIO()
    exif=piexif.dump({'Exif':{piexif.ExifIFD.DateTimeOriginal:b'2026:09:01 08:00:00'}})
    Image.new('RGB',(24,24),'blue').save(buffer,'JPEG',exif=exif)
    with patch.object(service.ai_service,'inspect_photo',return_value={'status':'ok','visible_date':'2026-08-31'}):
        item=service.create_supplement_photo(buffer.getvalue(),'Site','Employee','2026-09-01','a.jpg',
              project_id=env[2]['project_id'],employee_id=env[3]['employee_id'])
    assert item['inspection']['exif_datetime']=='2026-09-01 08:00:00'
    assert item['date_status']=='mismatch'
    assert not service.apply_to_attendance([item['id']])['success']
