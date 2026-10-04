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
    response=client.post('/api/supplement/batches',data={'project_id':env[2]['project_id'],'employee_id':env[3]['employee_id'],'supplement_date':'2026-09-01',
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
    stat=dest.stat()
    op['files'][0].update(device=stat.st_dev,inode=stat.st_ino)
    with service._connect() as c:
        c.execute('UPDATE supplement_operations SET document=? WHERE id=?',(json.dumps(op),'unfinished'))
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


def test_foreign_image_racing_publication_is_not_deleted(env, monkeypatch):
    import os
    service=env[0]
    item=record(env)
    dest=service.input_images_dir/'Site'/'2026-09-01'/(item['id']+'.jpg')
    original_open=Path.open
    original_link=os.link
    def race_open(path,*args,**kwargs):
        if path==dest and args and args[0]=='xb':
            path.write_bytes(b'foreign image')
        return original_open(path,*args,**kwargs)
    def race_link(source,target,*args,**kwargs):
        if Path(target)==dest:
            dest.write_bytes(b'foreign image')
        return original_link(source,target,*args,**kwargs)
    monkeypatch.setattr(Path,'open',race_open)
    monkeypatch.setattr(os,'link',race_link)
    result=service.apply_to_attendance([item['id']])
    assert not result['success'] and result['count']==0
    assert dest.read_bytes()==b'foreign image'


@pytest.mark.parametrize('field,value',[('employee_id',{'x':1}),('project_id',['x']),('employee_id',1),('project_id',True),('employee_id','')])
def test_patch_bad_identity_is_400(env,field,value):
    item=record(env)
    response=env[-1].patch('/api/supplement/records/'+item['id'],json={field:value})
    assert response.status_code==400


def test_upload_second_photo_failure_leaves_no_partial_records(env, monkeypatch):
    service=env[0]
    original=service.create_supplement_photo
    count=0
    def fail_second(*args,**kwargs):
        nonlocal count
        count+=1
        if count==2:
            raise RuntimeError('injected second intake failure')
        return original(*args,**kwargs)
    monkeypatch.setattr(service,'create_supplement_photo',fail_second)
    response=env[-1].post('/api/supplement/records',data={'project_id':env[2]['project_id'],
        'employee_id':env[3]['employee_id'],'configs':'[{"target_date":"2026-09-01"},{"target_date":"2026-09-01"}]',
        'photos':[(io.BytesIO(image_bytes()),'a.jpg'),(io.BytesIO(image_bytes('red')),'b.jpg')]})
    assert response.status_code==409
    assert service.get_staging_photos()==[]
    assert not list(service.raw_dir.iterdir()) and not list(service.staging_dir.iterdir())


def test_duplicate_destination_disappearing_before_commit_cannot_succeed(env, monkeypatch):
    one,two=record(env),record(env)
    service=env[0]
    first=service.apply_to_attendance([one['id']])
    dest=Path(first['applied'][0]['dest_path'])
    original_connect=service._connect
    connections=0
    def connect():
        nonlocal connections
        connections+=1
        if connections==5:
            dest.unlink(missing_ok=True)
        return original_connect()
    monkeypatch.setattr(service,'_connect',connect)
    result=service.apply_to_attendance([two['id']])
    if not dest.exists():
        assert not result['success']


def test_missing_supplement_metadata_never_bypasses_commit(env):
    from src.supplement_evidence import evidence_is_visible
    item=record(env)
    result=env[0].apply_to_attendance([item['id']])
    dest=Path(result['applied'][0]['dest_path'])
    dest.with_suffix(dest.suffix+'.json').unlink()
    assert not evidence_is_visible(dest)


def test_created_records_and_images_are_hidden_until_intake_commit(env,monkeypatch):
    service=env[0]
    original=service.create_supplement_photo
    observed=[]
    def intercept(*args,**kwargs):
        item=original(*args,**kwargs)
        observed.append(service.get_staging_photos())
        assert service.get_staging_image_path(item['id']) is None
        return item
    monkeypatch.setattr(service,'create_supplement_photo',intercept)
    response=env[-1].post('/api/supplement/records',data={'project_id':env[2]['project_id'],
        'employee_id':env[3]['employee_id'],'configs':'[{"target_date":"2026-09-01"}]',
        'photos':(io.BytesIO(image_bytes()),'a.jpg')})
    assert response.status_code==201
    assert observed==[[]]
    assert len(service.get_staging_photos())==1
    assert service.get_staging_image_path(response.json['items'][0]['id']).exists()


def test_prepared_intake_recovered_after_restart(env):
    service=env[0]
    op={'id':'intake-unfinished','kind':'create','files':[],'record_ids':['abc123abc123']}
    with service._connect() as c:
        c.execute('INSERT INTO supplement_operations VALUES(?,?,?,?)',(op['id'],'prepared',json.dumps(op),'time'))
    item=service.create_supplement_photo(image_bytes(),'Site','Employee','2026-09-01','a.jpg',
        project_id=env[2]['project_id'],employee_id=env[3]['employee_id'],record_id=op['record_ids'][0],intake_operation_id=op['id'])
    assert service.get_staging_photos()==[]
    assert (service.raw_dir/item['original_file_name']).is_file()
    reopened=FakePhotoService(str(service.data_dir),registry_provider=service.registry_provider,matcher_provider=service.matcher_provider)
    assert reopened.get_staging_photos()==[]
    assert not list(service.raw_dir.iterdir()) and not list(service.staging_dir.iterdir())


def test_migration_backup_preserves_legacy_and_does_not_invent_hash(tmp_path):
    import sqlite3
    root=tmp_path/'supplement_data'
    root.mkdir()
    raw=root/'supplement_staging_raw'
    raw.mkdir()
    (raw/'old_original.jpg').write_bytes(image_bytes())
    db=root/'supplement_batches.sqlite3'
    with sqlite3.connect(db) as c:
        c.execute('CREATE TABLE staging_photos(id TEXT PRIMARY KEY,project TEXT NOT NULL,employee TEXT NOT NULL,target_date TEXT NOT NULL,target_time TEXT NOT NULL,original_name TEXT NOT NULL,file_name TEXT NOT NULL,created_at TEXT NOT NULL)')
        c.execute('INSERT INTO staging_photos VALUES(?,?,?,?,?,?,?,?)',('old','Site','Employee','2026-09-01','','old.jpg','old.jpg','time'))
    service=FakePhotoService(str(root))
    backups=list((root/'backups').iterdir())
    assert len(backups)==1
    assert (backups[0]/'supplement_staging_raw'/'old_original.jpg').read_bytes()==image_bytes()
    with service._connect() as c:
        row=c.execute('SELECT source_sha256,target_time,source_provenance FROM staging_photos WHERE id=?',('old',)).fetchone()
        assert tuple(row)==('',None,'legacy_unverified')
    FakePhotoService(str(root))
    assert len(list((root/'backups').iterdir()))==1


def test_direct_image_route_hides_prepared_and_blocks_traversal(env):
    import ast
    from flask import jsonify,send_file
    from src.supplement_evidence import evidence_is_visible
    service=env[0]
    item=record(env)
    dest=service.input_images_dir/'Site'/'2026-09-01'/(item['id']+'.jpg')
    dest.parent.mkdir(parents=True)
    dest.write_bytes(image_bytes())
    dest.with_suffix('.jpg.json').write_text(json.dumps({'source':'supplement_original','operation_id':'missing','operation_db':str(service.db_path)}),encoding='utf-8')
    tree=ast.parse(Path('src/app.py').read_text(encoding='utf-8'))
    function=next(n for n in tree.body if isinstance(n,ast.FunctionDef) and n.name=='serve_image')
    function.decorator_list=[]
    namespace={'Path':Path,'INPUT_IMAGES_DIR':str(service.input_images_dir),'jsonify':jsonify,'send_file':send_file}
    exec(compile(ast.Module(body=[function],type_ignores=[]),'src/app.py','exec'),namespace)
    app=env[-1].application
    app.add_url_rule('/images/<path:filename>','direct_test',namespace['serve_image'])
    assert env[-1].get('/images/Site/2026-09-01/'+dest.name).status_code==404
    assert env[-1].get('/images/../identity.sqlite3').status_code==404


def test_daily_delete_uses_same_evidence_write_guard(env,monkeypatch):
    from src.daily_photo_routes import register_daily_photo_routes
    service=env[0]
    register_daily_photo_routes(env[-1].application,lambda:service.input_images_dir,
                  project_resolver=lambda reference:env[1].get_project(reference),evidence_service_provider=lambda:service)
    item=record(env)
    result=service.apply_to_attendance([item['id']])
    assert result['success']
    guarded=[]
    original=service.mutate_attendance
    def intercept(action):
        guarded.append(True)
        return original(action)
    monkeypatch.setattr(service,'mutate_attendance',intercept)
    response=env[-1].post('/api/photos/daily/delete',json={'project_id':env[2]['project_id'],
           'date':'2026-09-01','filename':result['applied'][0]['filename']})
    assert response.status_code==200 and response.json['deleted_count']==1
    assert guarded==[True]
    assert not service.apply_to_attendance([item['id']])['success']


def test_partial_delete_failure_restores_both_files(env,monkeypatch):
    import os
    service=env[0]
    item=record(env)
    original=os.replace
    count=0
    def fail_second(source,target):
        nonlocal count
        count+=1
        if count==2:
            raise OSError('injected rename failure')
        return original(source,target)
    monkeypatch.setattr(os,'replace',fail_second)
    with pytest.raises(RuntimeError):
        service.delete_staging_photos([item['id']])
    assert len(service.get_staging_photos())==1
    assert service.get_staging_image_path(item['id']).exists()
    assert (service.raw_dir/item['original_file_name']).exists()


def test_baseline_rebinding_is_not_reported_as_original_intake_hash(env):
    item=record(env)
    with env[0]._connect() as c:
        c.execute('UPDATE staging_photos SET source_sha256=?,source_provenance=? WHERE id=?',('','legacy_unverified',item['id']))
    response=env[-1].patch('/api/supplement/records/'+item['id'],json={'project_id':env[2]['project_id'],
        'employee_id':env[3]['employee_id'],'confirm_legacy_source':True})
    assert response.status_code==200
    assert response.json['item']['source_provenance']=='legacy_baseline'
    assert response.json['item']['integrity_status']=='baseline_verified'


def test_modified_derived_file_does_not_claim_processing_success(env):
    item=record(env)
    with env[0]._connect() as c:
        c.execute("UPDATE staging_photos SET watermark_status='completed',exif_status='completed' WHERE id=?",(item['id'],))
    env[0].get_staging_image_path(item['id']).write_bytes(image_bytes('red'))
    result=env[0].get_staging_photos()[0]
    assert result['storage_status']=='derived_changed'
    assert result['processing']['watermark']!='completed'
    assert result['can_apply']  # Original evidence remains independently valid.


def test_unverified_local_watermark_fallback_is_not_completed(env,monkeypatch):
    service=env[0]
    analysis={'status':'confirmed','suggested_lines':['1 Th9, 2026 08:00:00','Address'],
              'block_box_2d':[700,0,1000,1000]}
    def fallback(source,target,*args):
        Path(target).write_bytes(image_bytes('red'))
        return True
    monkeypatch.setattr(service.ai_service,'analyze_watermark_block',lambda *args:analysis)
    monkeypatch.setattr(service.ai_service,'is_configured',lambda:True)
    monkeypatch.setattr(service.ai_service,'generate_verified_watermark_crop',lambda *args:{'status':'verification_failed'})
    monkeypatch.setattr(service.ai_service,'verify_watermark_crop',lambda *args:False)
    monkeypatch.setattr(service.smart_replacer,'replace_timestamp',fallback)
    item=service.create_supplement_photo(image_bytes(),'Site','Employee','2026-09-01','a.jpg',target_time='09:00',
              project_id=env[2]['project_id'],employee_id=env[3]['employee_id'])
    assert item['processing']['watermark']=='verification_failed'
    assert service.get_staging_image_path(item['id']).read_bytes()==image_bytes()


def test_watermark_crop_hides_uncommitted_intake(env):
    service=env[0]
    item=record(env)
    with service._connect() as c:
        c.execute('INSERT INTO supplement_operations VALUES(?,?,?,?)',('crop_pending','prepared',json.dumps({'id':'crop_pending','kind':'create','files':[]}),'time'))
        c.execute('UPDATE staging_photos SET intake_operation_id=? WHERE id=?',('crop_pending',item['id']))
    assert service.get_watermark_crop_bytes(item['id']) is None


def test_old_unfinished_journal_preserves_metadata_and_reports_unresolved(env):
    service=env[0]
    dest=service.input_images_dir/'Site'/'2026-09-01'/'legacy.jpg'
    dest.parent.mkdir(parents=True)
    dest.write_bytes(image_bytes())
    sidecar=dest.with_suffix('.jpg.json')
    sidecar.write_text(json.dumps({'source':'supplement_original','operation_id':'old_op'}),encoding='utf-8')
    doc={'id':'old_op','kind':'apply','files':[{'dest':str(dest)}]}
    with service._connect() as c:
        c.execute('INSERT INTO supplement_operations VALUES(?,?,?,?)',('old_op','prepared',json.dumps(doc),'time'))
    with pytest.raises(RuntimeError,match='ownership'):
        service._recover_operations()
    assert dest.exists() and sidecar.exists()
    with service._connect() as c:
        assert c.execute('SELECT state FROM supplement_operations WHERE id=?',('old_op',)).fetchone()[0]=='prepared'


def test_final_intake_bytes_are_fsynced_before_commit(env,monkeypatch):
    import os
    service=env[0]
    sizes=[]
    original=os.fsync
    def track(fd):
        sizes.append(os.fstat(fd).st_size)
        return original(fd)
    monkeypatch.setattr(os,'fsync',track)
    service.create_records([(image_bytes(),'2026-09-01',None,'a.jpg')],env[2]['project_id'],env[3]['employee_id'])
    assert len([size for size in sizes if size==len(image_bytes())])>=2


@pytest.mark.parametrize('analysis_status,configured',[('needs_confirmation',False),('confirmed',False),('needs_confirmation',True)])
def test_local_fallback_verifies_expected_lines_before_reporting_success(env,monkeypatch,analysis_status,configured):
    service=env[0]
    analysis={'status':analysis_status,'suggested_lines':['1 Th9, 2026 08:00:00','Address'],'block_box_2d':[700,0,1000,1000]}
    calls=[]
    def fallback(source,target,*args):
        Path(target).write_bytes(image_bytes('red'))
        return True
    monkeypatch.setattr(service.ai_service,'analyze_watermark_block',lambda *args:analysis)
    monkeypatch.setattr(service.ai_service,'is_configured',lambda:configured)
    monkeypatch.setattr(service.ai_service,'verify_watermark_crop',lambda raw,lines:calls.append(lines) or True)
    monkeypatch.setattr(service.smart_replacer,'replace_timestamp',fallback)
    item=service.create_supplement_photo(image_bytes(),'Site','Employee','2026-09-01','a.jpg',target_time='09:00',project_id=env[2]['project_id'],employee_id=env[3]['employee_id'])
    assert calls==[['01 Th9, 2026 09:00:00','Address']]
    assert item['processing']['watermark']=='completed'


def test_archive_creation_requires_valid_stable_identity(env):
    response=env[-1].post('/api/supplement/batches',data={'employee':'A','supplement_date':'2026-09-01','photos':(io.BytesIO(image_bytes()),'a.jpg')})
    assert response.status_code==400
    response=env[-1].post('/api/supplement/batches',data={'project_id':'unknown','employee_id':env[3]['employee_id'],'supplement_date':'2026-09-01','photos':(io.BytesIO(image_bytes()),'a.jpg')})
    assert response.status_code==400


def test_staging_zip_rejects_changed_or_missing_file_as_whole(env):
    first=record(env)
    second=record(env,'green')
    env[0].get_staging_image_path(second['id']).write_bytes(image_bytes('red'))
    assert env[-1].post('/api/supplement/download-staging-zip',json={'ids':[first['id'],second['id']]}).status_code==409
    assert env[-1].get('/api/supplement/download-staging/'+second['id']).status_code==409
    env[0].get_staging_image_path(second['id']).unlink()
    assert env[-1].post('/api/supplement/download-staging-zip',json={'ids':[first['id'],second['id']]}).status_code==409


def test_download_delete_race_is_controlled_conflict(env,monkeypatch):
    service=env[0]
    item=record(env)
    original=service.get_staging_image_path
    def delete_between_reads(photo_id):
        path=original(photo_id)
        service.delete_staging_photos([photo_id])
        return path
    monkeypatch.setattr(service,'get_staging_image_path',delete_between_reads)
    with pytest.raises(RuntimeError):
        service.read_verified_derived(item['id'])
