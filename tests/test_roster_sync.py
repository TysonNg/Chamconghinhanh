import pytest
from src.identity_registry import IdentityRegistry
from src.employee_audit import save_roster, audit_project
from src.roster_sync import sync_roster


def test_sync_transfers_code_photos_and_retires_outside_roster(tmp_path):
    r=IdentityRegistry(tmp_path/'ids.db',tmp_path/'portraits')
    a,b=r.register_project('A'),r.register_project('B')
    eid=r.create_employee('Nguyen Van A')['employee_id']
    extra=r.create_employee('Outside Person')['employee_id']
    with r._connect() as c:
        r._assign(c,a['project_id'],eid,'OLD','2020-01-01',None)
        r._assign(c,b['project_id'],extra,'EXTRA','2020-01-01',None)
    root=r.project_portrait_dir(a['project_id']);root.mkdir(parents=True)
    image=root/'old.jpg';image.write_bytes(b'original photo')
    r.bind_portrait(a['project_id'],eid,'old.jpg','test')
    save_roster(r,b['project_id'],'month.xlsx',[
        {'name':'Nguyen Van A','payroll_code':'NEW'},
        {'name':'Tran Van B','payroll_code':'002'}])
    preview=sync_roster(r,b['project_id'],effective_date='2026-09-01')
    assert preview['counts']['transfer']==1 and preview['counts']['create']==1
    assert len(preview['outside_roster'])==1
    sync_roster(r,b['project_id'],True,preview['roster_token'],'2026-09-01')
    assert len(audit_project(r,b['project_id'])['employees'])==2
    assert not audit_project(r,a['project_id'])['employees']
    assert r.resolve_employee(b['project_id'],'NEW','2026-09-15').employee_id==eid
    assert r.resolve_employee(a['project_id'],'OLD','2026-08-15').employee_id==eid
    assert r.portrait_paths(b['project_id'],eid,'2026-09-15')[0].read_bytes()==image.read_bytes()
    assert sync_roster(r,b['project_id'],effective_date='2026-09-01')['counts']['existing']==2


def test_ambiguous_roster_cannot_partially_modify_directory(tmp_path):
    r=IdentityRegistry(tmp_path/'ids.db',tmp_path/'portraits');p=r.register_project('A')
    save_roster(r,p['project_id'],'month.xlsx',[{'name':'Person A','payroll_code':'001'}, {'name':'Person B','payroll_code':'001'}])
    preview=sync_roster(r,p['project_id'])
    assert preview['counts']['conflict']==2
    with pytest.raises(ValueError):sync_roster(r,p['project_id'],True,preview['roster_token'])
    assert not r.list_employees(p['project_id'])


def test_unique_project_code_renames_to_exact_roster_name(tmp_path):
    r=IdentityRegistry(tmp_path/'ids.db',tmp_path/'portraits');p=r.register_project('A')
    eid=r.create_employee('OLD NAME')['employee_id']
    with r._connect() as c:r._assign(c,p['project_id'],eid,'001','2000-01-01',None)
    save_roster(r,p['project_id'],'month.xlsx',[{'name':'Nguyễn Văn A','payroll_code':'001'}])
    preview=sync_roster(r,p['project_id'])
    assert preview['renamed']==1
    assert preview['rows'][0]['previous_name']=='OLD NAME'
    assert r.get_employee(eid)['display_name']=='OLD NAME'
    sync_roster(r,p['project_id'],True,preview['roster_token'])
    assert r.get_employee(eid)['display_name']=='Nguyễn Văn A'
    assert len(r.list_employees(p['project_id']))==1
    assert sync_roster(r,p['project_id'])['renamed']==0
    with r._connect() as c:
        assert c.execute('SELECT previous_name,new_name FROM employee_name_history').fetchone()[:] == ('OLD NAME','Nguyễn Văn A')


def test_choose_duplicate_profile_unlocks_sync(tmp_path):
    r=IdentityRegistry(tmp_path/'ids.db',tmp_path/'portraits');p=r.register_project('A')
    first=r.create_employee('Nguyen Van A')['employee_id']
    second=r.create_employee('Nguyễn Văn A')['employee_id']
    with r._connect() as c:
        r._assign(c,p['project_id'],first,'001','2000-01-01','2024-01-01')
        r._assign(c,p['project_id'],second,'001','2025-01-01','2025-02-01')
    save_roster(r,p['project_id'],'monthly.xlsx',[{'name':'Nguyen Van A','payroll_code':'001'}])
    preview=sync_roster(r,p['project_id'])
    assert preview['counts']['conflict']==1
    assert len(preview['rows'][0]['candidates'])==2
    selected=sync_roster(r,p['project_id'],selections={'001':first})
    assert selected['counts']['conflict']==0
    sync_roster(r,p['project_id'],True,selected['roster_token'],selections={'001':first})
    assert r.resolve_employee(p['project_id'],'001',selected['effective_date']).employee_id==first
    with pytest.raises(ValueError):sync_roster(r,p['project_id'],selections={'001':'unknown'})


def test_nx_prefix_names_are_not_empty_or_duplicate_codes(tmp_path):
    from src.employee_importer import normalize_vietnamese_name, clean_person_name
    assert clean_person_name('NXA4 Truong Van Kiem') == 'Truong Van Kiem'
    assert normalize_vietnamese_name('NXAV Huynh Thi Hoa') == 'huynh thi hoa'
    r=IdentityRegistry(tmp_path/'ids.db',tmp_path/'portraits');p=r.register_project('A')
    save_roster(r,p['project_id'],'monthly.pdf',[
        {'name':'NXA4 Truong Van Kiem','payroll_code':'00358'},
        {'name':'NXAV Huynh Thi Hoa','payroll_code':'00359'}])
    preview=sync_roster(r,p['project_id'])
    assert preview['counts']['conflict']==0
    assert preview['counts']['create']==2


def test_skipped_row_preserves_profile_and_scan_approval_persists(tmp_path):
    r=IdentityRegistry(tmp_path/'ids.db',tmp_path/'portraits');p=r.register_project('A')
    eid=r.create_employee('Old Name')['employee_id']
    with r._connect() as c:r._assign(c,p['project_id'],eid,'001','2000-01-01',None)
    save_roster(r,p['project_id'],'month.xlsx',[{'name':'New Name','payroll_code':'001'},{'name':'Person B','payroll_code':'002'}])
    preview=sync_roster(r,p['project_id'],effective_date='2026-06-01',skipped=['001|New Name'])
    assert preview['counts']['skip']==1
    assert preview['outside_roster']==[]
    assert not preview['sync_approved']
    sync_roster(r,p['project_id'],True,preview['roster_token'],'2026-06-01',skipped=['001|New Name'])
    check=sync_roster(r,p['project_id'])
    assert check['sync_approved'] and check['effective_date']=='2026-06-01'
    assert check['counts']['skip']==1
    assert r.get_employee(eid)['display_name']=='Old Name'
    assert r.list_employees(p['project_id'])[0]['active']==1
    save_roster(r,p['project_id'],'changed.xlsx',[{'name':'New Person','payroll_code':'003'}])
    assert not sync_roster(r,p['project_id'])['sync_approved']


def test_sync_approval_is_preserved_for_multiple_files_in_project(tmp_path):
    r=IdentityRegistry(tmp_path/'ids.db',tmp_path/'portraits');p=r.register_project('A')
    rows=[{'name':'Person A','payroll_code':'001'}]
    for filename in ('june.xlsx','july.xlsx'):
        save_roster(r,p['project_id'],filename,rows)
        preview=sync_roster(r,p['project_id'])
        sync_roster(r,p['project_id'],True,preview['roster_token'])
    save_roster(r,p['project_id'],'june.xlsx',rows)
    assert sync_roster(r,p['project_id'])['sync_approved']
