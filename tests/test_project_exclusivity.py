from src.identity_registry import IdentityRegistry
from src.employee_audit import audit_project, save_roster
from src.project_exclusivity import exclusive_assignment
import pytest


def test_history_is_distinct_and_exclusive_keeps_source_history(tmp_path):
    r = IdentityRegistry(tmp_path/'ids.db', tmp_path/'portraits')
    a, b, old = (r.register_project(n) for n in ['A','B','Old'])
    eid = r.create_employee('Huynh Quoc Duong')['employee_id']
    duplicate = r.create_employee('Huỳnh Quốc Dương')['employee_id']
    with r._connect() as c:
        r._assign(c, a['project_id'], eid, '001', '2000-01-01', None)
        r._assign(c, b['project_id'], eid, '002', '2000-01-01', None)
        r._assign(c, a['project_id'], duplicate, '', '2000-01-01', None)
        r._assign(c, old['project_id'], duplicate, '', '2000-01-01', '2020-01-01')
    save_roster(r, a['project_id'], 'monthly.xlsx', [{'name':'Huynh Quoc Duong','payroll_code':'001'}])
    historic = audit_project(r,a['project_id'])['employees'][1 if audit_project(r,a['project_id'])['employees'][0]['employee_id']==eid else 0]['other_projects']
    assert any(p['project']=='Old' and not p['current'] for p in historic)
    assert audit_project(r,old['project_id'])['employees']==[]
    preview = exclusive_assignment(r,a['project_id'],eid)
    assert len(preview['affected'])==2
    with pytest.raises(ValueError):
        exclusive_assignment(r,a['project_id'],eid,True,[])
    exclusive_assignment(r,a['project_id'],eid,True,preview['confirmation_keys'])
    assert [e['employee_id'] for e in audit_project(r,a['project_id'])['employees']]==[eid]
    assert audit_project(r,b['project_id'])['employees']==[]
    assert r.list_employees(b['project_id'])[0]['memberships'][0]['payroll_code']=='002'
    assert r.get_employee(duplicate)['active']==1
