"""Audit a SQLite snapshot, leaving the live identities and portraits untouched."""
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import sqlite3
import json
import hashlib
import html
from collections import defaultdict, Counter
from datetime import datetime
from src.identity_registry import IdentityRegistry
from src.employee_audit import audit_project, save_roster
from src.employee_importer import extract_employees_from_file, normalize_vietnamese_name


def main():
    base = Path(__file__).resolve().parents[1]
    output = base / 'results' / ('employee_audit_' + datetime.now().strftime('%Y%m%d_%H%M%S'))
    output.mkdir(parents=True)
    source = sqlite3.connect(f'file:{(base / "data/identity.sqlite3").as_posix()}?mode=ro', uri=True)
    snapshot = sqlite3.connect(output / 'identity_snapshot.sqlite3')
    source.backup(snapshot)
    source.close(); snapshot.close()
    portrait_root = base / '\u1ea2nh BV'
    registry = IdentityRegistry(output / 'identity_snapshot.sqlite3', portrait_root)
    projects = registry.list_projects()
    # Existing folders may not yet be registered. Discover on the snapshot only.
    for folder in portrait_root.iterdir():
        if folder.is_dir() and not folder.is_symlink():
            matches = [p for p in registry.list_projects(include_archived=True) if p['storage_dir'].casefold() == folder.name.casefold()]
            if not matches:
                registry.register_project(folder.name)
    projects = registry.list_projects()
    for p in projects:
        registry.import_legacy(p['project_id'])
    files = sorted((base / 'excel_uploads').glob('*.xlsx'))
    reference = next((f for f in files if f.name == 'T9.2026.xlsx'), None)
    rows = extract_employees_from_file(str(reference)) if reference else []
    reports = []
    for p in projects:
        if rows:
            save_roster(registry, p['project_id'], reference.name, rows)
        reports.append({'project': p['display_name'], **audit_project(registry, p['project_id'])})
    # Exact file duplicates are evidence of a shared image, not proof of identity.
    hashes = defaultdict(list)
    for p in projects:
        root = registry.project_portrait_dir(p['project_id'])
        with registry._connect() as c:
            bindings = c.execute('SELECT employee_id,relative_path FROM portrait_bindings WHERE project_id=? UNION SELECT employee_id,relative_path FROM legacy_sources WHERE project_id=?', (p['project_id'], p['project_id'])).fetchall()
            excluded = {(r[0], r[1]) for r in c.execute('SELECT employee_id,relative_path FROM portrait_exclusions WHERE project_id=?', (p['project_id'],))}
        seen = set()
        for binding in bindings:
            path = registry._bound_path(p['project_id'], binding['relative_path'])
            candidates = path.rglob('*') if path.is_dir() else [path]
            for image in candidates:
                if not image.is_file() or image.is_symlink() or not image.resolve().is_relative_to(root.resolve()) or image.suffix.lower() not in ('.jpg','.jpeg','.png','.bmp','.webp'):
                    continue
                eid = binding['employee_id']
                if (eid, image.relative_to(root).as_posix()) in excluded or (eid, image) in seen:
                    continue
                seen.add((eid, image))
                digest = hashlib.sha256(image.read_bytes()).hexdigest()
                hashes[digest].append({'project': p['display_name'], 'employee_id': eid, 'name': registry.get_employee(eid)['display_name'], 'path': str(image)})
    shared = [items for items in hashes.values() if len({(i['project'], i['employee_id']) for i in items}) > 1]
    stats = []
    for report in reports:
        people = report['employees']
        counts = Counter(e['name'].casefold() for e in people)
        similar_counts = Counter(normalize_vietnamese_name(e['name']) for e in people)
        stats.append({'project': report['project'], 'profiles': len(people),
                      'duplicate_name_groups': sum(n > 1 for n in counts.values()),
                      'similar_name_groups': sum(n > 1 for n in similar_counts.values()),
                      'no_code': sum(not e['codes'] for e in people),
                      'name_mismatch': sum('Tên hồ sơ khác tên trong bảng' in e['issues'] for e in people),
                      'code_absent': sum('Không tìm thấy mã trong bảng đã chọn' in e['issues'] for e in people),
                      'cross_project_candidates': sum(bool(e['other_projects']) for e in people)})
    data = {'reference_file': str(reference) if reference else None,
            'reference_scope': 'Chưa xác nhận bảng thuộc dự án nào; kết quả so với bảng chỉ là tham khảo, không kết luận hồ sơ sai.',
            'stats': stats, 'projects': reports, 'shared_images': shared}
    (output / 'audit.json').write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')
    esc = lambda x: html.escape(str(x))
    body = '<h1>Kiểm tra nhân viên hiện có</h1><p>' + esc(data['reference_scope']) + '</p>'
    body += '<p>Bảng tham khảo: ' + esc(reference.name if reference else 'Không có') + '. Cùng tên không đủ xác nhận cùng người. Ảnh giống hệt chỉ chứng minh dùng chung tệp ảnh. Dữ liệu được kiểm tra trên bản sao; hồ sơ gốc được giữ nguyên.</p>'
    body += '<table><tr><th>Dự án</th><th>Hồ sơ</th><th>Nhóm trùng tên</th><th>Nhóm tên tương tự</th><th>Thiếu mã</th><th>Tên lệch bảng</th><th>Mã vắng trong bảng</th><th>Có ứng viên ở dự án khác</th></tr>'
    for s in stats:
        body += '<tr>' + ''.join('<td>' + esc(v) + '</td>' for v in s.values()) + '</tr>'
    body += '</table>'
    for report in reports:
        body += '<h2>' + esc(report['project']) + '</h2><table><tr><th>Hồ sơ / Mã</th><th>Tên theo bảng</th><th>Cần kiểm tra</th><th>Dự án khác</th></tr>'
        for e in report['employees']:
            others = '; '.join(p['project'] + ': ' + p['name'] + (' (cùng ID)' if p['same_identity'] else ' (khác ID)') + ' — ' + ', '.join(m['payroll_code'] + ' [' + m['valid_from'] + ' → ' + (m['valid_to'] or 'chưa kết thúc') + ']' for m in p['memberships']) for p in e['other_projects'])
            body += '<tr>' + ''.join('<td>' + esc(v) + '</td>' for v in [e['name'] + ' / ' + ', '.join(e['codes']), '; '.join(r['name'] for r in e['roster_rows']), '; '.join(e['issues']) or 'Tên/mã khớp bảng tham khảo; chưa xác minh ảnh', others]) + '</tr>'
        body += '</table>'
    body += '<h2>Tệp ảnh dùng chung giữa hồ sơ</h2>'
    for group in shared:
        body += '<p>' + '<br>'.join(esc(i['project'] + ' / ' + i['name'] + ' / ' + i['path']) for i in group) + '</p>'
    if not shared:
        body += '<p>Không tìm thấy tệp ảnh có nội dung giống hệt giữa các hồ sơ đã liên kết.</p>'
    (output / 'audit.html').write_text('<!doctype html><html lang="vi"><meta charset="utf-8"><title>Kiểm tra nhân viên</title><style>body{font:15px system-ui;padding:24px;color:#172033}table{border-collapse:collapse;width:100%}td,th{border:1px solid #cbd5e1;padding:9px;text-align:left;vertical-align:top}th{background:#e8eef8}h2{margin-top:36px}</style>' + body + '</html>', encoding='utf-8')
    print(json.dumps({'output': str(output), 'stats': stats, 'shared_image_groups': len(shared)}, ensure_ascii=False))


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    main()
