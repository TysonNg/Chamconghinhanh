import io
import json
import zipfile
from urllib.parse import unquote

from docx import Document
from flask import Flask
from PIL import Image
import pytest

from src.daily_photo_routes import register_daily_photo_routes


def make_client(root, resolver=None):
    app = Flask(__name__)
    app.testing = True
    register_daily_photo_routes(app, lambda: root, project_resolver=resolver)
    return app.test_client()


def add_photo(root, project, day, name, shift=None, size=(80, 140)):
    folder = root / project / day
    folder.mkdir(parents=True, exist_ok=True)
    path = folder / name
    Image.new('RGB', size, name.encode('utf-8')[0]).save(path)
    if shift:
        path.with_suffix(path.suffix + '.json').write_text(
            json.dumps({'shift': shift, 'send_time': '16:05:59' if shift == 'afternoon' else '09:05:59'}), encoding='utf-8')
    return path


def report_text(doc):
    texts = [p.text for p in doc.paragraphs]
    texts.extend(c.text for t in doc.tables for r in t.rows for c in r.cells)
    return '\n'.join(texts).replace('\u200b', '')


def test_export_matches_gallery_and_is_scoped_to_project_and_date(tmp_path):
    client = make_client(tmp_path)
    for name, shift in [('sáng.png', 'morning'), ('chiều.png', 'afternoon'),
                        ('ngoài.png', 'outside'), ('cũ.png', None), ('lỗi.png', None)]:
        p = add_photo(tmp_path, 'Dự án A', '2026-09-18', name, shift)
        if name == 'lỗi.png':
            p.with_suffix('.png.json').write_text('bad metadata', encoding='utf-8')
    add_photo(tmp_path, 'Dự án B', '2026-09-18', 'other-project.png', 'morning')
    add_photo(tmp_path, 'Dự án A', '2026-08-18', 'other-month.png', 'morning')
    query = '?project=Dự án A&date=2026-09-18'
    response = client.get('/api/photos/daily/export-word' + query)
    assert response.status_code == 200
    assert response.mimetype.endswith('wordprocessingml.document')
    assert 'Anh_Dự án A_2026-09-18.docx' in unquote(response.headers['Content-Disposition'])
    doc = Document(io.BytesIO(response.data))
    text = report_text(doc)
    assert 'Tổng số: 2 ảnh' in text
    assert 'Ca sáng · 06:00–18:00 (1 ảnh)' in text
    assert 'Ca chiều · 18:00–06:00 (1 ảnh)' in text
    assert 'Ảnh khác' not in text
    assert 'Giờ gửi:' not in text
    listing = client.get('/api/photos/daily/2026-09-18?project=Dự án A').json
    for photo in listing['photos']:
        assert photo['filename'] not in text
    assert 'other-project' not in text and 'other-month' not in text
    assert len(doc.inline_shapes) == 2
    assert listing['total'] == 5


@pytest.mark.parametrize('shift', ['morning', 'afternoon', 'outside', None])
def test_single_group_and_empty_shifts(tmp_path, shift):
    add_photo(tmp_path, 'Site', '2026-09-18', 'a.png', shift)
    response = make_client(tmp_path).get('/api/photos/daily/export-word?project=Site&date=2026-09-18')
    doc = Document(io.BytesIO(response.data))
    text = report_text(doc)
    assert 'Ca sáng' in text and 'Ca chiều' in text
    assert text.count('Không có ảnh') == (1 if shift in ('morning', 'afternoon') else 2)
    assert 'Ảnh khác' not in text


@pytest.mark.parametrize('send_time,included', [
    ('05:40:15', False), ('06:59:59', False), ('07:00:00', True),
    ('07:00:01', True), ('15:00:00', True), ('15:00:01', False),
    (None, False), ('bad time', False), (123, False),
])
def test_word_morning_starts_at_seven_without_changing_gallery(tmp_path, send_time, included):
    path = add_photo(tmp_path, 'Site', '2026-09-18', 'am.png', 'morning')
    path.with_suffix('.png.json').write_text(json.dumps({'shift': 'morning', 'send_time': send_time}))
    add_photo(tmp_path, 'Site', '2026-09-18', 'pm.png', 'afternoon')
    client = make_client(tmp_path)
    response = client.get('/api/photos/daily/export-word?project=Site&date=2026-09-18')
    assert response.status_code == 200
    doc = Document(io.BytesIO(response.data))
    text = report_text(doc)
    count = int(included)
    assert f'Ca sáng · 06:00–18:00 ({count} ảnh)' in text
    assert 'Ca chiều · 18:00–06:00 (1 ảnh)' in text
    assert f'Tổng số: {1 + count} ảnh' in text
    assert len(doc.inline_shapes) == 1 + count
    assert client.get('/api/photos/daily/2026-09-18?project=Site').json['total'] == 2


def test_project_id_uses_display_name_and_mapped_legacy_day(tmp_path):
    add_photo(tmp_path, 'storage', '18', 'a.png', 'morning')
    (tmp_path / 'storage' / 'attendance-period.json').write_text(
        json.dumps({'version': 1, 'mappings': {'18': {
            'date': '2026-09-18', 'confirmed_by': 'test', 'confirmed_at': '2026-09-18T00:00:00Z'}}}))
    resolver = lambda ref: {'storage_dir': 'storage', 'display_name': 'Tên mới', 'active': True}
    client = make_client(tmp_path, resolver)
    response = client.get('/api/photos/daily/export-word?project_id=pid&date=2026-09-18')
    assert response.status_code == 200
    assert 'Tên mới' in report_text(Document(io.BytesIO(response.data)))
    assert 'Anh_Tên mới_2026-09-18.docx' in unquote(response.headers['Content-Disposition'])
    assert client.get('/api/photos/daily/export-word?project_id=pid&date=2026-08-18').status_code == 400


@pytest.mark.parametrize('morning,afternoon,expected_am,expected_pm', [
    (20, 20, 15, 15), (1, 20, 1, 15), (20, 0, 15, 0), (0, 20, 0, 15),
])
def test_daily_limit_is_fifteen_per_shift(tmp_path, morning, afternoon, expected_am, expected_pm):
    for shift, count in [('morning', morning), ('afternoon', afternoon)]:
        for i in range(count):
            add_photo(tmp_path, 'Site', '2026-09-18', f'{shift}{i:02}.png', shift)
    client = make_client(tmp_path)
    response = client.get('/api/photos/daily/export-word?project=Site&date=2026-09-18')
    assert response.status_code == 200
    doc = Document(io.BytesIO(response.data))
    assert len(doc.inline_shapes) == expected_am + expected_pm
    text = report_text(doc)
    assert f'Ca sáng · 06:00–18:00 ({expected_am} ảnh)' in text
    assert f'Ca chiều · 18:00–06:00 ({expected_pm} ảnh)' in text
    assert client.get('/api/photos/daily/2026-09-18?project=Site').json['total'] == morning + afternoon


@pytest.mark.parametrize('send_time,included', [
    ('15:59:59', False), ('16:00:00', True), ('23:59:59', True),
    ('00:00:00', False), (None, False), ('bad time', False),
])
def test_afternoon_export_starts_at_sixteen(tmp_path, send_time, included):
    path = add_photo(tmp_path, 'Site', '2026-09-18', 'pm.png', 'afternoon')
    path.with_suffix('.png.json').write_text(json.dumps({'shift': 'afternoon', 'send_time': send_time}))
    response = make_client(tmp_path).get('/api/photos/daily/export-word?project=Site&date=2026-09-18')
    assert response.status_code == 200
    doc = Document(io.BytesIO(response.data))
    assert len(doc.inline_shapes) == int(included)


def test_empty_invalid_ambiguous_and_broken_images(tmp_path):
    client = make_client(tmp_path)
    for query in ['project=Site&date=2026-09-18', 'project=Site&date=18',
                  'project=Site&date=2026-02-30', 'project=../bad&date=2026-09-18']:
        assert client.get('/api/photos/daily/export-word?' + query).status_code == 400
    path = add_photo(tmp_path, 'Site', '18', 'legacy.png')
    assert client.get('/api/photos/daily/export-word?project=Site&date=2026-09-18').status_code == 409
    path = add_photo(tmp_path, 'Site', '2026-09-19', 'broken.png', 'morning')
    path.write_bytes(b'broken image')
    response = client.get('/api/photos/daily/export-word?project=Site&date=2026-09-19')
    assert response.status_code == 400
    assert 'broken.png' in response.json['error']


def test_many_photos_grid_order_ratio_and_row_pagination(tmp_path):
    for i in range(31):
        add_photo(tmp_path, 'Site', '2026-09-18', f'{i:02}.png', 'morning',
                  size=(160, 80) if i % 2 else (80, 160))
    add_photo(tmp_path, 'Site', '2026-09-18', 'pm.webp', 'afternoon')
    response = make_client(tmp_path).get('/api/photos/daily/export-word?project=Site&date=2026-09-18')
    doc = Document(io.BytesIO(response.data))
    assert len(doc.inline_shapes) == 16
    assert [len(t.rows) for t in doc.tables] == [3, 1]
    assert all(len(r.cells) == 5 for r in doc.tables[0].rows)
    assert all(c.text == '' for t in doc.tables for r in t.rows for c in r.cells)
    assert all(len(c.paragraphs) == 1 for t in doc.tables for r in t.rows for c in r.cells)
    assert all(not p.paragraph_format.keep_with_next
               for t in doc.tables for r in t.rows for c in r.cells for p in c.paragraphs)
    assert doc.sections[0].page_width.cm == pytest.approx(21, abs=.01)
    assert doc.sections[0].page_height.cm == pytest.approx(29.7, abs=.01)
    for i, shape in enumerate(list(doc.inline_shapes)[:15]):
        assert shape.width / shape.height == pytest.approx(2 if i % 2 else .5, rel=.01)
    headings = [p for p in doc.paragraphs if p.style.name == 'Heading 1']
    assert not headings[0].paragraph_format.page_break_before
    assert headings[1].paragraph_format.page_break_before
    assert 'Ca chiều' in headings[1].text
    with zipfile.ZipFile(io.BytesIO(response.data)) as archive:
        xml = archive.read('word/document.xml').decode()
    assert xml.count('<w:cantSplit') == 4
    assert '<w:keepNext w:val="0"' in xml
    assert '<w:left w:w="40" w:type="dxa"' in xml
    assert doc.inline_shapes[1].width.cm == pytest.approx(3.45, abs=.01)


@pytest.mark.parametrize('period,day_count', [('2026-09', 30), ('2026-10', 31), ('2024-02', 29), ('2026-02', 28)])
def test_monthly_export_all_calendar_days_and_project_scope(tmp_path, period, day_count):
    add_photo(tmp_path, 'Site', f'{period}-{day_count}', 'last.png', 'afternoon')
    for i in range(17):
        add_photo(tmp_path, 'Site', f'{period}-01', f'{i:02}.png', 'morning')
    add_photo(tmp_path, 'Other', f'{period}-02', 'other.png', 'morning')
    add_photo(tmp_path, 'Site', '2025-01-01', 'other-month.png', 'morning')
    response = make_client(tmp_path).get(f'/api/photos/daily/export-month-word?project=Site&period={period}')
    assert response.status_code == 200
    assert response.mimetype == 'application/zip'
    assert f'Anh_Site_{period}.zip' in unquote(response.headers['Content-Disposition'])
    with zipfile.ZipFile(io.BytesIO(response.data)) as archive:
        assert archive.testzip() is None
        assert archive.namelist() == [f'{period}/Anh_Site_{period}-{day:02}.docx' for day in range(1, day_count + 1)]
        docs = [Document(io.BytesIO(archive.read(name))) for name in archive.namelist()]
    text = '\n'.join(report_text(doc) for doc in docs)
    dates = [p.text for doc in docs for p in doc.paragraphs if p.text.startswith('Ngày:')]
    assert len(dates) == day_count
    assert dates[0].startswith(f'Ngày: 01/{period[5:]}/{period[:4]}')
    assert dates[-1].startswith(f'Ngày: {day_count}/{period[5:]}/{period[:4]}')
    assert sum(len(doc.inline_shapes) for doc in docs) == 16  # Existing per-day 15-photo shift cap.
    assert text.count('Không có ảnh') == day_count * 2 - 2
    assert all(len([p for p in doc.paragraphs if p.text.startswith('Ngày:')]) == 1 for doc in docs)
    assert len(docs[0].inline_shapes) == 15
    assert len(docs[-1].inline_shapes) == 1


@pytest.mark.parametrize('period', ['', '2026', '2026-13', '2026-00', '2026-9', '../bad'])
def test_monthly_export_rejects_invalid_period(tmp_path, period):
    response = make_client(tmp_path).get('/api/photos/daily/export-month-word',
                                          query_string={'project': 'Site', 'period': period})
    assert response.status_code == 400


def test_monthly_export_empty_ambiguous_and_broken_images(tmp_path):
    client = make_client(tmp_path)
    url = '/api/photos/daily/export-month-word?project=Site&period=2026-09'
    assert client.get(url).status_code == 400
    legacy = add_photo(tmp_path, 'Site', '18', 'legacy.png', 'morning')
    assert client.get(url).status_code == 409
    legacy.rename(legacy.with_suffix('.txt'))
    # Mapping an old folder explicitly resolves its calendar identity.
    (tmp_path / 'Site' / 'attendance-period.json').write_text(json.dumps({
        'version': 1, 'mappings': {'18': {'date': '2026-09-18',
        'confirmed_by': 'test', 'confirmed_at': '2026-09-18T00:00:00Z'}}}))
    broken = add_photo(tmp_path, 'Site', '2026-09-01', 'broken.png', 'morning')
    broken.write_bytes(b'broken')
    response = client.get(url)
    assert response.status_code == 400
    assert 'broken.png' in response.json['error']


def test_monthly_export_project_id_and_mapped_legacy_day(tmp_path):
    add_photo(tmp_path, 'storage', '18', 'legacy.png', 'morning')
    (tmp_path / 'storage' / 'attendance-period.json').write_text(json.dumps({
        'version': 1, 'mappings': {'18': {'date': '2026-09-18',
        'confirmed_by': 'test', 'confirmed_at': '2026-09-18T00:00:00Z'}}}))
    resolver = lambda ref: {'storage_dir': 'storage', 'display_name': 'Tên mới', 'active': True}
    response = make_client(tmp_path, resolver).get(
        '/api/photos/daily/export-month-word?project_id=pid&period=2026-09')
    assert response.status_code == 200
    assert 'Anh_Tên mới_2026-09.zip' in unquote(response.headers['Content-Disposition'])
    with zipfile.ZipFile(io.BytesIO(response.data)) as archive:
        doc = Document(io.BytesIO(archive.read('2026-09/Anh_Tên mới_2026-09-18.docx')))
    assert 'Dự án: Tên mới' in report_text(doc)
    assert len(doc.inline_shapes) == 1


def test_monthly_progress_is_live_and_project_scoped(tmp_path, monkeypatch):
    import src.daily_photo_word_exporter as exporter
    add_photo(tmp_path, 'Site', '2026-09-01', 'a.png', 'morning')
    client = make_client(tmp_path)
    status_url = '/api/photos/daily/export-progress?project=Site&export_id=progress-test'
    assert client.get(status_url).status_code == 404
    original = exporter.export_monthly_photo_word
    def observed(project, period, days, progress_callback=None):
        progress_callback(4, len(days))
        state = client.get(status_url).json
        assert state['status'] == 'running' and state['completed_days'] == 4
        assert client.get(status_url.replace('project=Site', 'project=Other')).status_code == 404
        return original(project, period, days, progress_callback)
    monkeypatch.setattr(exporter, 'export_monthly_photo_word', observed)
    url = '/api/photos/daily/export-month-word?project=Site&period=2026-09&export_id=progress-test'
    assert client.get(url).status_code == 200
    state = client.get(status_url).json
    assert state['status'] == 'completed'
    assert state['completed_days'] == state['total_days'] == 30
    assert client.get(url).status_code == 400  # No duplicate heavy job for the same export ID.


def test_monthly_progress_reports_broken_image_error(tmp_path):
    path = add_photo(tmp_path, 'Site', '2026-09-01', 'broken.png', 'morning')
    path.write_bytes(b'broken')
    client = make_client(tmp_path)
    url = '/api/photos/daily/export-month-word?project=Site&period=2026-09&export_id=failed-test'
    assert client.get(url).status_code == 400
    state = client.get('/api/photos/daily/export-progress?project=Site&export_id=failed-test').json
    assert state['status'] == 'failed'
    assert 'broken.png' in state['error']


def test_daily_image_index_reuses_media_without_rehashing_prior_photos(tmp_path, monkeypatch):
    from datetime import date
    from docx.parts.image import ImagePart
    from src.daily_photo_word_exporter import export_daily_photo_word
    first = add_photo(tmp_path, 'Site', '2026-09-01', 'a.png', 'morning', size=(1600, 2400))
    second = add_photo(tmp_path, 'Site', '2026-09-02', 'b.png', 'morning', size=(1600, 2400))
    original_hash = ImagePart.sha1.fget
    hash_counts = {}
    def counted_hash(self):
        hash_counts[id(self)] = hash_counts.get(id(self), 0) + 1
        return original_hash(self)
    monkeypatch.setattr(ImagePart, 'sha1', property(counted_hash))
    photos = [{'path': path, 'shift': 'morning', 'send_time': '09:00:00'} for path in [first, second, first]]
    output, _ = export_daily_photo_word('Site', date(2026, 9, 1), photos)
    assert len(hash_counts) == 2
    assert all(count == 1 for count in hash_counts.values()), 'Only hash each image once, independent of later insertions'
    doc = Document(output)
    assert len(doc.inline_shapes) == 3
    output.seek(0)
    with zipfile.ZipFile(output) as archive:
        media = [name for name in archive.namelist() if name.startswith('word/media/')]
        assert len(media) == 2
        for name in media:
            with Image.open(io.BytesIO(archive.read(name))) as image:
                assert image.width <= 450 and image.height <= 600
