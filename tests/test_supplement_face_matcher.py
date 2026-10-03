from datetime import date
from pathlib import Path
from types import SimpleNamespace
import os

import pytest
import src.face_matcher as fm


@pytest.fixture
def setup(tmp_path, monkeypatch):
    portrait = tmp_path / 'portrait.jpg'
    camera = tmp_path / 'camera.jpg'
    portrait.write_bytes(b'portrait')
    camera.write_bytes(b'camera00')
    scope = ('project', 'employee', date(2026, 10, 4))
    registry = SimpleNamespace(
        is_member=lambda *args: args == scope,
        portrait_paths=lambda *args: [portrait] if args == scope else [])
    matcher = fm.FaceMatcher(str(tmp_path), identity_registry=registry, enforce_detection=False)
    responses = {b'portrait': [[1., 0.]], b'camera00': [[0., 1.], [1., 0.]]}
    calls = []

    def represent(**kwargs):
        calls.append(kwargs)
        assert kwargs['enforce_detection'] is True
        value = responses[Path(kwargs['img_path']).read_bytes()]
        if isinstance(value, Exception):
            raise value
        return [{'embedding': vector} for vector in value]

    monkeypatch.setattr(fm, 'get_deepface', lambda: SimpleNamespace(represent=represent))
    def match(**kwargs):
        return matcher.match_employee_in_photo(
            project_id=kwargs.pop('project_id', scope[0]),
            employee_id=kwargs.pop('employee_id', scope[1]),
            attendance_date=kwargs.pop('attendance_date', scope[2]),
            image_path=str(camera), **kwargs)
    return SimpleNamespace(**locals())


def test_second_face_matches_and_all_faces_are_cached(setup):
    s = setup
    result = s.match()
    assert result.status == 'matched'
    assert result.image_path == str(s.camera)
    assert result.distance == pytest.approx(0, abs=1e-6)
    assert result.reason
    assert s.match().status == 'matched'
    assert len(s.calls) == 2


def test_no_match_records_best_distance(setup):
    setup.responses[b'camera00'] = [[0., 1.]]
    result = setup.match()
    assert result.status == 'no_match'
    assert result.distance == pytest.approx(1.)
    assert result.reason


def test_empty_detection_is_no_face(setup):
    setup.responses[b'camera00'] = []
    assert setup.match().status == 'no_face'


@pytest.mark.parametrize('target', [b'camera00', b'portrait'])
def test_model_exception_is_error(setup, target):
    setup.responses[target] = RuntimeError('model failed')
    result = setup.match()
    assert result.status == 'error'
    assert 'model failed' in result.reason


def test_unavailable_deepface_is_error(setup, monkeypatch):
    monkeypatch.setattr(fm, 'get_deepface', lambda: None)
    assert setup.match().status == 'error'


@pytest.mark.parametrize('threshold', [0, -1, float('nan'), float('inf'), -float('inf')])
def test_invalid_threshold(setup, threshold):
    with pytest.raises(ValueError):
        setup.match(distance_threshold=threshold)
    assert not setup.calls


@pytest.mark.parametrize('target', ['camera', 'portrait'])
def test_same_size_same_mtime_content_invalidates_cache(setup, target):
    s = setup
    assert s.match().status == 'matched'
    path = getattr(s, target)
    stat = path.stat()
    replacement = b'changed0'
    path.write_bytes(replacement)
    os.utime(path, ns=(stat.st_atime_ns, stat.st_mtime_ns))
    s.responses[replacement] = [[-1., 0.]]
    assert s.match().status == 'no_match'
    assert len(s.calls) == 3


@pytest.mark.parametrize('attribute,value', [('model_name', 'Facenet512'), ('detector_backend', 'opencv')])
def test_model_and_backend_invalidate_cache(setup, attribute, value):
    assert setup.match().status == 'matched'
    setattr(setup.matcher, attribute, value)
    setup.responses[b'camera00'] = [[0., 1.]]
    assert setup.match().status == 'no_match'
    assert len(setup.calls) == 4


def test_no_confirmed_portrait(setup):
    setup.registry.portrait_paths = lambda *args: []
    assert setup.match().status == 'no_portrait'
    assert not setup.calls


@pytest.mark.parametrize('faces', [[], [[1., 0.], [0., 1.]]])
def test_portrait_requires_one_detected_face(setup, faces):
    setup.responses[b'portrait'] = faces
    assert setup.match().status == 'no_portrait'


@pytest.mark.parametrize('override', [dict(project_id='other'), dict(employee_id='other'), dict(attendance_date=date(2025, 1, 1))])
def test_wrong_membership_is_unresolved(setup, override):
    assert setup.match(**override).status == 'identity_unresolved'
    assert not setup.calls


def test_metric_and_threshold_are_used(setup):
    setup.matcher.distance_metric = 'euclidean'
    setup.responses[b'camera00'] = [[1., .5]]
    assert setup.match().status == 'no_match'
    result = setup.match(distance_threshold=.6)
    assert result.status == 'matched'
    assert result.distance == pytest.approx(.5)


@pytest.mark.parametrize('vector', [[0., 0.], [float('nan'), 1.]])
def test_invalid_embeddings_cannot_match(setup, vector):
    setup.responses[b'camera00'] = [vector]
    assert setup.match().status == 'error'


def test_detection_skip_backend_cannot_approve(setup):
    setup.matcher.detector_backend = 'skip'
    assert setup.match().status == 'error'
    assert not setup.calls
