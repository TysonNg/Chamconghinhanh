from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
import os

from PIL import Image
import pytest

from src.face_inference import usable_face_representations
from src.face_matcher import copy_to_ascii_path


@pytest.mark.parametrize('box,confidence,accepted', [
    ((100, 100, 100, 100), .99, True),
    ((20, 250, 55, 65), .99, False),  # avatar on a tall chat screenshot
    ((400, 500, 25, 30), .99, False),  # face on a photographed ID card
    ((100, 100, 100, 100), .5, False),
    ((-1, 100, 100, 100), .99, False),
    ((800, 1800, 100, 100), .99, False),
])
def test_only_large_confident_faces_can_supply_attendance_vectors(tmp_path, box, confidence, accepted):
    path = tmp_path / 'photo.png'
    Image.new('RGB', (869, 1884)).save(path)
    value = {'embedding': [1., 0.], 'face_confidence': confidence,
             'facial_area': dict(zip(('x', 'y', 'w', 'h'), box))}
    assert usable_face_representations([value], path) == ([value] if accepted else [])


def test_missing_detection_metadata_cannot_confirm_a_face(tmp_path):
    path = tmp_path / 'photo.png'
    Image.new('RGB', (200, 200)).save(path)
    assert usable_face_representations([{'embedding': [1., 0.]}], path) == []


def test_snapshot_uses_contents_even_if_size_and_mtime_are_unchanged(tmp_path):
    path = tmp_path / 'photo.jpg'
    path.write_bytes(b'first')
    stat = path.stat()
    original = copy_to_ascii_path(str(path))
    path.write_bytes(b'other')
    os.utime(path, ns=(stat.st_atime_ns, stat.st_mtime_ns))
    with ThreadPoolExecutor(max_workers=4) as pool:
        snapshots = list(pool.map(copy_to_ascii_path, [str(path)] * 4))
    assert len(set(snapshots)) == 1
    assert snapshots[0] != original
    assert Path(original).read_bytes() == b'first'
    assert Path(snapshots[0]).read_bytes() == b'other'
