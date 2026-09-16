"""Cache regression tests: subprocesses bound deadlock failures."""
import os
import pickle
import subprocess
import sys
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from threading import Barrier
from types import SimpleNamespace
from unittest.mock import patch

import numpy as np
import pytest
import src.face_matcher as fm

@pytest.mark.parametrize("dirty", [-1, 19])
def test_embedding_cache_does_not_deadlock(tmp_path, dirty):
    code = """
import sys, pickle
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch
import src.face_matcher as fm
root = Path(sys.argv[1])
fm._CACHE_FILE = str(root / 'cache.pkl')
dirty = int(sys.argv[2])
fm._DISK_CACHE = None if dirty == -1 else {}
fm._CACHE_DIRTY_COUNT = max(0, dirty)
p = root / 'photo.jpg'
p.write_bytes(b'image')
matcher = fm.FaceMatcher(str(root / 'portraits'))
def represent(**kwargs):
    assert fm._CACHE_LOCK.acquire(blocking=False), 'DeepFace ran under cache lock'
    fm._CACHE_LOCK.release()
    return [{'embedding': [1., 0.]}]
with patch.object(fm, 'get_deepface', return_value=SimpleNamespace(represent=represent)), patch.object(fm, 'copy_to_ascii_path', side_effect=lambda p:p):
    assert matcher._get_embedding(str(p)).tolist() == [1., 0.]
    assert matcher._get_embedding(str(p)).tolist() == [1., 0.]
if dirty == 19:
    assert len(pickle.loads(Path(fm._CACHE_FILE).read_bytes())) == 1
    assert fm._CACHE_DIRTY_COUNT == 0
print('cache-ok')
"""
    result = subprocess.run(
        [sys.executable, "-c", code, str(tmp_path), str(dirty)],
        cwd=Path(__file__).resolve().parents[1], capture_output=True,
        text=True, timeout=5, encoding="utf-8", errors="replace",
    )
    assert result.returncode == 0, result.stderr
    assert "cache-ok" in result.stdout

@pytest.fixture
def cache(tmp_path, monkeypatch):
    monkeypatch.setattr(fm, "_CACHE_FILE", str(tmp_path / "cache.pkl"))
    monkeypatch.setattr(fm, "_DISK_CACHE", {})
    monkeypatch.setattr(fm, "_CACHE_DIRTY_COUNT", 0)
    return tmp_path

def test_concurrent_same_image_does_not_double_count(cache, monkeypatch):
    p = cache / "image.jpg"
    p.write_bytes(b"test")
    barrier = Barrier(4)
    def represent(**kwargs):
        barrier.wait(timeout=3)
        return [{"embedding": [1., 0.]}]
    monkeypatch.setattr(fm, "get_deepface", lambda: SimpleNamespace(represent=represent))
    monkeypatch.setattr(fm, "copy_to_ascii_path", lambda p: p)
    matcher = fm.FaceMatcher(str(cache / "portraits"))
    with ThreadPoolExecutor(max_workers=4) as pool:
        values = list(pool.map(matcher._get_embedding, [str(p)] * 4))
    assert all(v.tolist() == [1., 0.] for v in values)
    assert fm._CACHE_DIRTY_COUNT == 1

def test_failed_replace_preserves_disk_and_dirty_state(cache, monkeypatch):
    p = Path(fm._CACHE_FILE)
    p.write_bytes(pickle.dumps({"old": [1]}))
    fm._DISK_CACHE["new"] = np.array([2.])
    fm._CACHE_DIRTY_COUNT = 1
    def fail(*args):
        raise OSError("disk failure")
    monkeypatch.setattr(fm.os, "replace", fail)
    fm.save_disk_cache(force=True)
    assert pickle.loads(p.read_bytes()) == {"old": [1]}
    assert fm._CACHE_DIRTY_COUNT == 1

def test_corrupt_cache_can_be_rebuilt(cache, monkeypatch):
    Path(fm._CACHE_FILE).write_bytes(b"broken")
    monkeypatch.setattr(fm, "_DISK_CACHE", None)
    assert fm.load_disk_cache() == {}
    fm._DISK_CACHE["valid"] = [1]
    fm._CACHE_DIRTY_COUNT = 1
    fm.save_disk_cache()
    assert pickle.loads(Path(fm._CACHE_FILE).read_bytes()) == {"valid": [1]}
