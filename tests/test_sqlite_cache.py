# -*- coding: utf-8 -*-
"""
Tests for SQLite Face Embeddings Cache & Migration
"""
import os
import pickle
import numpy as np
import pytest
from pathlib import Path

import src.face_matcher as fm


def test_sqlite_cache_roundtrip(tmp_path, monkeypatch):
    test_db = str(tmp_path / "test_embeddings.sqlite3")
    monkeypatch.setattr(fm, "_CACHE_DB_FILE", test_db)
    monkeypatch.setattr(fm, "_DISK_CACHE", None)
    monkeypatch.setattr(fm, "_CACHE_DIRTY_COUNT", 0)

    # 1. Load empty cache
    cache = fm.load_disk_cache()
    assert cache == {}

    # 2. Insert embedding into in-memory and save
    fake_vec = np.array([0.123, 0.456, 0.789], dtype=np.float32)
    key = ("C:/test/photo.jpg", 1024, 123456.78, "ArcFace", "retinaface")
    cache[key] = fake_vec
    fm._CACHE_DIRTY_COUNT = 1
    fm.save_disk_cache(force=True)

    # 3. Reload cache from SQLite
    monkeypatch.setattr(fm, "_DISK_CACHE", None)
    reloaded = fm.load_disk_cache()
    assert key in reloaded
    np.testing.assert_allclose(reloaded[key], fake_vec, rtol=1e-5)


def test_migrate_from_pickle_to_sqlite(tmp_path, monkeypatch):
    pkl_file = str(tmp_path / "old_cache.pkl")
    sqlite_file = str(tmp_path / "migrated.sqlite3")

    # Create dummy pickle file
    key1 = ("C:/img1.jpg", 500, 1000.0, "ArcFace", "retinaface")
    vec1 = np.array([1.0, 2.0, 3.0], dtype=np.float32)
    with open(pkl_file, "wb") as f:
        pickle.dump({key1: vec1}, f)

    monkeypatch.setattr(fm, "_CACHE_FILE", pkl_file)
    monkeypatch.setattr(fm, "_CACHE_DB_FILE", sqlite_file)
    monkeypatch.setattr(fm, "_DISK_CACHE", None)

    # Load should trigger migration
    cache = fm.load_disk_cache()
    assert key1 in cache
    np.testing.assert_allclose(cache[key1], vec1, rtol=1e-5)

    # pkl should be renamed to .bak
    assert not os.path.exists(pkl_file)
    assert os.path.exists(pkl_file + ".bak")
    assert os.path.exists(sqlite_file)


def test_clear_face_cache(tmp_path, monkeypatch):
    sqlite_file = str(tmp_path / "to_clear.sqlite3")
    monkeypatch.setattr(fm, "_CACHE_DB_FILE", sqlite_file)
    monkeypatch.setattr(fm, "_DISK_CACHE", None)

    cache = fm.load_disk_cache()
    key = ("C:/img.jpg", 100, 200.0, "ArcFace", "retinaface")
    cache[key] = np.array([0.5, 0.5], dtype=np.float32)
    fm._CACHE_DIRTY_COUNT = 1
    fm.save_disk_cache(force=True)

    assert len(fm.load_disk_cache()) == 1
    assert fm.clear_face_cache() is True
    assert len(fm.load_disk_cache()) == 0
