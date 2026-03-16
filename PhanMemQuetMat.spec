# -*- mode: python ; coding: utf-8 -*-

import os
import sys
import cv2

# Tìm đường dẫn cv2 data (chứa haarcascade XML files)
cv2_data_dir = os.path.join(os.path.dirname(cv2.__file__), 'data')

# Tìm đường dẫn tcl/tk data
python_dir = os.path.dirname(sys.executable)
tcl_dir = os.path.join(python_dir, 'tcl', 'tcl8.6')
tk_dir = os.path.join(python_dir, 'tcl', 'tk8.6')

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('templates', 'templates'),
        ('static', 'static'),
        (cv2_data_dir, os.path.join('cv2', 'data')),
        (tcl_dir, os.path.join('_tcl_data', 'tcl8.6')),
        (tk_dir, os.path.join('_tcl_data', 'tk8.6')),
    ],
    hiddenimports=[
        # Flask & web
        'flask', 'werkzeug', 'werkzeug.utils', 'werkzeug.serving',
        'werkzeug.debug', 'werkzeug.middleware', 'werkzeug.middleware.proxy_fix',
        'jinja2', 'markupsafe', 'itsdangerous', 'click', 'blinker',
        # Document processing
        'docx', 'docx.opc', 'docx.opc.constants', 'docx.oxml',
        'docx.oxml.ns', 'docx.shared', 'docx.enum', 'docx.enum.text',
        'openpyxl',
        # Image processing
        'cv2', 'numpy', 'PIL', 'PIL.Image',
        # Pandas
        'pandas', 'pandas.io.formats.excel',
        # DeepFace & AI
        'deepface', 'deepface.DeepFace', 'deepface.commons',
        'deepface.modules', 'deepface.models',
        'tf_keras', 'keras',
        'tensorflow', 'tensorflow.python', 'tensorflow.lite',
        'h5py', 'scipy', 'scipy.spatial', 'scipy.spatial.distance',
        # Other
        'queue', 'json', 'csv', 'threading', 'concurrent.futures',
        'logging', 'traceback', 'ctypes', 'tempfile', 'shutil',
        'unicodedata', 're', 'urllib.parse', 'xlrd',
        # src modules
        'src', 'src.app', 'src.config', 'src.face_detector',
        'src.face_matcher', 'src.attendance_processor',
        'src.word_exporter', 'src.pdf_extractor',
        'src.text_extractor', 'src.async_processor',
        'src.database_manager', 'src.excel_extractor',
        'src.excel_splitter', 'src.excel_face_analyzer', 'src.excel_list_word_exporter',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='PhanMemQuetMat',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='PhanMemQuetMat',
)
