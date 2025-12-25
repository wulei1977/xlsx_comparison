# -*- mode: python ; coding: utf-8 -*-
import sys
import os

# 获取 conda 环境的 DLL 路径
conda_prefix = os.path.dirname(sys.executable)
if '.venv' in conda_prefix:
    # 虚拟环境基于 conda，找到 base conda 路径
    conda_base = os.path.join(os.environ.get('CONDA_PREFIX_1', r'C:\Users\wulei\anaconda3\envs\py312'), 'Library', 'bin')
else:
    conda_base = os.path.join(conda_prefix, 'Library', 'bin')

# 需要的 DLL 文件
binaries_list = []
dll_names = ['ffi.dll', 'libexpat.dll', 'liblzma.dll', 'sqlite3.dll']
for dll in dll_names:
    dll_path = os.path.join(conda_base, dll)
    if os.path.exists(dll_path):
        binaries_list.append((dll_path, '.'))

a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=binaries_list,
    datas=[('templates', 'templates')],
    hiddenimports=[
        'pandas',
        'python_calamine',
        'flask',
        'jinja2',
        'werkzeug',
        'click',
        'ctypes',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='ExcelCompare',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='app.ico',
)
