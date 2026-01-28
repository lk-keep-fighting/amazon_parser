# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec file for Amazon Product Parser GUI.

Usage:
    pyinstaller amazon_parser.spec
"""

import os
from pathlib import Path

# 项目根目录
project_root = Path(SPECPATH).parent

# 数据文件 - 需要 Playwright 的浏览器驱动
# 注意：用户仍需运行 'playwright install' 来下载浏览器
datas = []

block_cipher = None

a = Analysis(
    [os.path.join(project_root, 'src', 'gui.py')],
    pathex=[str(project_root)],
    binaries=[],
    datas=datas,
    hiddenimports=[
        'playwright',
        'playwright.sync_api',
        'greenlet',
        'pyee',
        'certifi',
        'charset_normalizer',
        'idna',
        'urllib3',
        'openpyxl',
        'et_xmlfile',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='AmazonParser',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # GUI 应用，不显示控制台窗口
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # 可以添加 .ico 文件路径
    version_file=None,
)

# 如果需要创建目录版本（更小的启动文件，包含更多依赖文件）
# coll = COLLECT(
#     exe,
#     a.binaries,
#     a.zipfiles,
#     a.datas,
#     strip=False,
#     upx=True,
#     upx_exclude=[],
#     name='AmazonParser',
# )
