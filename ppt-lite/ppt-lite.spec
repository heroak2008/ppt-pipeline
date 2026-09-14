# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec：ppt-lite onedir 打包。

用法：uv run pyinstaller ppt-lite.spec --clean --noconfirm
产物：dist/ppt-lite/ppt-lite.exe（onedir 模式，启动快、便于排查）
"""
from pathlib import Path

block_cipher = None
ROOT = Path.cwd()                       # spec 在仓库根运行

a = Analysis(
    ["ppt_lite/runner.py"],
    pathex=[str(ROOT)],
    binaries=[],
    datas=[
        ("ppt_lite/templates", "ppt_lite/templates"),   # Jinja2 模板（冻结模式经 _MEIPASS 解析）
    ],
    hiddenimports=[
        "pptx", "pptx.util", "pptx.enum.shapes", "pptx.dml.color",
        "lxml", "olefile", "pypdfium2", "imagehash", "PIL", "PIL.Image",
        "uvicorn", "uvicorn.logging", "uvicorn.loops", "uvicorn.protocols",
        "uvicorn.protocols.http", "uvicorn.protocols.http.h11_impl",
        "uvicorn.lifespan", "uvicorn.lifespan.on",
        "jinja2", "multipart",  # python-multipart 提供 multipart
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["tkinter", "pytest", "httpx", "docx", "xlsxwriter", "scipy"],
    # 注意：numpy 不能排除——imagehash（phash 去重）依赖它
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="ppt-lite",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=True,               # 控制台窗口：显示启动横幅与日志
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
    upx=False,
    upx_exclude=[],
    name="ppt-lite",
)
