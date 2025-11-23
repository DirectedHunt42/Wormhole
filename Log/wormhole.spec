# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['E:\\Projects\\Programming\\Wormhole\\wormhole.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['customtkinter', 'tkinter', 'PIL', 'PIL.ImageTk', 'reportlab', 'pypdf', 'py7zr', 'docx', 'bs4', 'pptx', 'openpyxl', 'ezodf', 'odf', 'odf.opendocument', 'odf.text', 'striprtf', 'pydub', 'moviepy', 'moviepy.editor', 'imageio', 'imageio_ffmpeg', 'decorator', 'proglog', 'urllib.request', 'win32event', 'win32api', 'winerror'],
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
    a.binaries,
    a.datas,
    [],
    name='wormhole',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['E:\\Projects\\Programming\\Wormhole\\Icons\\Wormhole_Icon.ico'],
)
