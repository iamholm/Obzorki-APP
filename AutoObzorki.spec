# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['app_uii.py'],
    pathex=[],
    binaries=[],
    datas=[('Test.docx', '.'), ('autoobzorki.ico', '.')],
    hiddenimports=[],
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
    name='\u0410\u0432\u0442\u043e\u041e\u0431\u0437\u043e\u0440\u043a\u0438',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='autoobzorki.ico',
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='\u0410\u0432\u0442\u043e\u041e\u0431\u0437\u043e\u0440\u043a\u0438',
)
