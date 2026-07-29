# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['F:\\codex\\New project\\app.py'],
    pathex=[],
    binaries=[],
    datas=[('F:\\codex\\New project\\assets', 'assets'), ('F:\\codex\\New project\\installers_manifest.json', '.'), ('F:\\codex\\New project\\version.json', '.'), ('F:\\codex\\New project\\third_party\\open-shell\\portable\\PFiles\\Open-Shell', 'third_party\\open-shell\\PFiles\\Open-Shell'), ('F:\\codex\\New project\\third_party\\open-shell\\LICENSE.txt', 'third_party\\open-shell')],
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
    name='WGA',
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
    icon=['F:\\codex\\New project\\assets\\wga-icon.ico'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='WGA',
)
