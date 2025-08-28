# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['controle_cupons_sanar.py'],
    pathex=[],
    binaries=[],
    datas=[('logo_sanar.png', '.'), ('cupons_sanar.xlsx', '.'), ('lojas_cadastradas.csv', '.'), ('industrias_cadastradas.csv', '.')],
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
    a.binaries,
    a.datas,
    [],
    name='controle_cupons_sanar',
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
)
app = BUNDLE(
    exe,
    name='controle_cupons_sanar.app',
    icon=None,
    bundle_identifier=None,
)
