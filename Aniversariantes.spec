# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['script.py'],
    pathex=[],
    binaries=[],
    datas=[('modelos', 'modelos'), ('fontes', 'fontes'), ('outputs', 'outputs'), ('planilhas/*.xlsx', 'planilhas'), ('layouts', 'layouts'), ('layout_config.py', '.'), ('preferencias.py', '.'), ('servidor_licenca.py', '.'), ('version.py', '.')],
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
    name='Aniversariantes',
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
    icon='NONE',
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Aniversariantes',
)
