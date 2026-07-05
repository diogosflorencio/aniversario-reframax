# Build: .\build.ps1  ou  py -m PyInstaller Aniversariantes.spec --noconfirm --clean
# Pastas modelos/, fontes/, layouts/, planilhas/ e outputs/ ficam na raiz (ao lado do .exe).

from PyInstaller.utils.hooks import collect_submodules

block_cipher = None

hiddenimports = collect_submodules("PyQt6") + [
    "layout_config",
    "layout_editor",
    "preferencias",
    "servidor_licenca",
    "version",
    "win32print",
    "win32ui",
    "win32con",
    "win32api",
]

a = Analysis(
    ["script.py"],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="Aniversariantes",
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
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="Aniversariantes",
)
