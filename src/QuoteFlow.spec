# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['src/main.py'],
    pathex=['src'],
    binaries=[],
    datas=[('data/*.xlsx', 'data'), ('prices.db', '.')],
    hiddenimports=['ui.quotation_ui', 'database.price_loader', 'database.getsql', 'database.handlers.default_handler', 'database.handlers.header_handler', 'database.handlers.other_handler', 'excel_exporter'],
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
    name='QuoteFlow',
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
    name='QuoteFlow.app',
    icon=None,
    bundle_identifier=None,
)
