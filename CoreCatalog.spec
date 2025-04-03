# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['app\\main.py'],
    pathex=[],
    binaries=[],
    datas=[('C:\\Users\\156\\Desktop\\LOMO_PRJ\\LOMONOSOV\\resources\\my_icon.ico', 'resources'), ('C:\\Users\\156\\Desktop\\LOMO_PRJ\\LOMONOSOV\\resources\\scale.jpg', 'resources'), ('C:\\Users\\156\\Desktop\\LOMO_PRJ\\LOMONOSOV\\resources\\shkala.jpg', 'resources'), ('C:\\Users\\156\\Desktop\\LOMO_PRJ\\LOMONOSOV\\resources\\arial.ttf', 'resources')],
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
    name='CoreCatalog',
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
    icon=['C:\\Users\\156\\Desktop\\LOMO_PRJ\\LOMONOSOV\\resources\\my_icon.ico'],
)
