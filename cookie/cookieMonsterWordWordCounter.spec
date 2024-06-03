# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['cookieMonsterWordWordCounter.py'],
    pathex=[],
    binaries=[],
    datas=[('Data/Cookie Monster Image.png', 'Data'), ('Data/noun-stop-button-4906815-FFFFFF.png', 'Data'), ('Data/noun-play-button-6441783-FFFFFF.png', 'Data'), ("Data/Count's Laugh 1.mp3", 'Data'), ('Data/Documents Completed-[AudioTrimmer.com]-[AudioTrimmer.com].mp3', 'Data'), ('Data/Monster.ico', 'Data'), ('Data/Cookie Instructions.mp3', 'Data')],
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
    name='cookieMonsterWordWordCounter',
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
    icon=['Data\\Monster.ico'],
)
