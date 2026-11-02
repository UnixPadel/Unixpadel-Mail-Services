# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

datas, binaries, hiddenimports = collect_all('customtkinter')

hiddenimports += ['pymysql', 'updater', 'pandas', 'openpyxl', 'urllib.request', 'socket']

# dns.resolver dinamik import edildiği için PyInstaller onu kendiliğinden bulamaz;
# eksik olursa MX doğrulaması ve DNS sağlık testi yalnızca exe'de sessizce kapanır.
hiddenimports += ['dns', 'dns.resolver', 'dns.rdatatype', 'dns.exception']

a = Analysis(
    ['UnixPadel Mail Services.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
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
    name='UnixPadel Mail Services',
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

