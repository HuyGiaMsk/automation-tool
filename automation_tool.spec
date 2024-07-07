# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['src/gui/GUIApp.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('input', 'input'),
        ('output', 'output'),
        ('resource', 'resource'),
        ('script', 'script'),
        ('src', 'src'),
        ('test', 'test')
    ],
    hiddenimports=[
        'selenium.webdriver.chrome',
        'selenium.webdriver.support.expected_conditions',
        'selenium.webdriver.support.wait',
        'requests',
        'wget'
    ],
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
    name='automation_tool',
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
