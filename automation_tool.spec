# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['src/gui/GUIApp.py'],
    pathex=[],
    binaries=[],
    datas=[
        (".\\venv\\Lib\\site-packages\\autoit\\lib\\AutoItX3_x64.dll", "autoit\\lib"),
        ('resource', 'resource'),
        ('src', 'src')
    ],
    hiddenimports=[
        'selenium.webdriver.chrome',
        'selenium.webdriver.support.expected_conditions',
        'selenium.webdriver.support.wait',
        'requests',
        'wget',
        'xlwings',
        'autoit',
        'pdfplumber',
        'PyPDF2',
        "autoit.init",
        "autoit.autoit",
        "autoit.control",
        "autoit.process",
        "autoit.win",
        "pyautogui",
        "pywinauto",
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

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='automation_tool',
)