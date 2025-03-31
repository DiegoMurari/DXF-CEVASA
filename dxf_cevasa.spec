# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['main.py'],  # Nome do arquivo principal
    pathex=['.'],
    binaries=[],
    datas=[
    ('resources/excel/Planilha_template.xlsx', 'resources/excel'),
    ('icon.ico', '.'),
    ('create_shortcut.py', '.'),  # <- aqui
    ],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='DXF-CEVASA',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # ← Isto oculta o terminal
    icon='icon.ico'  # ← Ícone do executável
)
