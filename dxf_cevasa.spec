# -*- mode: python ; coding: utf-8 -*-

import os
from PyInstaller.utils.hooks import collect_submodules

project_name = "DXF-CEVASA"
icon_file = "icon.ico"

# Arquivos de dados (ex: templates, ícones)
datas = [
    ('resources/excel/Planilha_template.xlsx', 'resources/excel'),
    ('icon.ico', '.'),
    ('last_desenhista.txt', '.')
]

# Inclui todos os submódulos
hiddenimports = collect_submodules("ui") + collect_submodules("dxf")

a = Analysis(
    ['main.py'],
    pathex=[os.path.abspath('.')],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=None,
    noarchive=False
)

pyz = PYZ(a.pure, a.zipped_data, cipher=None)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name=project_name,
    icon=icon_file,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False  # True se quiser console; False para só GUI
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name=project_name
)
