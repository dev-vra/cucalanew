# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

# Dados da aplicação
app_name = 'Consolidador'
app_version = '1.0.0'
bundle_identifier = 'com.cucala.consolidador'

# Obter a arquitetura de destino do ambiente
import os
target_arch = os.environ.get('TARGET_ARCH', None)

# Configurações de análise
a = Analysis(
    ['consolidador.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('assets/*', 'assets'),
        ('data/*.json', 'data'),
    ],
    hiddenimports=[
        'pandas', 'openpyxl', 'customtkinter', 'PIL',
        'tkinter', 'queue', 'json', 'pathlib', 'configparser'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# Configurações do executável
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name=app_name,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=True,
    target_arch=target_arch,
    codesign_identity=None,
    entitlements_file=None,
    icon='assets/icon.icns',
)

# Configurações de coleta
app = BUNDLE(
    exe,
    name=f'{app_name}.app',
    icon='assets/icon.icns',
    bundle_identifier=bundle_identifier,
    info_plist={
        'CFBundleName': app_name,
        'CFBundleDisplayName': app_name,
        'CFBundleVersion': app_version,
        'CFBundleShortVersionString': app_version,
        'CFBundleExecutable': app_name,
        'CFBundleIdentifier': bundle_identifier,
        'NSHighResolutionCapable': 'True',
        'NSRequiresAquaSystemAppearance': 'False',
        'LSMinimumSystemVersion': '10.15.0',
    },
)