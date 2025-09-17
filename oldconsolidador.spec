# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['appconsolidador.py'],  # Nome do seu script principal
    pathex=[],
    binaries=[],
    datas=[
        ('assets', 'assets')  # Correto: Inclui a pasta de assets
    ],
    hiddenimports=[
        'pandas',
        'openpyxl',
        'customtkinter',
        'tkinter',
        'PIL',
        'dateutil.parser'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    cipher=block_cipher
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='Consolidador CUCALA',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  # Correto: Não abre um terminal
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Consolidador CUCALA'
)

# Seção que cria o .app para macOS
app = BUNDLE(
    coll,
    name='Consolidador CUCALA.app',
    icon='assets/icon.icns',  # Caminho para o ícone no formato .icns
    bundle_identifier='com.cucala.consolidador'  # Identificador único do app
)