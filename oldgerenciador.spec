# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['gerenciador.py'],  # Nome do seu script principal
    pathex=[],
    binaries=[],
    datas=[
        ('assets', 'assets')  # Inclui a pasta de assets (se tiver ícone/imagens)
    ],
    hiddenimports=[
        'pandas',
        'openpyxl',
        'customtkinter',
        'tkinter',
        'PIL'
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
    name='Gerenciador CUCALA',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False, # False para não abrir um terminal
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='assets/icon.icns' # Caminho para o ícone
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Gerenciador CUCALA'
)

# Seção que cria o .app para macOS
app = BUNDLE(
    coll,
    name='Gerenciador CUCALA.app',
    icon='assets/icon.icns',
    bundle_identifier='com.cucala.gerenciador' # Identificador único do app
)