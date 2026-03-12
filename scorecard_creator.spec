# scorecard_creator.spec
# PyInstaller build spec for Scorecard Creator desktop application.
#
# Build command (run from project root):
#   pyinstaller scorecard_creator.spec
#
# Output: dist/ScorecardCreator/ScorecardCreator.exe  (onedir mode)

import glob
import os
from PyInstaller.utils.hooks import collect_data_files

block_cipher = None

# ── Data files to bundle ──────────────────────────────────────────────────────
added_files = [
    ('templates', 'templates'),          # Flask/Jinja2 HTML templates
    ('static', 'static'),                # Bootstrap CSS/JS (offline)
    *collect_data_files('docx2pdf'),     # docx2pdf package data
    *collect_data_files('docx'),         # python-docx schema files
]

# ── Hidden imports PyInstaller misses via static analysis ─────────────────────
hidden_imports = [
    # Flask / Werkzeug
    'flask', 'flask.templating',
    'jinja2', 'jinja2.ext',
    'werkzeug', 'werkzeug.utils', 'werkzeug.routing', 'werkzeug.exceptions',
    # Windows COM / pywin32
    'comtypes', 'comtypes.client', 'comtypes.server',
    'pythoncom', 'pywintypes',
    'win32com', 'win32com.client', 'win32com.server',
    # pywebview (WebView2 / edgechromium backend)
    'webview', 'webview.platforms.edgechromium',
    'clr_loader', 'bottle', 'proxy_tools',
    # python-docx
    'docx', 'docx.oxml', 'docx.oxml.ns', 'docx.parts',
    # PyPDF2
    'PyPDF2',
    # docx2pdf
    'docx2pdf',
    # Standard lib helpers used at runtime
    'importlib.metadata',
    'email.mime.multipart',
    'email.mime.text',
]

# ── pywin32 system DLLs (must be placed in the root of the dist folder) ───────
pywin32_dlls = glob.glob(
    r'C:\Users\andre\AppData\Local\Programs\Python\Python312'
    r'\Lib\site-packages\pywin32_system32\*.dll'
)

a = Analysis(
    ['app_entry.py'],
    pathex=['.'],
    binaries=[(dll, '.') for dll in pywin32_dlls],
    datas=added_files,
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=['pyi_rth_win32.py'],
    excludes=[
        'matplotlib', 'numpy', 'pandas', 'scipy',
        'PIL', 'IPython', 'notebook',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,            # onedir mode (not onefile)
    name='ScorecardCreator',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,                        # DO NOT use UPX - corrupts pywin32 DLLs
    console=False,                    # No console window (windowed app)
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='installer/app_icon.ico',    # Create this before building
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='ScorecardCreator',          # Output: dist/ScorecardCreator/
)
