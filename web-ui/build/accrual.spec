# PyInstaller spec — cross-platform (run on target OS)
# Build: cd build && pyinstaller accrual.spec
import sys
from pathlib import Path
from PyInstaller.utils.hooks import collect_all, collect_submodules, collect_data_files

block_cipher = None

# Project layout: build/ is sibling of backend/ and frontend/
SPEC_DIR = Path(SPECPATH).resolve()
ROOT = SPEC_DIR.parent          # web-ui/
REPO = ROOT.parent              # iTech_Accrual_Updater/
BACKEND = ROOT / "backend"
FRONTEND_DIST = ROOT / "frontend" / "dist"
ACCRUAL_SRC = REPO              # accrual_updater.py + admin_fee_module_v18b.py

# Hidden imports — pywin32 + uvicorn + xlwings + openpyxl
hiddenimports = [
    "uvicorn.logging",
    "uvicorn.loops",
    "uvicorn.loops.auto",
    "uvicorn.protocols",
    "uvicorn.protocols.http",
    "uvicorn.protocols.http.auto",
    "uvicorn.protocols.websockets",
    "uvicorn.protocols.websockets.auto",
    "uvicorn.lifespan",
    "uvicorn.lifespan.on",
    "openpyxl",
    "xlrd",
    "pandas",
    "accrual_updater",
    "admin_fee_module_v18b",
]
hiddenimports += collect_submodules("uvicorn")
hiddenimports += collect_submodules("openpyxl")

if sys.platform == "win32":
    hiddenimports += [
        "win32com",
        "win32com.client",
        "win32com.gen_py",
        "pythoncom",
        "pywintypes",
    ]
    hiddenimports += collect_submodules("win32com")

# Bundle data
datas = []
# Frontend dist (if built)
if FRONTEND_DIST.exists():
    datas.append((str(FRONTEND_DIST), "frontend/dist"))
# AccrualUpdater Python source bundled at root of MEIPASS so launcher can import it
if ACCRUAL_SRC.exists():
    datas.append((str(ACCRUAL_SRC / "accrual_updater.py"), "."))
    datas.append((str(ACCRUAL_SRC / "admin_fee_module_v18b.py"), "."))

# xlwings ships .xlam addins
xlwings_datas, xlwings_binaries, xlwings_hidden = collect_all("xlwings")
datas += xlwings_datas
hiddenimports += xlwings_hidden

# openpyxl ships data templates
datas += collect_data_files("openpyxl")

a = Analysis(
    [str(BACKEND / "launcher.py")],
    pathex=[str(BACKEND), str(ACCRUAL_SRC)],
    binaries=xlwings_binaries,
    datas=datas,
    hiddenimports=hiddenimports,
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

# Single-file exe on Windows, .app bundle on macOS
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="7t-Accrual-Updater",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # no console window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

if sys.platform == "darwin":
    app = BUNDLE(
        exe,
        name="7t Accrual Updater.app",
        icon=None,
        bundle_identifier="ai.7t.accrual-updater",
        info_plist={
            "NSHighResolutionCapable": True,
            "LSBackgroundOnly": False,
        },
    )
