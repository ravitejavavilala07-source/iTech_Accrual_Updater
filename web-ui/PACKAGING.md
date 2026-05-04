# Packaging — 7t Accrual Updater

PyInstaller can't cross-compile. Build on each target OS.

## Layout

```
itech-web-ui/
├── backend/
│   ├── main.py          FastAPI app
│   ├── launcher.py      Entry point — starts uvicorn + opens browser
│   └── requirements.txt
├── frontend/
│   └── dist/            Built React (npm run build)
└── build/
    ├── accrual.spec     PyInstaller spec (cross-platform)
    ├── build-mac.sh     macOS build
    └── build-windows.ps1  Windows build
```

## macOS build (Ravi's machine)

```bash
cd build
./build-mac.sh
```

Output:
- `build/dist/7t-Accrual-Updater` — Unix executable (~36 MB)
- `build/dist/7t Accrual Updater.app` — clickable .app bundle

Run:
```bash
"./build/dist/7t Accrual Updater.app/Contents/MacOS/7t-Accrual-Updater"
```
Or double-click the `.app`. Browser opens at `http://127.0.0.1:8100` automatically.

## Windows build (for Kristina)

On a Windows machine with Python 3.10+ and Node 18+ installed:

```powershell
cd build
.\build-windows.ps1
```

Output: `build\dist\7t-Accrual-Updater.exe` (single file, ~50 MB).

Send this `.exe` to Kristina. She double-clicks it — uvicorn starts, default browser opens to the app.

## What ships inside the binary

- React frontend (compiled `dist/`)
- FastAPI backend (`main.py`)
- Both Python modules: `accrual_updater.py`, `admin_fee_module_v18b.py` (auto-located by `_locate_accrual_src()`)
- All Python deps: uvicorn, openpyxl, pandas, xlrd, xlwings
- **Windows only:** pywin32 (Outlook COM for bug reports), pythoncom

## Bug-report email flow on Windows

When Kristina submits a bug:
1. Backend logs to `%APPDATA%\AccrualUpdater\bug_reports.jsonl` (always)
2. Calls `Outlook.Application` via pywin32 COM
3. Mail addressed To: `ravi.vavilala@riseits.com`, Reply-To: her typed email, sent from her default Outlook account
4. If Outlook unavailable → frontend opens `mailto:` fallback

You can change the recipient by setting environment variable before launch:
```powershell
$env:BUG_REPORT_TO = "different@example.com"
.\7t-Accrual-Updater.exe
```

## Optional environment variables

| Var | Default | Purpose |
|---|---|---|
| `ACCRUAL_PORT` | `8100` | Server port |
| `ACCRUAL_ALLOW_MACROS` | `1` (set by launcher) | Allow `.xlsm` masters |
| `ACCRUAL_ALLOWED_ROOTS` | OS-specific | `;`-separated paths the user can browse |
| `BUG_REPORT_TO` | `ravi.vavilala@riseits.com` | Bug recipient |
| `SMTP_HOST` / `SMTP_USER` / `SMTP_PASS` | (unset) | Optional SMTP fallback for bug reports |

## Verifying the build

After launching the packaged exe:
1. Browser should open to `http://127.0.0.1:8100`
2. Check `/api/health` returns `{"status":"ok","accrual_available":true,"admin_fee_available":true}`
3. Click Browse → native OS picker should open (no console flash on Windows)
4. Help → Format Rules opens modal
5. Help → Bug Report sends mail (visible in Outlook Sent Items)

## Known limitations

- macOS Gatekeeper may block the .app on first run — right-click → Open
- Windows SmartScreen may warn — click "More info" → "Run anyway" (or sign with code-signing cert)
- The first launch unpacks PyInstaller bundle to a temp dir; subsequent launches are faster
- Concurrent runs are blocked (HTTP 409) — only one accrual run at a time
