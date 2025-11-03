```markdown
# iTech Accrual Updater — GUI

This project contains:
- accrual_updater.py — core parsing / updating logic.
- app_polished.py — polished PySide6 GUI (recommended).
- company_profiles.py — profiles manager for multiple companies.
- run_gui.command / run_gui.bat — simple launchers for macOS and Windows.
- requirements_polished.txt — dependencies for polished GUI.

Quick local test (macOS / Linux):
1. Open Terminal and change into the project folder:
   cd "/Users/youruser/path/to/iTech_Accrual_Updater"

2. Create & activate a virtual environment:
   python3 -m venv .venv
   source .venv/bin/activate

3. Install dependencies:
   pip install --upgrade pip
   pip install -r requirements_polished.txt

4. Run the GUI:
   python3 app_polished.py

5. Use the UI to browse for the Master file and Paysheets folder, add pay dates, test with dry run, then uncheck dry run and run to apply changes. The script will make a backup if you keep "Create backup" checked.

Packaging with PyInstaller (local build):
1. Ensure venv is active.
2. Install pyinstaller (if not already):
   pip install pyinstaller
3. Build:
   pyinstaller --noconfirm --onefile --windowed --name "iTechAccrual" --add-data "assets:assets" --icon assets/app_icon.icns app_polished.py
4. The built executable will be in dist/iTechAccrual.

Automated CI builds:
- There's an example GitHub Actions workflow earlier in the conversation that can produce macOS and Windows artifacts automatically. Commit and push the repo, then check Actions to download build artifacts.

Notes:
- For .xls support ensure xlrd==1.2.0 is installed.
- For macOS distribution you may need to code sign and notarize the app to avoid Gatekeeper warnings.
- Use the company_profiles.py manager to create and switch company profiles (iTech, Smartworks, etc.) for separate masters/paysheets/paydate configs.

If you want, I can:
- Provide a ready .spec file for PyInstaller bundling assets and setting working directory.
- Produce a GitHub Actions workflow tailored to the polished app and upload signed artifacts (you'll need to provide signing credentials).
- Create sample app icons (icns/ico) and a DMG/NSIS installer recipe.
```