#!/bin/bash
# Double-click this on macOS to run the polished GUI (opens Terminal)
cd "/Users/ravitejavavilala/iTech_Accrual_Updater"
if [ -f ".venv/bin/activate" ]; then
  source .venv/bin/activate
fi
python3 app_polished.py
echo "GUI exited. Close this window to finish."
read -n 1 -s -r -p "Press any key to close..."