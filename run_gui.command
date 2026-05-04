#!/bin/bash
# Double-click this on macOS to run the polished GUI (opens Terminal)
cd "/Users/ravitejavavilala/ml:ai projects/iTech_Accrual_Updater"

# Activate virtual environment
if [ -f ".venv/bin/activate" ]; then
  source .venv/bin/activate
else
  echo "ERROR: Virtual environment not found. Run: python -m venv .venv"
  read -n 1 -s -r -p "Press any key to close..."
  exit 1
fi

# Check if xlwings is installed
python3 -c "import xlwings" 2>/dev/null
if [ $? -ne 0 ]; then
  echo "Installing xlwings (required for Excel integration)..."
  pip install xlwings==0.33.1
fi

# Run the GUI
python3 app_polished.py
RESULT=$?

echo ""
echo "GUI exited with code $RESULT. Close this window to finish."
read -n 1 -s -r -p "Press any key to close..."
exit $RESULT