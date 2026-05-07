@echo off
setlocal
echo Installing build dependency...
python -m pip install -r requirements.txt
echo Building Windows executable...
python -m PyInstaller --noconfirm --onefile --windowed --name CSU_Steam_Bill_Automation steam_bill_automation_app.py
echo Build complete. Check the dist folder.
pause