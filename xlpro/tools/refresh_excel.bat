@echo off
@echo Restarting Excel...
cd xlpro/tools
"../../.venv/scripts/python.exe" "init_excel_session.py"
@echo Restart complete