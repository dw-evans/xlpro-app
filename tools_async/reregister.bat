@echo off
@echo Refreshing COM Server Registration
cd xlpro/tools
"../../.venv/scripts/python.exe" "../server.py" --unregister
"../../.venv/scripts/python.exe" "../server.py" --register
@echo Re-register complete