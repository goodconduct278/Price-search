@echo off
cd /d %~dp0
echo Starting import...
py db_import.py --db price_db.db --csvdir csv
echo.
pause
