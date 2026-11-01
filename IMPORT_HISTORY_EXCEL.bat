@echo off
setlocal
cd /d "%~dp0"
echo ============================================================
echo TOUCHLINE HISTORY MASTER EXCEL IMPORT
ECHO ============================================================
echo.
if not exist "HISTORY\TOUCHLINE_HISTORY_MASTER.xlsx" (
  echo [GRESKA] Nedostaje HISTORY\TOUCHLINE_HISTORY_MASTER.xlsx
  pause
  exit /b 1
)
python -c "from portal.sync_engine import init_db; init_db(); from portal.history_excel import import_master; print(import_master(force=True))"
if errorlevel 1 (
  echo [GRESKA] Import nije prosao.
  pause
  exit /b 1
)
echo.
echo [OK] History Master Excel je uvezen.
pause
