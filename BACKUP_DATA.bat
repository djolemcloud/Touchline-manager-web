@echo off
setlocal
cd /d "%~dp0"
if not exist "data\cm_data_hub.db" (
  echo [ERROR] data\cm_data_hub.db ne postoji.
  pause
  exit /b 1
)
if not exist "BACKUPS" mkdir "BACKUPS"
for /f "tokens=1-4 delims=/-. " %%a in ('date /t') do set D=%%d%%b%%c
for /f "tokens=1-2 delims=: " %%a in ('time /t') do set T=%%a%%b
copy /y "data\cm_data_hub.db" "BACKUPS\cm_data_hub_%D%_%T%.db" >nul
if errorlevel 1 (
  echo [ERROR] Backup nije napravljen.
) else (
  echo [OK] Backup baze je napravljen u BACKUPS folderu.
)
pause
endlocal
