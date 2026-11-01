@echo off
setlocal
cd /d "%~dp0"
title TOUCHLINE - INSTALL DEPENDENCIES
where py >nul 2>nul
if %errorlevel%==0 (
  py -3 -m pip install -r requirements.txt
  goto :done
)
where python >nul 2>nul
if %errorlevel%==0 (
  python -m pip install -r requirements.txt
  goto :done
)
echo [ERROR] Python nije pronadjen.
pause
exit /b 1
:done
echo.
echo [OK] Dependencies installed.
pause
endlocal
