@echo off
setlocal
cd /d "%~dp0"
title TOUCHLINE WEB PORTAL 1.7.1 - CLEAN FREEZE VERIFY
echo ============================================================
echo TOUCHLINE WEB PORTAL 1.7.1 - CLEAN FREEZE VERIFY
echo ============================================================
where py >nul 2>nul && (py -3 self_test.py & goto done)
where python >nul 2>nul && (python self_test.py & goto done)
python3 self_test.py
:done
pause
