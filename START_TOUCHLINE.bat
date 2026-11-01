@echo off
setlocal
cd /d "%~dp0"
title TOUCHLINE WEB PORTAL 1.7.1 CLEAN FREEZE
echo ============================================================
echo TOUCHLINE WEB PORTAL 1.7.1 - CLEAN FREEZE
echo SuperLiga + Prva liga dresovi
echo ============================================================
where py >nul 2>nul && (py -3 run.py & goto done)
where python >nul 2>nul && (python run.py & goto done)
python3 run.py
:done
pause
