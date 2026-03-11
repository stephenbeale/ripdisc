@echo off
title RipDisc Setup
echo.
echo  ========================================
echo    RipDisc - DVD ^& Blu-ray Ripping Tool
echo  ========================================
echo.
echo  This will set up RipDisc on your computer.
echo  It will check for required tools (MakeMKV and HandBrakeCLI)
echo  and help you install anything that's missing.
echo.
pause

powershell -ExecutionPolicy Bypass -NoProfile -File "%~dp0setup.ps1"

echo.
pause
