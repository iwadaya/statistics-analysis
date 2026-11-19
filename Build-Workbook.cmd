@echo off
rem Double-click to build Performance_Tracker.xlsm from Performance_Tracker.xlsx + vba\*.bas
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Build-Workbook.ps1" %*
pause
