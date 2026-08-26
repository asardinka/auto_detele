@echo off
chcp 65001 >nul
title AutoDelete
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File "%~dp0AutoDelete.ps1"
if errorlevel 1 pause
