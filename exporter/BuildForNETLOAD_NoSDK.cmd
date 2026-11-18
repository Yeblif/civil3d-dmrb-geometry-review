@echo off
title DMRB Superelevation Exporter - No-SDK Build
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0BuildForNETLOAD_NoSDK.ps1"
echo.
pause
