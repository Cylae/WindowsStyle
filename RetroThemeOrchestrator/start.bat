@echo off
REM This script launches the Retro Theme Assistant PowerShell application.
REM It bypasses the execution policy for this script only, making it easy to run.

powershell.exe -ExecutionPolicy Bypass -File "%~dp0main.ps1"