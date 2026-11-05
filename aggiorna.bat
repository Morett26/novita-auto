@echo off
rem Scarica le ultime notizie e apre il sito.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0aggiorna.ps1"
if errorlevel 1 pause
start "" "%~dp0index.html"
