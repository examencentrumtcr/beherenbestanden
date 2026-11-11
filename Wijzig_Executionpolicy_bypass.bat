:: Script om de Executionpolicy te wijzigen 
:: Laatst gewijzigd op 27-10-2026
:: Laatste wijziging: Werkt nu met PowerShell 7

@echo off
title PowerShell ExecutionPolicy wijzigen

echo ==========================================
echo   PowerShell ExecutionPolicy wijzigen
echo ==========================================
echo.
echo De ExecutionPolicy wordt ingesteld op Bypass
echo voor de huidige gebruiker.
echo.

pwsh -NoProfile -Command "Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope CurrentUser -Force"

if errorlevel 1 (
    echo.
    echo FOUT: De wijziging is mogelijk mislukt.
    echo Controleer of Groepsbeleid de instelling blokkeert.
    echo.
    pause
    exit /b 1
)

echo.
echo De effectieve ExecutionPolicy is:
pwsh -NoProfile -Command "Get-ExecutionPolicy"

echo.
echo Instellingen per scope:
pwsh -NoProfile -Command "Get-ExecutionPolicy -List"

echo.
pause
