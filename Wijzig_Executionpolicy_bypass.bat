:: Script om de ExecutionPolicy te wijzigen
:: Laatst gewijzigd op 19-11-2026
:: Werkt met PowerShell 7 en valt terug op Windows PowerShell 5.1

@echo off
setlocal
title PowerShell ExecutionPolicy wijzigen

echo ==========================================
echo   PowerShell ExecutionPolicy wijzigen
echo ==========================================
echo.

:: Controleer of PowerShell 7 aanwezig is
where pwsh.exe >nul 2>&1

if %errorlevel%==0 (
    set "PS=pwsh.exe"
    echo PowerShell 7 is gevonden.
) else (
    :: PowerShell 5.1 als alternatief
    set "PS=powershell.exe"
    echo PowerShell 7 niet gevonden.
    echo Windows PowerShell 5.1 wordt gebruikt.
)

echo.
echo Geselecteerde PowerShell:
echo %PS%
echo.
echo De ExecutionPolicy wordt ingesteld op Bypass
echo voor de huidige gebruiker.
echo.

:: ExecutionPolicy wijzigen
"%PS%" -NoProfile -Command "Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope CurrentUser -Force"

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
"%PS%" -NoProfile -Command "Get-ExecutionPolicy"

echo.
echo Instellingen per scope:
"%PS%" -NoProfile -Command "Get-ExecutionPolicy -List"

echo.
pause
endlocal
