@echo off
REM ============================================================
REM GEALAN PROXY INSTALLER - Universeel
REM Werkt op elke 3BM machine met Revit 2025
REM Rechtsklik - Als administrator uitvoeren
REM ============================================================
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo FOUT: Dit script vereist Administrator rechten!
    echo Rechtsklik - Als administrator uitvoeren
    pause
    exit /b 1
)

set ADDINS=C:\ProgramData\Autodesk\Revit\Addins\2025
set XDRIVE=X:\10_3BM_bouwkunde\50_Claude-Code-Projects\pyrevit\extensions\bouwkunde.extension\lib\gealan

echo.
echo === GEALAN PROXY INSTALLER ===
echo.

REM Backup origineel
if not exist "%ADDINS%\PlanersoftwareRevitPlugin.addin.ORIGINAL" (
    copy "%ADDINS%\PlanersoftwareRevitPlugin.addin" "%ADDINS%\PlanersoftwareRevitPlugin.addin.ORIGINAL"
    echo [OK] Backup gemaakt
) else (
    echo [OK] Backup bestaat al
)

REM Copy proxy DLL
if not exist "%ADDINS%\GealanProxy" mkdir "%ADDINS%\GealanProxy"
copy /Y "%XDRIVE%\GealanProxy.dll" "%ADDINS%\GealanProxy\"
echo [OK] Proxy DLL gekopieerd

REM Write proxy addin
(
echo ^<?xml version="1.0" encoding="utf-8"?^>
echo ^<RevitAddIns^>
echo   ^<AddIn Type="Application"^>
echo     ^<n^>PlanersoftwareRevitPlugin^</n^>
echo     ^<Assembly^>GealanProxy\GealanProxy.dll^</Assembly^>
echo     ^<AddInId^>30b29e5f-151b-4544-93b5-cbbc97500036^</AddInId^>
echo     ^<FullClassName^>GealanProxy.ProxyApplication^</FullClassName^>
echo     ^<VendorId^>Gealan^</VendorId^>
echo     ^<VendorDescription^>GEALAN Fenster-Systeme GmbH, www.gealan.com^</VendorDescription^>
echo   ^</AddIn^>
echo ^</RevitAddIns^>
) > "%ADDINS%\PlanersoftwareRevitPlugin.addin"
echo [OK] Addin aangepast naar proxy

echo.
echo KLAAR! Start Revit - je hebt nu:
echo  - Gealan Planersoftware (werkt normaal)
echo  - "Read K-merken" knop in Gealan Tools panel
echo  - "Gealan Rename" in Bouwkunde tab (na pyRevit reload)
echo.
pause
