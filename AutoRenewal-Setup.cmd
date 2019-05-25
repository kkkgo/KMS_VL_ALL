@echo off
set ClearKMSCache=1

:: #################################################################

set KMS_IP=172.16.0.2
set KMS_Port=1688
set KMS_Emulation=1

set "SysPath=%Windir%\System32"
if exist "%Windir%\Sysnative\reg.exe" (set "SysPath=%Windir%\Sysnative")
set "Path=%SysPath%;%Windir%;%SysPath%\Wbem;%SysPath%\WindowsPowerShell\v1.0\"

fsutil dirty query %SystemDrive% >nul 2>&1 || goto :E_Admin

set "_Nul1=1>nul"
set "_Nul2=2>nul"
set "_Nul6=2^>nul"
set "_Nul3=1>nul 2>nul"
set "_Pause=pause >nul"
set "_temp=%SystemRoot%\Temp"
set "_work=%~dp0"
if "%_work:~-1%"=="\" set "_work=%_work:~0,-1%"
set xOS=x64
if /i "%PROCESSOR_ARCHITECTURE%"=="x86" (if not defined PROCESSOR_ARCHITEW6432 set xOS=Win32)
set "IFEO=HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options"
set "OSPP=SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform"
set "SPPk=SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform"
set _Hook="%SystemRoot%\system32\SppExtComObjHook.dll"
set "_TaskEx=\Microsoft\Windows\SoftwareProtectionPlatform\SvcTrigger"
set "_TaskOs=\Microsoft\Windows\SoftwareProtectionPlatform\SvcRestartTaskLogon"
set "line============================================================="
color 07
title Auto Renewal Setup
mode con cols=98 lines=28
wmic path SoftwareLicensingProduct where (Description like '%%KMSCLIENT%%') get Name %_Nul2% | findstr /i Windows %_Nul1% && (set SppHook=1) || (set SppHook=0)
wmic path OfficeSoftwareProtectionService get Version %_Nul3% && (set OsppHook=1) || (set OsppHook=0)
setlocal EnableExtensions EnableDelayedExpansion

for /f "tokens=6 delims=[]. " %%G in ('ver') do set winbuild=%%G
    set OffVer="%IFEO%\osppsvc.exe"
if %winbuild% GEQ 9200 (
    set OSType=Win8
    set SppVer="%IFEO%\SppExtComObj.exe"
) else if %winbuild% GEQ 7600 (
    set OSType=Win7
    set SppVer="%IFEO%\sppsvc.exe"
) else (
    goto :UnsupportedVersion
)

echo.
if exist "%SystemRoot%\system32\SppExtComObj*.dll" (
dir /b /al %_Hook% %_Nul3% || goto :uninst
)
if not exist "!_work!\bin\!xOS!.dll" goto :E_DLL

:inst
choice /C YN /N /M "Local KMS Emulator will be installed on your computer. Continue? [y/n]: "
if errorlevel 2 exit
echo.
echo %line%
call :StopService sppsvc
if %OsppHook% NEQ 0 call :StopService osppsvc
if %winbuild% GEQ 9600 (
  WMIC /NAMESPACE:\\root\Microsoft\Windows\Defender PATH MSFT_MpPreference call Add ExclusionPath=%_Hook% %_Nul3% && set "AddExc= and Windows Defender exclusion"
)
echo.
echo Adding File%AddExc%...
echo %_Hook%
for %%# in (SppExtComObjHookAvrf.dll,SppExtComObjHook.dll,SppExtComObjPatcher.dll,SppExtComObjPatcher.exe) do (
  if exist "%SystemRoot%\system32\%%#" del /f /q "%SystemRoot%\system32\%%#" %_Nul3%
)
copy /y "!_work!\bin\!xOS!.dll" %_Hook% %_Nul3% || (echo Failed&goto :END)
echo.
echo Adding Registry Keys...
echo %SppVer%
if %OSType% EQU Win8 (
call :CreateIFEOEntry SppExtComObj.exe
)
if %OSType% EQU Win7 if %SppHook% NEQ 0 (
call :CreateIFEOEntry sppsvc.exe
)
echo %OffVer%
call :CreateIFEOEntry osppsvc.exe
if %winbuild% GEQ 9200 call :CreateTask
if %winbuild% GEQ 9200 schtasks /query /tn "%_TaskEx%" %_Nul3% && (
echo.
echo Adding Schedule Task...
echo "%_TaskEx%"
)
echo.
echo %line%
echo.
echo Done.
echo.
echo It is recommended to exclude this file in the Antivirus protection.
echo %_Hook%
echo.
echo Now run the file Activate.cmd to complete the Auto Renewal Activation.
goto :END

:uninst
choice /C YN /N /M "Local KMS Emulator will be removed from your computer. Continue? [y/n]: "
if errorlevel 2 exit
echo.
echo %line%
call :StopService sppsvc
if %OsppHook% NEQ 0 call :StopService osppsvc
if %winbuild% GEQ 9600 (
reg delete "HKLM\SOFTWARE\Policies\Microsoft\Windows NT\CurrentVersion\Software Protection Platform" /v "NoGenTicket" /f %_Nul3%
for %%# in (SppExtComObjHookAvrf.dll,SppExtComObjHook.dll,SppExtComObjPatcher.dll,SppExtComObjPatcher.exe) do (
  WMIC /NAMESPACE:\\root\Microsoft\Windows\Defender PATH MSFT_MpPreference call Remove ExclusionPath="%SystemRoot%\system32\%%#" %_Nul3% && set "RemExc= and Windows Defender exclusions"
  )
)
echo.
echo Removing Files%RemExc%...
for %%# in (SppExtComObjHookAvrf.dll,SppExtComObjHook.dll,SppExtComObjPatcher.dll,SppExtComObjPatcher.exe) do (
  if exist "%SystemRoot%\system32\%%#" (echo "%SystemRoot%\system32\%%#"&del /f /q "%SystemRoot%\system32\%%#")
)
echo.
echo Removing Registry Keys...
echo %SppVer%
if %OSType% EQU Win8 (
call :RemoveIFEOEntry SppExtComObj.exe
)
if %OSType% EQU Win7 if %SppHook% NEQ 0 (
call :RemoveIFEOEntry sppsvc.exe
)
echo %OffVer%
call :RemoveIFEOEntry osppsvc.exe
if %winbuild% GEQ 9200 schtasks /query /tn "%_TaskEx%" %_Nul3% && (
echo.
echo Removing Schedule Task...
echo "%_TaskEx%"
schtasks /delete /f /tn "%_TaskEx%" %_Nul3%
)
if %ClearKMSCache% EQU 1 (
echo.
echo Clearing KMS Cache...
call :cKMS SoftwareLicensingProduct SoftwareLicensingService %_Nul3%
if %OsppHook% NEQ 0 call :cKMS OfficeSoftwareProtectionProduct OfficeSoftwareProtectionService %_Nul3%
call :cREG %_Nul3%
)
echo.
echo %line%
echo.
echo Done.
goto :END

:StopService
sc query %1 | find /i "STOPPED" %_Nul1% || net stop %1 /y %_Nul3%
sc query %1 | find /i "STOPPED" %_Nul1% || sc stop %1 %_Nul3%
goto :eof

:CreateIFEOEntry
reg delete "%IFEO%\%1" /f /v Debugger %_Nul3%
reg add "%IFEO%\%1" /f /v VerifierDlls /t REG_SZ /d "SppExtComObjHook.dll" %_Nul3% || (echo Failed&del /f /q %_Hook%&goto :END)
reg add "%IFEO%\%1" /f /v GlobalFlag /t REG_DWORD /d 256 %_Nul3%
reg add "%IFEO%\%1" /f /v KMS_Emulation /t REG_DWORD /d %KMS_Emulation% %_Nul3%
if /i %1 EQU osppsvc.exe (
reg add "HKLM\%OSPP%" /f /v KeyManagementServiceName /t REG_SZ /d %KMS_IP% %_Nul3%
reg add "HKLM\%OSPP%" /f /v KeyManagementServicePort /t REG_SZ /d %KMS_Port% %_Nul3%
)
goto :eof

:RemoveIFEOEntry
if /i %1 NEQ osppsvc.exe (
reg delete "%IFEO%\%1" /f %_Nul3%
goto :eof
)
if %OsppHook% EQU 0 if /i %1 EQU osppsvc.exe (
reg delete "%IFEO%\%1" /f %_Nul3%
goto :eof
)
for %%A in (Debugger,VerifierDlls,GlobalFlag,KMS_Emulation,KMS_ActivationInterval,KMS_RenewalInterval,Office2010,Office2013,Office2016,Office2019) do reg delete "%IFEO%\%1" /f /v %%A %_Nul3%
goto :eof

:CreateTask
schtasks /query /tn "%_TaskEx%" %_Nul3% || (
  schtasks /query /tn "%_TaskOs%" %_Nul3% && (
    schtasks /query /tn "%_TaskOs%" /xml >"!_temp!\SvcTrigger.xml"
    schtasks /create /tn "%_TaskEx%" /xml "!_temp!\SvcTrigger.xml" /f %_Nul3%
    schtasks /change /tn "%_TaskEx%" /enable %_Nul3%
    del /f /q "!_temp!\SvcTrigger.xml" %_Nul3%
  )
)
schtasks /query /tn "%_TaskEx%" %_Nul3% || (
  if exist "!_work!\bin\SvcTrigger.xml" schtasks /create /tn "%_TaskEx%" /xml "!_work!\bin\SvcTrigger.xml" /f %_Nul3%
)
goto :eof

:cKMS
set spp=%1
set sps=%2
for /f "tokens=2 delims==" %%G in ('"wmic path %spp% where (Description like '%%KMSCLIENT%%') get ID /VALUE" %_Nul6%') do (set app=%%G&call :cAPP)
for /f "tokens=2 delims==" %%A in ('"wmic path %sps% get Version /VALUE"') do set ver=%%A
wmic path %sps% where version='%ver%' call ClearKeyManagementServiceMachine
wmic path %sps% where version='%ver%' call ClearKeyManagementServicePort
wmic path %sps% where version='%ver%' call DisableKeyManagementServiceDnsPublishing 1
wmic path %sps% where version='%ver%' call DisableKeyManagementServiceHostCaching 1
goto :eof

:cAPP
wmic path %spp% where ID='%app%' call ClearKeyManagementServiceMachine
wmic path %spp% where ID='%app%' call ClearKeyManagementServicePort
goto :eof

:cREG
reg delete "HKLM\%SPPk%\55c92734-d682-4d71-983e-d6ec3f16059f" /f
reg delete "HKLM\%SPPk%\0ff1ce15-a989-479d-af46-f275c6370663" /f
reg delete "HKLM\%SPPk%" /f /v KeyManagementServiceName
reg delete "HKLM\%SPPk%" /f /v KeyManagementServicePort
reg delete "HKU\S-1-5-20\%SPPk%\55c92734-d682-4d71-983e-d6ec3f16059f" /f
reg delete "HKU\S-1-5-20\%SPPk%\0ff1ce15-a989-479d-af46-f275c6370663" /f
reg delete "HKLM\%OSPP%\59a52881-a989-479d-af46-f275c6370663" /f
reg delete "HKLM\%OSPP%\0ff1ce15-a989-479d-af46-f275c6370663" /f
reg delete "HKLM\%OSPP%" /f /v KeyManagementServiceName
reg delete "HKLM\%OSPP%" /f /v KeyManagementServicePort
if %OsppHook% EQU 0 (
reg delete "HKLM\%OSPP%" /f
reg delete "HKU\S-1-5-20\%OSPP%" /f
)
goto :eof

:E_Admin
echo ==== ERROR ====
echo This script require administrator privileges.
echo To do so, right click on this script and select 'Run as administrator'
echo.
echo Press any key to exit.
pause >nul
goto :eof

:E_DLL
echo ==== ERROR ====
echo Required file !xOS!.dll is not found.
echo Make sure folder path is simple and Antivirus protection is OFF or excluded.
goto :END

:UnsupportedVersion
echo ==== ERROR ====
echo Unsupported OS version Detected.
echo Project is supported only for Windows 7/8/8.1/10 and their Server equivalent.
:END
echo.
echo Press any key to exit.
%_Pause%
goto :eof