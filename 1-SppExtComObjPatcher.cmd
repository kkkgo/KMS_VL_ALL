@echo off
set ClearKMSCache=1

set KMS_Emulation=1
set KMS_IP=172.16.0.2
set KMS_Port=1688
set KMS_ActivationInterval=120
set KMS_RenewalInterval=10080
set KMS_HWID=0x3A1C049600B60076

set "SysPath=%Windir%\System32"
if exist "%Windir%\Sysnative\reg.exe" (set "SysPath=%Windir%\Sysnative")
set "Path=%SysPath%;%Windir%;%SysPath%\Wbem;%SysPath%\WindowsPowerShell\v1.0\"

fsutil dirty query %systemdrive% >nul 2>&1 || goto :E_Admin

title SppExtComObjPatcher
set "_workdir=%~dp0"
if "%_workdir:~-1%"=="\" set "_workdir=%_workdir:~0,-1%"
set xOS=x64
if /i "%PROCESSOR_ARCHITECTURE%"=="x86" (if "%PROCESSOR_ARCHITEW6432%"=="" set xOS=Win32)
setlocal EnableExtensions EnableDelayedExpansion
set "IFEO=HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options"
set "OSPP=SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform"
set "SPPk=SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform"
wmic path SoftwareLicensingProduct where (Description like '%%KMSCLIENT%%') get Name 2>nul | findstr /i Windows 1>nul && (set SppHook=1) || (set SppHook=0)
wmic path OfficeSoftwareProtectionService get Version >nul 2>&1 && (set OsppHook=1) || (set OsppHook=0)

for /f "tokens=6 delims=[]. " %%G in ('ver') do set winbuild=%%G
if %winbuild% GEQ 9200 (
    set OSType=Win8
    set OffVer=2010
) else if %winbuild% GEQ 7600 (
    set OSType=Win7
    set OffVer=2010/2013/2016/2019
) else (
    goto :UnsupportedVersion
)

echo.
echo Microsoft (R) Windows Software Licensing.
echo Copyright (C) Microsoft Corporation. All rights reserved.
echo =========================================================
echo.
if exist "%SystemRoot%\system32\SppExtComObj*.dll" (
dir /b /al "%SystemRoot%\system32\SppExtComObjHook.dll" >nul 2>&1 || goto :uninst
)

:inst
choice /C YN /N /M "SppExtComObjPatcher will be installed on your computer. Continue? [y/n]: "
if errorlevel 2 exit
call :StopService sppsvc
if %OsppHook% NEQ 0 call :StopService osppsvc
if %winbuild% GEQ 9600 (
	WMIC /NAMESPACE:\\root\Microsoft\Windows\Defender PATH MSFT_MpPreference call Add ExclusionPath="%SystemRoot%\system32\SppExtComObjHook.dll" >nul 2>&1
)
echo.
echo Copying Files...
for %%# in (SppExtComObjHookAvrf.dll,SppExtComObjHook.dll,SppExtComObjPatcher.dll,SppExtComObjPatcher.exe) do (
	if exist "%SystemRoot%\system32\%%#" del /f /q "%SystemRoot%\system32\%%#" >nul 2>&1
)
copy /y "!_workdir!\!xOS!\SppExtComObjHook.dll" "%SystemRoot%\system32" >nul 2>&1 && (echo Copied SppExtComObjHook.dll) || (echo SppExtComObjHook.dll Failed&pause&exit /b)
echo.
echo Creating Registry Entries...
if %OSType% EQU Win8 (
    echo SppExtComObj.exe of Windows 8/8.1/10 - Office 2013/2016/2019
    call :CreateIFEOEntry SppExtComObj.exe
)
if %OSType% EQU Win7 if %SppHook% NEQ 0 (
    echo sppsvc.exe of Windows 7
    call :CreateIFEOEntry sppsvc.exe
)
    echo osppsvc.exe of Office %OffVer%
    call :CreateIFEOEntry osppsvc.exe
if %winbuild% GEQ 9200 (
	schtasks /query /tn "\Microsoft\Windows\SoftwareProtectionPlatform\SvcTrigger" >nul 2>&1 || schtasks /create /tn "\Microsoft\Windows\SoftwareProtectionPlatform\SvcTrigger" /xml "!_workdir!\Win32\SvcTrigger.xml" /f >nul 2>&1
)
goto :End

:uninst
choice /C YN /N /M "SppExtComObjPatcher will be removed from your computer. Continue? [y/n]: "
if errorlevel 2 exit
call :StopService sppsvc
if %OsppHook% NEQ 0 call :StopService osppsvc
echo.
echo Removing Installed Files...
for %%# in (SppExtComObjHookAvrf.dll,SppExtComObjHook.dll,SppExtComObjPatcher.dll,SppExtComObjPatcher.exe) do (
	if exist "%SystemRoot%\system32\%%#" (echo %%#&del /f /q "%SystemRoot%\system32\%%#")
)
echo.
echo Removing Registry Entries...
if %OSType% EQU Win8 (
    echo SppExtComObj.exe of Windows 8/8.1/10 - Office 2013/2016/2019
    call :RemoveIFEOEntry SppExtComObj.exe
)
if %OSType% EQU Win7 if %SppHook% NEQ 0 (
    echo sppsvc.exe of Windows 7
    call :RemoveIFEOEntry sppsvc.exe
)
    echo osppsvc.exe of Office %OffVer%
    call :RemoveIFEOEntry osppsvc.exe
if %ClearKMSCache% EQU 1 (
echo.
echo Clearing KMS Cache...
call :cKMS SoftwareLicensingProduct SoftwareLicensingService >nul 2>&1
if %OsppHook% NEQ 0 call :cKMS OfficeSoftwareProtectionProduct OfficeSoftwareProtectionService >nul 2>&1
call :cREG >nul 2>&1
)
if %winbuild% GEQ 9200 (
schtasks /query /tn "\Microsoft\Windows\SoftwareProtectionPlatform\SvcTrigger" >nul 2>&1 && schtasks /delete /f /tn "\Microsoft\Windows\SoftwareProtectionPlatform\SvcTrigger" >nul 2>&1
)
if %winbuild% GEQ 9600 (
reg delete "HKLM\SOFTWARE\Policies\Microsoft\Windows NT\CurrentVersion\Software Protection Platform" /v "NoGenTicket" /f >nul 2>&1
for %%# in (SppExtComObjHookAvrf.dll,SppExtComObjHook.dll,SppExtComObjPatcher.dll,SppExtComObjPatcher.exe) do (
	WMIC /NAMESPACE:\\root\Microsoft\Windows\Defender PATH MSFT_MpPreference call Remove ExclusionPath="%SystemRoot%\system32\%%#" >nul 2>&1
	)
)
goto :End

:StopService
sc query %1 | find /i "STOPPED" >nul || net stop %1 /y >nul 2>&1
sc query %1 | find /i "STOPPED" >nul || sc stop %1 >nul 2>&1
goto :eof

:CreateIFEOEntry
reg delete "%IFEO%\%1" /f /v Debugger >nul 2>&1
reg add "%IFEO%\%1" /f /v VerifierDlls /t REG_SZ /d "SppExtComObjHook.dll" >nul 2>&1
reg add "%IFEO%\%1" /f /v GlobalFlag /t REG_DWORD /d 256 >nul 2>&1
reg add "%IFEO%\%1" /f /v KMS_Emulation /t REG_DWORD /d %KMS_Emulation% >nul 2>&1
reg add "%IFEO%\%1" /f /v KMS_ActivationInterval /t REG_DWORD /d %KMS_ActivationInterval% >nul 2>&1
reg add "%IFEO%\%1" /f /v KMS_RenewalInterval /t REG_DWORD /d %KMS_RenewalInterval% >nul 2>&1
if /i %1 EQU SppExtComObj.exe if %winbuild% GEQ 9600 (
reg add "%IFEO%\%1" /f /v KMS_HWID /t REG_QWORD /d "%KMS_HWID%" >nul 2>&1
)
if /i %1 EQU osppsvc.exe (
reg add "HKLM\%OSPP%" /f /v KeyManagementServiceName /t REG_SZ /d %KMS_IP% >nul 2>&1
reg add "HKLM\%OSPP%" /f /v KeyManagementServicePort /t REG_SZ /d %KMS_Port% >nul 2>&1
)
goto :eof

:RemoveIFEOEntry
if /i %1 NEQ osppsvc.exe (
reg delete "%IFEO%\%1" /f >nul 2>&1
goto :eof
)
if %OsppHook% EQU 0 if /i %1 EQU osppsvc.exe (
reg delete "%IFEO%\%1" /f >nul 2>&1
goto :eof
)
for %%A in (VerifierDlls,GlobalFlag,Debugger,KMS_Emulation,KMS_ActivationInterval,KMS_RenewalInterval,Office2010,Office2013,Office2016,Office2019) do reg delete "%IFEO%\%1" /f /v %%A >nul 2>&1
goto :eof

:cKMS
set spp=%1
set sps=%2
for /f "tokens=2 delims==" %%G in ('"wmic path %spp% where (Description like '%%KMSCLIENT%%') get ID /VALUE" 2^>nul') do (set app=%%G&call :cAPP)
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
echo Press any key to exit...
pause >nul
goto :eof

:UnsupportedVersion
echo ==== ERROR ====
echo Unsupported OS version Detected.
echo Project is supported only for Windows 7/8/8.1/10 and their Server equivalent.
echo.
echo Press any key to exit...
pause >nul
goto :eof

:End
echo.
echo Done.
echo Press any key to exit...
pause >nul
goto :eof