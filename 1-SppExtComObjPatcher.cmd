@echo off
set ClearKMSCache=1

set KMS_Emulation=1
set KMS_IP=172.16.0.2
set KMS_Port=1688
set KMS_ActivationInterval=120
set KMS_RenewalInterval=10080
set KMS_HWID=0x3A1C049600B60076
set Windows=Random
set Office2010=Random
set Office2013=Random
set Office2016=Random
set Office2019=Random

%windir%\system32\reg.exe query "HKU\S-1-5-19" >nul 2>&1 || (
echo ==== ERROR ====
echo This script require administrator privileges.
echo To do so, right click on this script and select 'Run as administrator'
echo.
echo Press any key to exit...
pause >nul
goto :eof
)
title SppExtComObjPatcher
setlocal EnableExtensions EnableDelayedExpansion
cd /d "%~dp0"
IF /I "%PROCESSOR_ARCHITECTURE%" EQU "AMD64" (set xOS=x64) else (set xOS=Win32)
set "IFEO=HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options"
set "OSPP=HKLM\SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform"
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
IF EXIST "%SystemRoot%\system32\SppExtComObj*.dll" goto :uninst

:inst
choice /C YN /N /M "SppExtComObjPatcher will be installed on your computer. Continue? [y/n]: "
IF ERRORLEVEL 2 exit
call :StopService sppsvc
if %OsppHook% NEQ 0 call :StopService osppsvc
if %winbuild% GEQ 9600 (
WMIC /NAMESPACE:\\root\Microsoft\Windows\Defender PATH MSFT_MpPreference call Add ExclusionPath="%SystemRoot%\system32\SppExtComObjHook.dll" >nul 2>&1
)
echo.
echo Copying Files...
copy /y "%xOS%\SppExtComObjHook.dll" "%SystemRoot%\system32" >nul 2>&1 && (echo Copied SppExtComObjHook.dll) || (echo SppExtComObjHook.dll Failed&pause&exit)
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
goto :End

:uninst
choice /C YN /N /M "SppExtComObjPatcher will be removed from your computer. Continue? [y/n]: "
IF ERRORLEVEL 2 exit
call :StopService sppsvc
if %OsppHook% NEQ 0 call :StopService osppsvc
echo.
echo Removing Installed Files...
if exist "%SystemRoot%\system32\SppExtComObjHook.dll" (
	echo SppExtComObjHook.dll
	del /f /q "%SystemRoot%\system32\SppExtComObjHook.dll"
)
if exist "%SystemRoot%\system32\SppExtComObjPatcher.dll" (
	echo SppExtComObjPatcher.dll
	del /f /q "%SystemRoot%\system32\SppExtComObjPatcher.dll"
)
if exist "%SystemRoot%\system32\SppExtComObjPatcher.exe" (
	echo SppExtComObjPatcher.exe
	del /f /q "%SystemRoot%\system32\SppExtComObjPatcher.exe"
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
call :cKMS SoftwareLicensingProduct SoftwareLicensingService
if %OsppHook% NEQ 0 call :cKMS OfficeSoftwareProtectionProduct OfficeSoftwareProtectionService
)
if %winbuild% GEQ 9600 (
reg delete "HKLM\SOFTWARE\Policies\Microsoft\Windows NT\CurrentVersion\Software Protection Platform" /v "NoGenTicket" /f >nul 2>&1
schtasks /delete /f /tn "\Microsoft\Windows\SoftwareProtectionPlatform\SvcTrigger" >nul 2>&1
WMIC /NAMESPACE:\\root\Microsoft\Windows\Defender PATH MSFT_MpPreference call Remove ExclusionPath="%SystemRoot%\system32\SppExtComObjHook.dll" >nul 2>&1
WMIC /NAMESPACE:\\root\Microsoft\Windows\Defender PATH MSFT_MpPreference call Remove ExclusionPath="%SystemRoot%\system32\SppExtComObjPatcher.dll" >nul 2>&1
WMIC /NAMESPACE:\\root\Microsoft\Windows\Defender PATH MSFT_MpPreference call Remove ExclusionPath="%SystemRoot%\system32\SppExtComObjPatcher.exe" >nul 2>&1
)
goto :End

:StopService
sc query %1 | find /i "STOPPED" >nul || net stop %1 /y >nul 2>&1
sc query %1 | find /i "STOPPED" >nul || sc stop %1 >nul 2>&1
goto :EOF

:CreateIFEOEntry
reg add "%IFEO%\%1" /f /v Debugger /t REG_SZ /d "rundll32.exe SppExtComObjHook.dll,PatcherMain" >nul 2>&1
reg add "%IFEO%\%1" /f /v KMS_Emulation /t REG_DWORD /d %KMS_Emulation% >nul 2>&1
reg add "%IFEO%\%1" /f /v KMS_ActivationInterval /t REG_DWORD /d %KMS_ActivationInterval% >nul 2>&1
reg add "%IFEO%\%1" /f /v KMS_RenewalInterval /t REG_DWORD /d %KMS_RenewalInterval% >nul 2>&1
if /i %1 NEQ osppsvc.exe (
reg add "%IFEO%\%1" /f /v Windows /t REG_SZ /d "%Windows%" >nul 2>&1
if %winbuild% GEQ 9200 for %%A in (2013,2016,2019) do reg add "%IFEO%\%1" /f /v Office%%A /t REG_SZ /d "!Office%%A!" >nul 2>&1
)
if /i %1 EQU osppsvc.exe (
reg add "%IFEO%\%1" /f /v Office2010 /t REG_SZ /d "%Office2010%" >nul 2>&1
if %winbuild% LSS 9200 for %%A in (2013,2016,2019) do reg add "%IFEO%\%1" /f /v Office%%A /t REG_SZ /d "!Office%%A!" >nul 2>&1
reg add "%OSPP%" /f /v KeyManagementServiceName /t REG_SZ /d %KMS_IP% >nul 2>&1
reg add "%OSPP%" /f /v KeyManagementServicePort /t REG_SZ /d %KMS_Port% >nul 2>&1
)
if /i %1 EQU SppExtComObj.exe if %winbuild% GEQ 9600 (
reg add "%IFEO%\%1" /f /v KMS_HWID /t REG_QWORD /d "%KMS_HWID%" >nul 2>&1
)
goto :EOF

:RemoveIFEOEntry
if /i %1 NEQ osppsvc.exe (
reg delete "%IFEO%\%1" /f >nul 2>&1
goto :EOF
)
for %%A in (Debugger,KMS_Emulation,KMS_ActivationInterval,KMS_RenewalInterval,Office2010,Office2013,Office2016,Office2019) do reg delete "%IFEO%\%1" /f /v %%A >nul 2>&1
if %ClearKMSCache% EQU 1 (
reg delete "%OSPP%" /f /v KeyManagementServiceName >nul 2>&1
reg delete "%OSPP%" /f /v KeyManagementServicePort >nul 2>&1
)
goto :EOF

:cKMS
set spp=%1
set sps=%2
for /f "tokens=2 delims==" %%G in ('"wmic path %spp% where (Description like '%%KMSCLIENT%%') get ID /VALUE" 2^>nul') do (set app=%%G&call :Clear)
for /f "tokens=2 delims==" %%A in ('"wmic path %sps% get Version /VALUE"') do set ver=%%A
wmic path %sps% where version='%ver%' call ClearKeyManagementServiceMachine >nul 2>&1
wmic path %sps% where version='%ver%' call ClearKeyManagementServicePort >nul 2>&1
wmic path %sps% where version='%ver%' call DisableKeyManagementServiceDnsPublishing 1 >nul 2>&1
wmic path %sps% where version='%ver%' call DisableKeyManagementServiceHostCaching 1 >nul 2>&1
if /i %1 EQU SoftwareLicensingProduct (
reg delete "HKLM\%SPPk%\55c92734-d682-4d71-983e-d6ec3f16059f" /f >nul 2>&1
reg delete "HKLM\%SPPk%\0ff1ce15-a989-479d-af46-f275c6370663" /f >nul 2>&1
reg delete "HKEY_USERS\S-1-5-20\%SPPk%\55c92734-d682-4d71-983e-d6ec3f16059f" /f >nul 2>&1
reg delete "HKEY_USERS\S-1-5-20\%SPPk%\0ff1ce15-a989-479d-af46-f275c6370663" /f >nul 2>&1
) else (
reg delete "%OSPP%\59a52881-a989-479d-af46-f275c6370663" /f >nul 2>&1
reg delete "%OSPP%\0ff1ce15-a989-479d-af46-f275c6370663" /f >nul 2>&1
)
goto :EOF

:Clear
wmic path %spp% where ID='%app%' call ClearKeyManagementServiceMachine >nul 2>&1
wmic path %spp% where ID='%app%' call ClearKeyManagementServicePort >nul 2>&1
goto :EOF

:UnsupportedVersion
echo ==== ERROR ====
echo Unsupported OS version Detected.
echo Project is supported only for Windows 7/8/8.1/10 and their Server equivalent.
echo.
echo Press any key to exit...
pause >nul
goto :EOF

:End
echo.
echo Done.
echo Press any key to exit...
pause >nul
goto :EOF