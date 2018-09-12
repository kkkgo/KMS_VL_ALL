@echo off
set /a _Debug=0
::==========================================
:: Get Administrator Rights
set _Args=%*
if "%~1" NEQ "" (
  set _Args=%_Args:"=%
)
fltmc 1>nul 2>nul || (
  cd /d "%~dp0"
  cmd /u /c echo Set UAC = CreateObject^("Shell.Application"^) : UAC.ShellExecute "cmd.exe", "/k cd ""%~dp0"" && ""%~dpnx0"" ""%_Args%""", "", "runas", 1 > "%temp%\GetAdmin.vbs"
  "%temp%\GetAdmin.vbs"
  del /f /q "%temp%\GetAdmin.vbs" 1>nul 2>nul
  exit
)
::==========================================
:: Can be 0 (Delete Auto-Renewal-Task OR Manual-Mode) | 1 (Create Auto-Renewal-Task)
set /a _Task=0
:: Define switches
echo %*| find /i "-createtask" >nul&& set /a _Task=1
echo %*| find /i "-renewalonly" >nul&& set /a _Task=2
::==========================================
:: No Debug, Define the nul suppressors
if %_Debug% EQU 0 (
  set "_Nul_1=1>nul"
  set "_Nul_2=2>nul"
  set "_Nul_2e=2^>nul"
  set "_Nul_1_2=1>nul 2>nul"
  call :Begin
) else (
  REM Debug, Clear all nul suppressors, Call script redirecting output to log file
  set "_Nul_1="
  set "_Nul_2="
  set "_Nul_2e="
  set "_Nul_1_2="
  echo.
  echo Running in Debug Mode...
  echo The window will be closed when finished
  @echo on
  @prompt $G
  @call :Begin >"%~dpn0.tmp" 2>&1 &cmd /u /c type "%~dpn0.tmp">"%~dpn0_Debug.log"&del "%~dpn0.tmp"
)
exit
::==========================================
:Begin
:: Set Title of the Script; Color [Background][Text] in hex (0 to F)
title KMS-VL-ALL-7.2RC2 [2018-08-20T09:14Z]
color 07
:: Get Fully Qualified FileName of the Script
set "_FileName=%~f0"
:: Get Drive and Path containing the Script
set "_FileDir=%~dp0"
if "%_FileDir:~-1%"=="\" set "_FileDir=%_FileDir:~0,-1%"
:: Set Internal KMS Server Path
set "_ServerPath=%_FileDir%\32-bit\vlmcsd.exe"
:: Set Task Name for the Script
set "_TaskName=KMS_VL_ALL"
:: Set EnableExtensions and DelayedExpansion
setlocal EnableExtensions EnableDelayedExpansion
:: Can be 0 (Online Mode - Used for External KMS Server) | 1 (Offline Mode - Used for Internal KMS Server)
set /a _OfflineMode=1
:: Can be ONSTART | ONLOGON | MINUTE(1-1439) | HOURLY(1-23) | DAILY(1-365) | WEEKLY(1-52) | MONTHLY(1-12)
set "_TaskFrequency=ONLOGON"
:: Can be integers in the range shown above
set /a _TaskModifier=1
::==========================================
:: Set Parameters for KMS Server
:: Custom Windows ePID
set "_WindowsEPID=03612-00206-471-452343-03-1033-14393.0000-1082018"
:: Custom Windows 10 Enterprise G/GN ePID
set "_WindowsGEPID=03612-00206-471-452343-03-1033-14393.0000-1082018"
:: Custom Office 2010 ePID
set "_Office2010EPID=03612-00096-199-303490-03-1033-14393.0000-1082018"
:: Custom Office 2013 ePID
set "_Office2013EPID=03612-00206-234-394838-03-1033-14393.0000-1082018"
:: Custom Office 2016 ePID
set "_Office2016EPID=03612-00206-437-938923-03-1033-14393.0000-1082018"
:: Can be Custom HardwareID obtained from a Real KMS Server Host
set "_HardwareID=3A1C049600B60076"
:: Can be 0 (Custom ePIDs) | 1 (Randomized ePIDs for every Session) | 2 (Randomized ePIDs for every Request)
set /a _RandomLevel=0
:: Can be (15 to 43200) minutes; Default - 2 hours, Maximum - 30 days
set /a _KMSActivationInterval=43200
:: Can be (15 to 43200) minutes; Default - 7 days, Maximum - 30 days
set /a _KMSRenewalInterval=43200
::==========================================
:: Set Parameters for KMS Client
:: _KMSHost Can be (0-255.0-255.0-255.0-255), but NOT 127.x.x.x/Localhost IPs [Offline Mode] | KMS-ServerName/IP [Online Mode]
set "_KMSHost=172.16.0.4"
set "_KMSLocalHost=127.0.0.2"
:: Can be (1 to 65535) [Offline Mode]; 1688 [Online Mode]
set /a _KMSPort=1686
::==========================================
:: Set Registry Key for DLL Hook
set "_regKey=HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options"
::==========================================
:: Registry Keys for SPP and OSPP
set "_hkSPP=HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform"
set "_huSPP=HKEY_USERS\S-1-5-20\Software\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform"
set "_hkOSPP=HKLM\SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform"
::==========================================
:: Set KMS Genuine Ticket Validation Parameters
:: Can be 0 (Enable Genuine Ticket) | 1 (Disable Genuine Ticket)
set /a _KMSNoGenTicket=1
:: Registry Key for KMS Genuine Ticket
set "_KMSGenuineKey=HKLM\SOFTWARE\Policies\Microsoft\Windows NT\CurrentVersion\Software Protection Platform"
::==========================================
:: Go to the Path of the Script
pushd "%_FileDir%"

:: Check if [Office 2010 on Windows XP SP3 or Later] OR [Office 2013 or Later on Windows 7 / Server 2008 R2] is Installed
wmic path OfficeSoftwareProtectionService get Version %_Nul_1_2% && (
  set /a _OSPS=1
) || (
  set /a _OSPS=0
)

:: Check if Office products are ACTUALLY installed
for %%G in (14,15,16) do (
  call :OfficeDetect %%G
)

:: Get Architecture of the OS installed; OS Locale Independent from Windows XP / Server 2003 and Later
for /f "tokens=2 delims==" %%G in ('wmic path Win32_Processor get AddressWidth /value') do (
  set "_OSarch=%%G-bit"
)

:: Visual Studio Activation
call :VisualStudio "12.0" "InstallDir" "" "87DQC-G8CYR-CRPJ4-QX9K8-RFV2B" "06181" "2013 Ultimate"
call :VisualStudio "14.0" "InstallDir" "" "HM6NR-QXX7C-DFW2Y-8B82K-WTYJV" "07060" "2015 Enterprise"
call :VisualStudio "SxS\VS7" "15.0" "Common7\IDE" "NJVYC-BMHX2-G77MM-4XJMR-6Q8QF" "08860" "2017 Enterprise"

:: Get Windows OS build number
for /f "tokens=2 delims==" %%G in ('wmic path Win32_OperatingSystem get BuildNumber /value') do (
  set /a _WinBuild=%%G
)

:: Define installed Edition for Windows 10 build 1607 or later
if %_WinBuild% LSS 14393 goto :Main
:: Get Edition based on active CBS package, or fall back to Dism CurrentEdition
set "_CBS=HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\Packages"
set "_Pattern=Microsoft-Windows-*Edition~31bf3856ad364e35"
set "_EditionPkg=NUL"
for /f "tokens=8 delims=\" %%G in ('reg query "%_CBS%" /f "%_Pattern%" /k %_Nul_2e% ^| find /i "CurrentVersion"') do (
  reg query "%_CBS%\%%G" /v "CurrentState" %_Nul_2% | find /i "0x70" %_Nul_1% && (
    for /f "tokens=3 delims=-~" %%H in ('echo %%G') do set "_EditionPkg=%%H"
  )
)
if /i "%_EditionPkg:~-7%"=="Edition" (
  set "_Edition=%_EditionPkg:~0,-7%"
) else (
  for /f "tokens=3 delims=: " %%G in ('dism /English /Online /Get-CurrentEdition %_Nul_2e% ^| find /i "Current Edition :"') do (
    set "_Edition=%%G"
  )
)
:: Get Edition based on current installed product key
for /f "tokens=2 delims==" %%G in ('"wmic path SoftwareLicensingProduct where (Name like 'Windows%%' and PartialProductKey is not NULL) get LicenseFamily /value"') do if not errorlevel 1 (
  set "_EditionWMI=%%G"
)
if not defined _EditionWMI (
  if %_WinBuild% GEQ 17063 (
    for /f "skip=2 tokens=3" %%G in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionId') do (
      set "_Edition=%%G"
    )
  )
  goto :Main
)
:: Exclude Windows 10 S
for %%G in (Cloud,CloudN) do (
  if /i "%_EditionWMI%"=="%%G" goto :Main
)
set "_Edition=%_EditionWMI%"
::==========================================
:Main
:: Goto Main blocks according to Windows BuildNumber
if %_WinBuild% GEQ 9600 (
  REM NO parenthesis or brackets in echo messages
  echo Operating System: Windows 8.1 or Later
  goto :Win8.1AndLater
) else if %_WinBuild% GEQ 2600 (
  echo Operating System: Windows 8 or Earlier
  goto :Win8AndEarlier
) else (
  echo KMS_VL_ALL is NOT supported on this OS.
  echo.
  echo Closing in 5 Seconds...
  ping 127.0.0.1 -n 6 %_Nul_1_2%
  exit /b
)
::==========================================
:Close
:: Create/Delete Auto-Renewal Task based on parameter; Windows XP SP3 or Later Compatible
if %_Task% EQU 1 (
  schtasks /query /fo list %_Nul_2% | findstr /i "%_TaskName%" %_Nul_1% && (
    schtasks /delete /tn "%_TaskName%" /f %_Nul_1_2%
  )
  if /i %_TaskFrequency% EQU ONSTART (
    schtasks /create /tn "%_TaskName%" /ru "SYSTEM" /sc "%_TaskFrequency%" /tr "%_FileName% -renewalonly" %_Nul_1_2% && (
      echo.
      echo Auto-Renewal Task is Created.
    )
  ) else if /i %_TaskFrequency% EQU ONLOGON (
    schtasks /create /tn "%_TaskName%" /ru "SYSTEM" /sc "%_TaskFrequency%" /tr "%_FileName% -renewalonly" %_Nul_1_2% && (
      echo.
      echo Auto-Renewal Task is Created.
    )
  ) else (
    schtasks /create /tn "%_TaskName%" /ru "SYSTEM" /sc "%_TaskFrequency%" /mo "%_TaskModifier%" /tr "%_FileName% -renewalonly" %_Nul_1_2% && (
      echo.
      echo Auto-Renewal Task is Created.
    )
  )
) else if %_Task% EQU 0 (
  schtasks /query /fo list %_Nul_2% | findstr /i "%_TaskName%" %_Nul_1% && (
    schtasks /delete /tn "%_TaskName%" /f %_Nul_1_2%
    echo.
    echo Auto-Renewal Task is Deleted.
  )
)
if %_Debug% EQU 0 (
  echo.
  echo Closing in 5 Seconds...
  ping 127.0.0.1 -n 6 %_Nul_1_2%
)
exit /b
::==========================================
:Win8.1AndLater
if %_OfflineMode% EQU 1 (
  REM Stop 'sppsvc'
  call :StopService "sppsvc"
  REM Symlink the DLL Injection file to system32 folder based on OS architecture
  mklink "%SystemRoot%\system32\SECOPatcher.dll" "%_FileDir%\%_OSarch%\SECOPatcher.dll" %_Nul_1_2%
  REM Check Read/Execute security permissions and grant them if necessary
  icacls "%SystemRoot%\system32\SECOPatcher.dll" /findsid *S-1-5-32-545 %_Nul_2% | find /i "SECOPatcher.dll" %_Nul_1% || (
    icacls "%SystemRoot%\system32\SECOPatcher.dll" /grant *S-1-5-32-545:RX %_Nul_1_2%
  )
  REM Create registry keys for DLL Patcher
  call :CreateIFEOEntry "SppExtComObj.exe"
  REM Add Firewall Exceptions for VLMCSD and Start KMS Server
  call :AddFirewallRule
  call :StartKMS
)

:: Enable/Disable KMS Genuine Ticket Validation for Windows 8.1 and later
if %_WinBuild% GEQ 9600 (
  call :KMSGenuineTicket
)

:: Call Windows and Office Activation Main Functions
call :SLSActivation
if %_OSPS% NEQ 0 (
  if %_OfflineMode% EQU 1 (
    REM Localhost IP should be used for Office 2010 on Windows 8.1 and later
    set "_KMSHost=%_KMSLocalHost%"
  )
  call :OSPSActivation
)
if %_OfflineMode% EQU 1 (
  REM Stop 'sppsvc'
  call :StopService "sppsvc"
  REM Stop KMS Server and Remove Firewall Exceptions for VLMCSD
  call :StopKMS
  call :RemoveFirewallRule
  REM Restore original security permissions if changed
  icacls "%SystemRoot%\system32\SECOPatcher.dll" /reset %_Nul_1_2%
  REM Delete the DLL Injection symlink from system32 folder
  if exist "%SystemRoot%\system32\SECOPatcher.dll" del /f /q "%SystemRoot%\system32\SECOPatcher.dll" %_Nul_1_2%
  REM Remove registry keys for DLL Hook
  call :RemoveIFEOEntry "SppExtComObj.exe"
  REM Start 'sppsvc'
  sc start sppsvc trigger=timer;sessionid=0 %_Nul_1_2%
)
call :Close & exit /b
::==========================================
:Win8AndEarlier
:: Exit if No Office 2010 product is installed on Windows XP SP3/Server 2003 R2
if %_OSPS% EQU 0 (
  if %_WinBuild% LSS 6000 (
    echo.
    echo No Office 2010 Product Detected...
    call :Close & exit /b
  )
)

if %_OfflineMode% EQU 1 (
  REM Localhost IP can be used for Windows 8 and earlier
  set "_KMSHost=%_KMSLocalHost%"
  REM Windows Vista do not support SetKeyManagementServicePort, so revert to default KMS port
  if %_WinBuild% LSS 7600 (
    set /a _KMSPort=1688
  )
  REM Add Firewall Exceptions for VLMCSD and Start KMS Server
  call :AddFirewallRule
  call :StartKMS
)

:: Call Windows and Office Activation Main Functions
if %_WinBuild% GEQ 6000 (
  call :SLSActivation
)
if %_OSPS% NEQ 0 (
  call :OSPSActivation
)

if %_OfflineMode% EQU 1 (
  REM Stop KMS Server and Remove Firewall Exceptions for VLMCSD
  call :StopKMS
  call :RemoveFirewallRule
)
call :Close & exit /b
::==========================================
:AddFirewallRule
:: Add VLMCSD KMS Exception to Windows Firewall
if %_WinBuild% LSS 6000 (
  netsh firewall delete allowedprogram "%_ServerPath%" %_Nul_1_2%
  netsh firewall add allowedprogram "%_ServerPath%" "vlmcsd" %_Nul_1_2%
) else (
  netsh advfirewall firewall delete rule name="vlmcsd" %_Nul_1_2%
  netsh advfirewall firewall add rule name="vlmcsd" dir=in action=allow profile=any program="%_ServerPath%" %_Nul_1_2%
)
exit /b
::==========================================
:RemoveFirewallRule
:: Remove VLMCSD KMS Exception from Windows Firewall
if %_WinBuild% LSS 6000 (
  netsh firewall delete allowedprogram "%_ServerPath%" %_Nul_1_2%
) else (
  netsh advfirewall firewall delete rule name="vlmcsd" %_Nul_1_2%
)
exit /b
::==========================================
:StartKMS
:: Start VLMCSD KMS Server
if %_RandomLevel% EQU 0 (
  cmd /c start /b "" "%_ServerPath%" -P %_KMSPort% -0 %_Office2010EPID% -3 %_Office2013EPID% -6 %_Office2016EPID% -w %_WindowsEPID% -G %_WindowsGEPID% -H %_HardwareID% -R %_KMSRenewalInterval% -A %_KMSActivationInterval% -T0 -e %_Nul_1_2%
) else (
  cmd /c start /b "" "%_ServerPath%" -r %_RandomLevel% -P %_KMSPort% -H %_HardwareID% -R %_KMSRenewalInterval% -A %_KMSActivationInterval% -T0 -e %_Nul_1_2%
)
:: Mind boggling BUG Fix; Windows Vista or earlier takes some time to start KMS Server which prevents Activation; So add delay for it to successfully start
if %_WinBuild% LSS 7600 (
  ping 127.0.0.1 -n 12 %_Nul_1_2%
)
exit /b
::==========================================
:StopKMS
:: Stop VLMCSD KMS Server
taskkill /t /f /IM vlmcsd.exe %_Nul_1_2%
exit /b
::==========================================
:KMSGenuineTicket
:: Enable/Disable KMS Genuine Ticket Validation registry key based on user parameter
reg add "%_KMSGenuineKey%" /v NoGenTicket /t REG_DWORD /d %_KMSNoGenTicket% /f %_Nul_1_2%
exit /b
::==========================================
:CreateIFEOEntry
:: Create DLL Injection Registry key based on parameter
reg add "%_regKey%\%~1" /f /v "Debugger" /t REG_SZ /d "rundll32.exe SECOPatcher.dll,PatcherMain" %_Nul_1_2%
exit /b
::==========================================
:RemoveIFEOEntry
:: Remove DLL Injection Registry key based on parameter
if '%~1' NEQ 'osppsvc.exe' (
  reg delete "%_regKey%\%~1" /f %_Nul_1_2%
)
if '%~1' EQU 'osppsvc.exe' (
  reg delete "%_regKey%\%~1" /f /v "Debugger" %_Nul_1_2%
)
exit /b
::==========================================
:StopService
:: Stop service based on parameter
sc query "%1" | findstr /i "STOPPED" %_Nul_1_2% || (
  net stop "%1" /y %_Nul_1_2%
)
sc query "%1" | findstr /i "STOPPED" %_Nul_1_2% || (
  sc stop "%1" %_Nul_1_2%
)
exit /b
::==========================================
:SLSActivation
reg delete "%_hkSPP%\55c92734-d682-4d71-983e-d6ec3f16059f" /f %_Nul_1_2%
reg delete "%_hkSPP%\0ff1ce15-a989-479d-af46-f275c6370663" /f %_Nul_1_2%
set "_MicrosoftProduct=SoftwareLicensingProduct"
set "_MicrosoftService=SoftwareLicensingService"

:: Detect if Office 2013 [Volume Licensed] or Later is Installed
wmic path %_MicrosoftProduct% where (Description like '%%KMSCLIENT%%') get Name /value %_Nul_2% | findstr /i "Office" %_Nul_1% && (
  set /a _OfficeVL=1
) || (
  set /a _OfficeVL=0
  if %_WinBuild% GEQ 9200 (
    echo.
    echo No Office 2013 or Later Volume License Product Detected.
  )
)

:: Detect if installed Windows supports KMS Activation; Exit if there are no VolumeLicensed Windows or Office
wmic path %_MicrosoftProduct% where (Description like '%%KMSCLIENT%%') get Name /value %_Nul_2% | findstr /i "Windows" %_Nul_1% || (
  echo.
  echo No Windows Volume License Product Detected.
  if %_OfficeVL% EQU 0 (
    exit /b
  )
)

:: Check if GVLK is installed for Windows
wmic path %_MicrosoftProduct% where (Description like '%%KMSCLIENT%%' and PartialProductKey is not NULL) get Name /value %_Nul_2% | findstr /i "Windows" %_Nul_1% && (
  set /a _WindowsGVLK=1
) || (
  set /a _WindowsGVLK=0
)

:: Call Common Core Activation Routines
call :CommonSLSandOSPS
reg delete "%_hkSPP%\55c92734-d682-4d71-983e-d6ec3f16059f" /f %_Nul_1_2%
reg delete "%_hkSPP%\0ff1ce15-a989-479d-af46-f275c6370663" /f %_Nul_1_2%
reg delete "%_huSPP%\55c92734-d682-4d71-983e-d6ec3f16059f" /f %_Nul_1_2%
reg delete "%_huSPP%\0ff1ce15-a989-479d-af46-f275c6370663" /f %_Nul_1_2%
exit /b
::==========================================
:OSPSActivation
reg delete "%_hkOSPP%\59a52881-a989-479d-af46-f275c6370663" /f %_Nul_1_2%
reg delete "%_hkOSPP%\0ff1ce15-a989-479d-af46-f275c6370663" /f %_Nul_1_2%
set "_MicrosoftProduct=OfficeSoftwareProtectionProduct"
set "_MicrosoftService=OfficeSoftwareProtectionService"

:: Determine if installed Office product is Retail or VL version; Exit if no VolumeLicensed Office is detected
wmic path %_MicrosoftProduct% where (Description like '%%KMSCLIENT%%') get Name %_Nul_1_2% || (
  if %_WinBuild% LSS 9200 (
    echo.
    echo No Office 2010 or Later Volume License Product Detected.
    exit /b
  ) else (
    echo.
    echo No Office 2010 Volume License Product Detected.
    exit /b
  )
)

:: Call Common Core Activation Routines
call :CommonSLSandOSPS
reg delete "%_hkOSPP%\59a52881-a989-479d-af46-f275c6370663" /f %_Nul_1_2%
reg delete "%_hkOSPP%\0ff1ce15-a989-479d-af46-f275c6370663" /f %_Nul_1_2%
exit /b
::==========================================
:CommonSLSandOSPS
:: Get SoftwareLicensingService/OfficeSoftwareProtectionService version to set 'KMSHost' and 'KMSPort' values
for /f "tokens=2 delims==" %%G in ('"wmic path %_MicrosoftService% get Version /value"') do (
  set "_Version=%%G"
)
wmic path %_MicrosoftService% where version='%_Version%' call SetKeyManagementServiceMachine MachineName="%_KMSHost%" %_Nul_1_2%
wmic path %_MicrosoftService% where version='%_Version%' call SetKeyManagementServicePort %_KMSPort% %_Nul_1_2%
:: This is available only on SoftwareLicensingService version 6.2 and later; Not available for OfficeSoftwareProtectionService
if %_WinBuild% GEQ 9200 (
  wmic path %_MicrosoftService% where version='%_Version%' call SetVLActivationTypeEnabled 2 %_Nul_1_2%
)

:: For all the supported KMS Clients in SoftwareLicensingProduct/OfficeSoftwareProtectionProduct call 'CheckProduct'
for /f "tokens=2 delims==" %%G in ('"wmic path %_MicrosoftProduct% where (Description like '%%KMSCLIENT%%') get ID /value"') do (
  set "_ActivationID=%%G"
  call :CheckProduct
)

:: Clear KMS Server details from KMS Client, for Offline mode only
if %_OfflineMode% EQU 1 (
  wmic path %_MicrosoftService% where version='%_Version%' call ClearKeyManagementServiceMachine %_Nul_1_2%
  wmic path %_MicrosoftService% where version='%_Version%' call ClearKeyManagementServicePort %_Nul_1_2%
  wmic path %_MicrosoftService% where version='%_Version%' call DisableKeyManagementServiceDnsPublishing 1 %_Nul_1_2%
  wmic path %_MicrosoftService% where version='%_Version%' call DisableKeyManagementServiceHostCaching 1 %_Nul_1_2%
  REM This is available only on SoftwareLicensingService version 6.2 and later; Not available for OfficeSoftwareProtectionService
  if %_WinBuild% GEQ 9200 (
    wmic path %_MicrosoftService% where version='%_Version%' call ClearVLActivationTypeEnabled %_Nul_1_2%
  )
)
exit /b
::==========================================
:CheckProduct
:: Detect Office Products
set /a _OfficeSLP=0
wmic path %_MicrosoftProduct% where ID='%_ActivationID%' get Name /value | findstr /i "Office" %_Nul_1% && (
  set /a _OfficeSLP=1
)
:: If Detected KMS Client is already activated earlier OR has GVLK, call Activate function
if %_OfficeSLP% EQU 0 wmic path %_MicrosoftProduct% where ID='%_ActivationID%' get LicenseStatus | findstr "1" %_Nul_1_2% && (
  call :Activate %_ActivationID%
  exit /b
)
wmic path %_MicrosoftProduct% where (PartialProductKey is not NULL) get ID | findstr /i "%_ActivationID%" %_Nul_1_2% && (
  call :Activate %_ActivationID%
  exit /b
)
:: Skip for Unnecessary Products
if %_OfficeSLP% EQU 0 (
  if %_WindowsGVLK% EQU 1 (
    exit /b
  )
)
:: Ugly hack for multiple Windows 10 SKU-IDs
for %%G in (
  b71515d9-89a2-4c60-88c8-656fbcca7f3a
  5b2add49-b8f4-42e0-a77c-adad4efeeeb1
  af43f7f0-3b1e-4266-a123-1fdb53f4323b
  075aca1f-05d7-42e5-a3ce-e349e7be7078
  2cf5af84-abab-4ff0-83f8-f040fb2576eb
  11a37f09-fb7f-4002-bd84-f3ae71d11e90
  43f2ab05-7c87-4d56-b27c-44d0f9a3dabd
  6ae51eeb-c268-4a21-9aae-df74c38b586d
  ff808201-fec6-4fd4-ae16-abbddade5706
  34260150-69ac-49a3-8a0d-4a403ab55763
  903663f7-d2ab-49c9-8942-14aa9e0a9c72
  4dfd543d-caa6-4f69-a95f-5ddfe2b89567
  5fe40dd6-cf1f-4cf2-8729-92121ac2e997
  2cc171ef-db48-4adc-af09-7c574b37f139
) do (
  if /i '%_ActivationID%' EQU '%%G' (
    exit /b
  )
)

:: If Detected KMS Client do not have GVLK, do checks for permanent activation, then install GVLK and activate it
for /f "tokens=3 delims==, " %%G in ('"wmic path %_MicrosoftProduct% where ID='%_ActivationID%' get Name /value"') do (
  set "_ProductName=%%G"
)
if '%_ProductName%' EQU '19' (
  if %_Office16% EQU 0 (
    exit /b
  )
  call :CheckOffice19 %_ActivationID%
  exit /b
) else if '%_ProductName%' EQU '16' (
  if %_Office16% EQU 0 (
    exit /b
  )
  call :CheckOffice16 %_ActivationID%
  exit /b
) else if '%_ProductName%' EQU '15' (
  if %_Office15% EQU 0 (
    exit /b
  )
  call :CheckOffice15 %_ActivationID%
  exit /b
) else if '%_ProductName%' EQU '14' (
  if %_Office14% EQU 0 (
    exit /b
  )
  call :CheckOffice14 %_ActivationID%
  exit /b
)

:: Pre Windows 10 build 1607 do not have combined editions
if not defined _Edition (
  call :CheckWindows %_ActivationID%
  exit /b
)
:: Ugly hack for combined editions in Windows 10 installation
if /i '%_ActivationID%' EQU '2de67392-b7a7-462a-b1ca-108dd189f588' (
  if /i %_Edition% NEQ Professional (
    exit /b
  )
)
if /i '%_ActivationID%' EQU 'a80b5abf-76ad-428b-b05d-a47d2dffeebf' (
  if /i %_Edition% NEQ ProfessionalN (
    exit /b
  )
)
if /i '%_ActivationID%' EQU '82bbc092-bc50-4e16-8e18-b74fc486aec3' (
  if /i %_Edition% NEQ ProfessionalWorkstation (
    exit /b
  )
)
if /i '%_ActivationID%' EQU '4b1571d3-bafb-4b40-8087-a961be2caf65' (
  if /i %_Edition% NEQ ProfessionalWorkstationN (
    exit /b
  )
)
if /i '%_ActivationID%' EQU '3f1afc82-f8ac-4f6c-8005-1d233e606eee' (
  if /i %_Edition% NEQ ProfessionalEducation (
    exit /b
  )
)
if /i '%_ActivationID%' EQU '5300b18c-2e33-4dc2-8291-47ffcec746dd' (
  if /i %_Edition% NEQ ProfessionalEducationN (
    exit /b
  )
)
if /i '%_ActivationID%' EQU '73111121-5638-40f6-bc11-f1d7b0d64300' (
  if /i %_Edition% NEQ Enterprise (
    exit /b
  )
)
if /i '%_ActivationID%' EQU 'e272e3e2-732f-4c65-a8f0-484747d0d947' (
  if /i %_Edition% NEQ EnterpriseN (
    exit /b
  )
)
if /i '%_ActivationID%' EQU 'e0c42288-980c-4788-a014-c080d2e1926e' (
  if /i %_Edition% NEQ Education (
    exit /b
  )
)
if /i '%_ActivationID%' EQU '3c102355-d027-42c6-ad23-2e7ef8a02585' (
  if /i %_Edition% NEQ EducationN (
    exit /b
  )
)
if /i '%_ActivationID%' EQU 'e4db50ea-bda1-4566-b047-0ca50abc6f07' (
  if /i %_Edition% NEQ ServerRdsh (
    exit /b
  )
)
if /i '%_ActivationID%' EQU 'ec868e65-fadf-4759-b23e-93fe37f2cc29' (
  if /i %_Edition% NEQ ServerRdsh (
    exit /b
  )
)
if /i '%_ActivationID%' EQU '58e97c99-f377-4ef1-81d5-4ad5522b5fd8' (
  if /i %_Edition% NEQ Core (
    exit /b
  )
)
if /i '%_ActivationID%' EQU 'cd918a57-a41b-4c82-8dce-1a538e221a83' (
  if /i %_Edition% NEQ CoreSingleLanguage (
    exit /b
  )
)
call :CheckWindows %_ActivationID%
exit /b
::==========================================
:CheckWindows
wmic path %_MicrosoftProduct% where (LicenseStatus='1' and GracePeriodRemaining='0') get Name %_Nul_2% | findstr /i "Windows" %_Nul_1% && (
  echo.
  echo Detected Windows %_ProductName% is permanently activated.
  exit /b
)

:: If Windows is not permanently activated, Install GVLK and Activate
call :SelectKey %1
exit /b
::==========================================
:CheckOffice19
:: Ugly check for old volume licenses of Office 2019 Pro SKUs
if /i '%1' EQU '0bc88885-718c-491d-921f-6f214349e79c' (
  exit /b
)
if /i '%1' EQU 'fc7c4d0c-2e85-4bb9-afd4-01ed1476b5e9' (
  exit /b
)
if /i '%1' EQU '500f6619-ef93-4b75-bcb4-82819998a3ca' (
  exit /b
)
if /i '%1' EQU '85dd8b5f-eaa4-4af3-a628-cce9e77c9a03' (
  wmic path %_MicrosoftProduct% where 'PartialProductKey is not NULL' get ID | findstr /i "0bc88885-718c-491d-921f-6f214349e79c" %_Nul_1_2% && (
    exit /b
  )
)
if /i '%1' EQU '2ca2bf3f-949e-446a-82c7-e25a15ec78c4' (
  wmic path %_MicrosoftProduct% where 'PartialProductKey is not NULL' get ID | findstr /i "fc7c4d0c-2e85-4bb9-afd4-01ed1476b5e9" %_Nul_1_2% && (
    exit /b
  )
)
if /i '%1' EQU '5b5cf08f-b81a-431d-b080-3450d8620565' (
  wmic path %_MicrosoftProduct% where 'PartialProductKey is not NULL' get ID | findstr /i "500f6619-ef93-4b75-bcb4-82819998a3ca" %_Nul_1_2% && (
    exit /b
  )
)
if /i '%1' EQU '85dd8b5f-eaa4-4af3-a628-cce9e77c9a03' (
  call :CheckOffice "%1" "19ProPlus2019VL_MAK_AE" "Office ProPlus 2019" "19ProPlus2019XC2RVL_MAKC2R" "Office ProPlus 2019 C2R"
  exit /b
)
if /i '%1' EQU '6912a74b-a5fb-401a-bfdb-2e3ab46f4b02' (
  call :CheckOffice "%1" "19Standard2019VL_MAK_AE" "Office Standard 2019"
  exit /b
)
if /i '%1' EQU '2ca2bf3f-949e-446a-82c7-e25a15ec78c4' (
  call :CheckOffice "%1" "19ProjectPro2019VL_MAK_AE" "Project Pro 2019" "19ProjectPro2019XC2RVL_MAKC2R" "Project Pro 2019 C2R"
  exit /b
)
if /i '%1' EQU '1777f0e3-7392-4198-97ea-8ae4de6f6381' (
  call :CheckOffice "%1" "19ProjectStd2019VL_MAK_AE" "Project Standard 2019"
  exit /b
)
if /i '%1' EQU '5b5cf08f-b81a-431d-b080-3450d8620565' (
  call :CheckOffice "%1" "19VisioPro2019VL_MAK_AE" "Visio Pro 2019" "19VisioPro2019XC2RVL_MAKC2R" "Visio Pro 2019 C2R"
  exit /b
)
if /i '%1' EQU 'e06d7df3-aad0-419d-8dfb-0ac37e2bdf39' (
  call :CheckOffice "%1" "19VisioStd2019VL_MAK_AE" "Visio Standard 2019"
  exit /b
)
call :SelectKey %1
exit /b
::==========================================
:CheckOffice16
if /i '%1' EQU '9caabccb-61b1-4b4b-8bec-d10a3c3ac2ce' (
  call :CheckOffice "%1" "16MondoVL_MAK" "Office Mondo 2016"
  exit /b
)
if /i '%1' EQU 'd450596f-894d-49e0-966a-fd39ed4c4c64' (
  call :CheckOffice "%1" "16ProPlusVL_MAK" "Office ProPlus 2016"
  exit /b
)
if /i '%1' EQU 'dedfa23d-6ed1-45a6-85dc-63cae0546de6' (
  call :CheckOffice "%1" "16StandardVL_MAK" "Office Standard 2016"
  exit /b
)
if /i '%1' EQU '4f414197-0fc2-4c01-b68a-86cbb9ac254c' (
  call :CheckOffice "%1" "16ProjectProVL_MAK" "Project Pro 2016"
  exit /b
)
if /i '%1' EQU 'da7ddabc-3fbe-4447-9e01-6ab7440b4cd4' (
  call :CheckOffice "%1" "16ProjectStdVL_MAK" "Project Standard 2016"
  exit /b
)
if /i '%1' EQU '6bf301c1-b94a-43e9-ba31-d494598c47fb' (
  call :CheckOffice "%1" "16VisioProVL_MAK" "Visio Pro 2016"
  exit /b
)
if /i '%1' EQU 'aa2a7821-1827-4c2c-8f1d-4513a34dda97' (
  call :CheckOffice "%1" "16VisioStdVL_MAK" "Visio Standard 2016"
  exit /b
)
if /i '%1' EQU '829b8110-0e6f-4349-bca4-42803577788d' (
  call :CheckOffice "%1" "16ProjectProXC2RVL_MAKC2R" "Project Pro 2016 C2R"
  exit /b
)
if /i '%1' EQU 'cbbaca45-556a-4416-ad03-bda598eaa7c8' (
  call :CheckOffice "%1" "16ProjectStdXC2RVL_MAKC2R" "Project Standard 2016 C2R"
  exit /b
)
if /i '%1' EQU 'b234abe3-0857-4f9c-b05a-4dc314f85557' (
  call :CheckOffice "%1" "16VisioProXC2RVL_MAKC2R" "Visio Pro 2016 C2R"
  exit /b
)
if /i '%1' EQU '361fe620-64f4-41b5-ba77-84f8e079b1f7' (
  call :CheckOffice "%1" "16VisioStdXC2RVL_MAKC2R" "Visio Standard 2016 C2R"
  exit /b
)
call :SelectKey %1
exit /b
::==========================================
:CheckOffice15
if /i '%1' EQU 'dc981c6b-fc8e-420f-aa43-f8f33e5c0923' (
  call :CheckOffice "%1" "MondoVL_MAK" "Office Mondo 2013"
  exit /b
)
if /i '%1' EQU 'b322da9c-a2e2-4058-9e4e-f59a6970bd69' (
  call :CheckOffice "%1" "ProPlusVL_MAK" "Office ProPlus 2013"
  exit /b
)
if /i '%1' EQU 'b13afb38-cd79-4ae5-9f7f-eed058d750ca' (
  call :CheckOffice "%1" "StandardVL_MAK" "Office Standard 2013"
  exit /b
)
if /i '%1' EQU '4a5d124a-e620-44ba-b6ff-658961b33b9a' (
  call :CheckOffice "%1" "ProjectProVL_MAK" "Project Pro 2013"
  exit /b
)
if /i '%1' EQU '427a28d1-d17c-4abf-b717-32c780ba6f07' (
  call :CheckOffice "%1" "ProjectStdVL_MAK" "Project Standard 2013"
  exit /b
)
if /i '%1' EQU 'e13ac10e-75d0-4aff-a0cd-764982cf541c' (
  call :CheckOffice "%1" "VisioProVL_MAK" "Visio Pro 2013"
  exit /b
)
if /i '%1' EQU 'ac4efaf0-f81f-4f61-bdf7-ea32b02ab117' (
  call :CheckOffice "%1" "VisioStdVL_MAK" "Visio Standard 2013"
  exit /b
)
call :SelectKey %1
exit /b
::==========================================
:CheckOffice14
set "_VisioPremium="
set "_VisioPro="
for /f "tokens=2 delims==" %%G in ('"wmic path %_MicrosoftProduct% where (Name like '%%OfficeVisioPrem-MAK%%') get LicenseStatus /value" %_Nul_2e%') do (
  set /a _VisioPremium=%%G
)
for /f "tokens=2 delims==" %%G in ('"wmic path %_MicrosoftProduct% where (Name like '%%OfficeVisioPro-MAK%%') get LicenseStatus /value" %_Nul_2e%') do (
  set /a _VisioPro=%%G
)
if /i '%1' EQU '09ed9640-f020-400a-acd8-d7d867dfd9c2' (
  call :CheckOffice "%1" "Mondo-MAK" "Office Mondo 2010"
  exit /b
)
if /i '%1' EQU '6f327760-8c5c-417c-9b61-836a98287e0c' (
  call :CheckOffice "%1" "ProPlus-MAK" "Office ProPlus 2010" "ProPlusAcad-MAK" "Office Professional Academic 2010"
  exit /b
)
if /i '%1' EQU '9da2a678-fb6b-4e67-ab84-60dd6a9c819a' (
  call :CheckOffice "%1" "Standard-MAK" "Office Standard 2010"
  exit /b
)
if /i '%1' EQU 'ea509e87-07a1-4a45-9edc-eba5a39f36af' (
  call :CheckOffice "%1" "SmallBusBasics-MAK" "Office Home and Business 2010"
  exit /b
)
if /i '%1' EQU 'df133ff7-bf14-4f95-afe3-7b48e7e331ef' (
  call :CheckOffice "%1" "ProjectPro-MAK" "Project Pro 2010"
  exit /b
)
if /i '%1' EQU '5dc7bf61-5ec9-4996-9ccb-df806a2d0efe' (
  call :CheckOffice "%1" "ProjectStd-MAK" "Project Standard 2010"
  exit /b
)
if /i '%1' EQU '92236105-bb67-494f-94c7-7f7a607929bd' (
  call :CheckOffice "%1" "VisioPrem-MAK" "Visio Premium 2010" "VisioPro-MAK" "Visio Pro 2010"
  exit /b
)
if defined _VisioPremium exit /b
if /i '%1' EQU 'e558389c-83c3-4b29-adfe-5e4d7f46c358' (
  call :CheckOffice "%1" "VisioPro-MAK" "Visio Pro 2010" "VisioStd-MAK" "Visio Standard 2010"
  exit /b
)
if defined _VisioPro exit /b
if /i '%1' EQU '9ed833ff-4f92-4f36-b370-8683a4f13275' (
  call :CheckOffice "%1" "VisioStd-MAK" "Visio Standard 2010"
  exit /b
)
call :SelectKey %1
exit /b
::==========================================
:CheckOffice
set /a _License=0
set /a _License2=0
for /f "tokens=2 delims==" %%G in ('"wmic path %_MicrosoftProduct% where (Name like '%%Office%~2%%') get LicenseStatus /VALUE" %_Nul_2e%') do (
  set /a _License=%%G
)
if "%~4" NEQ "" (
  for /f "tokens=2 delims==" %%G in ('"wmic path %_MicrosoftProduct% where (Name like '%%Office%~4%%') get LicenseStatus /VALUE" %_Nul_2e%') do (
    set /a _License2=%%G
  )
)
if "%_License2%" EQU "1" (
  echo Detected %5 is permanently MAK activated.
  exit /b
)
if "%_License%" EQU "1" (
  echo Detected %3 is permanently MAK activated.
  exit /b
)
:: If Office product is not permanently activated, Install GVLK and Activate
call :SelectKey %1
exit /b
::==========================================
:OfficeDetect
set _Office%1=0
for /f "tokens=2*" %%G in ('"reg query HKLM\SOFTWARE\Microsoft\Office\%1.0\Common\InstallRoot /v Path" %_Nul_2e%') do if exist "%%H\OSPP.VBS" (
  set _Office%1=1
)
for /f "tokens=2*" %%G in ('"reg query HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\%1.0\Common\InstallRoot /v Path" %_Nul_2e%') do if exist "%%H\OSPP.VBS" (
  set _Office%1=1
)
if exist "%ProgramFiles%\Microsoft Office\Office%1\OSPP.VBS" (
  set _Office%1=1
)
if exist "%ProgramFiles(x86)%\Microsoft Office\Office%1\OSPP.VBS" (
  set _Office%1=1
)
exit /b
::==========================================
:VisualStudio
:: Clear variable to avoid confliction if multiple Visual Studio versions installed
set "_VS="
if "%_OSarch%" EQU "64-bit" (
  set "_H64=SOFTWARE\WOW6432Node"
) else (
  set "_H64=SOFTWARE"
)
for /f "skip=2 tokens=2*" %%G in ('"reg query HKLM\%_H64%\Microsoft\VisualStudio\%~1 /v %~2" %_Nul_2e%') do (
  set "_VS=%%H%~3"
)
if not exist "%_VS%\StorePID.exe" (
  exit /b
)
start "" /b "%_VS%\StorePID.exe" %~4 %~5 && (
  echo Visual Studio %~6 activated successfully.
  echo.
) || (
  echo Visual Studio %~6 activation failed.
  echo.
)
exit /b
::==========================================
:Activate
:: Clear any manually set KMSHostIP and KMSPort with /skms or /sethst; Since they override KMSHostIP and KMSPort values set for SLS/OSPS
wmic path %_MicrosoftProduct% where ID='%1' call ClearKeyManagementServiceMachine %_Nul_1_2%
wmic path %_MicrosoftProduct% where ID='%1' call ClearKeyManagementServicePort %_Nul_1_2%

:: Call Activate method of the corresponding KMS Client
for /f "tokens=2 delims==" %%G in ('"wmic path %_MicrosoftProduct% where ID='%1' get Name /value"') do (
  echo.
  echo Attempting to activate %%G
)
wmic path %_MicrosoftProduct% where ID='%1' call Activate %_Nul_1_2%
set ERRORCODE=%ERRORLEVEL%

:: Get Remaining Grace Period of the KMS Client
for /f "tokens=2 delims==" %%G in ('"wmic path %_MicrosoftProduct% where ID='%1' get GracePeriodRemaining /value"') do (
  set /a _gprMinutes=%%G
  set /a _gprDays=%%G/1440
)
if %_gprMinutes% EQU 43200 (
  if %_WinBuild% EQU 9200 if %_OfficeSLP% EQU 0 (
    echo Windows 8 Core/ProfessionalWMC Activation Successful
    echo Remaining Period: %_gprDays% days ^(%_gprMinutes% minutes^)
    exit /b
  )
)
if %_gprMinutes% EQU 64800 (
  echo Windows Core/ProfessionalWMC Activation Successful
  echo Remaining Period: %_gprDays% days ^(%_gprMinutes% minutes^)
  exit /b
)
if %_gprMinutes% EQU 216000000 (
  if %_WinBuild% GEQ 15063 (
    echo Windows 10 Enterprise G/GN Activation Successful
    echo Remaining Period: %_gprDays% days ^(%_gprMinutes% minutes^)
    exit /b
  )
)
if %_gprMinutes% EQU 259200 (
  echo Activation Successful
) else (
  call cmd /c exit /b %ERRORCODE%
  echo Activation Failed: 0x%=ExitCode%
)
echo Remaining Period: %_gprDays% days ^(%_gprMinutes% minutes^)
exit /b
::==========================================
:SelectKey
:: Select GenericVolumeLicenseKey based on Activation-ID (SKU-ID) and Install it, if found
for /f "tokens=2 delims==" %%G in ('"wmic path %_MicrosoftProduct% where ID='%1' get Name /value"') do (
  set "_Name=%%G"
  echo.
  echo Searching GenericVolumeLicenseKey for %%G
  goto :%1 %_Nul_2% || goto :KeyNotFound
)
::==========================================
:: Office 2019 Professional Plus
:85dd8b5f-eaa4-4af3-a628-cce9e77c9a03
set "_key=NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP"
goto :InstallKey
:: Office 2019 Standard
:6912a74b-a5fb-401a-bfdb-2e3ab46f4b02
set "_key=6NWWJ-YQWMR-QKGCB-6TMB3-9D9HK"
goto :InstallKey
:: Project 2019 Professional
:2ca2bf3f-949e-446a-82c7-e25a15ec78c4
set "_key=B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B"
goto :InstallKey
:: Project 2019 Standard
:1777f0e3-7392-4198-97ea-8ae4de6f6381
set "_key=C4F7P-NCP8C-6CQPT-MQHV9-JXD2M"
goto :InstallKey
:: Visio 2019 Professional
:5b5cf08f-b81a-431d-b080-3450d8620565
set "_key=9BGNQ-K37YR-RQHF2-38RQ3-7VCBB"
goto :InstallKey
:: Visio 2019 Standard
:e06d7df3-aad0-419d-8dfb-0ac37e2bdf39
set "_key=7TQNQ-K3YQQ-3PFH7-CCPPM-X4VQ2"
goto :InstallKey
:: Access 2019
:9e9bceeb-e736-4f26-88de-763f87dcc485
set "_key=9N9PT-27V4Y-VJ2PD-YXFMF-YTFQT"
goto :InstallKey
:: Excel 2019
:237854e9-79fc-4497-a0c1-a70969691c6b
set "_key=TMJWT-YYNMB-3BKTF-644FC-RVXBD"
goto :InstallKey
:: Outlook 2019
:c8f8a301-19f5-4132-96ce-2de9d4adbd33
set "_key=7HD7K-N4PVK-BHBCQ-YWQRW-XW4VK"
goto :InstallKey
:: PowerPoint 2019
:3131fd61-5e4f-4308-8d6d-62be1987c92c
set "_key=RRNCX-C64HY-W2MM7-MCH9G-TJHMQ"
goto :InstallKey
:: Publisher 2019
:9d3e4cca-e172-46f1-a2f4-1d2107051444
set "_key=G2KWX-3NW6P-PY93R-JXK2T-C9Y9V"
goto :InstallKey
:: Skype for Business 2019
:734c6c6e-b0ba-4298-a891-671772b2bd1b
set "_key=NCJ33-JHBBY-HTK98-MYCV8-HMKHJ"
goto :InstallKey
:: Word 2019
:059834fe-a8ea-4bff-b67b-4d006b5447d3
set "_key=PBX3G-NWMT6-Q7XBW-PYJGG-WXD33"
goto :InstallKey
:: Office 2019 Professional Plus C2R-P
:0bc88885-718c-491d-921f-6f214349e79c
set "_key=VQ9DP-NVHPH-T9HJC-J9PDT-KTQRG"
goto :InstallKey
:: Project 2019 Professional C2R-P
:fc7c4d0c-2e85-4bb9-afd4-01ed1476b5e9
set "_key=XM2V9-DN9HH-QB449-XDGKC-W2RMW"
goto :InstallKey
:: Visio 2019 Professional C2R-P
:500f6619-ef93-4b75-bcb4-82819998a3ca
set "_key=N2CG9-YD3YK-936X4-3WR82-Q3X4H"
goto :InstallKey
::==========================================
:: Office 2016 Mondo
:9caabccb-61b1-4b4b-8bec-d10a3c3ac2ce
set "_key=HFTND-W9MK4-8B7MJ-B6C4G-XQBR2"
goto :InstallKey
:: Office 2016 Professional Plus
:d450596f-894d-49e0-966a-fd39ed4c4c64
set "_key=XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99"
goto :InstallKey
:: Office 2016 Standard
:dedfa23d-6ed1-45a6-85dc-63cae0546de6
set "_key=JNRGM-WHDWX-FJJG3-K47QV-DRTFM"
goto :InstallKey
:: Project 2016 Professional
:4f414197-0fc2-4c01-b68a-86cbb9ac254c
set "_key=YG9NW-3K39V-2T3HJ-93F3Q-G83KT"
goto :InstallKey
:: Project 2016 Standard
:da7ddabc-3fbe-4447-9e01-6ab7440b4cd4
set "_key=GNFHQ-F6YQM-KQDGJ-327XX-KQBVC"
goto :InstallKey
:: Visio 2016 Professional
:6bf301c1-b94a-43e9-ba31-d494598c47fb
set "_key=PD3PC-RHNGV-FXJ29-8JK7D-RJRJK"
goto :InstallKey
:: Visio 2016 Standard
:aa2a7821-1827-4c2c-8f1d-4513a34dda97
set "_key=7WHWN-4T7MP-G96JF-G33KR-W8GF4"
goto :InstallKey
:: Access 2016
:67c0fc0c-deba-401b-bf8b-9c8ad8395804
set "_key=GNH9Y-D2J4T-FJHGG-QRVH7-QPFDW"
goto :InstallKey
:: Excel 2016
:c3e65d36-141f-4d2f-a303-a842ee756a29
set "_key=9C2PK-NWTVB-JMPW8-BFT28-7FTBF"
goto :InstallKey
:: OneNote 2016
:d8cace59-33d2-4ac7-9b1b-9b72339c51c8
set "_key=DR92N-9HTF2-97XKM-XW2WJ-XW3J6"
goto :InstallKey
:: Outlook 2016
:ec9d9265-9d1e-4ed0-838a-cdc20f2551a1
set "_key=R69KK-NTPKF-7M3Q4-QYBHW-6MT9B"
goto :InstallKey
:: PowerPoint 2016
:d70b1bba-b893-4544-96e2-b7a318091c33
set "_key=J7MQP-HNJ4Y-WJ7YM-PFYGF-BY6C6"
goto :InstallKey
:: Publisher 2016
:041a06cb-c5b8-4772-809f-416d03d16654
set "_key=F47MM-N3XJP-TQXJ9-BP99D-8K837"
goto :InstallKey
:: Skype for Business 2016
:83e04ee1-fa8d-436d-8994-d31a862cab77
set "_key=869NQ-FJ69K-466HW-QYCP2-DDBV6"
goto :InstallKey
:: Word 2016
:bb11badf-d8aa-470e-9311-20eaf80fe5cc
set "_key=WXY84-JN2Q9-RBCCQ-3Q3J3-3PFJ6"
goto :InstallKey
:: Project 2016 Professional C2R-P
:829b8110-0e6f-4349-bca4-42803577788d
set "_key=WGT24-HCNMF-FQ7XH-6M8K7-DRTW9"
goto :InstallKey
:: Project 2016 Standard C2R-P
:cbbaca45-556a-4416-ad03-bda598eaa7c8
set "_key=D8NRQ-JTYM3-7J2DX-646CT-6836M"
goto :InstallKey
:: Visio 2016 Professional C2R-P
:b234abe3-0857-4f9c-b05a-4dc314f85557
set "_key=69WXN-MBYV6-22PQG-3WGHK-RM6XC"
goto :InstallKey
:: Visio 2016 Standard C2R-P
:361fe620-64f4-41b5-ba77-84f8e079b1f7
set "_key=NY48V-PPYYH-3F4PX-XJRKJ-W4423"
goto :InstallKey
:: Office 2016 MondoR Automation
:e914ea6e-a5fa-4439-a394-a9bb3293ca09
set "_key=DMTCJ-KNRKX-26982-JYCKT-P7KB6"
goto :InstallKey
::==========================================
:: Office 2013 Mondo
:dc981c6b-fc8e-420f-aa43-f8f33e5c0923
set "_key=42QTK-RN8M7-J3C4G-BBGYM-88CYV"
goto :InstallKey
:: Office 2013 Professional Plus
:b322da9c-a2e2-4058-9e4e-f59a6970bd69
set "_key=YC7DK-G2NP3-2QQC3-J6H88-GVGXT"
goto :InstallKey
:: Office 2013 Standard
:b13afb38-cd79-4ae5-9f7f-eed058d750ca
set "_key=KBKQT-2NMXY-JJWGP-M62JB-92CD4"
goto :InstallKey
:: Project 2013 Professional
:4a5d124a-e620-44ba-b6ff-658961b33b9a
set "_key=FN8TT-7WMH6-2D4X9-M337T-2342K"
goto :InstallKey
:: Project 2013 Standard
:427a28d1-d17c-4abf-b717-32c780ba6f07
set "_key=6NTH3-CW976-3G3Y2-JK3TX-8QHTT"
goto :InstallKey
:: Visio 2013 Professional
:e13ac10e-75d0-4aff-a0cd-764982cf541c
set "_key=C2FG9-N6J68-H8BTJ-BW3QX-RM3B3"
goto :InstallKey
:: Visio 2013 Standard
:ac4efaf0-f81f-4f61-bdf7-ea32b02ab117
set "_key=J484Y-4NKBF-W2HMG-DBMJC-PGWR7"
goto :InstallKey
:: Access 2013
:6ee7622c-18d8-4005-9fb7-92db644a279b
set "_key=NG2JY-H4JBT-HQXYP-78QH9-4JM2D"
goto :InstallKey
:: Excel 2013
:f7461d52-7c2b-43b2-8744-ea958e0bd09a
set "_key=VGPNG-Y7HQW-9RHP7-TKPV3-BG7GB"
goto :InstallKey
:: OneDrive for Business 2013 (Groove)
:fb4875ec-0c6b-450f-b82b-ab57d8d1677f
set "_key=H7R7V-WPNXQ-WCYYC-76BGV-VT7GH"
goto :InstallKey
:: InfoPath 2013
:a30b8040-d68a-423f-b0b5-9ce292ea5a8f
set "_key=DKT8B-N7VXH-D963P-Q4PHY-F8894"
goto :InstallKey
:: Lync 2013
:1b9f11e3-c85c-4e1b-bb29-879ad2c909e3
set "_key=2MG3G-3BNTT-3MFW9-KDQW3-TCK7R"
goto :InstallKey
:: OneNote 2013
:efe1f3e6-aea2-4144-a208-32aa872b6545
set "_key=TGN6P-8MMBC-37P2F-XHXXK-P34VW"
goto :InstallKey
:: Outlook 2013
:771c3afa-50c5-443f-b151-ff2546d863a0
set "_key=QPN8Q-BJBTJ-334K3-93TGY-2PMBT"
goto :InstallKey
:: PowerPoint 2013
:8c762649-97d1-4953-ad27-b7e2c25b972e
set "_key=4NT99-8RJFH-Q2VDH-KYG2C-4RD4F"
goto :InstallKey
:: Publisher 2013
:00c79ff1-6850-443d-bf61-71cde0de305f
set "_key=PN2WF-29XG2-T9HJ7-JQPJR-FCXK4"
goto :InstallKey
:: Word 2013
:d9f5b1c6-5386-495a-88f9-9ad6b41ac9b3
set "_key=6Q7VD-NX8JD-WJ2VH-88V73-4GBJ7"
goto :InstallKey
::==========================================
:: Office 2010 Professional Plus
:6f327760-8c5c-417c-9b61-836a98287e0c
set "_key=VYBBJ-TRJPB-QFQRF-QFT4D-H3GVB"
goto :InstallKey
:: Office 2010 Standard
:9da2a678-fb6b-4e67-ab84-60dd6a9c819a
set "_key=V7QKV-4XVVR-XYV4D-F7DFM-8R6BM"
goto :InstallKey
:: Access 2010
:8ce7e872-188c-4b98-9d90-f8f90b7aad02
set "_key=V7Y44-9T38C-R2VJK-666HK-T7DDX"
goto :InstallKey
:: Excel 2010
:cee5d470-6e3b-4fcc-8c2b-d17428568a9f
set "_key=H62QG-HXVKF-PP4HP-66KMR-CW9BM"
goto :InstallKey
:: SharePoint Workspace 2010 (Groove)
:8947d0b8-c33b-43e1-8c56-9b674c052832
set "_key=QYYW6-QP4CB-MBV6G-HYMCJ-4T3J4"
goto :InstallKey
:: InfoPath 2010
:ca6b6639-4ad6-40ae-a575-14dee07f6430
set "_key=K96W8-67RPQ-62T9Y-J8FQJ-BT37T"
goto :InstallKey
:: OneNote 2010
:ab586f5c-5256-4632-962f-fefd8b49e6f4
set "_key=Q4Y4M-RHWJM-PY37F-MTKWH-D3XHX"
goto :InstallKey
:: Outlook 2010
:ecb7c192-73ab-4ded-acf4-2399b095d0cc
set "_key=7YDC2-CWM8M-RRTJC-8MDVC-X3DWQ"
goto :InstallKey
:: PowerPoint 2010
:45593b1d-dfb1-4e91-bbfb-2d5d0ce2227a
set "_key=RC8FX-88JRY-3PF7C-X8P67-P4VTT"
goto :InstallKey
:: Project 2010 Professional
:df133ff7-bf14-4f95-afe3-7b48e7e331ef
set "_key=YGX6F-PGV49-PGW3J-9BTGG-VHKC6"
goto :InstallKey
:: Project 2010 Standard
:5dc7bf61-5ec9-4996-9ccb-df806a2d0efe
set "_key=4HP3K-88W3F-W2K3D-6677X-F9PGB"
goto :InstallKey
:: Publisher 2010
:b50c4f75-599b-43e8-8dcd-1081a7967241
set "_key=BFK7F-9MYHM-V68C7-DRQ66-83YTP"
goto :InstallKey
:: Word 2010
:2d0882e7-a4e7-423b-8ccc-70d91e0158b1
set "_key=HVHB3-C6FV7-KQX9W-YQG79-CRY7T"
goto :InstallKey
:: Visio 2010 Premium
:92236105-bb67-494f-94c7-7f7a607929bd
set "_key=D9DWC-HPYVV-JGF4P-BTWQB-WX8BJ"
goto :InstallKey
:: Visio 2010 Professional
:e558389c-83c3-4b29-adfe-5e4d7f46c358
set "_key=7MCW8-VRQVK-G677T-PDJCM-Q8TCP"
goto :InstallKey
:: Visio 2010 Standard
:9ed833ff-4f92-4f36-b370-8683a4f13275
set "_key=767HD-QGMWX-8QTDB-9G3R2-KHFGJ"
goto :InstallKey
:: Office 2010 Home and Business
:ea509e87-07a1-4a45-9edc-eba5a39f36af
set "_key=D6QFG-VBYP2-XQHM7-J97RH-VVRCK"
goto :InstallKey
:: Office 2010 Mondo
:09ed9640-f020-400a-acd8-d7d867dfd9c2
set "_key=YBJTT-JG6MD-V9Q7P-DBKXJ-38W9R"
goto :InstallKey
:: Office 2010 Mondo
:ef3d4e49-a53d-4d81-a2b1-2ca6c2556b2c
set "_key=7TC2V-WXF6P-TD7RT-BQRXR-B8K32"
goto :InstallKey
::==========================================
:: Windows 10 Home
:58e97c99-f377-4ef1-81d5-4ad5522b5fd8
set "_key=TX9XD-98N7V-6WMQ6-BX7FG-H8Q99"
goto :InstallKey
:: Windows 10 Home N
:7b9e1751-a8da-4f75-9560-5fadfe3d8e38
set "_key=3KHY7-WNT83-DGQKR-F7HPR-844BM"
goto :InstallKey
:: Windows 10 Home Single Language
:cd918a57-a41b-4c82-8dce-1a538e221a83
set "_key=7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH"
goto :InstallKey
:: Windows 10 Home Country Specific
:a9107544-f4a0-4053-a96a-1479abdef912
set "_key=PVMJN-6DFY6-9CCP6-7BKTT-D3WVR"
goto :InstallKey
:: Windows 10 Professional
:2de67392-b7a7-462a-b1ca-108dd189f588
set "_key=W269N-WFGWX-YVC9B-4J6C9-T83GX"
goto :InstallKey
:: Windows 10 Professional N
:a80b5abf-76ad-428b-b05d-a47d2dffeebf
set "_key=MH37W-N47XK-V7XM9-C7227-GCQG9"
goto :InstallKey
:: Windows 10 Professional Education
:3f1afc82-f8ac-4f6c-8005-1d233e606eee
set "_key=6TP4R-GNPTD-KYYHQ-7B7DP-J447Y"
goto :InstallKey
:: Windows 10 Professional Education N
:5300b18c-2e33-4dc2-8291-47ffcec746dd
set "_key=YVWGF-BXNMC-HTQYQ-CPQ99-66QFC"
goto :InstallKey
:: Windows 10 Professional Workstation
:82bbc092-bc50-4e16-8e18-b74fc486aec3
set "_key=NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J"
goto :InstallKey
:: Windows 10 Professional Workstation N
:4b1571d3-bafb-4b40-8087-a961be2caf65
set "_key=9FNHH-K3HBT-3W4TD-6383H-6XYWF"
goto :InstallKey
:: Windows 10 Education
:e0c42288-980c-4788-a014-c080d2e1926e
set "_key=NW6C2-QMPVW-D7KKK-3GKT6-VCFB2"
goto :InstallKey
:: Windows 10 Education N
:3c102355-d027-42c6-ad23-2e7ef8a02585
set "_key=2WH4N-8QGBV-H22JP-CT43Q-MDWWJ"
goto :InstallKey
:: Windows 10 Enterprise
:73111121-5638-40f6-bc11-f1d7b0d64300
set "_key=NPPR9-FWDCX-D2C8J-H872K-2YT43"
goto :InstallKey
:: Windows 10 Enterprise N
:e272e3e2-732f-4c65-a8f0-484747d0d947
set "_key=DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4"
goto :InstallKey
:: Windows 10 Enterprise G
:e0b2d383-d112-413f-8a80-97f373a5820c
set "_key=YYVX9-NTFWV-6MDM3-9PT4T-4M68B"
goto :InstallKey
:: Windows 10 Enterprise G N
:e38454fb-41a4-4f59-a5dc-25080e354730
set "_key=44RPN-FTY23-9VTTB-MP9BX-T84FV"
goto :InstallKey
:: Windows 10 Enterprise 2015 LTSB
:7b51a46c-0c04-4e8f-9af4-8496cca90d5e
set "_key=WNMTR-4C88C-JK8YV-HQ7T2-76DF9"
goto :InstallKey
:: Windows 10 Enterprise 2015 LTSB N
:87b838b7-41b6-4590-8318-5797951d8529
set "_key=2F77B-TNFGY-69QQF-B8YKP-D69TJ"
goto :InstallKey
:: Windows 10 Enterprise 2016 LTSB
:2d5a5a60-3040-48bf-beb0-fcd770c20ce0
set "_key=DCPHK-NFMTC-H88MJ-PFHPY-QJ4BJ"
goto :InstallKey
:: Windows 10 Enterprise 2016 LTSB N
:9f776d83-7156-45b2-8a5c-359b9c9f22a3
set "_key=QFFDN-GRT3P-VKWWX-X7T3R-8B639"
goto :InstallKey
:: Windows 10 Enterprise LTSC 2018
:32d2fab3-e4a8-42c2-923b-4bf4fd13e6ee
set "_key=M7XTQ-FN8P6-TTKYV-9D4CC-J462D"
goto :InstallKey
:: Windows 10 Enterprise LTSC 2018 N
:7103a333-b8c8-49cc-93ce-d37c09687f92
set "_key=92NFX-8DJQP-P6BBQ-THF9C-7CG2H"
goto :InstallKey
:: Windows 10 Enterprise Remote Server
:e4db50ea-bda1-4566-b047-0ca50abc6f07
set "_key=7NBT4-WGBQX-MP4H7-QXFF8-YP3KX"
goto :InstallKey
:: Windows 10 Enterprise for Remote Sessions
:ec868e65-fadf-4759-b23e-93fe37f2cc29
set "_key=CPWHC-NT2C7-VYW78-DHDB2-PG3GK"
goto :InstallKey
:: Windows 10 Lean
:0df4f814-3f57-4b8b-9a9d-fddadcd69fac
set "_key=NBTWJ-3DR69-3C4V8-C26MC-GQ9M6"
goto :InstallKey
::==========================================
:: Windows Server 2019 Essentials
:034d3cbb-5d4b-4245-b3f8-f84571314078
set "_key=WVDHN-86M7X-466P6-VHXV7-YY726"
goto :InstallKey
:: Windows Server 2019 Standard
:de32eafd-aaee-4662-9444-c1befb41bde2
set "_key=N69G4-B89J2-4G8F4-WWYCC-J464C"
goto :InstallKey
:: Windows Server 2019 Datacenter
:34e1ae55-27f8-4950-8877-7a03be5fb181
set "_key=WMDGN-G9PQG-XVVXX-R3X43-63DFG"
goto :InstallKey
:: Windows Server 2019 Standard ACor
:73e3957c-fc0c-400d-9184-5f7b6f2eb409
set "_key=N2KJX-J94YW-TQVFB-DG9YT-724CC"
goto :InstallKey
:: Windows Server 2019 Datacenter ACor
:90c362e5-0da1-4bfd-b53b-b87d309ade43
set "_key=6NMRW-2C8FM-D24W7-TQWMY-CWH2D"
goto :InstallKey
:: Windows Server 2019 Azure Core
:a99cc1f0-7719-4306-9645-294102fbff95
set "_key=FDNH6-VW9RW-BXPJ7-4XTYG-239TB"
goto :InstallKey
:: Windows Server 2019 ARM64
:8de8eb62-bbe0-40ac-ac17-f75595071ea3
set "_key=GRFBW-QNDC4-6QBHG-CCK3B-2PR88"
goto :InstallKey
::==========================================
:: Windows Server 2016 Essentials
:2b5a1b0f-a5ab-4c54-ac2f-a6d94824a283
set "_key=JCKRF-N37P4-C2D82-9YXRT-4M63B"
goto :InstallKey
:: Windows Server 2016 Standard
:8c1c5410-9f39-4805-8c9d-63a07706358f
set "_key=WC2BQ-8NRM3-FDDYY-2BFGV-KHKQY"
goto :InstallKey
:: Windows Server 2016 Datacenter
:21c56779-b449-4d20-adfc-eece0e1ad74b
set "_key=CB7KF-BWN84-R7R2Y-793K2-8XDDG"
goto :InstallKey
:: Windows Server 2016 Standard ACor
:61c5ef22-f14f-4553-a824-c4b31e84b100
set "_key=PTXN8-JFHJM-4WC78-MPCBR-9W4KR"
goto :InstallKey
:: Windows Server 2016 Datacenter ACor
:e49c08e7-da82-42f8-bde2-b570fbcae76c
set "_key=2HXDN-KRXHB-GPYC7-YCKFJ-7FVDG"
goto :InstallKey
:: Windows Server 2016 Cloud Storage
:7b4433f4-b1e7-4788-895a-c45378d38253
set "_key=QN4C6-GBJD2-FB422-GHWJK-GJG2R"
goto :InstallKey
:: Windows Server 2016 Azure Core
:3dbf341b-5f6c-4fa7-b936-699dce9e263f
set "_key=VP34G-4NPPG-79JTQ-864T4-R3MQX"
goto :InstallKey
:: Windows Server 2016 ARM64
:43d9af6e-5e86-4be8-a797-d072a046896c
set "_key=K9FYF-G6NCK-73M32-XMVPY-F9DRR"
goto :InstallKey
::==========================================
:: Windows 8.1 Professional
:c06b6981-d7fd-4a35-b7b4-054742b7af67
set "_key=GCRJD-8NW9H-F2CDX-CCM8D-9D6T9"
goto :InstallKey
:: Windows 8.1 Professional N
:7476d79f-8e48-49b4-ab63-4d0b813a16e4
set "_key=HMCNV-VVBFX-7HMBH-CTY9B-B4FXY"
goto :InstallKey
:: Windows 8.1 Enterprise
:81671aaf-79d1-4eb1-b004-8cbbe173afea
set "_key=MHF9N-XY6XB-WVXMC-BTDCT-MKKG7"
goto :InstallKey
:: Windows 8.1 Enterprise N
:113e705c-fa49-48a4-beea-7dd879b46b14
set "_key=TT4HM-HN7YT-62K67-RGRQJ-JFFXW"
goto :InstallKey
:: Windows 8.1 Professional WMC
:096ce63d-4fac-48a9-82a9-61ae9e800e5f
set "_key=789NJ-TQK6T-6XTH8-J39CJ-J8D3P"
goto :InstallKey
:: Windows 8.1 Core
:fe1c3238-432a-43a1-8e25-97e7d1ef10f3
set "_key=M9Q9P-WNJJT-6PXPY-DWX8H-6XWKK"
goto :InstallKey
:: Windows 8.1 Core N
:78558a64-dc19-43fe-a0d0-8075b2a370a3
set "_key=7B9N3-D94CG-YTVHR-QBPX3-RJP64"
goto :InstallKey
:: Windows 8.1 Core ARM
:ffee456a-cd87-4390-8e07-16146c672fd0
set "_key=XYTND-K6QKT-K2MRH-66RTM-43JKP"
goto :InstallKey
:: Windows 8.1 Core Single Language
:c72c6a1d-f252-4e7e-bdd1-3fca342acb35
set "_key=BB6NG-PQ82V-VRDPW-8XVD2-V8P66"
goto :InstallKey
:: Windows 8.1 Core Country Specific
:db78b74f-ef1c-4892-abfe-1e66b8231df6
set "_key=NCTT7-2RGK8-WMHRF-RY7YQ-JTXG3"
goto :InstallKey
:: Windows 8.1 Embedded Industry
:0ab82d54-47f4-4acb-818c-cc5bf0ecb649
set "_key=NMMPB-38DD4-R2823-62W8D-VXKJB"
goto :InstallKey
:: Windows 8.1 Embedded Industry Enterprise
:cd4e2d9f-5059-4a50-a92d-05d5bb1267c7
set "_key=FNFKF-PWTVT-9RC8H-32HB2-JB34X"
goto :InstallKey
:: Windows 8.1 Embedded Industry Automotive
:f7e88590-dfc7-4c78-bccb-6f3865b99d1a
set "_key=VHXM3-NR6FT-RY6RT-CK882-KW2CJ"
goto :InstallKey
:: Windows 8.1 Core Connected (with Bing)
:e9942b32-2e55-4197-b0bd-5ff58cba8860
set "_key=3PY8R-QHNP9-W7XQD-G6DPH-3J2C9"
goto :InstallKey
:: Windows 8.1 Core Connected N (with Bing)
:c6ddecd6-2354-4c19-909b-306a3058484e
set "_key=Q6HTR-N24GM-PMJFP-69CD8-2GXKR"
goto :InstallKey
:: Windows 8.1 Core Connected Single Language (with Bing)
:b8f5e3a3-ed33-4608-81e1-37d6c9dcfd9c
set "_key=KF37N-VDV38-GRRTV-XH8X6-6F3BB"
goto :InstallKey
:: Windows 8.1 Core Connected Country Specific (with Bing)
:ba998212-460a-44db-bfb5-71bf09d1c68b
set "_key=R962J-37N87-9VVK2-WJ74P-XTMHR"
goto :InstallKey
:: Windows 8.1 Professional Student
:e58d87b5-8126-4580-80fb-861b22f79296
set "_key=MX3RK-9HNGX-K3QKC-6PJ3F-W8D7B"
goto :InstallKey
:: Windows 8.1 Professional Student N
:cab491c7-a918-4f60-b502-dab75e334f40
set "_key=TNFGH-2R6PB-8XM3K-QYHX2-J4296"
goto :InstallKey
::==========================================
:: Windows Server 2012 R2 Standard
:b3ca044e-a358-4d68-9883-aaa2941aca99
set "_key=D2N9P-3P6X9-2R39C-7RTCD-MDVJX"
goto :InstallKey
:: Windows Server 2012 R2 Datacenter
:00091344-1ea4-4f37-b789-01750ba6988c
set "_key=W3GGN-FT8W3-Y4M27-J84CP-Q3VJ9"
goto :InstallKey
:: Windows Server 2012 R2 Essentials
:21db6ba4-9a7b-4a14-9e29-64a60c59301d
set "_key=KNC87-3J2TX-XB4WP-VCPJV-M4FWM"
goto :InstallKey
:: Windows Server 2012 R2 Cloud Storage
:b743a2be-68d4-4dd3-af32-92425b7bb623
set "_key=3NPTF-33KPT-GGBPR-YX76B-39KDD"
goto :InstallKey
::==========================================
:: Windows 8 Professional
:a98bcd6d-5343-4603-8afe-5908e4611112
set "_key=NG4HW-VH26C-733KW-K6F98-J8CK4"
goto :InstallKey
:: Windows 8 Professional N
:ebf245c1-29a8-4daf-9cb1-38dfc608a8c8
set "_key=XCVCF-2NXM9-723PB-MHCB7-2RYQQ"
goto :InstallKey
:: Windows 8 Enterprise
:458e1bec-837a-45f6-b9d5-925ed5d299de
set "_key=32JNW-9KQ84-P47T8-D8GGY-CWCK7"
goto :InstallKey
:: Windows 8 Enterprise N
:e14997e7-800a-4cf7-ad10-de4b45b578db
set "_key=JMNMF-RHW7P-DMY6X-RF3DR-X2BQT"
goto :InstallKey
:: Windows 8 Professional WMC
:a00018a3-f20f-4632-bf7c-8daa5351c914
set "_key=GNBB8-YVD74-QJHX6-27H4K-8QHDG"
goto :InstallKey
:: Windows 8 Core
:c04ed6bf-55c8-4b47-9f8e-5a1f31ceee60
set "_key=BN3D2-R7TKB-3YPBD-8DRP2-27GG4"
goto :InstallKey
:: Windows 8 Core N
:197390a0-65f6-4a95-bdc4-55d58a3b0253
set "_key=8N2M2-HWPGY-7PGT9-HGDD8-GVGGY"
goto :InstallKey
:: Windows 8 Core Single Language
:8860fcd4-a77b-4a20-9045-a150ff11d609
set "_key=2WN2H-YGCQR-KFX6K-CD6TF-84YXQ"
goto :InstallKey
:: Windows 8 Core Country Specific
:9d5584a2-2d85-419a-982c-a00888bb9ddf
set "_key=4K36P-JN4VD-GDC6V-KDT89-DYFKP"
goto :InstallKey
:: Windows 8 Core ARM
:af35d7b7-5035-4b63-8972-f0b747b9f4dc
set "_key=DXHJF-N9KQX-MFPVR-GHGQK-Y7RKV"
goto :InstallKey
:: Windows 8 Embedded Industry Professional
:10018baf-ce21-4060-80bd-47fe74ed4dab
set "_key=RYXVT-BNQG7-VD29F-DBMRY-HT73M"
goto :InstallKey
:: Windows 8 Embedded Industry Enterprise
:18db1848-12e0-4167-b9d7-da7fcda507db
set "_key=NKB3R-R2F8T-3XCDP-7Q2KW-XWYQ2"
goto :InstallKey
::==========================================
:: Windows Server 2012 Standard
:f0f5ec41-0d55-4732-af02-440a44a3cf0f
set "_key=XC9B7-NBPP2-83J2H-RHMBY-92BT4"
goto :InstallKey
:: Windows Server 2012 Datacenter
:d3643d60-0c42-412d-a7d6-52e6635327f6
set "_key=48HP8-DN98B-MYWDG-T2DCC-8W83P"
goto :InstallKey
:: Windows Server 2012 MultiPoint Standard
:7d5486c7-e120-4771-b7f1-7b56c6d3170c
set "_key=HM7DN-YVMH3-46JC3-XYTG7-CYQJJ"
goto :InstallKey
:: Windows Server 2012 MultiPoint Premium
:95fd1c83-7df5-494a-be8b-1300e1c9d1cd
set "_key=XNH6W-2V9GX-RGJ4K-Y8X6F-QGJ2G"
goto :InstallKey
::==========================================
:: Windows 7 Professional
:b92e9980-b9d5-4821-9c94-140f632f6312
set "_key=FJ82H-XT6CR-J8D7P-XQJJ2-GPDD4"
goto :InstallKey
:: Windows 7 Professional N
:54a09a0d-d57b-4c10-8b69-a842d6590ad5
set "_key=MRPKT-YTG23-K7D7T-X2JMM-QY7MG"
goto :InstallKey
:: Windows 7 Professional E
:5a041529-fef8-4d07-b06f-b59b573b32d2
set "_key=W82YF-2Q76Y-63HXB-FGJG9-GF7QX"
goto :InstallKey
:: Windows 7 Enterprise
:ae2ee509-1b34-41c0-acb7-6d4650168915
set "_key=33PXH-7Y6KF-2VJC9-XBBR8-HVTHH"
goto :InstallKey
:: Windows 7 Enterprise N
:1cb6d605-11b3-4e14-bb30-da91c8e3983a
set "_key=YDRBP-3D83W-TY26F-D46B2-XCKRJ"
goto :InstallKey
:: Windows 7 Enterprise E
:46bbed08-9c7b-48fc-a614-95250573f4ea
set "_key=C29WB-22CC8-VJ326-GHFJW-H9DH4"
goto :InstallKey
:: Windows 7 Embedded POS Ready
:db537896-376f-48ae-a492-53d0547773d0
set "_key=YBYF6-BHCR3-JPKRB-CDW7B-F9BK4"
goto :InstallKey
:: Windows 7 Embedded ThinPC
:aa6dd3aa-c2b4-40e2-a544-a6bbb3f5c395
set "_key=73KQT-CD9G6-K7TQG-66MRP-CQ22C"
goto :InstallKey
:: Windows 7 Embedded Standard
:e1a8296a-db37-44d1-8cce-7bc961d59c54
set "_key=XGY72-BRBBT-FF8MH-2GG8H-W7KCW"
goto :InstallKey
::==========================================
:: Windows Server 2008 R2 Web
:a78b8bd9-8017-4df5-b86a-09f756affa7c
set "_key=6TPJF-RBVHG-WBW2R-86QPH-6RTM4"
goto :InstallKey
:: Windows Server 2008 R2 HPC edition
:cda18cf3-c196-46ad-b289-60c072869994
set "_key=TT8MH-CG224-D3D7Q-498W2-9QCTX"
goto :InstallKey
:: Windows Server 2008 R2 Standard
:68531fb9-5511-4989-97be-d11a0f55633f
set "_key=YC6KT-GKW9T-YTKYR-T4X34-R7VHC"
goto :InstallKey
:: Windows Server 2008 R2 Enterprise
:620e2b3d-09e7-42fd-802a-17a13652fe7a
set "_key=489J6-VHDMP-X63PK-3K798-CPX3Y"
goto :InstallKey
:: Windows Server 2008 R2 Datacenter
:7482e61b-c589-4b7f-8ecc-46d455ac3b87
set "_key=74YFP-3QFB3-KQT8W-PMXWJ-7M648"
goto :InstallKey
:: Windows Server 2008 R2 for Itanium-based Systems
:8a26851c-1c7e-48d3-a687-fbca9b9ac16b
set "_key=GT63C-RJFQ3-4GMB6-BRFB9-CB83V"
goto :InstallKey
:: Windows MultiPoint Server 2010
:f772515c-0e87-48d5-a676-e6962c3e1195
set "_key=736RG-XDKJK-V34PF-BHK87-J6X3K"
goto :InstallKey
::==========================================
:: Windows Vista Business
:4f3d1606-3fea-4c01-be3c-8d671c401e3b
set "_key=YFKBB-PQJJV-G996G-VWGXY-2V3X8"
goto :InstallKey
:: Windows Vista Business N
:2c682dc2-8b68-4f63-a165-ae291d4cf138
set "_key=HMBQG-8H2RH-C77VX-27R82-VMQBT"
goto :InstallKey
:: Windows Vista Enterprise
:cfd8ff08-c0d7-452b-9f60-ef5c70c32094
set "_key=VKK3X-68KWM-X2YGT-QR4M6-4BWMV"
goto :InstallKey
:: Windows Vista Enterprise N
:d4f54950-26f2-4fb4-ba21-ffab16afcade
set "_key=VTC42-BM838-43QHV-84HX6-XJXKV"
goto :InstallKey
::==========================================
:: Windows Server 2008 Web
:ddfa9f7c-f09e-40b9-8c1a-be877a9a7f4b
set "_key=WYR28-R7TFJ-3X2YQ-YCY4H-M249D"
goto :InstallKey
:: Windows Server 2008 Standard
:ad2542d4-9154-4c6d-8a44-30f11ee96989
set "_key=TM24T-X9RMF-VWXK6-X8JC9-BFGM2"
goto :InstallKey
:: Windows Server 2008 Standard without Hyper-V
:2401e3d0-c50a-4b58-87b2-7e794b7d2607
set "_key=W7VD6-7JFBR-RX26B-YKQ3Y-6FFFJ"
goto :InstallKey
:: Windows Server 2008 Enterprise
:c1af4d90-d1bc-44ca-85d4-003ba33db3b9
set "_key=YQGMW-MPWTJ-34KDK-48M3W-X4Q6V"
goto :InstallKey
:: Windows Server 2008 Enterprise without Hyper-V
:8198490a-add0-47b2-b3ba-316b12d647b4
set "_key=39BXF-X8Q23-P2WWT-38T2F-G3FPG"
goto :InstallKey
:: Windows Server 2008 HPC (Compute Cluster)
:7afb1156-2c1d-40fc-b260-aab7442b62fe
set "_key=RCTX3-KWVHP-BR6TB-RB6DM-6X7HP"
goto :InstallKey
:: Windows Server 2008 Datacenter
:68b6e220-cf09-466b-92d3-45cd964b9509
set "_key=7M67G-PC374-GR742-YH8V4-TCBY3"
goto :InstallKey
:: Windows Server 2008 Datacenter without Hyper-V
:fd09ef77-5647-4eff-809c-af2b64659a45
set "_key=22XQ2-VRXRG-P8D42-K34TD-G3QQC"
goto :InstallKey
:: Windows Server 2008 for Itanium-Based Systems
:01ef176b-3e0d-422a-b4f8-4ea880035e8f
set "_key=4DWFP-JF3DJ-B7DTH-78FJB-PDRHK"
goto :InstallKey
::==========================================
:KeyNotFound
:: If GVLK is not found for the current SKU-ID, attempt Activation as GVLK might be present in OS by default
echo.
echo GVLK for %_Name%
echo with SKU-ID %1 Not Found
echo.
echo If Activation Fails Now, Please enter GVLK for this Product manually and re-run KMS-VL-ALL.cmd
call :Activate %1
exit /b
::==========================================
:InstallKey
:: Call InstallProductKey method of SLS/OSPS to install GVLK
echo Installing Key...
wmic path %_MicrosoftService% where version='%_Version%' call InstallProductKey ProductKey="%_key%" %_Nul_1_2%
call :Activate %1
exit /b
::==========================================