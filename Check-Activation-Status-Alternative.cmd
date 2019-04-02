@echo off
set spp=SoftwareLicensingProduct
set sps=SoftwareLicensingService
set ospp=OfficeSoftwareProtectionProduct
set osps=OfficeSoftwareProtectionService
set winApp=55c92734-d682-4d71-983e-d6ec3f16059f
set o14App=59a52881-a989-479d-af46-f275c6370663
set o15App=0ff1ce15-a989-479d-af46-f275c6370663
for %%# in (spp_get,ospp_get,W1nd0ws,sppw,0ff1ce15,sppo,osppsvc,ospp14,ospp15) do set "%%#="
for /f "tokens=6 delims=[]. " %%# in ('ver') do set winbuild=%%#
set "spp_get=Description, DiscoveredKeyManagementServiceMachineName, DiscoveredKeyManagementServiceMachinePort, EvaluationEndDate, GracePeriodRemaining, ID, KeyManagementServiceMachine, KeyManagementServicePort, KeyManagementServiceProductKeyID, LicenseStatus, LicenseStatusReason, Name, PartialProductKey, ProductKeyID, VLActivationInterval, VLRenewalInterval"
set "ospp_get=%spp_get%"
if %winbuild% geq 9200 set "spp_get=%spp_get%, DiscoveredKeyManagementServiceMachineIpAddress, KeyManagementServiceLookupDomain, ProductKeyChannel, VLActivationTypeEnabled"

set "SysPath=%Windir%\System32"
if exist "%Windir%\Sysnative\reg.exe" (set "SysPath=%Windir%\Sysnative")
set "Path=%SysPath%;%Windir%;%SysPath%\Wbem;%SysPath%\WindowsPowerShell\v1.0\"

call :PKey %spp% %winApp% W1nd0ws sppw
if %winbuild% geq 9200 call :PKey %spp% %o15App% 0ff1ce15 sppo
wmic path %osps% get Version 1>nul 2>nul && (
call :PKey %ospp% %o14App% osppsvc ospp14
if %winbuild% lss 9200 call :PKey %ospp% %o15App% osppsvc ospp15
)

:SPP
echo ************************************************************
echo ***                   Windows Status                     ***
echo ************************************************************
if not defined W1nd0ws (
echo.
echo Error: product key not found.
goto :SPPo
)
for /f "tokens=2 delims==" %%# in ('"wmic path %spp% where (ApplicationID='%winApp%' and PartialProductKey is not null) get ID /value"') do (
  set "chkID=%%#"
  call :Property "%spp%" "%sps%" "%spp_get%"
  call :Output
  echo ____________________________________________________________________________
  echo.
)

:SPPo
set verbose=1
if not defined 0ff1ce15 (
if defined osppsvc goto :OSPP
goto :End
)
echo ************************************************************
echo ***                   Office Status                      ***
echo ************************************************************
for /f "tokens=2 delims==" %%# in ('"wmic path %spp% where (ApplicationID='%o15App%' and PartialProductKey is not null) get ID /value"') do (
  set "chkID=%%#"
  call :Property "%spp%" "%sps%" "%spp_get%"
  call :Output
  echo ____________________________________________________________________________
  echo.
)
set verbose=0
if defined osppsvc goto :OSPP
goto :End

:OSPP
if %verbose%==1 (
echo ************************************************************
echo ***                   Office Status                      ***
echo ************************************************************
)
if defined ospp15 for /f "tokens=2 delims==" %%# in ('"wmic path %ospp% where (ApplicationID='%o15App%' and PartialProductKey is not null) get ID /value"') do (
  set "chkID=%%#"
  call :Property "%ospp%" "%osps%" "%ospp_get%"
  call :Output
  echo ____________________________________________________________________________
  echo.
)
if defined ospp14 for /f "tokens=2 delims==" %%# in ('"wmic path %ospp% where (ApplicationID='%o14App%' and PartialProductKey is not null) get ID /value"') do (
  set "chkID=%%#"
  call :Property "%ospp%" "%osps%" "%ospp_get%"
  call :Output
  echo ____________________________________________________________________________
  echo.
)
goto :End

:PKey
wmic path %1 where (ApplicationID='%2' and PartialProductKey is not null) get ID /value 2>nul | findstr /i ID 1>nul && (set %3=1&set %4=1)
exit /b

:Property
for %%# in (%~3) do set "%%#="
if %~1 equ %ospp% for %%# in (DiscoveredKeyManagementServiceMachineIpAddress, KeyManagementServiceLookupDomain, ProductKeyChannel, VLActivationTypeEnabled) do set "%%#="
set "KmsClient="
for /f "tokens=* delims=" %%# in ('"wmic path %~1 where (ID='%chkID%') get %~3 /value" ^| findstr ^=') do set "%%#"

set /a gprDays=%GracePeriodRemaining%/1440
echo %Description%| findstr /i VOLUME_KMSCLIENT 1>nul && (set KmsClient=1)
call cmd /c exit /b %LicenseStatusReason%
set "LicenseReason=%=ExitCode%"

if %LicenseStatus%==0 (
set "License=Unlicensed"
set "LicenseMsg="
)
if %LicenseStatus%==1 (
set "License=Licensed"
set "LicenseMsg="
if not %GracePeriodRemaining%==0 set "LicenseMsg=Volume activation expiration: %GracePeriodRemaining% minute(s) (%gprDays% day(s))"
)
if %LicenseStatus%==2 (
set "License=Initial grace period"
set "LicenseMsg=Time remaining: %GracePeriodRemaining% minute(s) (%gprDays% day(s))"
)
if %LicenseStatus%==3 (
set "License=Additional grace period (KMS license expired or hardware out of tolerance)"
set "LicenseMsg=Time remaining: %GracePeriodRemaining% minute(s) (%gprDays% day(s))"
)
if %LicenseStatus%==4 (
set "License=Non-genuine grace period."
set "LicenseMsg=Time remaining: %GracePeriodRemaining% minute(s) (%gprDays% day(s))"
)
if %LicenseStatus%==6 (
set "License=Extended grace period"
set "LicenseMsg=Time remaining: %GracePeriodRemaining% minute(s) (%gprDays% day(s))"
)
if %LicenseStatus%==5 (
set "License=Notification"
  if "%LicenseReason%"=="C004F200" (set "LicenseMsg=Notification Reason: 0xC004F200 (non-genuine)."
  ) else if "%LicenseReason%"=="C004F009" (set "LicenseMsg=Notification Reason: 0xC004F009 (grace time expired)."
  ) else (set "LicenseMsg=Notification Reason: 0x%LicenseReason%"
  )
)
if %LicenseStatus% gtr 6 (
set "License=Unknown"
set "LicenseMsg="
)
if not defined KmsClient exit /b

if %KeyManagementServicePort%==0 set KeyManagementServicePort=1688
set "KmsReg=Registered KMS machine name: %KeyManagementServiceMachine%:%KeyManagementServicePort%"
if "%KeyManagementServiceMachine%"=="" set "KmsReg=Registered KMS machine name: KMS name not available"

if %DiscoveredKeyManagementServiceMachinePort%==0 set DiscoveredKeyManagementServiceMachinePort=1688
set "KmsDns=KMS machine name from DNS: %DiscoveredKeyManagementServiceMachineName%:%DiscoveredKeyManagementServiceMachinePort%"
if "%DiscoveredKeyManagementServiceMachineName%"=="" set "KmsDns=DNS auto-discovery: KMS name not available"

for /f "tokens=* delims=" %%# in ('"wmic path %~2 get ClientMachineID, KeyManagementServiceHostCaching /value" ^| findstr ^=') do set "%%#"
if /i %KeyManagementServiceHostCaching%==True (set KeyManagementServiceHostCaching=Enabled) else (set KeyManagementServiceHostCaching=Disabled)

if %winbuild% lss 9200 exit /b
if %~1 equ %ospp% exit /b

if "%DiscoveredKeyManagementServiceMachineIpAddress%"=="" set "DiscoveredKeyManagementServiceMachineIpAddress=not available"

if "%KeyManagementServiceLookupDomain%"=="" set "KeyManagementServiceLookupDomain="

if %VLActivationTypeEnabled%==3 (
set VLActivationType=Token
) else if %VLActivationTypeEnabled%==2 (
set VLActivationType=KMS
) else if %VLActivationTypeEnabled%==1 (
set VLActivationType=AD
) else (
set VLActivationType=All
)
exit /b

:Output
echo.
echo Name: %Name%
echo Description: %Description%
echo Activation ID: %ID%
echo Extended PID: %ProductKeyID%
if defined ProductKeyChannel echo Product Key Channel: %ProductKeyChannel%
echo Partial Product Key: %PartialProductKey%
echo License Status: %License%
if defined LicenseMsg echo %LicenseMsg%
if not %LicenseStatus%==0 if not %EvaluationEndDate:~0,8%==16010101 echo Evaluation End Date: %EvaluationEndDate:~0,4%-%EvaluationEndDate:~4,2%-%EvaluationEndDate:~6,2% %EvaluationEndDate:~8,2%:%EvaluationEndDate:~10,2% UTC
if not defined KmsClient exit /b
if defined VLActivationTypeEnabled echo Configured Activation Type: %VLActivationType%
echo.
if not %LicenseStatus%==1 (
echo Please activate the product in order to update KMS client information values.
exit /b
)
echo Most recent activation information:
echo Key Management Service client information
echo.    Client Machine ID (CMID): %ClientMachineID%
echo.    %KmsDns%
echo.    %KmsReg%
if defined DiscoveredKeyManagementServiceMachineIpAddress echo.    KMS machine IP address: %DiscoveredKeyManagementServiceMachineIpAddress%
echo.    KMS machine extended PID: %KeyManagementServiceProductKeyID%
echo.    Activation interval: %VLActivationInterval% minutes
echo.    Renewal interval: %VLRenewalInterval% minutes
echo.    KMS host caching: %KeyManagementServiceHostCaching%
if defined KeyManagementServiceLookupDomain echo.    KMS SRV record lookup domain: %KeyManagementServiceLookupDomain%
exit /b

:End
echo.
echo Press any key to exit...
pause >nul
exit /b
