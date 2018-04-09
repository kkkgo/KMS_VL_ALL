@echo off
::=============================================================
chcp 437 >nul
::=============================================================
:: Get Fully Qualified FileName of the script
set "_FileName=%~f0"
::=============================================================
:: Get Drive and Path containing the script
set "_FileDir=%~dp0"
::=============================================================
setlocal EnableExtensions EnableDelayedExpansion
::=============================================================
:: Get Administrator Rights
fltmc >nul 2>&1 || (
  echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\GetAdmin.vbs"
  echo UAC.ShellExecute "!_FileName!", "", "", "runas", 1 >> "%temp%\GetAdmin.vbs"
  cmd /u /c type "%temp%\GetAdmin.vbs">"%temp%\GetAdminUnicode.vbs"
  cscript //nologo "%temp%\GetAdminUnicode.vbs"
  del /f /q "%temp%\GetAdmin.vbs" >nul 2>&1
  del /f /q "%temp%\GetAdminUnicode.vbs" >nul 2>&1
  exit
)
::=============================================================
:: Go to the Path of the Script
pushd "!_FileDir!"
::=============================================================
:: Get Windows OS build
for /f "tokens=2 delims==" %%G in ('wmic path Win32_OperatingSystem get BuildNumber /value') do (
  set /a _WinBuild=%%G
)
::=============================================================
:: Get Architecture of the OS
for /f "tokens=2 delims==" %%G in ('wmic path Win32_Processor get AddressWidth /value') do (
  set "_OSarch=%%G-bit"
)
::=============================================================
:: Set SLMGR alias
set "_SLMGR=%SystemRoot%\System32\slmgr.vbs"

:: Get Office C2R installation path
for /f "tokens=2*" %%G in ('"reg query HKLM\SOFTWARE\Microsoft\Office\ClickToRun /v InstallPath" 2^>nul') do (
  set "_OSPP=%%H\Office16\OSPP.VBS"
  set "_LicPath=%%H\root\Licenses16"
)

sc query ClickToRunSvc >nul 2>&1

if %ERRORLEVEL% EQU 1060 (
  echo.
  echo Could not detect Office 2016 ClickToRun service...
  goto :end
)

if not exist "%_LicPath%\*.xrm-ms" (
  echo.
  echo Could not detect Office 2016 Licenses files...
  goto :end
)

if %_WinBuild% LSS 9200 if not exist "%_OSPP%" (
  echo.
  echo Could not detect Licensing tool OSPP.VBS...
  goto :end
)
::=============================================================
:Check
set /a _VL=0
set /a _TB=0
set /a _Grace=0

if %_WinBuild% GEQ 9200 (
  set _MicrosoftProduct=SoftwareLicensingProduct
  set _MicrosoftService=SoftwareLicensingService
) else (
  set _MicrosoftProduct=OfficeSoftwareProtectionProduct
  set _MicrosoftService=OfficeSoftwareProtectionService
)

for /f "tokens=2 delims==" %%G in ('"wmic path %_MicrosoftService% get version /format:list" 2^>nul') do (
  set "_ver=%%G"
)

wmic path %_MicrosoftProduct% where (Description like '%%KMSCLIENT%%' and LicenseFamily != 'Office16MondoR_KMS_Automation') get Name /value 2>nul | findstr /i /C:"Office 16" 1>nul && (
  set /a _VL=1
)

wmic path %_MicrosoftProduct% where (Description like '%%TIMEBASED%%') get Name /value 2>nul | findstr /i /C:"Office 16" 1>nul && (
  set /a _TB=1
)

wmic path %_MicrosoftProduct% where (Description like '%%Grace%%') get Name /format:list 2>nul | findstr /i /C:"Office 16" 1>nul && (
  set /a _Grace=1
)

if %_TB% EQU 0 if %_Grace% EQU 0 if %_VL% EQU 1 (
  echo.
  echo No Conversion or Cleanup Required...
  goto :end
)

echo.
echo Cleaning Current Office 2016 Licenses...
cd /d "%~dp0"

%_OSarch%\cleanospp.exe >nul 2>&1
::=============================================================
:Ret2VL
echo.
echo Installing Office 2016 Volume Licenses...
cd /d "%_LicPath%"

for /f "delims=" %%G in ('dir /b /on client-issuance-*.xrm-ms') do (
  if %_WinBuild% GEQ 9200 (
    cscript //Nologo //B %_SLMGR% /ilc %%G
  ) else (
    cscript //Nologo //B "%_OSPP%" /inslic:%%G
  )
)

if %_WinBuild% GEQ 9200 (
  cscript //Nologo //B %_SLMGR% /ilc pkeyconfig-office.xrm-ms
) else (
  cscript //Nologo //B "%_OSPP%" /inslic:pkeyconfig-office.xrm-ms
)

set SkuIds=(mondo,proplus,projectpro,visiopro,standard,projectstd,visiostd)
set ProSkuIds=(access,excel,onenote,outlook,powerpoint,publisher,skypeforbusiness,word)
set ProSkuId2=(o365proplus,professional)
set StdSkuIds=(excel,onenote,outlook,powerpoint,publisher,word)

for /d %%G in %SkuIds% do (
  set /a _%%G=0
)

for /d %%G in %ProSkuIds% do (
  set /a _%%G=0
)

for /d %%G in %ProSkuId2% do (
  set /a _%%G=0
)

for /f "tokens=2,*" %%G in ('"reg query HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration /v ProductReleaseIds" 2^>nul') do (
  set "ProductIds=%%H"
)

for /d %%G in %SkuIds% do (
  echo %ProductIds% | findstr /I /C:"%%Gretail" 1>nul && (
    set /a _%%G=1
  )
)

for /d %%G in %ProSkuIds% do (
  echo %ProductIds% | findstr /I /C:"%%Gretail" 1>nul && (
    set /a _%%G=1
  )
)

for /d %%G in %ProSkuId2% do (
  echo %ProductIds% | findstr /I /C:"%%Gretail" 1>nul && (
    set /a _%%G=1
  )
)

if %_mondo% EQU 1 (
  call :InsLic mondo
  goto :GVLK
)

for /d %%G in %SkuIds% do (
  if !_%%G! EQU 1 (
    call :InsLic %%G
  )
)

for /d %%G in %ProSkuId2% do (
  if !_%%G! EQU 1 if !_proplus! EQU 0 call :InsLic proplus
)

for /d %%G in %StdSkuIds% do (
  if !_%%G! EQU 1 if !_proplus! EQU 0 if !_standard! EQU 0 if !_professional! EQU 0 call :InsLic %%G
)

for /d %%G in (skypeforbusiness) do (
  if !_%%G! EQU 1 if !_proplus! EQU 0 call :InsLic %%G
)

for /d %%G in (access) do (
  if !_%%G! EQU 1 if !_proplus! EQU 0 !_professional! EQU 0 call :InsLic %%G
)
::=============================================================
:GVLK
echo.
echo Installing KMS Client Keys...
echo.
for /f "tokens=2 delims==" %%G in ('"wmic path %_MicrosoftProduct% where (Description like '%%Office 16, VOLUME_KMSCLIENT%%') get ID /value"') do (
  set "ActID=%%G"
  call :InsKey
)
goto :end
::=============================================================
:InsLic
for /f "delims=" %%G in ('dir /b /on %1VL_*.xrm-ms') do (
  if %_WinBuild% GEQ 9200 (
    cscript //Nologo //B %_SLMGR% /ilc %%G
  ) else (
    cscript //Nologo //B "%_OSPP%" /inslic:%%G
  )
)
exit /b
::=============================================================
:InsKey
for /f "tokens=2 delims==" %%G in ('"wmic path %_MicrosoftProduct% where ID='%ActID%' get Name /format:list"') do (
  echo %%G
)
(echo edition = "%ActID%"
echo Set keys = CreateObject ^("Scripting.Dictionary"^)
echo keys.Add "9caabccb-61b1-4b4b-8bec-d10a3c3ac2ce", "HFTND-W9MK4-8B7MJ-B6C4G-XQBR2"
echo keys.Add "d450596f-894d-49e0-966a-fd39ed4c4c64", "XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99"
echo keys.Add "dedfa23d-6ed1-45a6-85dc-63cae0546de6", "JNRGM-WHDWX-FJJG3-K47QV-DRTFM"
echo keys.Add "4f414197-0fc2-4c01-b68a-86cbb9ac254c", "YG9NW-3K39V-2T3HJ-93F3Q-G83KT"
echo keys.Add "da7ddabc-3fbe-4447-9e01-6ab7440b4cd4", "GNFHQ-F6YQM-KQDGJ-327XX-KQBVC"
echo keys.Add "6bf301c1-b94a-43e9-ba31-d494598c47fb", "PD3PC-RHNGV-FXJ29-8JK7D-RJRJK"
echo keys.Add "aa2a7821-1827-4c2c-8f1d-4513a34dda97", "7WHWN-4T7MP-G96JF-G33KR-W8GF4"
echo keys.Add "67c0fc0c-deba-401b-bf8b-9c8ad8395804", "GNH9Y-D2J4T-FJHGG-QRVH7-QPFDW"
echo keys.Add "c3e65d36-141f-4d2f-a303-a842ee756a29", "9C2PK-NWTVB-JMPW8-BFT28-7FTBF"
echo keys.Add "d8cace59-33d2-4ac7-9b1b-9b72339c51c8", "DR92N-9HTF2-97XKM-XW2WJ-XW3J6"
echo keys.Add "ec9d9265-9d1e-4ed0-838a-cdc20f2551a1", "R69KK-NTPKF-7M3Q4-QYBHW-6MT9B"
echo keys.Add "d70b1bba-b893-4544-96e2-b7a318091c33", "J7MQP-HNJ4Y-WJ7YM-PFYGF-BY6C6"
echo keys.Add "041a06cb-c5b8-4772-809f-416d03d16654", "F47MM-N3XJP-TQXJ9-BP99D-8K837"
echo keys.Add "83e04ee1-fa8d-436d-8994-d31a862cab77", "869NQ-FJ69K-466HW-QYCP2-DDBV6"
echo keys.Add "bb11badf-d8aa-470e-9311-20eaf80fe5cc", "WXY84-JN2Q9-RBCCQ-3Q3J3-3PFJ6"
echo if keys.Exists^(edition^) then
echo WScript.Echo keys.Item^(edition^)
echo End If)>"%temp%\key.vbs"
set "key=Unknown"
for /f %%G in ('cscript /nologo "%temp%\key.vbs"') do set key=%%G
del /f /q "%temp%\key.vbs" >nul 2>&1
if %key%==Unknown (
  echo.
  echo Could not find matching KMS Client key
  exit /b
)
wmic path %_MicrosoftService% where version='%_ver%' call InstallProductKey ProductKey="%key%" >nul 2>&1
exit /b
::=============================================================
:end
echo.
echo.
echo Press any key to exit...
pause >nul
exit