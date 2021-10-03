@setlocal DisableDelayedExpansion
@set uivr=v44
@echo off
:: change to 0 to keep configured KMS cache upon removal (recommended only if you plan to reinstall)
set ClearKMSCache=1

:: change to 1 to enable debug mode
set _Debug=0

:: change to 1 to suppress any output
set Silent=0

:: change to 1 to redirect output to a text file, works only with Silent=1
set Logger=0

:: Notice for advanced users on Windows 64-bit (x64 / ARM64):
:: when you bundle KMS_VL_ALL script(s) inside self-extracting program or run it from another command script
:: if the exe pack or the caller script is running as 32-bit (x86) process
:: KMS_VL_ALL script(s) will close then relaunch itself using 64-bit (x64 / ARM64) cmd.exe
:: in that case, be advised not to proceed your pack or caller script depending on KMS_VL_ALL script(s) closure
:: instead, make sure the exe pack or the other caller script are already 64-bit (x64 / ARM64) process

:: ###################################################################
:: # NORMALLY THERE IS NO NEED TO CHANGE ANYTHING BELOW THIS COMMENT #
:: ###################################################################

set KMS_IP=0.0.0.0
set KMS_Port=1688
set KMS_Emulation=1
set Unattend=0

set "_Null=1>nul 2>nul"

set "_cmdf=%~f0"
if exist "%SystemRoot%\Sysnative\cmd.exe" (
setlocal EnableDelayedExpansion
start %SystemRoot%\Sysnative\cmd.exe /c ""!_cmdf!" %*"
exit /b
)
if exist "%SystemRoot%\SysArm32\cmd.exe" if /i %PROCESSOR_ARCHITECTURE%==AMD64 (
setlocal EnableDelayedExpansion
start %SystemRoot%\SysArm32\cmd.exe /c ""!_cmdf!" %*"
exit /b
)

set ForceIns=0
set ForceRem=0
set "_args=%*"
if not defined _args goto :NoProgArgs
if "%~1"=="" set "_args="&goto :NoProgArgs

set _args=%_args:"=%
for %%A in (%_args%) do (
if /i "%%A"=="/d" (set _Debug=1
) else if /i "%%A"=="/u" (set Unattend=1
) else if /i "%%A"=="/s" (set Silent=1
) else if /i "%%A"=="/l" (set Logger=1
) else if /i "%%A"=="/i" (set ForceIns=1&set ForceRem=0
) else if /i "%%A"=="/r" (set ForceIns=0&set ForceRem=1
) else if /i "%%A"=="/k" (set ClearKMSCache=0
)
)
if %ForceIns% EQU 1 set Unattend=1
if %ForceRem% EQU 1 set Unattend=1

:NoProgArgs
if %Silent% EQU 1 set Unattend=1
set "_run=nul"
if %Logger% EQU 1 set _run="%~dpn0_Silent.log"

set "SysPath=%SystemRoot%\System32"
if exist "%SystemRoot%\Sysnative\reg.exe" (set "SysPath=%SystemRoot%\Sysnative")
set "Path=%SysPath%;%SystemRoot%;%SysPath%\Wbem;%SysPath%\WindowsPowerShell\v1.0\"
set "_err===== ERROR ===="
set "_psc=powershell -nop -c"
set "_buf={$W=$Host.UI.RawUI.WindowSize;$B=$Host.UI.RawUI.BufferSize;$W.Height=27;$B.Height=300;$Host.UI.RawUI.WindowSize=$W;$Host.UI.RawUI.BufferSize=$B;}"
set "o_x64=684103f5c312ae956e66a02b965d9aad59710745"
set "o_x86=da8f931c7f3bc6643e20063e075cd8fa044b53ae"
set "o_arm=1139ae6243934ca621e6d4ed2e2f34cc130ef88a"
if /i "%PROCESSOR_ARCHITECTURE%"=="amd64" set "xBit=x64"&set "xOS=x64"&set "_orig=%o_x64%"
if /i "%PROCESSOR_ARCHITECTURE%"=="arm64" set "xBit=x86"&set "xOS=A64"&set "_orig=%o_arm%"
if /i "%PROCESSOR_ARCHITECTURE%"=="x86" if "%PROCESSOR_ARCHITEW6432%"=="" set "xBit=x86"&set "xOS=x86"&set "_orig=%o_x86%"
if /i "%PROCESSOR_ARCHITEW6432%"=="amd64" set "xBit=x64"&set "xOS=x64"&set "_orig=%o_x64%"
if /i "%PROCESSOR_ARCHITEW6432%"=="arm64" set "xBit=x86"&set "xOS=A64"&set "_orig=%o_arm%"

set _invpth=0
set "param=%~f0"
cmd /v:on /c echo(^^!param^^!| findstr /R "[| ` ~ ! @ %% \^ & ( ) \[ \] { } + = ; ' , |]*^" 1>nul 2>nul
if %errorlevel% EQU 0 set _invpth=1
reg query HKU\S-1-5-19 1>nul 2>nul || goto :E_Admin

set "_temp=%SystemRoot%\Temp"
set "_log=%~dpn0"
set "_work=%~dp0"
if "%_work:~-1%"=="\" set "_work=%_work:~0,-1%"
set _UNC=0
if "%_work:~0,2%"=="\\" set _UNC=1
for /f "skip=2 tokens=2*" %%a in ('reg query "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /v Desktop') do call set "_dsk=%%b"
if exist "%PUBLIC%\Desktop\desktop.ini" set "_dsk=%PUBLIC%\Desktop"
setlocal EnableDelayedExpansion

if %_Debug% EQU 0 (
  set "_Nul1=1>nul"
  set "_Nul2=2>nul"
  set "_Nul6=2^>nul"
  set "_Nul3=1>nul 2>nul"
  set "_Pause=pause >nul"
  if %Unattend% EQU 1 set "_Pause="
  if %Silent% EQU 0 (call :Begin) else (call :Begin >!_run! 2>&1)
) else (
  set "_Nul1="
  set "_Nul2="
  set "_Nul6="
  set "_Nul3="
  set "_Pause="
  copy /y nul "!_work!\#.rw" 1>nul 2>nul && (if exist "!_work!\#.rw" del /f /q "!_work!\#.rw") || (set "_log=!_dsk!\%~n0")
  if %Silent% EQU 0 (
  echo.
  echo Running in Debug Mode...
  if not defined _args (echo The window will be closed when finished) else (echo please wait...)
  echo.
  echo writing debug log to:
  echo "!_log!_Debug.log"
  )
  @echo on
  @prompt $G
  @call :Begin >"!_log!_tmp.log" 2>&1 &cmd /u /c type "!_log!_tmp.log">"!_log!_Debug.log"&del "!_log!_tmp.log"
)
@color 07
@title %ComSpec%
@echo off
@exit /b

:Begin
if %_Debug% EQU 1 (
if defined _args echo %_args%
)
set "_wApp=55c92734-d682-4d71-983e-d6ec3f16059f"
set "_oApp=0ff1ce15-a989-479d-af46-f275c6370663"
set "_oA14=59a52881-a989-479d-af46-f275c6370663"
set "IFEO=HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options"
set "OSPP=SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform"
set "OPPk=SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform"
set "SPPk=SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform"
set _Hook="%SysPath%\SppExtComObjHook.dll"
set w7inf=%SystemRoot%\Migration\WTR\KMS_VL_ALL.inf
set "_TaskEx=\Microsoft\Windows\SoftwareProtectionPlatform\SvcTrigger"
set "_TaskOs=\Microsoft\Windows\SoftwareProtectionPlatform\SvcRestartTaskLogon"
set "line3=____________________________________________________________"
for /f "tokens=6 delims=[]. " %%G in ('ver') do set winbuild=%%G
set SSppHook=0
for /f %%A in ('dir /b /ad %SysPath%\spp\tokens\skus') do (
  if %winbuild% GEQ 9200 if exist "%SysPath%\spp\tokens\skus\%%A\*GVLK*.xrm-ms" set SSppHook=1
  if %winbuild% LSS 9200 if exist "%SysPath%\spp\tokens\skus\%%A\*VLKMS*.xrm-ms" set SSppHook=1
  if %winbuild% LSS 9200 if exist "%SysPath%\spp\tokens\skus\%%A\*VL-BYPASS*.xrm-ms" set SSppHook=1
)
set OsppHook=1
sc query osppsvc %_Nul3%
if %errorlevel% equ 1060 set OsppHook=0

set ESU_KMS=0
if %winbuild% LSS 9200 for /f %%A in ('dir /b /ad %SysPath%\spp\tokens\channels') do (
  if exist "%SysPath%\spp\tokens\channels\%%A\*VL-BYPASS*.xrm-ms" set ESU_KMS=1
)
set ESU_EDT=0
if %ESU_KMS% EQU 1 for %%A in (%ESUEditions%) do (
  if exist "%SysPath%\spp\tokens\skus\Security-SPP-Component-SKU-%%A\*.xrm-ms" set ESU_EDT=1
)
if %ESU_EDT% EQU 1 set SSppHook=1

if %winbuild% GEQ 9200 (
  set OSType=Win8
  set SppVer=SppExtComObj.exe
) else if %winbuild% GEQ 7600 (
  set OSType=Win7
  set SppVer=sppsvc.exe
) else (
  goto :UnsupportedVersion
)
if %OSType% EQU Win8 reg query "%IFEO%\sppsvc.exe" %_Nul3% && (
reg delete "%IFEO%\sppsvc.exe" /f %_Nul3%
call :StopService sppsvc
)

color 07
if %Unattend% EQU 0 title Auto Renewal Setup %uivr%
if %Silent% EQU 0 if %_Debug% EQU 0 mode con cols=100 lines=28

if %ForceIns% EQU 1 goto :inst
if %ForceRem% EQU 1 goto :remv
if exist %_Hook% dir /b /al %_Hook% %_Nul3% || goto :remv
reg query "%IFEO%\%SppVer%" /v VerifierFlags %_Nul3% && goto :remv
reg query "%IFEO%\osppsvc.exe" /v VerifierFlags %_Nul3% && goto :remv
if not exist "!_work!\bin\%xOS%.dll" goto :E_DLL

:inst
echo.
if %_Debug% NEQ 0 goto :pinst
if %Unattend% NEQ 0 (
echo Mode: Installation
goto :pinst
)
choice /C YN /N /M "Local KMS Emulator will be installed on your computer. Continue? [y/n]: "
if errorlevel 2 exit /b
:pinst
echo.
echo %line3%
call :StopService sppsvc
if %OsppHook% NEQ 0 call :StopService osppsvc
if %winbuild% GEQ 9600 (
  reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows NT\CurrentVersion\Software Protection Platform" /v NoGenTicket /t REG_DWORD /d 1 /f %_Nul3%
  if %winbuild% EQU 14393 reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows NT\CurrentVersion\Software Protection Platform" /v NoAcquireGT /t REG_DWORD /d 1 /f %_Nul3%
  WMIC /NAMESPACE:\\root\Microsoft\Windows\Defender PATH MSFT_MpPreference call Add ExclusionPath=%_Hook% Force=True %_Nul3% && set "AddExc= and Windows Defender exclusion"
)
echo.
echo Adding File%AddExc%...
echo %SystemRoot%\System32\SppExtComObjHook.dll
for %%# in (SppExtComObjHookAvrf.dll,SppExtComObjHook.dll,SppExtComObjPatcher.dll,SppExtComObjPatcher.exe) do (
  if exist "%SysPath%\%%#" del /f /q "%SysPath%\%%#" %_Nul3%
  if exist "%SystemRoot%\SysWOW64\%%#" del /f /q "%SystemRoot%\SysWOW64\%%#" %_Nul3%
)
for /f "skip=1 tokens=* delims=" %%# in ('certutil -hashfile "!_work!\bin\%xOS%.dll" SHA1^|findstr /i /v CertUtil') do set "_hash=%%#"
set "_hash=%_hash: =%"
if /i not "%_hash%"=="%_orig%" (
echo.
echo === WARNING ===
echo SHA1 hash verification mismatch.
echo "bin\%xOS%.dll"
echo Expected: %_orig%
echo Detected: %_hash%
echo.
echo If you compiled the file yourself, then ignore this message.
)
copy /y "!_work!\bin\%xOS%.dll" %_Hook% %_Nul3% || (echo Failed&del /f /q %_Hook%&goto :TheEnd)
echo.
echo Adding Registry Keys...
if %SSppHook% NEQ 0 call :CreateIFEOEntry %SppVer%
call :CreateIFEOEntry osppsvc.exe
if %OSType% EQU Win7 (
call :CreateIFEOEntry SppExtComObj.exe
if %SSppHook% NEQ 0 if not exist %w7inf% (
  echo.&echo Adding migration fail-safe...&echo %w7inf%
  if not exist "%SystemRoot%\Migration\WTR" md "%SystemRoot%\Migration\WTR"
  (
  echo [WTR]
  echo Name="KMS_VL_ALL"
  echo.
  echo [WTR.W8]
  echo NotifyUser="No"
  echo.
  echo [System.Registry]
  echo "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\sppsvc.exe [*]"
  )>%w7inf%
  )
)
if %OSType% EQU Win8 call :CreateTask
if not exist "!_work!\Activate.cmd" (
echo %line3%
echo.
echo %_err%
echo Activate.cmd is missing, skipping activation...
goto :einst
)
if %Silent% EQU 0 if %_Debug% EQU 0 (
%_Nul3% %_psc% "&%_buf%"
if %Unattend% EQU 0 title Auto Renewal Setup %uivr%
)
echo.
echo %line3%
set "_para=/u"
if %_Debug% EQU 1 set "_para=!_para! /d"
if %Silent% EQU 1 set "_para=!_para! /s"
if %Logger% EQU 1 set "_para=!_para! /l"
cmd.exe /c ""!_work!\Activate.cmd" !_para!"
if %Unattend% EQU 0 title Auto Renewal Setup %uivr%
:einst
echo %line3%
echo.
echo Done.
echo Make sure to exclude this file in the Antivirus protection.
echo %SystemRoot%\System32\SppExtComObjHook.dll
goto :TheEnd

:remv
echo.
if %_Debug% NEQ 0 goto :premv
if %Unattend% NEQ 0 (
echo Mode: Removal
goto :premv
)
choice /C YN /N /M "Local KMS Emulator will be removed from your computer. Continue? [y/n]: "
if errorlevel 2 exit /b
:premv
echo.
echo %line3%
call :StopService sppsvc
if %OsppHook% NEQ 0 call :StopService osppsvc
if %winbuild% GEQ 9600 (
  for %%# in (NoGenTicket,NoAcquireGT) do reg delete "HKLM\SOFTWARE\Policies\Microsoft\Windows NT\CurrentVersion\Software Protection Platform" /v %%# /f %_Null%
  WMIC /NAMESPACE:\\root\Microsoft\Windows\Defender PATH MSFT_MpPreference call Remove ExclusionPath=%_Hook% Force=True %_Nul3% && set "RemExc= and Windows Defender exclusions"
)
echo.
echo Removing Files%RemExc%...
for %%# in (SppExtComObjHookAvrf.dll,SppExtComObjHook.dll,SppExtComObjPatcher.dll,SppExtComObjPatcher.exe) do if exist "%SysPath%\%%#" (
  echo %SystemRoot%\System32\%%#
  del /f /q "%SysPath%\%%#" %_Nul3%
)
for %%# in (SppExtComObjHookAvrf.dll,SppExtComObjHook.dll,SppExtComObjPatcher.dll,SppExtComObjPatcher.exe) do if exist "%SystemRoot%\SysWOW64\%%#" (
  echo %SystemRoot%\SysWOW64\%%#
  del /f /q "%SystemRoot%\SysWOW64\%%#" %_Nul3%
)
if exist %w7inf% (
	echo %w7inf%
	del /f /q %w7inf%
)
echo.
echo Removing Registry Keys...
for %%# in (SppExtComObj.exe,sppsvc.exe,osppsvc.exe) do reg query "%IFEO%\%%#" %_Nul3% && (
  call :RemoveIFEOEntry %%#
)
if %OSType% EQU Win8 schtasks /query /tn "%_TaskEx%" %_Nul3% && (
echo.
echo Removing Schedule Task...
echo %_TaskEx%
schtasks /delete /f /tn "%_TaskEx%" %_Nul3%
)
if %ClearKMSCache% EQU 1 (call :cCache) else (call :cREG %_Nul3%)
echo.
echo %line3%
echo.
echo Done.
goto :TheEnd

:cREG
reg add "HKLM\%SPPk%" /f /v KeyManagementServiceName /t REG_SZ /d "0.0.0.0"
reg add "HKLM\%SPPk%" /f /v KeyManagementServicePort /t REG_SZ /d "1688"
reg delete "HKLM\%SPPk%" /f /v DisableDnsPublishing
reg delete "HKLM\%SPPk%" /f /v DisableKeyManagementServiceHostCaching
reg delete "HKLM\%SPPk%\%_wApp%" /f
if %winbuild% GEQ 9200 (
if not %xOS%==x86 (
reg add "HKLM\%SPPk%" /f /v KeyManagementServiceName /t REG_SZ /d "0.0.0.0" /reg:32
reg add "HKLM\%SPPk%" /f /v KeyManagementServicePort /t REG_SZ /d "1688" /reg:32
reg delete "HKLM\%SPPk%\%_oApp%" /f /reg:32
reg add "HKLM\%SPPk%\%_oApp%" /f /v KeyManagementServiceName /t REG_SZ /d "0.0.0.0" /reg:32
reg add "HKLM\%SPPk%\%_oApp%" /f /v KeyManagementServicePort /t REG_SZ /d "1688" /reg:32
)
reg delete "HKLM\%SPPk%\%_oApp%" /f
reg add "HKLM\%SPPk%\%_oApp%" /f /v KeyManagementServiceName /t REG_SZ /d "0.0.0.0"
reg add "HKLM\%SPPk%\%_oApp%" /f /v KeyManagementServicePort /t REG_SZ /d "1688"
)
if %winbuild% GEQ 9600 (
reg delete "HKU\S-1-5-20\%SPPk%\%_wApp%" /f
reg delete "HKU\S-1-5-20\%SPPk%\%_oApp%" /f
)
if %OsppHook% EQU 0 (
goto :eof
)
reg add "HKLM\%OPPk%" /f /v KeyManagementServiceName /t REG_SZ /d "0.0.0.0"
reg delete "HKLM\%OPPk%" /f /v KeyManagementServicePort
reg delete "HKLM\%OPPk%" /f /v DisableDnsPublishing
reg delete "HKLM\%OPPk%" /f /v DisableKeyManagementServiceHostCaching
reg delete "HKLM\%OPPk%\%_oA14%" /f
reg delete "HKLM\%OPPk%\%_oApp%" /f
goto :eof

:rREG
reg delete "HKLM\%SPPk%" /f /v KeyManagementServiceName
reg delete "HKLM\%SPPk%" /f /v KeyManagementServicePort
reg delete "HKLM\%SPPk%" /f /v DisableDnsPublishing
reg delete "HKLM\%SPPk%" /f /v DisableKeyManagementServiceHostCaching
reg delete "HKLM\%SPPk%\%_wApp%" /f
if %winbuild% GEQ 9200 (
if not %xOS%==x86 (
reg delete "HKLM\%SPPk%" /f /v KeyManagementServiceName /reg:32
reg delete "HKLM\%SPPk%" /f /v KeyManagementServicePort /reg:32
reg delete "HKLM\%SPPk%\%_oApp%" /f /reg:32
)
reg delete "HKLM\%SPPk%\%_oApp%" /f
)
if %winbuild% GEQ 9600 (
reg delete "HKU\S-1-5-20\%SPPk%\%_wApp%" /f
reg delete "HKU\S-1-5-20\%SPPk%\%_oApp%" /f
)
reg delete "HKLM\%OPPk%" /f /v KeyManagementServiceName
reg delete "HKLM\%OPPk%" /f /v KeyManagementServicePort
reg delete "HKLM\%OPPk%" /f /v DisableDnsPublishing
reg delete "HKLM\%OPPk%" /f /v DisableKeyManagementServiceHostCaching
reg delete "HKLM\%OPPk%\%_oA14%" /f
reg delete "HKLM\%OPPk%\%_oApp%" /f
goto :eof

:cCache
echo.
echo Clearing KMS Cache...
call :rREG %_Nul3%
set "_C16R="
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\ClickToRun /v InstallPath" %_Nul6%') do if exist "%%b\root\Licenses16\ProPlus*.xrm-ms" set "_C16R=1"
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\ClickToRun /v InstallPath /reg:32" %_Nul6%') do if exist "%%b\root\Licenses16\ProPlus*.xrm-ms" set "_C16R=1"
if %winbuild% GEQ 9200 if defined _C16R (
echo.
echo ## Notice ##
echo.
echo To make sure Office programs do not show a non-genuine banner
echo please apply manual or auto-renewal activation, and don't uninstall afterward.
)
goto :eof

:StopService
sc query %1 | find /i "STOPPED" %_Nul1% || net stop %1 /y %_Nul3%
sc query %1 | find /i "STOPPED" %_Nul1% || sc stop %1 %_Nul3%
goto :eof

:CreateIFEOEntry
echo [%IFEO%\%1]
reg delete "%IFEO%\%1" /f /v Debugger %_Null%
reg add "%IFEO%\%1" /f /v VerifierDlls /t REG_SZ /d "SppExtComObjHook.dll" %_Nul3% || (echo Failed&del /f /q %_Hook%&goto :TheEnd)
reg add "%IFEO%\%1" /f /v VerifierDebug /t REG_DWORD /d 0x00000000 %_Nul3%
reg add "%IFEO%\%1" /f /v VerifierFlags /t REG_DWORD /d 0x80000000 %_Nul3%
reg add "%IFEO%\%1" /f /v GlobalFlag /t REG_DWORD /d 0x00000100 %_Nul3%
reg add "%IFEO%\%1" /f /v KMS_Emulation /t REG_DWORD /d %KMS_Emulation% %_Nul3%
if /i %1 EQU osppsvc.exe (
reg add "HKLM\%OSPP%" /f /v KeyManagementServiceName /t REG_SZ /d %KMS_IP% %_Nul3%
reg add "HKLM\%OSPP%" /f /v KeyManagementServicePort /t REG_SZ /d %KMS_Port% %_Nul3%
)
goto :eof

:RemoveIFEOEntry
echo [%IFEO%\%1]
if /i %1 NEQ osppsvc.exe (
reg delete "%IFEO%\%1" /f %_Null%
goto :eof
)
if %OsppHook% EQU 0 (
reg delete "%IFEO%\%1" /f %_Null%
)
if %OsppHook% NEQ 0 for %%A in (Debugger,VerifierDlls,VerifierDebug,VerifierFlags,GlobalFlag,KMS_Emulation,KMS_ActivationInterval,KMS_RenewalInterval,Office2010,Office2013,Office2016,Office2019) do reg delete "%IFEO%\%1" /v %%A /f %_Null%
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
schtasks /query /tn "%_TaskEx%" %_Nul3% && (
echo.
echo Adding Schedule Task...
echo %_TaskEx%
)
goto :eof

:E_Admin
echo %_err%
echo This script requires administrator privileges.
echo To do so, right-click on this script and select 'Run as administrator'
echo.
if %_invpth% EQU 1 (
echo.
echo === WARNING ===
echo Disallowed special characters are detected in the file path name.
echo Before you can use 'Run as administrator' successfully,
echo make sure the path do not have any of the following characters:
echo ^` ^~ ^! ^@ %% ^^ ^& ^( ^) [ ] { } ^+ ^= ^; ^' ^,
echo.
)
echo Press any key to exit.
if %_Debug% EQU 1 goto :eof
if %Unattend% EQU 1 goto :eof
pause >nul
goto :eof

:E_DLL
echo.
echo %_err%
echo Required file bin\%xOS%.dll is not found.
echo Verify that Antivirus protection is OFF or the current folder is excluded.
goto :TheEnd

:UnsupportedVersion
echo %_err%
echo Unsupported OS version Detected.
echo Project is supported only for Windows 7/8/8.1/10/11 and their Server equivalent.
:TheEnd
echo.
if %Unattend% EQU 0 echo Press any key to exit.
%_Pause%
goto :eof