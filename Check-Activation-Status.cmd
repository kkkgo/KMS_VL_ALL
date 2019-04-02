@echo off
set "SysPath=%Windir%\System32"
if exist "%Windir%\Sysnative\reg.exe" (set "SysPath=%Windir%\Sysnative")
set "Path=%SysPath%;%Windir%;%SysPath%\Wbem;%SysPath%\WindowsPowerShell\v1.0\"
set "_tempdir=%SystemRoot%\Temp"
set bit=64&set wow=1
if /i "%PROCESSOR_ARCHITECTURE%"=="x86" (if "%PROCESSOR_ARCHITEW6432%"=="" set bit=32&set wow=0)
set "line=************************************************************"
setlocal EnableDelayedExpansion
echo %line%
echo ***                   Windows Status                     ***
echo %line%
copy /y %Windir%\System32\slmgr.vbs "!_tempdir!\slmgr.vbs" >nul 2>&1
cscript //nologo "!_tempdir!\slmgr.vbs" /dli || (echo Error executing slmgr.vbs&del /f /q "!_tempdir!\slmgr.vbs"&goto :End)
cscript //nologo "!_tempdir!\slmgr.vbs" /xpr
del /f /q "!_tempdir!\slmgr.vbs" >nul 2>&1
echo ____________________________________________________________________________

:office2016
set office=
for /f "tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\16.0\Common\InstallRoot /v Path" 2^>nul') do (set "office=%%b")
if exist "!office!\OSPP.VBS" (
echo.
echo %line%
echo ***              Office 2016 %bit%-bit Status               ***
echo %line%
cscript //nologo "!office!\OSPP.VBS" /dstatus
)
if %wow%==0 goto :office2013
set office=
for /f "tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\16.0\Common\InstallRoot /v Path" 2^>nul') do (set "office=%%b")
if exist "!office!\OSPP.VBS" (
echo.
echo %line%
echo ***              Office 2016 32-bit Status               ***
echo %line%
cscript //nologo "!office!\OSPP.VBS" /dstatus
)

:office2013
set office=
for /f "tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\15.0\Common\InstallRoot /v Path" 2^>nul') do (set "office=%%b")
if exist "!office!\OSPP.VBS" (
echo.
echo %line%
echo ***              Office 2013 %bit%-bit Status               ***
echo %line%
cscript //nologo "!office!\OSPP.VBS" /dstatus
)
if %wow%==0 goto :office2010
set office=
for /f "tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\15.0\Common\InstallRoot /v Path" 2^>nul') do (set "office=%%b")
if exist "!office!\OSPP.VBS" (
echo.
echo %line%
echo ***              Office 2013 32-bit Status               ***
echo %line%
cscript //nologo "!office!\OSPP.VBS" /dstatus
)

:office2010
set office=
for /f "tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\14.0\Common\InstallRoot /v Path" 2^>nul') do (set "office=%%b")
if exist "!office!\OSPP.VBS" (
echo.
echo %line%
echo ***              Office 2010 %bit%-bit Status               ***
echo %line%
cscript //nologo "!office!\OSPP.VBS" /dstatus
)
if %wow%==0 goto :office2016C2R
set office=
for /f "tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\14.0\Common\InstallRoot /v Path" 2^>nul') do (set "office=%%b")
if exist "!office!\OSPP.VBS" (
echo.
echo %line%
echo ***              Office 2010 32-bit Status               ***
echo %line%
cscript //nologo "!office!\OSPP.VBS" /dstatus
)

:office2016C2R
reg query HKLM\SOFTWARE\Microsoft\Office\ClickToRun /v InstallPath >nul 2>&1 || goto :office2013C2R
set office=
for /f "tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\ClickToRun /v InstallPath" 2^>nul') do (set "office=%%b\Office16")
if exist "!office!\OSPP.VBS" (
echo.
echo %line%
echo ***              Office 2016/2019 C2R Status             ***
echo %line%
cscript //nologo "!office!\OSPP.VBS" /dstatus
)

:office2013C2R
reg query HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun /v InstallPath >nul 2>&1 || goto :office2010C2R
set office=
if exist "%ProgramFiles%\Microsoft Office\Office15\OSPP.VBS" (
  set "office=%ProgramFiles%\Microsoft Office\Office15"
) else if exist "%ProgramFiles(x86)%\Microsoft Office\Office15\OSPP.VBS" (
  set "office=%ProgramFiles(x86)%\Microsoft Office\Office15"
)
if exist "!office!\OSPP.VBS" (
echo.
echo %line%
echo ***                Office 2013 C2R Status                ***
echo %line%
cscript //nologo "!office!\OSPP.VBS" /dstatus
)

:office2010C2R
reg query HKLM\SOFTWARE\Microsoft\Office\14.0\ClickToRun /v InstallPath >nul 2>&1 || goto :End
set office=
if exist "%ProgramFiles%\Microsoft Office\Office14\OSPP.VBS" (
  set "office=%ProgramFiles%\Microsoft Office\Office14"
) else if exist "%ProgramFiles(x86)%\Microsoft Office\Office14\OSPP.VBS" (
  set "office=%ProgramFiles(x86)%\Microsoft Office\Office14"
)
if exist "!office!\OSPP.VBS" (
echo.
echo %line%
echo ***                Office 2010 C2R Status                ***
echo %line%
cscript //nologo "!office!\OSPP.VBS" /dstatus
)

:End
echo.
echo Press any key to exit...
PAUSE >NUL
EXIT /B