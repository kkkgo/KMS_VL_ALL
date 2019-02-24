@ECHO OFF
set "SysPath=%Windir%\System32"
if exist "%Windir%\Sysnative\reg.exe" (set "SysPath=%Windir%\Sysnative")
set "Path=%SysPath%;%Windir%;%SysPath%\Wbem;%SysPath%\WindowsPowerShell\v1.0\"
set "_tempdir=%temp%"
setlocal EnableExtensions EnableDelayedExpansion
ECHO ************************************************************
ECHO ***                   Windows Status                     ***
ECHO ************************************************************
COPY /Y %Windir%\System32\slmgr.vbs "!_tempdir!\slmgr.vbs" >NUL 2>&1
cscript //nologo "!_tempdir!\slmgr.vbs" /dli || (ECHO Error running vbs script&DEL /F /Q "!_tempdir!\slmgr.vbs"&GOTO :End)
cscript //nologo "!_tempdir!\slmgr.vbs" /xpr
DEL /F /Q "!_tempdir!\slmgr.vbs" >NUL 2>&1
ECHO ____________________________________________________________________________

:office2016
IF EXIST %Windir%\SysWOW64\cmd.exe (SET bit=64&SET wow=1) ELSE (SET bit=32&SET wow=0)
SET office=
FOR /F "tokens=2*" %%a IN ('"REG QUERY HKLM\SOFTWARE\Microsoft\Office\16.0\Common\InstallRoot /v Path" 2^>NUL') DO (SET "office=%%b")
IF EXIST "%office%\OSPP.VBS" (
ECHO.
ECHO ************************************************************
ECHO ***              Office 2016 %bit%-bit Status               ***
ECHO ************************************************************
cscript //nologo "%office%\OSPP.VBS" /dstatus
)
IF %wow%==0 GOTO :office2013
SET office=
FOR /F "tokens=2*" %%a IN ('"REG QUERY HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\16.0\Common\InstallRoot /v Path" 2^>NUL') DO (SET "office=%%b")
IF EXIST "%office%\OSPP.VBS" (
ECHO.
ECHO ************************************************************
ECHO ***              Office 2016 32-bit Status               ***
ECHO ************************************************************
cscript //nologo "%office%\OSPP.VBS" /dstatus
)

:office2013
SET office=
FOR /F "tokens=2*" %%a IN ('"REG QUERY HKLM\SOFTWARE\Microsoft\Office\15.0\Common\InstallRoot /v Path" 2^>NUL') DO (SET "office=%%b")
IF EXIST "%office%\OSPP.VBS" (
ECHO.
ECHO ************************************************************
ECHO ***              Office 2013 %bit%-bit Status               ***
ECHO ************************************************************
cscript //nologo "%office%\OSPP.VBS" /dstatus
)
IF %wow%==0 GOTO :office2010
SET office=
FOR /F "tokens=2*" %%a IN ('"REG QUERY HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\15.0\Common\InstallRoot /v Path" 2^>NUL') DO (SET "office=%%b")
IF EXIST "%office%\OSPP.VBS" (
ECHO.
ECHO ************************************************************
ECHO ***              Office 2013 32-bit Status               ***
ECHO ************************************************************
cscript //nologo "%office%\OSPP.VBS" /dstatus
)

:office2010
SET office=
FOR /F "tokens=2*" %%a IN ('"REG QUERY HKLM\SOFTWARE\Microsoft\Office\14.0\Common\InstallRoot /v Path" 2^>NUL') DO (SET "office=%%b")
IF EXIST "%office%\OSPP.VBS" (
ECHO.
ECHO ************************************************************
ECHO ***              Office 2010 %bit%-bit Status               ***
ECHO ************************************************************
cscript //nologo "%office%\OSPP.VBS" /dstatus
)
IF %wow%==0 GOTO :office2016C2R
SET office=
FOR /F "tokens=2*" %%a IN ('"REG QUERY HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\14.0\Common\InstallRoot /v Path" 2^>NUL') DO (SET "office=%%b")
IF EXIST "%office%\OSPP.VBS" (
ECHO.
ECHO ************************************************************
ECHO ***              Office 2010 32-bit Status               ***
ECHO ************************************************************
cscript //nologo "%office%\OSPP.VBS" /dstatus
)

:office2016C2R
REG QUERY HKLM\SOFTWARE\Microsoft\Office\ClickToRun /v InstallPath >NUL 2>&1 || GOTO :office2013C2R
SET office=
for /f "tokens=2*" %%a IN ('"REG QUERY HKLM\SOFTWARE\Microsoft\Office\ClickToRun /v InstallPath" 2^>NUL') do (set "office=%%b\Office16")
IF EXIST "%office%\OSPP.VBS" (
ECHO.
ECHO ************************************************************
ECHO ***              Office 2016/2019 C2R Status             ***
ECHO ************************************************************
cscript //nologo "%office%\OSPP.VBS" /dstatus
)

:office2013C2R
REG QUERY HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun /v InstallPath >NUL 2>&1 || GOTO :office2010C2R
SET office=
IF EXIST "%ProgramFiles%\Microsoft Office\Office15\OSPP.VBS" (
  set "office=%ProgramFiles%\Microsoft Office\Office15"
) else IF EXIST "%ProgramFiles(x86)%\Microsoft Office\Office15\OSPP.VBS" (
  set "office=%ProgramFiles(x86)%\Microsoft Office\Office15"
)
IF EXIST "%office%\OSPP.VBS" (
ECHO.
ECHO ************************************************************
ECHO ***                Office 2013 C2R Status                ***
ECHO ************************************************************
cscript //nologo "%office%\OSPP.VBS" /dstatus
)

:office2010C2R
REG QUERY HKLM\SOFTWARE\Microsoft\Office\14.0\ClickToRun /v InstallPath >NUL 2>&1 || GOTO :End
SET office=
IF EXIST "%ProgramFiles%\Microsoft Office\Office14\OSPP.VBS" (
  set "office=%ProgramFiles%\Microsoft Office\Office14"
) else IF EXIST "%ProgramFiles(x86)%\Microsoft Office\Office14\OSPP.VBS" (
  set "office=%ProgramFiles(x86)%\Microsoft Office\Office14"
)
IF EXIST "%office%\OSPP.VBS" (
ECHO.
ECHO ************************************************************
ECHO ***                Office 2010 C2R Status                ***
ECHO ************************************************************
cscript //nologo "%office%\OSPP.VBS" /dstatus
)

:End
ECHO.
ECHO Press any key to exit...
PAUSE >NUL
EXIT /B