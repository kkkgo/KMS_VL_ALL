@echo off
cd /d "%~dp0" && ( if exist "%temp%\getadmin.vbs" del "%temp%\getadmin.vbs" ) && fsutil dirty query %systemdrive% 1>nul 2>nul || (  cmd /u /c echo Set UAC = CreateObject^("Shell.Application"^) : UAC.ShellExecute "cmd.exe", "/k cd ""%~dp0"" && ""%~dpnx0""", "", "runas", 1 >> "%temp%\getadmin.vbs" && "%temp%\getadmin.vbs" 1>nul 2>nul && exit /B )

echo ************************************************************
echo ***                   Windows Status                     ***
echo ************************************************************
copy /y %systemroot%\System32\slmgr.vbs "%temp%\slmgr.vbs" >nul 2>&1
cscript //nologo "%temp%\slmgr.vbs" /dli
cscript //nologo "%temp%\slmgr.vbs" /xpr
echo ____________________________________________________________________________
echo.

set verb=0
set spp=SoftwareLicensingProduct
for /f "tokens=2 delims==" %%G in ('"wmic path %spp% where (PartialProductKey is not null) get ID /value"') do (set app=%%G&call :chk)
del /f /q "%temp%\slmgr.vbs" >nul 2>&1
echo.
echo Press any key to exit...
pause >nul
exit

:chk
wmic path %spp% where ID='%app%' get Name /value | findstr /i "Windows" 1>nul && (exit /b)
if %verb%==0 (
set verb=1
echo ************************************************************
echo ***                   Office Status                      ***
echo ************************************************************
)
cscript //nologo "%temp%\slmgr.vbs" /dli %app%
cscript //nologo "%temp%\slmgr.vbs" /xpr %app%
echo ____________________________________________________________________________
echo.
exit /b