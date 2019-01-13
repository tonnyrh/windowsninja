@IF NOT DEFINED ECHO_ON ECHO OFF
@IF DEFINED ECHO_ON ECHO ON
@ECHO.
@ECHO WindowsNinja
@ECHO Tonny Roger Holm
@ECHO.
REM Version 1.1.0.68
SET COPYCMD=/Y
SET LOGPATH=%~dp0..
SET RESTART=%1

:: Special for choice: Only start if directory exist
:: if not exist "%~dp0..\Pictures\Info\%computername%" exit

IF NOT EXIST "%TEMP%\%~n0-tmp" MKDIR "%TEMP%\%~n0-tmp"
:: Check if new version. If so kill and restart

XCOPY /Y /D /L "%~dp0WindowsNinja.exe" "%TEMP%\%~n0-tmp\*.*"|find /i "1 Fil"
IF "%ERRORLEVEL%"=="0" SET RESTART=RESTART
ECHO RESTARTMODUS:%RESTART%
IF "%RESTART%"=="RESTART" TASKKILL /F /FI "USERNAME eq %USERDOMAIN%\%USERNAME%" /IM %~n0.exe

tasklist|find /i "%~n0.exe" >NUL:
IF "%ERRORLEVEL%"=="0" GOTO :EOF

XCOPY /S /D /H "%~dp0*.*" "%TEMP%\%~n0-tmp\*.*"
START  "%~n0" /D "%~dp0" /LOW "%TEMP%\%~n0-tmp\%~n0.exe" "%LOGPATH%"
rem START  "%~n0" /D "%~dp0" "%TEMP%\%~n0-tmp\%~n0.exe" "%LOGPATH%"
