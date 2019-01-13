:: V1.4
ECHO on

SET EXEPath=%~dp0
Set FiName=%~n0
FOR /F "eol=; tokens=1,2,3,4,5,6 delims=:, " %%i in ('ECHO %DATE%-%TIME%') do set filetime=%%i-%%j_%%k
::SET LOGGDIR=C:\temp\_LOG
::SET LOGG="%LOGGDIR%\%computername%-%username%-%finame%-%filetime%.log"
SET TMPFOLDER=%temp%\%FiName%



set subject="Error in %service%"
set body="An error was found, Se attached logfile"
Set FileToSend="%1"
Set Priority=1
Set bodyf="%2"

:: Account Settings
Set smtpserver=smtp.choice.no


IF NOT EXIST "%TMPFOLDER%" MKDIR "%TMPFOLDER%"




XCOPY /Y "%~dp0*.*" "%TMPFOLDER%\*.*" > NUL 2>&1



"%EXEPath%blat.exe" -install %smtpserver% %sender%
"%EXEPath%blat.exe" -attacht %FileToSend% -to %DestinationAddress% -subject %subject% -bodyf %bodyf% -Priority %Priority%
:: > %LOGG% 2>&1

:: Deleting logfiles older than 30 days
::IF EXIST "%LOGGDIR%" CSCRIPT /NOLOGO "%TMPFOLDER%\DeleteOldFiles.vbs" /DaysOld:30 /Folder:"%LOGGDIR%" /FilePrefix:"" /Filesuffix:.log /EXECUTE

:: Deleting archived files older than 30 days
::IF EXIST "%ArchiveDir%" CSCRIPT /NOLOGO "%TMPFOLDER%\DeleteOldFiles.vbs" /DaysOld:30 /Folder:"%ArchiveDir%" /FilePrefix:"" /Filesuffix:".sia" /EXECUTE

