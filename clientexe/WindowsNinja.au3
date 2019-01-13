#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Res_Description=WindowsNinja for monitoring onscreen activity
#AutoIt3Wrapper_Res_Fileversion=1.1.0.68
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_LegalCopyright=Tonny Roger Holm
#AutoIt3Wrapper_Res_Icon_Add=WindowsNinja.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

; "Dude, quit playing window ninja over there. Why can't you just use this tool and be chill?"

; Todo


#include <MsgBoxConstants.au3>
#include <FileConstants.au3>
#include <MsgBoxConstants.au3>
#include <WinAPIFiles.au3>
#include <File.au3>
#include <Array.au3>
#include <String.au3>
#include <ScreenCapture.au3>
#include <WinAPISys.au3>
#include <Date.au3>
#include <EventLog.au3>
#include <WinAPIProc.au3>
#include <TrayConstants.au3>
#include <GUIConstantsEx.au3>
#include <FontConstants.au3>
#include <Inet.au3>

# JSON implementation
; #include "JSON.au3"
; #include "JSON_Translate.au3"


Global $bDebugFlag = False
Global $bDebugFlagArray = False
Global $bDebugFlagMouse = False
Global $bListQwinstaInCurrentLogg = False
Global $bEnforceBringToConsole = False
Global $aActionArrayMissingWindowList[2000] ;Storage for announced missingList.

Global $bPrintWindowsTextInLogg = False
Global $bPrintAllActivity = False
Global $Version = FileGetVersion(@AutoItExe)
Global $bList[1][10] ; This is the list for historic data cache in the GetAllWindowsInfo Function
Global $bMainExist = False
Global $bErrorCaptureWindowOnly = False
Global $bFirstrun = True
Global $bScreensaverDisable = False

;Get Environment Variables
Global $sEnvTemp = EnvGet("TEMP")
Global $sEnvComputername = EnvGet("COMPUTERNAME")
Global $sEnvUserName = EnvGet("USERNAME")
Global $sEnvSESSIONNAME = EnvGet("SESSIONNAME")

Global $sEnvDEBUG=EnvGet("DEBUG")
if $sEnvDEBUG="TRUE" then Global $bDebugFlag = True
Global $sDebugFile = $sEnvTemp & "\WindowsNinja-debug.log"
If FileExists($sDebugFile) then FileDelete($sDebugFile)
Global $hDebugFile = FileOpen($sDebugFile, $FO_READ + $FO_OVERWRITE)
Global $iDebugMaxLines=20000
Global $iDebugLinesCount=0

Global $sWindowsNinjaPath = ""
$sWindowsNinjaPath = EnvGet("WindowsNinjaPath")
Global $sBatBane = _PathFull(@ScriptDir)
DebugLogger('- ' & @ScriptLineNumber & " --> " & "Batbane:" & $sBatBane & @CRLF)
Global $sFileName = @YEAR & @MON & @MDAY & "_WindowsNinja+" & $sEnvComputername & "+" & $sEnvUserName & ".csv"
Global $sDirLogsStub = "Logs"
Global $sDirPicturesBaseStub = "Pictures"
Global $sDirPictureErrorStub = $sDirPicturesBaseStub & "\Error"
Global $sDirPictureHistoricStub = $sDirPicturesBaseStub & "\Historic\" & $sEnvComputername
Global $sDirPictureCurrentStub = $sDirPicturesBaseStub & "\Current"
Global $DirStatusInfoStub = "Status"
Global $iPicturSleepTime = 0
Global $bNoAction = False
Global $iSleepTime = 3 ; Default sleep time
Global $SleepTimeWindow=1000 ; When in sleep this sets the window for polling other routines inside the sleep loop.
Global $iPeakProcessMemoryLimit = 100000000 ; No more than 100 Kb
Global $sPictureFileLastWritten = ""
Global $MousexHistory = MouseGetPos(0)
Global $MouseyHistory = MouseGetPos(1)
Global $TrayTipHeader = "WindowsNinja"
Global $bCSVConfigLoaded = False
Global $bUserIsConnected = False
Global $bUserIsConnectedHistory = False
Global $bNoActionWhileUserConnected = False
Global $bMailEnable=False
Global $BlatSenderSuffix="_noreply.WindowsNinja@" & EnvGet("USERDNSDOMAIN")
Global $BlatSmtpServer
Global $BlatSender
Global $BlatAttach
Global $BlatTo
Global $BlatSubject
Global $BlatBody
Global $BlatPriority
Global $bMouseMoved=False
Global $bTsconToConsoleInitiated=false
Global $iMouseMovedDetectPauseTime=0
Global $bMouseMovedArmed=False
Global $iMouseMovedDetectPauseTimeLeft=0
Global $MousexHistoryActivity = MouseGetPos(0)
Global $MouseyHistoryActivity = MouseGetPos(1)
Global $bSleepEnable = False ; Set to false to skip sleep interval
Global $sAu3InfoExe = StringLeft(@AutoItExe, StringInStr(@AutoItExe, "\", 0, -1) - 1) & "\Tools\Au3Info.exe"
Global $sCurlExe = StringLeft(@AutoItExe, StringInStr(@AutoItExe, "\", 0, -1) - 1) & "\Tools\Curl\curl.exe"
Global $sTestPicture = StringLeft(@AutoItExe, StringInStr(@AutoItExe, "\", 0, -1) - 1) & "\Tools\ArtWork\WindowsNinja-Test.jpg"
Global $sErrorPicture = StringLeft(@AutoItExe, StringInStr(@AutoItExe, "\", 0, -1) - 1) & "\Tools\ArtWork\WindowsNinja-Error.jpg"
; CPU limits
Global $iPeakCPULimit = 0
Global $bCheckCPULimit = False
Global $bCPULoadOK=True
Global $CPULoadPercent

; Slack parameters
Global $sSlackToken
Global $sSlackUploadURL="https://slack.com/api/files.upload"
Global $bSlack=False
Global $sSlackErrorChannel="#general"
Global $sSlackInfoChannel="#general"


Opt("WinDetectHiddenText", 0)


Global $sDirConfigStub = "Config"
If $CmdLine[0] > 0 Then
	Global $sBaseWorkDir = $CmdLine[1]

Else
	Global $sBaseWorkDir = $sBatBane & "\.."
EndIf

If Not $sWindowsNinjaPath = "" Then $sBaseWorkDir = $sWindowsNinjaPath

If Not FileExists($sBaseWorkDir & "\" & $sDirLogsStub) Then DirCreate($sBaseWorkDir & "\" & $sDirLogsStub)
Global $sLoggfilePath = $sBaseWorkDir & "\" & $sDirLogsStub & "\" & $sFileName

Global $sPictureErrorPath = $sBaseWorkDir & "\" & $sDirPictureErrorStub
If Not FileExists($sPictureErrorPath) Then DirCreate($sPictureErrorPath)
; Flush the directory for old errors.
FileDelete($sPictureErrorPath & "\*" & $sEnvComputername & "+" & $sEnvUserName & ".jpg")


Global $sPictureHistoricPath = $sBaseWorkDir & "\" & $sDirPictureHistoricStub
Global $sPictureCurrentPath = $sBaseWorkDir & "\" & $sDirPictureCurrentStub
Global $sPictureBasePath = $sBaseWorkDir & "\" & $sDirPicturesBaseStub
If Not FileExists($sPictureHistoricPath) Then DirCreate($sPictureHistoricPath)
If Not FileExists($sPictureCurrentPath) Then DirCreate($sPictureCurrentPath)

If Not FileExists($sBaseWorkDir & "\" & $DirStatusInfoStub) Then DirCreate($sBaseWorkDir & "\" & $DirStatusInfoStub)
Global $StatusFile = $sBaseWorkDir & "\" & $DirStatusInfoStub & "\" & $sEnvComputername & "+" & $sEnvUserName & ".csv"
Global $StatusText = ""


If $CmdLine[0] = 2 Then
	$iSleepTime = $CmdLine[2]
EndIf
If Not FileExists($sBaseWorkDir & "\" & $sDirConfigStub) Then DirCreate($sBaseWorkDir & "\" & $sDirConfigStub)


Global $hTimerPicture = TimerInit()
Global $sPictureFileHistoricBackup = ""
Global $sPictureFileHistoric = ""
Global $bTerminalserverSetting = _GetTerminalServerSetting()
Global $iDelFilesOlderThanCounterLimit=20
Global $iDelFilesOlderThanCounter=$iDelFilesOlderThanCounterLimit

Eventlogg(False,4, "Starting up:" & @AutoItExe & @CRLF & "Version:" & $Version & @CRLF & "Workdir: " & $sBaseWorkDir)

;Look to delete all very old files for this computer

DelFilesOlderThan($sPictureHistoricPath, "*" & $sEnvComputername & "*.jpg", 0, 0, 31)
DelFilesOlderThan($sPictureCurrentPath, "*" & $sEnvComputername & "*.jpg", 0, 0, 3)
DelFilesOlderThan($sPictureErrorPath, "*" & $sEnvComputername & "*.jpg", 0, 0, 31)
DelFilesOlderThan($sBaseWorkDir & "\" & $DirStatusInfoStub & "\", "*" & $sEnvComputername & "*.csv", 0, 0, 31)
DelFilesOlderThan($sBaseWorkDir & "\" & $sDirLogsStub & "\", "*" & $sEnvComputername & "*.csv", 0, 0, 31)

;Prepeare Tray


Trayinit()


Do ;Main loop
	Global $aActionArray[1][7]

	;Prepear StatusFile
	$StatusText = Chr(34) & "COMPUTERNAME" & Chr(34) & ";" & Chr(34) & "USERNAME" & Chr(34) & ";" & Chr(34) & "ITEM" & Chr(34) & ";" & Chr(34) & "DATA" & Chr(34) & ";" & @CRLF
	AddStatusTextEntry("DATE", @MDAY & "." & @MON & "." & @YEAR)
	AddStatusTextEntry("TIME", @HOUR & ":" & @MIN & ":" & @SEC)
	AddStatusTextEntry("COMPUTERNAME", $sEnvComputername)
	AddStatusTextEntry("USER", $sEnvUserName)
	AddStatusTextEntry("SESSIONNAME", $sEnvSESSIONNAME)
	Local $aData = _WinAPI_GetProcessMemoryInfo()
	Global $iPeakProcessMemory = $aData[1]
	AddStatusTextEntry("NO_ACTION_FLAG", $bNoAction)
	AddStatusTextEntry("MEMORYUSAGE", $iPeakProcessMemory)
	AddStatusTextEntry("EXECUTIONPATH", @AutoItExe)
	AddStatusTextEntry("EXECUTIONVERSION", $Version)
	AddStatusTextEntry("CPUArch", @CPUArch)
	AddStatusTextEntry("KBLayout", @KBLayout)
	AddStatusTextEntry("OSVersion", @OSVersion)
	AddStatusTextEntry("OSServicePack", @OSServicePack)

	;AddStatusTextEntry("AutoItPID",@AutoItPID)
	If Not (@IPAddress1 = "127.0.0.1" Or @IPAddress1 = "::1:") Then
		AddStatusTextEntry("IPadress", @IPAddress1)
	Else
		AddStatusTextEntry("IPadress", @IPAddress2)
	EndIf


	GetParametersFromCSV($sBaseWorkDir & "\" & $sDirConfigStub & "\WindowsNinjaConfig.csv")
	GetParametersFromCSV($sBaseWorkDir & "\" & $sDirConfigStub & "\" & $sEnvComputername & "-WindowsNinjaConfig.csv")
	GetParametersFromCSV($sBaseWorkDir & "\" & $sDirConfigStub & "\" & $sEnvUserName & "-WindowsNinjaConfig.csv")

	If $bCSVConfigLoaded = False Then
		Eventlogg(True,2, "No Configuration files found in path " & $sBaseWorkDir & "\" & $sDirConfigStub & " Pause for 15 minutes...")
		Sleep(15 * 60 * 1000)
	EndIf

	; Clean log Area for old files each x run
	$iDelFilesOlderThanCounter=$iDelFilesOlderThanCounter+1
	if $iDelFilesOlderThanCounter>$iDelFilesOlderThanCounterLimit then
	  DelFilesOlderThan($sBaseWorkDir & "\" & $sDirLogsStub, "*_WindowsNinja+" & $sEnvComputername & "+" & $sEnvUserName & ".csv", 0, 0, 7)
	  $iDelFilesOlderThanCounter=0
   EndIf

	_ControlRDPStatus() ; Poll QWINSTA (RDP connect status). Check if user is connected and try to grab Console
	;_MouseMovedSensorTimerCheck() ; Decrease timer for the Mouse Timer
	_MouseMovedSensor() ; Check if mouse has moved and sets pause/unpause.
	_NoActionWhileUserConnectedSensor() ; Handles pause/unpause depending on if user is detected active or not via QWINSTA (RDP Conect status)

	; Check CPU usage
	AddStatusTextEntry("CheckCPU", $bCheckCPULimit)
	If $bCheckCPULimit Then
		$CPULoadPercent=int(_GetTotalCPULoad())
		if $CPULoadPercent > $iPeakCPULimit Then
			$bCPULoadOK=False
		Else
			$bCPULoadOK=True
		EndIf
	AddStatusTextEntry("CPULoadOk", $bCPULoadOK)
	AddStatusTextEntry("CPULoadpercent", $CPULoadPercent)
	AddStatusTextEntry("iPeakCPULimit", $iPeakCPULimit)
	EndIf



	; Calling essential routines if OK (bNoAction=False)
	if $bCPULoadOK then
		If Not $bNoAction Then
			_KeepAliveScreen()
			If $bScreensaverDisable Then _ScreenSaverActive(False)
			GetAllWindowsInfo()
			GetCurrentScreenshot(False)
			Local $NextTime= _DateAdd('s', $iSleepTime/1000, _NowCalc())
			TraySetToolTip("WindowsNinja is capturing, next schedule: " & _DateTimeFormat( $NextTime, 3 ))
		Else
			DebugLogger('+ ' & @ScriptLineNumber & " --> " & "$bNoAction set!" & @CRLF)
		EndIf
	Endif

	If $iPeakProcessMemory > $iPeakProcessMemoryLimit Then
		$bMainExist = True
		Eventlogg(True,1, "Process Memory Over limit - quitting. Currently memory usage: " & $iPeakProcessMemory & " limit " & $iPeakProcessMemoryLimit)
		AddStatusTextEntry("FATAL_MEMORY_OVER_LIMIT", $iPeakProcessMemoryLimit)
	EndIf

	Global $bFirstrun = False
	Global $hFileStatus = FileOpen($StatusFile, $FO_READ + $FO_OVERWRITE)
	FileWrite($hFileStatus, $StatusText)
	FileClose($hFileStatus)

	; Sleep logic START
	Local $RoundsSleeping=($iSleepTime)/$SleepTimeWindow
	For $iCountSleepRounds=1 to $RoundsSleeping
	if $bSleepEnable then
		Sleep($SleepTimeWindow)
		_MouseMovedSensorTimerCheck($SleepTimeWindow)
	Endif
	; Routines to poll more frequent START

	_MouseMovedSensor() ; Check if mouse has moved and sets pause/unpause.

	; Routines to poll more frequent END
	Next

	$bSleepEnable = True
	; Sleep logic STOP

Until $bMainExist
FileClose($hDebugFile)

; End


Func AddStatusTextEntry($sItem, $sData)
	DebugLogger('+ ' & @ScriptLineNumber & " --> " & "AddStatusTextEntry: " & $StatusText & @CRLF)
	$StatusText = $StatusText & Chr(34) & $sEnvComputername & Chr(34) & ";" & Chr(34) & $sEnvUserName & Chr(34) & ";" & Chr(34) & $sItem & Chr(34) & ";" & Chr(34) & $sData & Chr(34) & ";" & @CRLF
EndFunc   ;==>AddStatusTextEntry



Func GetAllWindowsInfo()


	; Prepare the loggfile


	$bFileExist = FileExists($sLoggfilePath)
	Local $hFileOpen = FileOpen($sLoggfilePath, $FO_READ + $FO_APPEND)
	If Not $bFileExist Then
		FileWrite($hFileOpen, "DATE;TIME;COMPUTERNAME;USERNAME;HANDLE;WINDOWS_TITLE;VISIBLE_TEXT;PID;PROCESSNAME;STATUS;GETWINDOWACTION" & @CRLF)
		Local $iBlistTmp = $bList[0][0]
		If $iBlistTmp < 1 Then $iBlistTmp = 1
		_ArrayColDelete($bList, 0)
		_ArrayColDelete($bList, 0)
		_ArrayColDelete($bList, 0)
		ReDim $bList[$iBlistTmp][3]
	EndIf
	Local $aWinlist = WinList()
	Local $aProcesslist = ProcessList()
	Local $alist[1][10]
	Local $ContActiveWindows = 0
	; _ArrayColInsert($aWinlist, 2)

	; Loop through the array displaying only visable windows with a title.
	If $bDebugFlagArray Then _ArrayDisplay($aWinlist, "aWinlist", Default, 8)
	If $bDebugFlagArray Then _ArrayDisplay($bList, "blist", Default, 8)
	For $i = 1 To $aWinlist[0][0] ; Loop MAIN!
		Local $sTmpErrorPictureFile = ""
		Local $sProcessName = ""
		Local $iPID = 0
		Local $sStatus = "OK_Start"
		Local $sGetWindowAction = ""
		Local $sText = ""
		If $aWinlist[$i][0] <> "" And BitAND(WinGetState($aWinlist[$i][1]), 2) And WinExists($aWinlist[$i][0]) Then
			If WinExists($aWinlist[$i][0]) Then $sText = WinGetText($aWinlist[$i][0])
			If WinExists($aWinlist[$i][0]) Then $iPID = WinGetProcess($aWinlist[$i][0])
			;DebugLogger('- ' & @ScriptLineNumber & " --> " & "PID:" & $iPID & @CRLF)
			Local $ProcessPointer = 0
			$ProcessPointer = _ArraySearch($aProcesslist, $iPID, 0, 0, 0, 0, 1, 1, False)
			If $ProcessPointer > 0 Then $sProcessName = $aProcesslist[$ProcessPointer][0]
			If $bPrintWindowsTextInLogg Then
				$stmpText = $sText
			Else
				$stmpText = ProperTXT($sText)
			EndIf
			; Update alist
			$ContActiveWindows += 1
			$alist[0][0] = $ContActiveWindows
			If $ContActiveWindows > 0 Then ReDim $alist[UBound($alist, $UBOUND_ROWS) + 1][UBound($alist, $UBOUND_COLUMNS)]
			$alist[$ContActiveWindows][0] = $aWinlist[$i][0]
			$alist[$ContActiveWindows][1] = $aWinlist[$i][1]
			$alist[$ContActiveWindows][2] = ProperTXT($sText)
			$alist[$ContActiveWindows][3] = $iPID
			$alist[$ContActiveWindows][4] = $sProcessName

			; Lets see if any match from old array
			Local $iIndexInBList = _ArraySearch($bList, $aWinlist[$i][1], 0, 0, 0, 0, 1, 1, False)
			; DebugLogger('- ' & @ScriptLineNumber & " --> " & "$iIndex:" & $aWinlist[$i][1] & " - " & $iIndexInBList & @CRLF)
			; Write to file if not done before:
			Local $bPrintActivity = False
			If $bPrintAllActivity Then $bPrintActivity = True
			If Not $bPrintAllActivity Then
				If $iIndexInBList = -1 Then
					$bPrintActivity = True
				EndIf
			EndIf


			; Now scan if the list consist of any hits from the actionArray !!!! From here.
			Local $bFoundActionArray = False
			For $Itemp3 = 0 To UBound($aActionArray, $UBOUND_ROWS) - 1
				; Check if ScreenSaver and if flagged to terminate
				If SearchTXT($sProcessName, ".scr") Then
					If $bScreensaverDisable Then
						WinKill($aWinlist[$i][1], "")
						$sStatus = "Screensaver_Killed_Process"
					EndIf
				EndIf


				;Checking WINDOWS_TITLE if text added and if there is a match
				If SearchTXT($aWinlist[$i][0], $aActionArray[$Itemp3][1]) Then
					;DebugLogger('+ ' & @ScriptLineNumber & " --> " & "Checking ActionArray WINDOWS_TITLE :" & $aActionArray[$Itemp3][1] & " DID MATCH against " & $aWinlist[$i][0] & @CRLF)
					;Checking VISIBLE_TEXT.
					If SearchTXT($sText, $aActionArray[$Itemp3][2]) Then
						DebugLogger('+ ' & @ScriptLineNumber & " --> " & "Checking ActionArray VISIBLE_TEXT :" & $aActionArray[$Itemp3][2] & " DID MATCH against " & $stmpText & @CRLF)
						If SearchTXT($sProcessName, $aActionArray[$Itemp3][3]) Then
							; Both search lines WINDOWS_TITLE + VISIBLE_TEXT + PROCESSNAMEcan not be empty
							If StringLen($aActionArray[$Itemp3][1] & $aActionArray[$Itemp3][2] & $aActionArray[$Itemp3][3]) > 0 Then
								DebugLogger('* ' & @ScriptLineNumber & " --> " & "Checking ActionArray WINDOWS_TITLE + VISIBLE_TEXT + ProcesName can not be empty OK! '" & $aActionArray[$Itemp3][0] & "' " & $ContActiveWindows & " " & @CRLF)
								$bFoundActionArray = True
								;Update alist of number of hits
								$alist[0][1] = $alist[0][1] + 1
								;Update $aActionArray rule for number of hits for that rule
								$aActionArray[$Itemp3][6] = $aActionArray[$Itemp3][6] + 1
								;Here we must check if blist has any data, if so skip.
								If $iIndexInBList > 0 Then
									If Not $bList[$iIndexInBList][5] = $aActionArray[$Itemp3][0] Then
										If Not $aActionArray[$Itemp3][5] = True Then
											;Update hitted in the table
											$alist[$ContActiveWindows][5] = $aActionArray[$Itemp3][0]
											$alist[$ContActiveWindows][6] = NoDate()
											$alist[$ContActiveWindows][7] = NoTime()
											; Logg the error
											$bPrintActivity = True
											$sStatus = "Error UnwantedWindow"
											$sGetWindowAction = $aActionArray[$Itemp3][0]
											; Throw Eventlogg Error.
											If SearchTXT($aActionArray[$Itemp3][4], "Eventlogg") Then
												Eventlogg(True,1, "Detected Window " & Chr(34) & $aWinlist[$i][0] & Chr(34) & " with process " & Chr(34) & $sProcessName & Chr(34) & " and matches rule " & Chr(34) & $sGetWindowAction & Chr(34) & " and thus is unwanted on this machine")
											EndIf

											; Take Snapshot
											If SearchTXT($aActionArray[$Itemp3][4], "Snapshot") Then
												$sTmpErrorPictureFile = $sPictureErrorPath & "\" & @YEAR & @MON & @MDAY & @HOUR & @MIN & @SEC & "_ERR_" & $aActionArray[$Itemp3][0] & "+" & $sEnvComputername & "+" & $sEnvUserName & ".jpg"
												DebugLogger('/ ' & @ScriptLineNumber & " --> " & "Capture ERROR Screen to path: '" & $sTmpErrorPictureFile & "'" & @CRLF)
												DebugLogger('/ ' & @ScriptLineNumber & " --> " & "Capture ERROR Screen WindowOnly: '" & $bErrorCaptureWindowOnly & "'" & @CRLF)
												If $bErrorCaptureWindowOnly Then
													DebugLogger('/ ' & @ScriptLineNumber & " --> " & "Setting Focus on '" & $aWinlist[$i][0] & "'" & $aWinlist[$i][1] & "'" & @CRLF)
													WinSetState($aWinlist[$i][0], "", @SW_SHOW)
													WinSetState($aWinlist[$i][0], "", @SW_RESTORE)
													;WinSetState($aWinlist[$i][0], "", @SW_MAXIMIZE)
													WinSetState($aWinlist[$i][0], "", @SW_ENABLE)
													_WinAPI_BringWindowToTop($aWinlist[$i][0])
													WinActive($aWinlist[$i][0])
													WinSetOnTop($aWinlist[$i][0], "", $WINDOWS_ONTOP)
													_ScreenCapture_CaptureWnd($sTmpErrorPictureFile, $aWinlist[$i][0])
													WinSetOnTop($aWinlist[$i][0], "", $WINDOWS_NOONTOP)
												Else
													_ScreenCapture_Capture($sTmpErrorPictureFile)
												EndIf
												$alist[$ContActiveWindows][8] = $sTmpErrorPictureFile
											EndIf
										;Send Mail
											if $bMailEnable then
												If SearchTXT($aActionArray[$Itemp3][4], "Mail") Then
													BlatMailSend($BlatSmtpServer,@ComputerName & $BlatSenderSuffix, $sTmpErrorPictureFile ,$BlatTo,"WindowsNinja Detected an error" , "Detected Window " & Chr(34) & $aWinlist[$i][0] & Chr(34) & " with process " & Chr(34) & $sProcessName & Chr(34) & " and matches rule " & Chr(34) & $sGetWindowAction & Chr(34) & " and thus is unwanted on this machine" ,"1")
												Endif
											Endif
										; Handle Slack integration
											if $bSlack Then
												If SearchTXT($aActionArray[$Itemp3][4], "Slack") Then
													SlackPostAttachementCurl($sSlackErrorChannel,"Error detected on " & @ComputerName,$sTmpErrorPictureFile,$sSlackToken,$sTmpErrorPictureFile, ProperTXTCMD("Detected Window '" & $aWinlist[$i][0] & "' with process '" & $sProcessName & "' and matches rule '" & $sGetWindowAction & "' and thus is unwanted on this machine"), $sSlackUploadURL)
												Endif
											Endif
										EndIf

									Else
										;Insert copy from blist here

										$alist[$ContActiveWindows][5] = $aActionArray[$Itemp3][0]
										$alist[$ContActiveWindows][6] = $bList[$iIndexInBList][6]
										$alist[$ContActiveWindows][7] = $bList[$iIndexInBList][7]
										$alist[$ContActiveWindows][8] = $bList[$iIndexInBList][8]
									EndIf
								EndIf
							EndIf
						EndIf
					EndIf
				EndIf
			Next
			If $bPrintActivity Then FileWrite($hFileOpen, Chr(34) & @MDAY & "." & @MON & "." & @YEAR & Chr(34) & ";" & Chr(34) & @HOUR & ":" & @MIN & ":" & @SEC & Chr(34) & ";" & Chr(34) & $sEnvComputername & Chr(34) & ";" & Chr(34) & $sEnvUserName & Chr(34) & ";" & Chr(34) & $aWinlist[$i][1] & Chr(34) & ";" & Chr(34) & $aWinlist[$i][0] & Chr(34) & ";" & Chr(34) & $stmpText & Chr(34) & ";" & Chr(34) & $iPID & Chr(34) & ";" & Chr(34) & $sProcessName & Chr(34) & ";" & Chr(34) & $sStatus & Chr(34) & ";" & Chr(34) & $sGetWindowAction & Chr(34) & @CRLF)
		EndIf
	Next
	If $bDebugFlagArray Then _ArrayDisplay($alist, "alist", Default, 8)
	If $bDebugFlagArray Then _ArrayDisplay($aActionArray, "$aActionArray", Default, 8)
	;Handle Blist cleanups ----------------------------------------------------------------------------------------------------------------
	;OK Lets see if any errorpictures from Blist should have been cleanup up (not existing anymore)
	If UBound($bList, $UBOUND_ROWS) > 0 Then
		If $bList[0][1] > 0 Then ; Any Activity logged?
			DebugLogger('- ' & @ScriptLineNumber & " --> " & "Starting checkup of Blist:" & $bList[0][0] & @CRLF)
			$bPrintActivity = False
			For $itmp = 1 To $bList[0][0]
				DebugLogger('- ' & @ScriptLineNumber & " --> " & $itmp & ":: Blist check: " & $bList[$itmp][8] & " - Stringlen:" & StringLen($bList[$itmp][8]) & " AlistLookup: " & _ArraySearch($alist, $bList[$itmp][8], 0, 0, 0, 0, 1, 8, False) & @CRLF)
				If StringLen($bList[$itmp][8]) > 0 Then
					If _ArraySearch($alist, $bList[$itmp][8], 0, 0, 0, 0, 1, 8, False) = -1 Then
						$sStatus = "OK_PictureDeleted"
						$bPrintActivity = True
						If $bPrintActivity Then FileWrite($hFileOpen, Chr(34) & @MDAY & "." & @MON & "." & @YEAR & Chr(34) & ";" & Chr(34) & @HOUR & ":" & @MIN & ":" & @SEC & Chr(34) & ";" & Chr(34) & $sEnvComputername & Chr(34) & ";" & Chr(34) & $sEnvUserName & Chr(34) & ";" & Chr(34) & $bList[$itmp][1] & Chr(34) & ";" & Chr(34) & $bList[$itmp][0] & Chr(34) & ";" & Chr(34) & $bList[$itmp][2] & Chr(34) & ";" & Chr(34) & $bList[$itmp][3] & Chr(34) & ";" & Chr(34) & $bList[$itmp][4] & Chr(34) & ";" & Chr(34) & $sStatus & Chr(34) & ";" & Chr(34) & $bList[$itmp][5] & Chr(34) & @CRLF)
						If FileExists($bList[$itmp][8]) Then
							DebugLogger('* ' & @ScriptLineNumber & " --> " & "Delete old ErrorPicture" & $bList[$itmp][8] & @CRLF)
							FileDelete($bList[$itmp][8])
						EndIf
					Else
						AddStatusTextEntry("ERRORPICTUREFILE", $bList[$itmp][8])
					EndIf
				EndIf
			Next
		EndIf
		;Peek at ActionArray to see if there is any includefilters to ditch
		For $Itemp4 = 0 To UBound($aActionArray, $UBOUND_ROWS) - 1
			If $aActionArray[$Itemp4][5] = True Then
				If $aActionArray[$Itemp4][6] < 1 Then
					If $aActionArrayMissingWindowList[$Itemp4] = "" Then ;Checking if earlier notified
						$aActionArrayMissingWindowList[$Itemp4] = "DETECTED"
						;OK, includefilter but no hits> Must be error.
						AddStatusTextEntry("ERRORMISSINGWINDOWRULE", $aActionArray[$Itemp4][0])
						$sStatus = "Error_MissingWindow"
						$bPrintActivity = True
						If $bPrintActivity Then FileWrite($hFileOpen, Chr(34) & @MDAY & "." & @MON & "." & @YEAR & Chr(34) & ";" & Chr(34) & @HOUR & ":" & @MIN & ":" & @SEC & Chr(34) & ";" & Chr(34) & $sEnvComputername & Chr(34) & ";" & Chr(34) & $sEnvUserName & Chr(34) & ";" & Chr(34) & "" & Chr(34) & ";" & Chr(34) & $aActionArray[$Itemp4][1] & Chr(34) & ";" & Chr(34) & $aActionArray[$Itemp4][2] & Chr(34) & ";" & Chr(34) & "" & Chr(34) & ";" & Chr(34) & $aActionArray[$Itemp4][3] & Chr(34) & ";" & Chr(34) & $sStatus & Chr(34) & ";" & Chr(34) & $aActionArray[$Itemp4][0] & Chr(34) & @CRLF)
						DebugLogger('* ' & @ScriptLineNumber & " --> " & "Missing process in Includefilter " & Chr(34) & @MDAY & "." & @MON & "." & @YEAR & Chr(34) & ";" & Chr(34) & @HOUR & ":" & @MIN & ":" & @SEC & Chr(34) & ";" & Chr(34) & $sEnvComputername & Chr(34) & ";" & Chr(34) & $sEnvUserName & Chr(34) & ";" & Chr(34) & "" & Chr(34) & ";" & Chr(34) & $aActionArray[$Itemp4][1] & Chr(34) & ";" & Chr(34) & $aActionArray[$Itemp4][2] & Chr(34) & ";" & Chr(34) & "" & Chr(34) & ";" & Chr(34) & $aActionArray[$Itemp4][3] & Chr(34) & ";" & Chr(34) & $sStatus & Chr(34) & ";" & Chr(34) & $aActionArray[$Itemp4][0] & Chr(34) & @CRLF)
						If SearchTXT($aActionArray[$Itemp4][4], "Eventlogg") Then
							Eventlogg(True,1, "Error MissingWindow from rule " & Chr(34) & $aActionArray[$Itemp4][0] & Chr(34) & " which is required on this machine")
						EndIf
						if $bMailEnable then
							If SearchTXT($aActionArray[$Itemp4][4], "Mail") Then
								BlatMailSend($BlatSmtpServer,@ComputerName & $BlatSenderSuffix, "" ,$BlatTo,"WindowsNinja Detected an error" ,"Missing process in Includefilter " & Chr(34) & $aActionArray[$Itemp4][1] & Chr(34) ,"1")
							Endif
						Endif

						; Insert Slack here
						If SearchTXT($aActionArray[$Itemp4][4], "Slack") Then
						   SlackPostAttachementCurl($sSlackErrorChannel,"Error detected on " & @ComputerName,$sErrorPicture,$sSlackToken,$sErrorPicture, ProperTXTCMD("Missing process in Includefilter " & Chr(34) & $aActionArray[$Itemp4][1] & Chr(34)), $sSlackUploadURL)
						Endif
					Else
						DebugLogger('* ' & @ScriptLineNumber & " --> " & "Missing process in Includefilter (Already Notfied) " & Chr(34) & @MDAY & "." & @MON & "." & @YEAR & Chr(34) & ";" & Chr(34) & @HOUR & ":" & @MIN & ":" & @SEC & Chr(34) & ";" & Chr(34) & $sEnvComputername & Chr(34) & ";" & Chr(34) & $sEnvUserName & Chr(34) & ";" & Chr(34) & "" & Chr(34) & ";" & Chr(34) & $aActionArray[$Itemp4][1] & Chr(34) & ";" & Chr(34) & $aActionArray[$Itemp4][2] & Chr(34) & ";" & Chr(34) & "" & Chr(34) & ";" & Chr(34) & $aActionArray[$Itemp4][3] & Chr(34) & ";" & Chr(34) & $sStatus & Chr(34) & ";" & Chr(34) & $aActionArray[$Itemp4][0] & Chr(34) & @CRLF)
					EndIf
				Else
					$aActionArrayMissingWindowList[$Itemp4] = ""
				EndIf
			Else
				$aActionArrayMissingWindowList[$Itemp4] = ""
			EndIf
		Next
		If $bDebugFlagArray Then _ArrayDisplay($aActionArrayMissingWindowList, "$aActionArrayMissingWindowList", Default, 8)
	EndIf

	FileClose($hFileOpen)

	ReDim $bList[$alist[0][0]][3]
	; Roll the list to keep history.
	$bList = $alist
EndFunc   ;==>GetAllWindowsInfo

Func GetParametersFromCSV($sLoggfilePathCSV)
	Local $aRetArray[1000][4]
	Local $iActionArrayCount = 0
	DebugLogger('- ' & @ScriptLineNumber & " --> " & "$sLoggfilePathCSV:" & $sLoggfilePathCSV & @CRLF)
	If FileExists($sLoggfilePathCSV) Then
		$bCSVConfigLoaded = True
		AddStatusTextEntry("CSVCONFIG", $sLoggfilePathCSV)
		_FileReadToArray(GetTmpFile($sLoggfilePathCSV), $aRetArray, $FRTA_NOCOUNT, ";")
		If $bDebugFlagArray Then _ArrayDisplay($aRetArray, "CSV Parameters $aRetArray", Default, 8)
		Local $iRows = UBound($aRetArray, $UBOUND_ROWS)
		Local $iCols = UBound($aRetArray, $UBOUND_COLUMNS)
		DebugLogger('- ' & @ScriptLineNumber & " --> " & "Rows:" & $iRows & ",Cols:" & $iCols & @CRLF)
		For $t1 = 1 To $iRows - 1
			Local $IncludeItem = False
			If SearchTXT($sEnvComputername, $aRetArray[$t1][0]) Then
				If SearchTXT($sEnvUserName, $aRetArray[$t1][1]) Then $IncludeItem = True
			EndIf
			;DebugLogger('- ' & @ScriptLineNumber & " --> " & "Array:" & $aRetArray[$t1][0] & @CRLF)
			If $IncludeItem Then
				Select
					Case $aRetArray[$t1][2] = "Sleeptime"
						$iSleepTime = ($aRetArray[$t1][3])*1000
						DebugLogger('- ' & @ScriptLineNumber & " --> " & "$iSleeptime:" & $iSleepTime & @CRLF)
					Case $aRetArray[$t1][2] = "PrintWindowsTextInLogg"
						If $aRetArray[$t1][3] = "TRUE" Then $bPrintWindowsTextInLogg = True
					Case $aRetArray[$t1][2] = "PrintAllActivity"
						If $aRetArray[$t1][3] = "TRUE" Then $bPrintAllActivity = True
					Case $aRetArray[$t1][2] = "GetWindowAction"
						If $iActionArrayCount > 0 Then
							ReDim $aActionArray[UBound($aActionArray, $UBOUND_ROWS) + 1][UBound($aActionArray, $UBOUND_COLUMNS)]
						EndIf
						For $itemp2 = 3 To 8
							$aActionArray[$iActionArrayCount][$itemp2 - 3] = $aRetArray[$t1][$itemp2]
						Next
						$iActionArrayCount += 1
					Case $aRetArray[$t1][2] = "StatusCaptureTimer"
						If $iPicturSleepTime = 0 Then $hTimerPicture = TimerInit()
						$iPicturSleepTime = $aRetArray[$t1][3]*1000
						DebugLogger('- ' & @ScriptLineNumber & " --> " & "$iPicturSleepTime:" & $iPicturSleepTime & @CRLF)
					Case $aRetArray[$t1][2] = "ErrorCaptureWindowOnly"
						If $aRetArray[$t1][3] = "TRUE" Then $bErrorCaptureWindowOnly = True

					Case $aRetArray[$t1][2] = "NoAction"
						If $aRetArray[$t1][3] = "TRUE" Then $bNoAction = True

					Case $aRetArray[$t1][2] = "PeakProcessMemoryLimit"
						$iPeakProcessMemoryLimit = $aRetArray[$t1][3]

					Case $aRetArray[$t1][2] = "PeakCPULimit"
						$iPeakCPULimit = $aRetArray[$t1][3]
						$bCheckCPULimit = True

					Case $aRetArray[$t1][2] = "EnforceBringToConsole"
						If $aRetArray[$t1][3] = "true" Then $bEnforceBringToConsole = True

					Case $aRetArray[$t1][2] = "ListQwinsta"
						If $aRetArray[$t1][3] = "true" Then $bListQwinstaInCurrentLogg = True

					Case $aRetArray[$t1][2] = "DisableScreenSaver"
						If $aRetArray[$t1][3] = "true" Then $bScreensaverDisable = True

					Case $aRetArray[$t1][2] = "NoActionWhileUserConnected"
						If $aRetArray[$t1][3] = "true" Then $bNoActionWhileUserConnected = True

					Case $aRetArray[$t1][2] = "MoseMovedDetectPauseTime"
						$iMouseMovedDetectPauseTime=($aRetArray[$t1][3])*1000
						if $iMouseMovedDetectPauseTime="" then $iMouseMovedDetectPauseTime=0

					Case $aRetArray[$t1][2] = "SmtpServer"
						$BlatSmtpServer = $aRetArray[$t1][3]
						Global $bMailEnable=True

					Case $aRetArray[$t1][2] = "smtpMailSuffix"
					$BlatSenderSuffix = $aRetArray[$t1][3]
					if $BlatSenderSuffix="" then $BlatSenderSuffix="_noreply.WindowsNinja@" & EnvGet("USERDNSDOMAIN")

					Case $aRetArray[$t1][2] = "smtpMailTo"
					$BlatTo = $aRetArray[$t1][3]

					Case $aRetArray[$t1][2] = "SlackToken"
					$sSlackToken = $aRetArray[$t1][3]
					$bSlack=True

				    Case $aRetArray[$t1][2] = "SlackInfoChannel"
					$sSlackInfoChannel = $aRetArray[$t1][3]

					Case $aRetArray[$t1][2] = "SlackErrorChannel"
					$sSlackErrorChannel = $aRetArray[$t1][3]


				EndSelect
			EndIf
		Next
	EndIf
EndFunc   ;==>GetParametersFromCSV


Func SearchTXT($sTextToSearchIn, $sString)

	; Version 1.0: Match text (ignorecase) and blanks also
	; Version 1.1: Rebuld to StringRegExp
    ; Version 1.2: Adding permanent Case insensitive
	; Example: (?i)(?=Test)(^((?!NotThis)[\s\S])*$)

	; (?i) = Ignore Case
	; (?=Test) = Include if "test" but not if:
	; (^((?!NotThis)[\s\S])*$) = (not) Including "NotThis"


	Local $bReturn = False
	DebugLogger('+ ' & @ScriptLineNumber & " --> " & "StringRegExp '" & $sTextToSearchIn & "' '" & $sString & "'" & StringRegExp($sTextToSearchIn, $sString)& @CRLF)
	DebugLogger('+ ' & @ScriptLineNumber & " --> " & "StringInStr  '" & $sTextToSearchIn & "' '" & $sString & "'" & StringInStr($sTextToSearchIn, $sString) & @CRLF)
	If StringRegExp($sTextToSearchIn, "(?i)" & $sString)>0 Then
		$bReturn = True
	EndIf
	If $sString = "" Then $bReturn = True
	Return $bReturn
EndFunc   ;==>SearchTXT

Func ProperTXT($sTextToReplace)
	;Chops text down to 128 characters and drops special characters
	Local $sReturn = ""
	; Local $iCount = 0
	Local $iMax = 180
	$sReturn = $sTextToReplace
	$sReturn = StringLeft($sReturn, $iMax)
	$sReturn = StringReplace($sReturn, Chr(34), "'", 0, 0)
	$sReturn = StringReplace($sReturn, Chr(10), "|", 0, 0)
	$sReturn = StringReplace($sReturn, Chr(13), " ", 0, 0)
	$sReturn = StringLeft($sReturn, $iMax)

	;DebugLogger('+ ' & @ScriptLineNumber & " --> " & "ProperTXT '" & $sReturn & "'" & @CRLF)

	Return $sReturn
EndFunc   ;==>ProperTXT

Func ProperTXTCMD($sTextToReplace)
	;Chops text down to 128 characters and drops special characters
	Local $sReturn = ""
	; Local $iCount = 0
	Local $iMax = 300
	$sReturn = $sTextToReplace
	$sReturn = StringLeft($sReturn, $iMax)
	$sReturn = StringReplace($sReturn, Chr(34), "'", 0, 0)
	$sReturn = StringReplace($sReturn, Chr(10), "*", 0, 0)
	$sReturn = StringReplace($sReturn, Chr(13), " ", 0, 0)
	$sReturn = StringReplace($sReturn, "|", "*", 0, 0)

	;DebugLogger('+ ' & @ScriptLineNumber & " --> " & "ProperTXT '" & $sReturn & "'" & @CRLF)

	Return $sReturn
EndFunc   ;==>ProperTXT



Func NoDate()
	Return @MDAY & "." & @MON & "." & @YEAR
EndFunc   ;==>NoDate

Func NoTime()
	Return @HOUR & ":" & @MIN & ":" & @SEC
EndFunc   ;==>NoTime



Func DelFilesOlderThan($Basedir, $Filematch, $MinutesOld, $HoursOld, $DaysOld)

	Local $Counttmp1
	Local $DateDifftmp
	Local $bFileDelete
	Local $dTimeStamp
	Local $FoundFiles = True
	DebugLogger('/ ' & @ScriptLineNumber & " --> " & "Searching in path " & $Basedir & " for files " & $Filematch & @CRLF)
	Local $aFileList = _FileListToArray($Basedir, $Filematch)
	If @error = 1 Then
		$FoundFiles = False
		DebugLogger('/ ' & @ScriptLineNumber & " --> " & "Invalid Path " & $Basedir & @CRLF)
	EndIf
	If @error = 4 Then
		$FoundFiles = False
		DebugLogger('/ ' & @ScriptLineNumber & " --> " & "No files found in " & $Basedir & @CRLF)
	EndIf
	; Display the results returned by _FileListToArray.
	If Not UBound($aFileList, $UBOUND_ROWS) > 0 Then $FoundFiles = False
	If $bDebugFlagArray Then _ArrayDisplay($aFileList, "$aFileList")
	If $FoundFiles Then
		For $Conttmp1 = 1 To $aFileList[0]
			; Display the modified timestamp of the file and return as a string in the format YYYYMMDDHHMMSS
			Local $stmpFileTarget = $Basedir & "\" & $aFileList[$Conttmp1]
			Local $aTimeStamp = FileGetTime($stmpFileTarget, $FT_MODIFIED, 0)
			Local $dTimeStamp = $aTimeStamp[0] & "/" & $aTimeStamp[1] & "/" & $aTimeStamp[2] & " " & $aTimeStamp[3] & ":" & $aTimeStamp[4] & ":" & $aTimeStamp[5]
			$DateDifftmp = _NowCalc()
			;DebugLogger('- ' & @ScriptLineNumber & " --> " & "_NowCalc:1 " & $DateDifftmp  & @CRLF)
			$DateDifftmp = _DateAdd('d', -$DaysOld, $DateDifftmp)
			;DebugLogger('- ' & @ScriptLineNumber & " --> " & "_NowCalc:2 " & $DateDifftmp  & @CRLF)
			$DateDifftmp = _DateAdd('h', -$HoursOld, $DateDifftmp)
			;DebugLogger('- ' & @ScriptLineNumber & " --> " & "_NowCalc:3 " & $DateDifftmp  & @CRLF)
			$DateDifftmp = _DateAdd('n', -$MinutesOld, $DateDifftmp)
			;DebugLogger('- ' & @ScriptLineNumber & " --> " & "_NowCalc:4 " & $DateDifftmp  & @CRLF)
			;DebugLogger('- ' & @ScriptLineNumber & " --> " & "_DateDiff " & _Datediff("n",$DateDifftmp,$dTimeStamp)  & @CRLF)
			Local $dTimediff = _DateDiff("n", $dTimeStamp, $DateDifftmp)
			If $dTimediff > 0 Then
				$bFileDelete = True
			Else
				$bFileDelete = False
			EndIf
			DebugLogger('- ' & @ScriptLineNumber & " --> " & "Filepath:" & $stmpFileTarget & " Datemodify:" & $dTimeStamp & " $DateDifftmp: " & $DateDifftmp & " To be deleted : " & $bFileDelete & " _Datediff(minutes):" & $dTimediff & @CRLF)
			If FileExists($stmpFileTarget) Then
				If $bFileDelete Then
					FileDelete($stmpFileTarget)
				EndIf
			EndIf
		Next
	EndIf
EndFunc   ;==>DelFilesOlderThan

Func _ControlRDPStatus()
	Local $aWTSSessions = _Qwinsta() ; Get the Qwinsta data
	;Evaluation and rules to keep alive screen:


	;Typical Qwinsta Output:
	; 1: SESSIONNAME       USERNAME                 ID  STATE   TYPE        DEVICE
	; 2: services                                    0  Disc
	; 3:>console           tonny.r.holm              1  Active
	; 4: rdp-tcp                                 65536  Listen
	Local $bCurrentConsole = False
	Local $bCurrentActive = False
	Local $iCurrentUserRow = 0
	Local $iCurrentUserID = 0
	Local $bConsoleIsActive = False
	Local $iConsoleRow = 0
	Local $sConsoleUserActive = ""
	$iCurrentUserRow = _ArraySearch($aWTSSessions, ">", 0, 0, 0, 0, 1, 0, False)
	If $iCurrentUserRow > 0 Then ;OK, if current session cant be found something is really wrong.
		If $aWTSSessions[$iCurrentUserRow][1] = "console" Then $bCurrentConsole = True
		if not $bCurrentConsole then Global $bTsconToConsoleInitiated=false
		DebugLogger('- ' & @ScriptLineNumber & " --> $aWTSSessions[$iCurrentUserRow][1]:" & $aWTSSessions[$iCurrentUserRow][1] & @CRLF)
		If $aWTSSessions[$iCurrentUserRow][4] = "Active" Then
			$bCurrentActive = True
			$bUserIsConnected = True
		Else
			$bUserIsConnected = False
		Endif
		DebugLogger('- ' & @ScriptLineNumber & " --> $bUserIsConnected:" & $bUserIsConnected & @CRLF)
		$iCurrentUserID = $aWTSSessions[$iCurrentUserRow][3]
		;Check if Console is active
		;AddStatusTextEntry("CurrentUserID",$iCurrentUserID)
		;AddStatusTextEntry("CurrentActive",$bCurrentActive)
		If $bEnforceBringToConsole Then _TsconToConsole($iCurrentUserID)
		If Not $bCurrentActive Then
			If Not $bCurrentConsole Then
				$iConsoleRow = _ArraySearch($aWTSSessions, "console", 0, 0, 0, 0, 1, 1, False)
				If $iConsoleRow > 0 Then
					$sConsoleUserActive = $aWTSSessions[$iConsoleRow][2]
					If $aWTSSessions[$iConsoleRow][4] = "Active" Then
						$bConsoleIsActive = True
						Global $bMouseMovedArmed=True
					EndIf
				EndIf
				If Not $bConsoleIsActive Then
					_TsconToConsole($iCurrentUserID)
					Global $bTsconToConsoleInitiated=true
					Global $bMouseMovedArmed=True
					;Eventlogg(False,4, "Routine grabs the console from user: " & $sConsoleUserActive & @CRLF & " $bCurrentConsole:" & $bCurrentConsole & " $bCurrentActive:" & $bCurrentActive & " $iCurrentUserRow:" & $iCurrentUserRow & " $iCurrentUserID:'" & $iCurrentUserID & "'" & " $bConsoleIsActive:'" & $bConsoleIsActive & "'")
				Else
					Global $bTsconToConsoleInitiated=False
					Eventlogg(True,2, "Routine tried to grab the console, but was logged on with user: " & $sConsoleUserActive & @CRLF & " $bCurrentConsole:" & $bCurrentConsole & " $bCurrentActive:" & $bCurrentActive & " $iCurrentUserRow:" & $iCurrentUserRow & " $iCurrentUserID:'" & $iCurrentUserID & "'" & " $bConsoleIsActive:'" & $bConsoleIsActive & "'")
				EndIf
			EndIf
		EndIf

	EndIf
	DebugLogger('- ' & @ScriptLineNumber & " --> $bCurrentConsole:" & $bCurrentConsole & " $bCurrentActive:" & $bCurrentActive & " $iCurrentUserRow:" & $iCurrentUserRow & " $iCurrentUserID:'" & $iCurrentUserID & "'" & " $bConsoleIsActive:'" & $bConsoleIsActive & "'" & @CRLF)


EndFunc   ;==>_ControlRDPStatus


Func GetCurrentScreenshot($Enforce)


	If ($iPicturSleepTime > 0) Or $bFirstrun or $Enforce Then
		If (TimerDiff($hTimerPicture) > ($iPicturSleepTime)) Or $bFirstrun or $Enforce Then


			Global $hTimerPicture = TimerInit()
			; Global $sPictureFile = $sPictureCurrentPath & "\" & $sEnvComputername & "+" & $sEnvUserName & "_" & @YEAR & @MON & @MDAY & @HOUR & @MIN & @SEC & ".jpg"
			Global $sPictureFile = $sPictureCurrentPath & "\" & $sEnvComputername & "+" & $sEnvUserName & ".jpg"
			Global $sPictureFileHistoric = $sPictureHistoricPath & "\HISTORIC-" & @HOUR & "+" & $sEnvComputername & "+" & $sEnvUserName & ".jpg"
			DelFilesOlderThan($sPictureCurrentPath, "*" & $sEnvComputername & "+" & $sEnvUserName & "*.jpg", 0, 1, 0)

			DebugLogger('/ ' & @ScriptLineNumber & " --> " & "Capture INFO Screen to path: '" & $sPictureFile & "'" & @CRLF)
			_ScreenCapture_Capture($sPictureFile)
			$sPictureFileLastWritten = $sPictureFile
			DebugLogger('/ ' & @ScriptLineNumber & " --> " & "Picture Historic: " & @CRLF & "'" & $sPictureFileHistoricBackup & "'" & @CRLF & "'" & $sPictureFileHistoric & "'" & @CRLF)
			If Not SearchTXT($sPictureFileHistoricBackup, $sPictureFileHistoric) Then
				;if $bDebugFlag then

				If FileExists($sPictureFileHistoric) Then
					FileDelete($sPictureFileHistoric)
					DebugLogger('/ ' & @ScriptLineNumber & " --> " & "Deltefile: '" & $sPictureFileHistoric & "'" & @CRLF)
				EndIf
				FileCopy($sPictureFile, $sPictureFileHistoric, $FC_OVERWRITE + $FC_CREATEPATH)
				Global $sPictureFileHistoricBackup = $sPictureFileHistoric

			EndIf
		EndIf
	EndIf

		AddStatusTextEntry("PICTUREFILE", $sPictureFileLastWritten)
		AddStatusTextEntry("PICTUREFILEHISTORICLAST", $sPictureFileHistoricBackup)
		AddStatusTextEntry("PICTUREFILEHISTORIC", $sPictureFileHistoric)


EndFunc   ;==>GetCurrentScreenshot



; Func EventloggOLD($bShowTray,$iErrortype, $sErrorText)
	; Version 1.2
	;Event type. This can be one of the following values:
	;    0 - Success event
	;    1 - Error event
	;    2 - Warning event
	;    4 - Information event
	;    8 - Success audit event
	;    16 - Failue audit event
	; https://support.microsoft.com/en-us/kb/131008
;	if $TrayTipHeader="" then $TrayTipHeader=@ScriptName
;	Local $hEventLog, $aData[4] = [3, 1, 2, 3]
;	$hEventLog = _EventLog__Open("", @ScriptName)
;	_EventLog__Report($hEventLog, $iErrortype, 0, 2, EnvGet("USERNAME"), _Now() & " : " & @ScriptName & " : (" & $Version & ") : " & EnvGet("COMPUTERNAME") & " : " & @CRLF & $sErrorText, $aData)
;	DebugLogger('- ' & @ScriptLineNumber & " --> " & "Sending to Eventlogg: " & $sErrorText & @CRLF)
;	_EventLog__Close($hEventLog)

; 	local $iTrayTipCode=0
;	if $iErrortype=1 then $iTrayTipCode=3
;	if $iErrortype=2 then $iTrayTipCode=2
;	if $iErrortype=16 then $iTrayTipCode=3

;	if $bShowTray then TrayTip($TrayTipHeader, $sErrorText, 3, $iTrayTipCode)
; EndFunc   ;==>Eventlogg





Func Eventlogg($bShowTray,$iErrortype, $sErrorText)
	; Version 1.3
	; Using Batch instead to throw via VBSCRIPT.
	;Event type. This can be one of the following values:
	;    0 - Success event
	;    1 - Error event
	;    2 - Warning event
	;    4 - Information event
	;    8 - Success audit event
	;    16 - Failue audit event
	; https://support.microsoft.com/en-us/kb/131008
	if $TrayTipHeader="" then $TrayTipHeader=@ScriptName

	Local $tmpvbs = @TempDir & "\" & @ScriptName & "_eventhandler.vbs"
	Local $EventMsg= @ScriptName & " : (" & $Version & ") : " & ProperTXTCMD($sErrorText)
	Local $EventVBS="echo set objShell = CreateObject(" & Chr(34) & "WScript.Shell" & Chr(34) & ") : objShell.LogEvent " & $iErrortype & "," & Chr(34) & $EventMsg & Chr(34)
    DebugLogger('/ ' & @ScriptLineNumber & " --> " & "Eventlogg message: '" & $EventMsg &"'" & @CRLF)
	DebugLogger('/ ' & @ScriptLineNumber & " --> " & "Eventlogg Command: '" & $EventVBS &"'" & @CRLF)
	CmdExecSTDOUT($EventVBS & " >" & $tmpvbs)
	CmdExecSTDOUT("cscript /nologo " & Chr(34) & $tmpvbs & Chr(34))

	; $iErrortype, 0, 2, EnvGet("USERNAME"), _Now() & " : " & @ScriptName & " : (" & $Version & ") : " & EnvGet("COMPUTERNAME") & " : " & @CRLF & $sErrorText, $aData)


	DebugLogger('- ' & @ScriptLineNumber & " --> " & "Sending to Eventlogg: " & $sErrorText & @CRLF)


    ; $TIP_ICONNONE (0) = No icon (default)
    ; $TIP_ICONASTERISK (1) = Info icon
    ; $TIP_ICONEXCLAMATION (2) = Warning icon
    ; $TIP_ICONHAND (3) = Error icon
    ; $TIP_NOSOUND (16) = Disable sound
	local $iTrayTipCode=0
	if $iErrortype=1 then $iTrayTipCode=3
	if $iErrortype=2 then $iTrayTipCode=2
	if $iErrortype=16 then $iTrayTipCode=3

	if $bShowTray then TrayTip($TrayTipHeader, $sErrorText, 3, $iTrayTipCode)
EndFunc   ;==>Eventlogg

Func GetTmpFile($FileNameTmp)

	Local $aFilePart = StringSplit($FileNameTmp, "\")
	Local $sFilePart = $aFilePart[$aFilePart[0]]
	Local $tmpDir = @TempDir & "\" & @ScriptName
	Local $Destfile = $tmpDir & "\" & $sFilePart

	CmdExec("if not exist " & Chr(34) & $tmpDir & Chr(34) & " mkdir " & Chr(34) & $tmpDir & Chr(34))
	CmdExec("xcopy /Y /H /D " & Chr(34) & $FileNameTmp & Chr(34) & " " & Chr(34) & $tmpDir & Chr(34))
	DebugLogger('- ' & @ScriptLineNumber & " NewLocalCopy:" & $Destfile & @CRLF)
	Return ($Destfile)
EndFunc   ;==>GetTmpFile

Func CmdExec($Command)
	;Replaces ' to "
	$Command = StringReplace($Command, "'", Chr(34))
	DebugLogger('- ' & @ScriptLineNumber & " RUN " & @ComSpec & " /c " & $Command & @CRLF)
	Local $iReturn = RunWait(@ComSpec & " /c " & $Command, "",  _WinAPI_ExpandEnvironmentStrings("%windir%"), @SW_HIDE)
	DebugLogger('- ' & @ScriptLineNumber & " --> " & "Command: " & $Command & @CRLF)
	DebugLogger('- ' & @ScriptLineNumber & " --> " & "Returned: " & $iReturn & @CRLF)
	Return ($iReturn)
EndFunc   ;==>CmdExec

Func SlackPostAttachementCurl($sChannel,$sTitle,$sFileName,$sToken,$sFilepathToPost, $sInitialComment, $sSlackUploadURL)
	; Get Token for Slack here: https://api.slack.com/web
	; Expected return is "ok":true
	Local $sOKString=Chr(34) & "ok" & Chr(34) & ":true"
	Local $sSlackCommand = Chr(34) & $sCurlExe & Chr(34) & " -k -i -X POST -H " & Chr(34) & "Content-Type: multipart/form-data" & Chr(34) & " -F " & Chr(34) & "channels=" & $sChannel  & Chr(34) & " -F " & Chr(34) & "title=" & ProperTXTCMD($sTitle) & Chr(34) & " -F " & Chr(34) & "filename=" & $sFileName & Chr(34) & " -F " & Chr(34) & "token=" & $sToken & Chr(34) & " -F " & Chr(34) & "file=@" & $sFilepathToPost  & Chr(34) & " -F " & Chr(34) & "initial_comment=" & ProperTXTCMD($sInitialComment) & Chr(34) & " " &  $sSlackUploadURL
	; $CurlCall = CmdExecSTDOUT("start /w " & Chr(34) & "SlackPostAttachementCurl" & Chr(34) & " /MIN " & $sSlackCommand)
    DebugLogger('- ' & @ScriptLineNumber & " SlackPostAttachementCurl " &$sSlackCommand & @CRLF)

	Local $iPID = Run($sSlackCommand, _WinAPI_ExpandEnvironmentStrings("%windir%"), @SW_HIDE, $STDERR_MERGED)
	ProcessWaitClose($iPID)
	$CurlCall = StdoutRead($iPID)
   DebugLogger('- ' & @ScriptLineNumber & " SlackPostAttachementCurl:Result: " & $CurlCall &  @CRLF)
	if StringInStr($CurlCall,$sOKString) Then
		Eventlogg(False,4, "Posted Slack Message OK: " & $sTitle & " - " &  $sInitialComment & ", File: " & $sFilepathToPost )
	Else
		Eventlogg(False,1, "Posted Slack Message FAILED " & $sTitle & " - " &  $sInitialComment & ", File: " & $sFilepathToPost & ". Errormessage: '" & $CurlCall & "'")
	EndIf
	Return $CurlCall
EndFunc




Func CmdExecSTDOUT($Command)
	DebugLogger('- ' & @ScriptLineNumber & " RUN " & @ComSpec & " /c " & $Command & @CRLF)
	Local $iPID = Run(@ComSpec & " /c " & $Command, _WinAPI_ExpandEnvironmentStrings("%windir%"), @SW_HIDE, $STDERR_MERGED)
	;Local $iPID = Run($Command, "", @SW_HIDE, $STDOUT_CHILD)
	DebugLogger('- ' & @ScriptLineNumber & " PID " & $iPID & @CRLF)
	ProcessWaitClose($iPID)
	Local $sOutput = StdoutRead($iPID)
	DebugLogger('- ' & @ScriptLineNumber & " $sOutput " & @CRLF & $sOutput & @CRLF)
	Return $sOutput
EndFunc   ;==>CmdExecSTDOUT




Func _Qwinsta()
	; This Function returns an array with qwinsta data.
	;C:\Documents and Settings\Administrator>qwinsta %USERNAME%|find /i "%USERNAME%"
	;>rdp-tcp#6         Administrator             2  Active  rdpwd

	Local $sSessionname = EnvGet("SESSIONNAME")
	If @OSArch = "X86" Then
		Local $sQwinsta = CmdExecSTDOUT("qwinsta.exe") ;32 Bit
	Else
		Local $sQwinsta = CmdExecSTDOUT("%windir%\Sysnative\qwinsta.exe") ;64 Bit
	EndIf
	Local $aQwinstaTmp = StringSplit(StringTrimRight(StringStripCR($sQwinsta), StringLen(@CRLF)), @CRLF)
	DebugLogger('- ' & @ScriptLineNumber & " ubound($aQwinstaTmp,$UBOUND_ROWS):" & UBound($aQwinstaTmp, $UBOUND_ROWS) & @CRLF)
	Local $aQwinsta[UBound($aQwinstaTmp, $UBOUND_ROWS)][7]
	;$aQwinsta[0][0]="Active"
	;$aQwinsta[0][1]="SessionName"
	;$aQwinsta[0][2]="Username"
	;$aQwinsta[0][3]="ID"
	;$aQwinsta[0][4]="STATE"
	;$aQwinsta[0][5]="TYPE"
	;$aQwinsta[0][6]="DEVICE"

	For $Itemp5 = 0 To (UBound($aQwinstaTmp, $UBOUND_ROWS)) - 1
		DebugLogger('- ' & @ScriptLineNumber & " aQwinstaTmp:" & ":" & $Itemp5 & ":" & $aQwinstaTmp[$Itemp5] & @CRLF)
		If $bListQwinstaInCurrentLogg Then AddStatusTextEntry("Qwinsta", $aQwinstaTmp[$Itemp5])
		$aQwinsta[$Itemp5][0] = _StringFixSpaces(StringMid($aQwinstaTmp[$Itemp5], 1, 1))
		$aQwinsta[$Itemp5][1] = _StringFixSpaces(StringMid($aQwinstaTmp[$Itemp5], 2, 17))
		$aQwinsta[$Itemp5][2] = _StringFixSpaces(StringMid($aQwinstaTmp[$Itemp5], 19, 20))
		$aQwinsta[$Itemp5][3] = _StringFixSpaces(StringMid($aQwinstaTmp[$Itemp5], 39, 9))
		$aQwinsta[$Itemp5][4] = _StringFixSpaces(StringMid($aQwinstaTmp[$Itemp5], 48, 8))
		$aQwinsta[$Itemp5][5] = _StringFixSpaces(StringMid($aQwinstaTmp[$Itemp5], 56, 12))
		$aQwinsta[$Itemp5][6] = _StringFixSpaces(StringMid($aQwinstaTmp[$Itemp5], 68))
	Next
	If $bDebugFlagArray Then _ArrayDisplay($aQwinsta, "$aQwinsta", Default, 8)
	Return $aQwinsta

EndFunc   ;==>_Qwinsta

Func _StringFixSpaces($StringToFix)
	Return StringStripWS($StringToFix, $STR_STRIPLEADING + $STR_STRIPTRAILING + $STR_STRIPSPACES)
EndFunc   ;==>_StringFixSpaces




Func _TsconToConsole($SessionID)

	; %windir%\System32\tscon.exe 2 /dest:console
	; Method will throw the RDP session to Console (warning, will disconnect RDP sessions)

	If @OSArch = "X86" Then
		Local $iPIDC = Run(@ComSpec & " /c " & "%windir%\System32\tscon.exe " & $SessionID & " /dest:console",  _WinAPI_ExpandEnvironmentStrings("%windir%"), @SW_HIDE, $STDERR_MERGED)
	Else
		Local $iPIDC = Run(@ComSpec & " /c " & "%windir%\Sysnative\tscon.exe " & $SessionID & " /dest:console",  _WinAPI_ExpandEnvironmentStrings("%windir%"), @SW_HIDE, $STDERR_MERGED)
	Endif
	DebugLogger('- ' & @ScriptLineNumber & " _TsconToConsole($SessionID) " & $SessionID & @CRLF)
	ProcessWaitClose($iPIDC)
	Local $TsConsoleOut = StdoutRead($iPIDC)
	Eventlogg(False,4, "_TsconToConsole: $SessionID'" & $SessionID & "'" & @CRLF & @ComSpec & " /c " & "%windir%\System32\tscon.exe " & $SessionID & " /dest:console" & @CRLF & $TsConsoleOut )

EndFunc   ;==>_TsconToConsole



Func _KeepAliveScreen()
	; This moves the mouse a bit for keepalive operations
	If $MousexHistory = MouseGetPos(0) and $MouseyHistory = MouseGetPos(1) Then
		Local $Mousex = MouseGetPos(0)
		Local $Mousey = MouseGetPos(1)
		MouseMove($Mousex, $Mousey + 1, 0)
		MouseMove($Mousex, $Mousey - 1, 0)
		MouseMove($Mousex, $Mousey, 0)
	EndIf
	Global $MousexHistory = MouseGetPos(0)
	Global $MouseyHistory = MouseGetPos(1)
	DebugLogger('- ' & @ScriptLineNumber & " --> " & "_KeepAliveScreen! '" & @CRLF)
EndFunc   ;==>_KeepAliveScreen


;Not used Functions -----------------------------------------------


Func _GetTerminalServerSetting()
	$LocalOutput = ""
	$strComputer = @ComputerName
	$objWMIService = ObjGet("winmgmts:\\" & $strComputer & "\root\cimv2")
	DebugLogger('- ' & @ScriptLineNumber & " Win32_TerminalServiceSetting, starting!" & @CRLF)
	If IsObj($objWMIService) Then
		$colSettings = $objWMIService.ExecQuery("Select * from Win32_TerminalServiceSetting")
		If IsObj($colSettings) Then
			For $objItem In $colSettings
				$LocalOutput = $objItem.AllowTSConnections
			Next
			Return $LocalOutput
			DebugLogger('- ' & @ScriptLineNumber & " Win32_TerminalServiceSetting:" & $LocalOutput & @CRLF)
		EndIf
	EndIf
EndFunc   ;==>_GetTerminalServerSetting


Func _WinAPI_Base64Decode($sB64String)
	Local $aCrypt = DllCall("Crypt32.dll", "bool", "CryptStringToBinaryA", "str", $sB64String, "dword", 0, "dword", 1, "ptr", 0, "dword*", 0, "ptr", 0, "ptr", 0)
	If @error Or Not $aCrypt[0] Then Return SetError(1, 0, "")
	Local $tBuffer = DllStructCreate("byte[" & $aCrypt[5] & "]")
	$aCrypt = DllCall("Crypt32.dll", "bool", "CryptStringToBinaryA", "str", $sB64String, "dword", 0, "dword", 1, "struct*", $tBuffer, "dword*", $aCrypt[5], "ptr", 0, "ptr", 0)
	If @error Or Not $aCrypt[0] Then Return SetError(2, 0, "")
	Return DllStructGetData($tBuffer, 1)
EndFunc   ;==>_WinAPI_Base64Decode

Func _WinAPI_LZNTDecompress(ByRef $tInput, ByRef $tOutput, $iBufferSize)
	$tOutput = DllStructCreate("byte[" & $iBufferSize & "]")
	If @error Then Return SetError(1, 0, 0)
	Local $aRet = DllCall("ntdll.dll", "uint", "RtlDecompressBuffer", "ushort", 0x0002, "struct*", $tOutput, "ulong", $iBufferSize, "struct*", $tInput, "ulong", DllStructGetSize($tInput), "ulong*", 0)
	If @error Then Return SetError(2, 0, 0)
	If $aRet[0] Then Return SetError(3, $aRet[0], 0)
	Return $aRet[6]
EndFunc   ;==>_WinAPI_LZNTDecompress

; ------------------------------------------------- Inactive Functions end

Func _ScreenSaverActive($bBoolean)
	Local Const $SPI_SETSCREENSAVEACTIVE = 17
	Local $lActiveFlag

	Dim $lActiveFlag
	Dim $retvaL

	If $bBoolean Then
		$lActiveFlag = 1
	Else
		$lActiveFlag = 0
	EndIf

	$dll = DllOpen("user32.dll")
	$retvaL = DllCall($dll, "long", "SystemParametersInfo", "long", $SPI_SETSCREENSAVEACTIVE, "long", $lActiveFlag, "long", 0, "long", 0)
	DllClose($dll)
	DebugLogger('- ' & @ScriptLineNumber & "_ScreenSaverActive " & $bBoolean & @CRLF)
EndFunc   ;==>_ScreenSaverActive




; TrayHandling --------------------------

Func Trayinit()
   $rc1 = TraySetIcon(@ScriptDir & "\WindowsNinja.ico", -1)
	Opt("TrayMenuMode", 3) ; The default tray menu items will not be shown and items are not checked when selected. These are options 1 and 2 for TrayMenuMode.
	Opt("TrayOnEventMode", 1) ; Enable TrayOnEventMode.
	; TrayTip($TrayTipHeader, "This screen session is monitored..", 3)

	Local $About=TrayCreateItem("About")
	TrayItemSetOnEvent($About, "TrayAboutScript")

	Local $ScreenShot=TrayCreateItem("Take a ScreenShot")
	TrayItemSetOnEvent($ScreenShot, "TrayInstantScreenShot")


    Local $TrayDebug=TrayCreateMenu("Debug")

	Global $TrayDebugStart=TrayCreateItem("Start Debug File",$TrayDebug)
	TrayItemSetOnEvent($TrayDebugStart, "TrayDebugStart")

	Global $TrayDebugStop=TrayCreateItem("Stop Debug File",$TrayDebug)
	TrayItemSetOnEvent($TrayDebugStop, "TrayDebugStop")

	TrayItemSetState($TrayDebugStop,$TRAY_CHECKED)


	Local $TrayDebugView=TrayCreateItem("View Debug File",$TrayDebug)
	TrayItemSetOnEvent($TrayDebugView, "TrayDebugView")


	Local $TrayAu3Info=TrayCreateItem("Launch Au3Info",$TrayDebug)
	TrayItemSetOnEvent($TrayAu3Info, "TrayAu3Info")

	Local $ParametersScript=TrayCreateItem("Parameters",$TrayDebug)
	TrayItemSetOnEvent($ParametersScript, "TrayParametersScript")

	Local $PictureDirectory=TrayCreateItem("Open Screenshot Directory",$TrayDebug)
	TrayItemSetOnEvent($PictureDirectory, "TrayPictureDirectory")


	Local $WindowsScript=TrayCreateItem("Windows Detected",$TrayDebug)
	TrayItemSetOnEvent($WindowsScript, "TrayWindowsScript")

	Local $SendMail=TrayCreateItem("Send Test Mail",$TrayDebug)
	TrayItemSetOnEvent($SendMail, "TraySendMail")

	Local $TrayStatusText=TrayCreateItem("Status",$TrayDebug)
	TrayItemSetOnEvent($TrayStatusText, "TrayStatusText")

	Local $EnableMouseDetect=TrayCreateItem("Enforce Mouse Detect",$TrayDebug)
	TrayItemSetOnEvent($EnableMouseDetect, "TrayEnableMouseDetect")

	Local $TrayQWINSTA=TrayCreateItem("QWINSTA",$TrayDebug)
	TrayItemSetOnEvent($TrayQWINSTA, "TrayQWINSTA")


	Local $TraySlackTest=TrayCreateItem("Test Slack",$TrayDebug)
	TrayItemSetOnEvent($TraySlackTest, "TraySlackTest")


	TrayCreateItem("") ; Create a separator line.

    Global $TrayPause=TrayCreateMenu("Pause")

    Global $TrayPauseOn=TrayCreateItem("On",$TrayPause)
	TrayItemSetOnEvent($TrayPauseOn, "TrayPauseOn")

    Global $TrayPauseOff=TrayCreateItem("Off",$TrayPause)
	TrayItemSetOnEvent($TrayPauseOff, "TrayPauseOff")

    Global $TrayPauseOffEnforce=TrayCreateItem("Off(StartNow)",$TrayPause)
	TrayItemSetOnEvent($TrayPauseOffEnforce, "TrayPauseOffEnforce")

	Local $Exit=TrayCreateItem("Exit")
	TrayItemSetOnEvent($Exit, "TrayExitScript")



EndFunc

Func TrayDebugStart()
	 TrayItemSetState($TrayDebugStart,$TRAY_CHECKED)
	 TrayItemSetState($TrayDebugStop,$TRAY_UNCHECKED)
	global $bDebugFlag=True
EndFunc

Func TrayDebugStop()
	 TrayItemSetState($TrayDebugStop,$TRAY_CHECKED)
	 TrayItemSetState($TrayDebugStart,$TRAY_UNCHECKED)
	global $bDebugFlag=False
EndFunc

Func TrayDebugView()
Run(@ComSpec & " /c start notepad.exe " & Chr(34) & $sDebugFile & Chr(34),  _WinAPI_ExpandEnvironmentStrings("%windir%"), @SW_SHOWDEFAULT,  $RUN_CREATE_NEW_CONSOLE)
EndFunc


Func TraySlackTest()
	$sSlackStatustmp=SlackPostAttachementCurl($sSlackInfoChannel,"Test Message from " & @ComputerName,$sTestPicture,$sSlackToken,$sTestPicture, "Cowabunga, it works!", $sSlackUploadURL )
	;DecodeJsonFromCurlOutput($sSlackStatustmp, "id")
	MsgBox($MB_SYSTEMMODAL + $MB_ICONINFORMATION, "Slack Test" , $sSlackStatustmp)
EndFunc

Func TraySendMail()
	Local $BlatStatus=BlatMailSend($BlatSmtpServer,@ComputerName & $BlatSenderSuffix, $sPictureFileHistoric ,$BlatTo,"WindowsNinja Test Mail" ,"See attached picture" ,"0")
	MsgBox($MB_SYSTEMMODAL + $MB_ICONINFORMATION, "SendMailStatus" , $BlatStatus)
EndFunc

Func TrayQWINSTA()
	 _ArrayDisplay(_Qwinsta(), "$Qwinsta", Default, 8)
EndFunc

Func TrayEnableMouseDetect()
	Global $bMouseMovedArmed=True
EndFunc

Func TrayInstantScreenShot()
	local $Pitch
	for $Pitch = 3000 to 3400 step 100
		Beep($Pitch,25)
	Next
	GetCurrentScreenshot(True)
EndFunc

Func TrayAu3Info()
	Run($sAu3InfoExe, _WinAPI_ExpandEnvironmentStrings("%windir%"), @SW_SHOWDEFAULT,  $RUN_CREATE_NEW_CONSOLE)
EndFunc

Func TrayPictureDirectory()
	Run(@ComSpec & " /c explorer.exe " & Chr(34) & $sPictureBasePath & Chr(34),  _WinAPI_ExpandEnvironmentStrings("%windir%"), @SW_SHOWDEFAULT,  $RUN_CREATE_NEW_CONSOLE)
EndFunc

Func TrayParametersScript()

;If $bDebugFlag Then _ArrayDisplay($alist, "alist", Default, 8)
_ArrayDisplay($aActionArray, "Parameters", Default, 8)

EndFunc

Func TrayWindowsScript()

_ArrayDisplay($blist, "Windows Detected", Default, 8)

EndFunc

Func TrayExitScript()
   Eventlogg(False,4, $sEnvUserName &  " did a manual Exit")
	Exit
EndFunc

Func TrayAboutScript()
MsgBox($MB_SYSTEMMODAL + $MB_ICONINFORMATION, "WindowsNinja" & " V" & $Version, "Created by Tonny Roger Holm" & @CRLF & @CRLF & _
					"This tool will look for configured error messages on windows dialogue boxes, and notify if present" & @CRLF & @CRLF & _
					"Normally no activity will be monitored while the user is logged on." & @CRLF & @CRLF & _
		Chr(34) &	"Dude, quit playing window ninja over there. Why can't you just use this tool and be chill?" & Chr(34) & @CRLF & @CRLF & _
					"Path: " & StringLeft(@AutoItExe, StringInStr(@AutoItExe, "\", 0, -1) - 1))


EndFunc

Func TrayPauseOn()
	  TrayInfoPause(True)
	  Eventlogg(True,4, $sEnvUserName &  " triggered manual pause")
EndFunc


Func TrayPauseOff()

	  TrayInfoPause(False)
	  Global $bMouseMovedArmed=False
	  Eventlogg(True,4, $sEnvUserName &  " triggered manual start")
EndFunc



Func TrayPauseOffEnforce()
	  Global $bSleepEnable = False
	  TrayInfoPause(False)
	  Global $bMouseMovedArmed=False
	  Eventlogg(True,4, $sEnvUserName &  " triggered manual start + Enforce")
EndFunc



Func TrayInfoPause($bPause)
   if $bPause then
	$bNoAction =True
	  TraySetToolTip("WindowsNinja is sleeping")
		  TraySetState($TRAY_ICONSTATE_FLASH)
		  TrayItemSetState($TrayPauseOn,$TRAY_CHECKED)
		  TrayItemSetState($TrayPauseOff,$TRAY_UNCHECKED)
	  ;$rc1 = TraySetIcon(@ScriptDir & "\WindowsNinja-Pause.ico", -1)
   Else
	  ;$rc1 = TraySetIcon(@ScriptDir & "\WindowsNinja.ico", -1)
	 $bNoAction =False
	  TraySetToolTip("WindowsNinja is capturing data")
		  TraySetState($TRAY_ICONSTATE_STOPFLASH)
		  TrayItemSetState($TrayPauseOn,$TRAY_UNCHECKED)
		  TrayItemSetState($TrayPauseOff,$TRAY_CHECKED)
EndIf
EndFunc

Func TrayStatusText()


   MsgBox($MB_SYSTEMMODAL, "Window Ninja Status", $StatusText)
EndFunc

; TrayHandling --------------------------




Func BlatMailSend($BlatSmtpServer,$BlatSender, $BlatAttach,$BlatTo,$BlatSubject,$BlatBody,$BlatPriority)
Local $BlatExe=_PathFull(@ScriptDir) & "\Tools\blat\blat.exe"

$BlatSubject=ProperTXT($BlatSubject)
$BlatBody=ProperTXT($BlatBody)



if $BlatAttach="" Then
	Local $BlatAttachStub=""
Else
	Local $BlatAttachStub=" -attach " & Chr(34) & $BlatAttach & Chr(34)
Endif



Local $BlatMessage=" -to " & $BlatTo & " -subject " & Chr(34) & $BlatSubject & Chr(34) & " -body " & Chr(34) & $BlatBody & Chr(34) &  " -Priority " & $BlatPriority & $BlatAttachStub
Local $BlatInit = CmdExecSTDOUT(Chr(34) & $BlatExe & Chr(34) & " -install " & $BlatSmtpServer & " " & $BlatSender)
Eventlogg(False,4, "Initiating Mail " & @CRLF & Chr(34) & $BlatExe & Chr(34) & " -install " & $BlatSmtpServer & " " & $BlatSender & @CRLF & "Feedback was:" & @CRLF & $BlatInit)


	Local $iPID = Run(Chr(34) & $BlatExe & Chr(34) & $BlatMessage,  _WinAPI_ExpandEnvironmentStrings("%windir%"), @SW_HIDE, $STDERR_MERGED)
	DebugLogger('- ' & @ScriptLineNumber & " PID " & $iPID & @CRLF)
	ProcessWaitClose($iPID)
	Local $BlatSendt = StdoutRead($iPID)

; Local $BlatSendt = CmdExecSTDOUT(Chr(34) & $BlatExe & Chr(34) & $BlatMessage)
Eventlogg(False,4, "Sending Mail " & @CRLF & Chr(34) & $BlatExe & Chr(34) & $BlatMessage &  @CRLF & "Feedback was:" & @CRLF & $BlatSendt )




AddStatusTextEntry("MAILINIT", $BlatExe & " -install " & $BlatSmtpServer & " " & $BlatSender)
AddStatusTextEntry("MAILSEND", $BlatExe & $BlatMessage)


; "%EXEPath%blat.exe" -install %smtpserver% %sender%
; "%EXEPath%blat.exe" -attacht %FileToSend% -to %DestinationAddress% -subject %subject% -body %body% -Priority %Priority%

Return "-----INITIALIZE-----" & @CRLF  & $BlatInit & @CRLF & @CRLF & "-----RETURN----- " & @CRLF & $BlatSendt & @CRLF

EndFunc

Func DebugLogger($WriteOut)
	if $bDebugFlag Then
		ConsoleWrite($WriteOut)
		FileWrite($hDebugFile, $WriteOut)
		$iDebugLinesCount=$iDebugLinesCount+1
		if $iDebugLinesCount>$iDebugMaxLines Then
			$bDebugFlag=False
			TrayDebugStop()
			Global $iDebugLinesCount=0
			Eventlogg(False,4, "Max debug lines reached (" & $iDebugLinesCount & "). Stopping debug")
		Endif
	EndIf

EndFunc

Func _MouseMovedSensorTimerCheck($iDeltaSleepTimeFactor)
	if $iMouseMovedDetectPauseTime>0 Then
		if $iMouseMovedDetectPauseTimeLeft>0 then $iMouseMovedDetectPauseTimeLeft=$iMouseMovedDetectPauseTimeLeft-$iDeltaSleepTimeFactor
	Endif
EndFunc

Func _MouseMovedSensor()
	;Sensor for MouseMoved

	if $iMouseMovedDetectPauseTime>0 Then
		If $bMouseMovedArmed then
			if ($MousexHistoryActivity = MouseGetPos(0)) and ($MouseyHistoryActivity = MouseGetPos(1)) then
				$bMouseMoved=False
			Else
				$bMouseMoved=True
			EndIf

			DebugLogger('+ ' & @ScriptLineNumber & " --> " & "iMoseMovedDetectPauseTime: " &  $iMouseMovedDetectPauseTime & ",iMouseMovedDetectPauseTimeLeft:" &  $iMouseMovedDetectPauseTimeLeft & ",bMouseMoved:" & $bMouseMoved & ",$bMouseMovedArmed:" & $bMouseMovedArmed & ",MouseGetPos:" & MouseGetPos(0) & ",$bNoAction:" &  $bNoAction & @CRLF)

	; If property MoseMovedDetectPauseTime is set above 0 and the $bMouseMovedArmed mechanism is set then check if $bMouseMoved is true. If so do a count down before enabling again.
			if $bMouseMoved Then
					$iMouseMovedDetectPauseTimeLeft=$iMouseMovedDetectPauseTime
					if not $bNoAction then
						AddStatusTextEntry("MoseMovedDetectPauseTime", "MOUSEMOVEDETECTED:" & $iMouseMovedDetectPauseTimeLeft)
						Eventlogg(True,4, $sEnvUserName &  " is using screen. Pause capturing")
						TrayInfoPause(True)
						$bNoAction=True
					Endif
				Else
					if 	$bNoAction then
						if $iMouseMovedDetectPauseTimeLeft<=0 Then ;Timer run Out, Enable again
							AddStatusTextEntry("MoseMovedDetectPauseTime", "TIMEOUT-ENABLECAPTURE")
							Eventlogg(True,4, $sEnvUserName &  " is idle, re-enable capture")
							TrayInfoPause(False)
							$bNoAction=False
						EndIf
					Endif
				EndIf
		EndIf
		$MousexHistoryActivity = MouseGetPos(0)
		$MouseyHistoryActivity = MouseGetPos(1)
	Endif

EndFunc



Func _NoActionWhileUserConnectedSensor()

	; Sensor for NoActionWhileUserConnected
	IF $bNoActionWhileUserConnected then
	 if not $bUserIsConnected=$bUserIsConnectedHistory then ;ok, something has happend
		if $bUserIsConnected=True then ;User is connected, stop monitoring
		   $bNoAction=True
		   if not $bTsconToConsoleInitiated then ;This means: If NOT a switch to console has happend execute the following (Console is always active, and RDP throw to Console happend is unlikely to be a truly active user)
			Eventlogg(False,4, $sEnvUserName &  " is logged on - pause monitoring")
			AddStatusTextEntry("NoActionWhileUserConnected", "IN_PAUSE_MODE")
			TrayInfoPause(True)
		   Else
			AddStatusTextEntry("NoActionWhileUserConnected", "IN_SWITCHTOCONSOLE_MODE")
			TrayInfoPause(False)
		   Endif

		Else
		   $bNoAction=False
		   Eventlogg(False,4, $sEnvUserName &  " is not logged on - start monitoring")
		   TrayInfoPause(False)
		EndIf
	 $bUserIsConnectedHistory=$bUserIsConnected
	   EndIf
	EndIf
EndFunc

Func DecodeJsonFromCurlOutput($sCurlRawReturn, $sVariable)
; By Tonny Roger Holm
; Lightweight scan of returned values from Curl to get result from an variable in JSON format
; Probably should have used RFC4627 https://www.autoitscript.com/forum/topic/104150-json-udf-library-fully-rfc4627-compliant/
; ....but lets try a shortcut and if works....
; We dont care much about sections at this stage...
MsgBox($MB_SYSTEMMODAL + $MB_ICONINFORMATION, "$sCurlRawReturn" , ProperTXT($sCurlRawReturn))
;First lets find the JSON Section "{"
$sJSONString=StringRight($sCurlRawReturn,StringInStr($sCurlRawReturn,"{"))
MsgBox($MB_SYSTEMMODAL + $MB_ICONINFORMATION, "$sJSONString" , ProperTXT($sJSONString))
$sJSONString1=StringRight($sJSONString,StringInStr($sJSONString,Chr(34) & $sVariable & Chr(34) & ":" & Chr(34)))
MsgBox($MB_SYSTEMMODAL + $MB_ICONINFORMATION, "$sJSONString1" , ProperTXT($sJSONString1))
;$sJSONString2=StringLeft(StringInStr($sJSONString1,Chr(34))
;MsgBox(0, Default, $sJSONString2)
EndFunc


Func _GetTotalCPULoad()
; wmic cpu get loadpercentage|find /i /v "Load"
$GetCPUPercent = CmdExecSTDOUT("wmic cpu get loadpercentage|find /i /v " & Chr(34) & "Load" & Chr(34) )
Return $GetCPUPercent
DebugLogger('+ ' & @ScriptLineNumber & " --> " & "GetTotalCPULoad: " & $GetCPUPercent & @CRLF)
EndFunc










