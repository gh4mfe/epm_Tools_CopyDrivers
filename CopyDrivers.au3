#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Outfile_x64=CopyDrivers_x64_3.3.14.2.exe
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Res_Fileversion=1.6.0.0
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#Region AutoIt3Wrapper directives section
#EndRegion AutoIt3Wrapper directives section

#cs ----------------------------------------------------------------------------
	
	AutoIt Version: 3.2.12.1
	Author:         Jan Buelens, Landesk Software (idea stolen from Sergio Ribeiro)
	
	Script Function:
	CopyDrivers. This script copies machine dependent drivers to a provisioning / OSD client machine.
	A mapping table (CopyDrivers.ini) is used to map the WMI machine model to a driver source folder.
	
	A GUI is included to build CopyDrivers.ini
	
	Change History:
	   V1.5.1 05 Dec 2008. This is the version that came out with the HII V9 document
	   V1.6   07 July 2009.
	            * Some people are copying dozens over even hundreds of MB. Progress info is therefore desirable. Picked up a copy function with progress bar
				* logging
				* Wildcard matching. Example:
				
				[Models]
				*=common
				Precision Workstation T3400=T3400
				2007FVG=ThinkPad T60
				HP Compaq dc5100 SFF*=DC5100
				
				Now suppose a target machine has a WMI model name = "HP Compaq dc5100 SFF(PZ579UA)". The first line (*) matches all models, therefore the subfolders called "common" will be copied. 
				The last line also matches, therefore subfolder DC5100 will also be copied.	
	
	
#ce ----------------------------------------------------------------------------

#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <EditConstants.au3>
#include <StaticConstants.au3>

$progname = "CopyDrivers V1.6"
Dim $logfilename = ""			; log file (from /log command line parameter)
Dim $log = -1

Dim $iniFilename = ""
Dim $Manufacturer = ""
Dim $Model = ""
Dim $Version = ""
Dim $subfolder = ""
Dim $SourceFolder = ""
Dim $TargetFolder = ""
Dim $sysprep = "C:\sysprep\sysprep.inf"

Dim $bVerbose = False
Dim $bRunOnce = True
Dim $bCmdLines = True

; ===========================================================================================
; Validate command line parameters and verify that copydrivers.ini exists
; ===========================================================================================

If $CmdLine[0] = 0 Then DoGui() ; no command line parameters - do gui

If $CmdLine[0] > 0 And ($CmdLine[1] = "/?" Or $CmdLine[1] = "-?" Or $CmdLine[1] = "help") Then
	Usage()
EndIf

For $n = 1 To $CmdLine[0]
	$s = ""
	If $CmdLine[$n] = "/s" Or $CmdLine[$n] = "-s" Then
		; "legacy" syntax: /s <sourcefolder>
		If $n >= $CmdLine[0] Then Usage()
		$SourceFolder = $CmdLine[$n + 1]
		$n = $n + 1
	ElseIf ValParam($CmdLine[$n], "s", $s) Then
		; "new" syntax: /s=<sourcefolder>
		$SourceFolder = $s
	ElseIf $CmdLine[$n] = "/d" Or $CmdLine[$n] = "-d" Then
		; "legacy" syntax: /d <targetfolder>
		If $n >= $CmdLine[0] Then Usage()
		$TargetFolder = $CmdLine[$n + 1]
		$n = $n + 1
	ElseIf ValParam($CmdLine[$n], "d", $s) Then
		; "new" syntax: /d <targetfolder>
		$TargetFolder = $s
	ElseIf ValParam($CmdLine[$n], "log", $s) Then
		$logfilename = $s
	ElseIf $CmdLine[$n] = "/c" Or $CmdLine[$n] = "-c" Then
		; if no other command line parameters are required, use /c to copy drivers rather than launch GUI
	ElseIf $CmdLine[$n] = "/v" Or $CmdLine[$n] = "-v" Then
		$bVerbose = True;
		;If there is a reason for not doing the cmdlines and GuiRunOnce stuff, uncomment these lines
	; ElseIf $CmdLine[$n] = "/cmdlines" Or $CmdLine[$n] = "-cmdlines"  Then
	;	$bCmdLines = False
	; ElseIf $CmdLine[$n] = "/RunOnce" Or $CmdLine[$n] = "-RunOnce"  Then
	;	$bRunOnce = False
	Else
		Usage()
	EndIf
Next

LogOpen($logfilename)

$iniFilename = PathConcat(@ScriptDir, "copydrivers.ini") ; @ScriptDir is folder in which this script (or compiled program) resides
LogIniSection($iniFilename, "Config")
If Not FileExists($iniFilename) Then ErrorExit("File not found: " & $iniFilename, 2)

; If no source and target folders were defined on the command line, take them from the [Config] section of copydrivers.ini

If $SourceFolder = "" Then $SourceFolder = IniRead($iniFilename, "Config", "DriversSource", "")
If $TargetFolder = "" Then $TargetFolder = IniRead($iniFilename, "Config", "DriversTarget", "")
If $SourceFolder = "" Then ErrorExit("No Drivers Source Folder defined", 3)
If $TargetFolder = "" Then ErrorExit("No Drivers Target Folder defined", 4)

LogMessage("source: " & $SourceFolder & ", target: " & $TargetFolder)

; ===========================================================================================
; Read Manufacturer, Model and Version from WMI. We only use Model.
; ===========================================================================================

ReadWmi($Manufacturer, $Model, $Version)
LogMessage("WMI info: Manufacturer=" & $Manufacturer & ", Model=" & $Model & ", Version=" & $Version)
If $bVerbose Then MsgBox(0, $progname, "Manufacturer: " & $Manufacturer & @CRLF & "Model: " & $Model & @CRLF & "Version: " & $Version)

; ===========================================================================================
; Copy the driver files. Do a wildcard match of the WMI model with each line in [Models] section of copydrivers.ini
; and copy all subfolders that match
; ===========================================================================================

DirCreate($TargetFolder)
If Not IsFolder($TargetFolder) Then ErrorExit("Unable to create target folder: " & $TargetFolder, 8)

;IniReadSection returns a 2 dimensional array of keywords and values; $ini[n][0] is key # n, $ini[n][1] is value # n; $ini[0][0] is the number of elements

LogIniSection($iniFilename, "Models")
$ini = IniReadSection($iniFilename, "Models")
If @error Or $ini[0][0] = 0 Then ErrorExit("There is no [Models] section in " & $iniFilename, 5)


For $n = 1 To $ini[0][0]
	; if $Model = $ini[$n][0] Then	; if you want a plain compare rather than wildcard match, uncomment this line and comment out next line
	If WildcardMatch($Model, $ini[$n][0]) Then
		$subfolder = $ini[$n][1]
		$src = PathConcat($SourceFolder, $subfolder)
		LogMessage("Match on line " & $n & ": " & $ini[$n][0])
		LogMessage(GetTime() & " Start copy from " & $src & " to " & $TargetFolder)
		If Not IsFolder($src) Then ErrorExit("Source folder not found: " & $src, 7)
		; Previous versions used the built-in DirCopy() function. Replaced with _CopyDirWithProgress(). If the new function causes trouble, use DirCopy again
		; If Not DirCopy($src, $TargetFolder, 1) Then ErrorExit("Unable to copy folder: " & $TargetFolder, 9)	; 1 on DirCopy means overwrite existing files
		$stat = _CopyDirWithProgress($src, $TargetFolder)
		LogMessage(GetTime() & " End copy from " & $src & " to " & $TargetFolder)
		If $stat Then ErrorExit("Error copying driver files", 9)
	EndIf
Next
If $subfolder = "" Then ErrorExit("No match found for Model """ & $Model & """ in " & $iniFilename, 6)


; ===========================================================================================
; Handle RunOnce and CmdLines
; ===========================================================================================

If $bRunOnce Then DoRunOnce()
If $bCmdLines Then DoCmdLines()

; ===========================================================================================
; Done
; ===========================================================================================

; ===========================================================================================
Func DoCmdLines()
	; Run at deployment time if the $bCmdLines is true. If there is a cmdlines.txt file in the drivers folder that we just copied, set up
	; sysprep.inf such that it will be processed at mini-setup time. If sysprep.inf already refers to a cmdlines.txt file, merge it. The cmdlines.txt file must be
	; in the format as described in the sysprep documentation. Example:
	;
	;  [cmdlines]
	;  "c:\drivers\setup\driver1\setup.exe"
	;
	; This program also has a GUI that allows cmdlines.txt to be edited in a convenient way, without the user being aware of the format or the location of the file.
	; ===========================================================================================

	Local $MyBase = $TargetFolder
	Local $MyCmdLines = PathConcat($MyBase, "cmdlines.txt")
	Local $OemCmdLines = ""

	If Not FileExists($MyCmdLines) Then
		LogMessage("File not found: " & $MyCmdLines)
		Return
	EndIf
	
	LogMessage("File exists: " & $MyCmdLines)

	If Not FileExists($sysprep) Then ErrorExit("File not found: " & $sysprep, 2)
	Local $InstallFilesPath = IniRead($sysprep, "unattended", "InstallFilesPath", "")
	If $InstallFilesPath = "" Then
		; no InstallFilesPath in sysprep.inf - create one
		$InstallFilesPath = $MyBase
		IniWrite($sysprep, "unattended", "InstallFilesPath", $InstallFilesPath)
		LogMessage("Added to sysprep.inf [unattended]: InstallFilesPath=" & $InstallFilesPath)
	EndIf

	$OemCmdLines = PathConcat($InstallFilesPath, "$oem$\cmdlines.txt")
	
	LogFileContents($OemCmdLines, $OemCmdLines & " before merge:")

	If Not FileExists($OemCmdLines) Then
		; no $oem$\cmdlines.txt exist - just copy ours
		Local $success = FileCopy($MyCmdLines, $OemCmdLines, 8) ; 8 = create folders
		If $success = 0 Then ErrorExit("Copy " & $MyCmdLines & " to " & $OemCmdLines & " failed", 3)
	Else
		; A $oem$\cmdlines.txt already exists - append ours
		Local $file1 = FileOpen($MyCmdLines, 0) ; 0 = read
		If $file1 = -1 Then ErrorExit("Error opening " & $MyCmdLines, 5)
		Local $file2 = FileOpen($OemCmdLines, 1) ; 1 = append
		If $file2 = -1 Then ErrorExit("Error opening " & $OemCmdLines, 4)
		While 1
			$line = FileReadLine($file1)
			If @error Then ExitLoop
			If StringStripWS($line, 8) <> "[commands]" Then ; StringStripWS($line,8) strips all white space
				FileWriteLine($file2, $line)
			EndIf
		WEnd
		FileClose($file1)
		FileClose($file2)
	EndIf
	
	LogFileContents($OemCmdLines, $OemCmdLines & " after merge:")

EndFunc   ;==>DoCmdLines


; ===========================================================================================
Func DoRunOnce()
; Run at deployment time if the $bRunonce is true. If there is a file called GuiRunOnce.ini in the drivers
; folder that we just copied, merge its GuiRunOnce section with the sysprep.inf GuiRunOnce section. The GuiRunOnce.ini file must be
; in the format as described in the sysprep documentation. Example:
;
;  [GuiRunOnce]
;  Command0="c:\drivers\driver1\setup.exe"
;  Command1="c:\drivers\driver2\setup.exe"
;
; This program also has a GUI that allows GuiRunOnce.ini to be edited in a convenient way, without the user to be aware of the format or the location of the file.
; ===========================================================================================


	Local $MyRunOnce = PathConcat($TargetFolder, "GuiRunOnce.ini")

	If Not FileExists($MyRunOnce) Then
		LogMessage("File not found: " & $MyRunOnce)
		Return
	EndIf
	
	LogMessage("File exists: " & $MyRunOnce)
	
	LogIniSection($sysprep, "GuiRunOnce", "sysprep.inf [GuiRunOnce] section before merge:")

	If Not FileExists($sysprep) Then ErrorExit("File not found: " & $sysprep, 2)
	Local $section1[1][1]
	$section1[0][0] = 0
	$section1 = IniReadSection($sysprep, "GuiRunOnce")
	If @error Then
		Dim $section1[1][1]
		$section1[0][0] = 0
	EndIf
	Local $section2 = IniReadSection($MyRunOnce, "GuiRunOnce")
	If @error Or $section2[0][0] = 0 Then
		LogMessage("No [GuiRunOnce] section in " & $MyRunOnce)
		Return
	EndIf

	Local $count, $i
	$count = $section1[0][0]
	For $i = 1 To $section2[0][0]
		$count = $count + 1
		ReDim $section1[$count + 1][2]
		$section1[$count][0] = $section2[$i][0]
		$section1[$count][1] = $section2[$i][1]
	Next
	$section1[0][0] = $count
	For $i = 1 To $section1[0][0]
		$section1[$i][0] = "Command" & ($i - 1)
	Next

	IniWriteSection($sysprep, "GuiRunOnce", $section1)
	
	LogIniSection($sysprep, "GuiRunOnce", "sysprep.inf [GuiRunOnce] section after merge:")


EndFunc   ;==>DoRunOnce

; ===========================================================================================
; Wilcard match. Autoit has native support for regular expression matching, but not for wildcard matching. This function massages the pattern so it can
; be used as a regular expression. Copied from http://www.autoitscript.com/forum/index.php?showtopic=78620.
Func WildcardMatch($str, $pattern)
; ===========================================================================================
	Local $sChar, $sChars = '\.+[^]$(){}=!<>|:'
	For $i = 1 To StringLen($sChars)
		$sChar = StringMid($sChars, $i, 1)
		$pattern = StringReplace($pattern, $sChar, '\' & $sChar)
	Next
	$pattern = StringReplace($pattern, '?', '.{1}')
	$pattern = StringReplace($pattern, '*', '.*')
	Return StringRegExp($str, $pattern)
EndFunc   ;==>WildcardMatch


; ===========================================================================================
; Set the 3 WMI attributes mentioned. We only use the Model, but feel free to organise things differently
Func ReadWmi(ByRef $Manufacturer, ByRef $Model, ByRef $Version)
; ===========================================================================================

	$objWMIService = ObjGet("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
	If $objWMIService = 0 Then ErrorExit("Failed to connect to WMI", 10)

	$colRows = $objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
	For $row In $colRows
		$Manufacturer = StringStripWS($row.Manufacturer, 3)
		$Model = StringStripWS($row.Model, 3)
	Next

	$colRows = $objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystemProduct")
	For $row In $colRows
		$Version = StringStripWS($row.Version, 3)
	Next

	Return True

EndFunc   ;==>ReadWmi

; ===========================================================================================
Func Usage()
; ===========================================================================================

	MsgBox("0", $progname, _
			"This program copies machine specific drivers from a source folder on a" _
			 & @CRLF & "server to a destination folder on the local machine." & @CRLF _
			 & @CRLF & "Optional command line parameters:" _
			 & @CRLF & "  /s=<sourcedir> : base path for machine specfic driver folders" _
			 & @CRLF & "                   (overrides DriversSource in copydrivers.ini)" _
			 & @CRLF & "  /d=<destdir> : target folder on the local machine." _
			 & @CRLF & "                   (overrides DriversTarget in copydrivers.ini)" _
			 & @CRLF & "  /log=<logfile> : log file." _
			 & @CRLF & "  /v : verbose" _
			 & @CRLF & "  /c : copy" _
			 & @CRLF & "" _
			 & @CRLF & "CopyDrivers requires a list that associates machine types with driver folders." _
			 & @CRLF & "This list is to be supplied in the [Models] section of copydrivers.ini, which" _
			 & @CRLF & "associates WMI model names with specific subfolders of <sourcedir>." & @CRLF _
			 & @CRLF & "When invoked without command line switches, CopyDrivers opens a GUI that allows" _
			 & @CRLF & "copydrivers.ini to be edited. Use /c to do the copying rather than show the GUI." _
			)

	Exit 1

EndFunc   ;==>Usage

; ===========================================================================================
Func GetTime()
; ===========================================================================================
	Local $s = @HOUR & ":" & @MIN & ":" & @SEC
	return $s
EndFunc

; ===========================================================================================
Func LogOpen(ByRef $logfilename)
; ===========================================================================================

	Local $scriptName = StringTrimRight(@ScriptName, 4)
	While 1
		If $logfilename <> "" Then
			; log filename specified on command line
			$log = FileOpen($logfilename, 10)	; 10 = 2 (write, create) + 8 (create path)
			ExitLoop
		EndIf			

		; No /log command line parameter. If there is a copydrivers.ini file with a DriversTarget parameter (typically c:\drivers), create the log in there
		Local $iniFilename = PathConcat(@ScriptDir, "copydrivers.ini")
		Local $DriverPath = IniRead($iniFilename, "Config", "DriversTarget", "")
		if $DriverPath <> "" Then
			$DriverPath = StringStripWS($DriverPath, 3)
			if StringRight($DriverPath, 1) <> "\" Then $DriverPath &= "\"
			$logfilename = $DriverPath & $scriptName & ".log"
			$log = FileOpen($logfilename, 10)	; 10 = 2 (write, create) + 8 (create path)
			if $log <> -1 Then ExitLoop
		EndIf
		
		; No /log command parameter and no copydrivers.ini. If running from a local path, create log in folder of running program		
		If Not IsRemote(@ScriptFullPath) Then
			$logfilename = StringTrimRight(@ScriptFullPath, 3) & "log"
			$log = FileOpen($logfilename, 2)
			if $log <> -1 Then ExitLoop
		EndIf
		
		; failed to create log
		$log = -1
		$logfilename = ""
		return
	Wend
	
	LogCmdLine()
EndFunc

; ===========================================================================================
Func LogMessage($msg)
; ===========================================================================================
	FileWriteLine($log, $msg)
EndFunc

; ===========================================================================================
Func LogCmdLine()
; ===========================================================================================
	Local $n
	LogMessage($progname & ", command line parameter(s): " & $CmdLine[0])
	For $n = 1 to $CmdLine[0]
		LogMessage("   " & $CmdLine[$n])
	Next
EndFunc

; ===========================================================================================
Func LogIniSection($inifilename, $inisection, $msg = Default)
; ===========================================================================================
	Local $i
	if $msg = Default Then
		LogMessage($inifilename & ",section [" & $inisection & "]:")
	Else
		LogMessage($msg)
	EndIf
	Local $section = IniReadSection($inifilename, $inisection)
	if @error Then
		if not FileExists($inifilename) Then
			LogMessage("   File does not exist: " & $inifilename)
			Return
		EndIf
		LogMessage("   " & $inifilename & " includes no [" & $inisection & "] section")
		Return
	EndIf
	For $i = 1 to $section[0][0]
		LogMessage("   " & $section[$i][0] & " = " & $section[$i][1])
	Next
EndFunc

; ===========================================================================================
Func LogFileContents($filename, $msg = Default)
; ===========================================================================================
	if $msg = Default Then
		LogMessage($filename & ":")
	Else
		LogMessage($msg)
	EndIf
	Local $f = FileOpen($filename, 0)
	if @error Then
		if not FileExists($filename) Then
			LogMessage("   File does not exist: " & $filename)
			Return
		EndIf
		LogMessage("   Error opening " & $filename)
		Return
	EndIf
	While 1
		Local $line = FileReadLine($f)
		if @error Then ExitLoop
		LogMessage("   " & $line)
	WEnd
	FileClose($f)
EndFunc

; ===========================================================================================
; Return true if $s is a network path. Must be full path.
Func IsRemote($s)
; ===========================================================================================
	If StringLeft($s, 2) = "\\" Then Return True
	Local $drive = StringLeft($s, 3)
	if DriveGetType($drive) = "Network" Then Return True
	Return False
EndFunc

; ===========================================================================================
; Return true if $s is a folder
Func IsFolder($s)
; ===========================================================================================
	If Not FileExists($s) Or Not StringInStr(FileGetAttrib($s), "D") Then Return False
	Return True
EndFunc

; ===========================================================================================
; Concatenate a filename ($s) with a base path
Func PathConcat($base, $s)
; ===========================================================================================
	$base = StringStripWS($base,3)
	$s = StringStripWS($s,3)
	if StringRight($base,1) <> "\" Then $base &= "\"
	if StringLeft($s,1) = "\" Then $s = StringTrimLeft($s,1)
	Return $base & $s
EndFunc

; ===========================================================================================
; Return true if running under WinPE
Func IsWinPE()
; ===========================================================================================
	If EnvGet("SystemDrive") = "X:" Then Return True
	Return False
EndFunc

; ===========================================================================================
Func ErrorExit($msg, $exitcode)
; ===========================================================================================
	LogMessage($msg)
	FileClose($log)
	MsgBox(0x40010, $progname, $msg, 10)	; 10 is timeout, i.e. the msgbox closes after 10 seconds
	Exit $exitcode
EndFunc

; ===========================================================================================
; parse command line parameter such as /keyw=something. Examples:
;    ValParam("/path=c:\temp", "path", $value) sets $value to "c:\temp" and returns True
;    ValParam("-path=c:\temp", "path", $value) sets $value to "c:\temp" and returns True
;    ValParam("/path=c:\temp", "dir", $value) sets $value to "" and returns False
Func ValParam($param, $keyword, ByRef $value)
; ===========================================================================================
	$value = ""
	Local $p1 = "/" & $keyword & "="
	Local $p2 = "-" & $keyword & "="
	Local $len = StringLen($p1)
	if StringLen($param) < ($len + 1) Then Return False
	Local $t = StringLeft($param, $len)
	if ($t <> $p1) And ($t <> $p2) Then Return False
	$value = StringMid($param, $len + 1)	; 1 based
	Return True

EndFunc

#cs -----------------------------------------------------------------------------------------
	
	Here comes the GUI stuff. It does nothing that you can't do by simply editing the CopyDrivers.ini file
	
#ce -----------------------------------------------------------------------------------------

; ===========================================================================================
; Main GUI function called when program is invoked with no command line parameters
Func DoGui()
; ===========================================================================================
	$iniFilename = @ScriptDir & "\copydrivers.ini" ; @ScriptDir is folder in which this script (or compiled program) resides
	$SourceFolder = IniRead($iniFilename, "Config", "DriversSource", "")
	$TargetFolder = IniRead($iniFilename, "Config", "DriversTarget", "")

	#Region ### START Koda GUI section ### Form=z:\install\autoit\koda_1.7.0.1\forms\myform1.kxf
	$Form_Main = GUICreate($progname, 452, 322)
	$BtnOK = GUICtrlCreateButton("OK", 16, 280, 97, 25, 0)
	$EditSource = GUICtrlCreateInput("", 13, 32, 305, 21)
	GUICtrlSetState(-1, $GUI_DISABLE)
	$BtnConfig = GUICtrlCreateButton("Edit", 335, 32, 57, 21, 0)
	$ListView1 = GUICtrlCreateListView("WMI Model|Subfolder", 13, 120, 305, 145)
	GUICtrlSendMsg(-1, 0x101E, 0, 150)
	GUICtrlSendMsg(-1, 0x101E, 1, 150)
	; GUICtrlSetTip(-1, "abc")
	$BtnAdd = GUICtrlCreateButton("Add", 335, 126, 57, 21, 0)
	$BtnEdit = GUICtrlCreateButton("Edit", 335, 157, 57, 21, 0)
	$BtnDelete = GUICtrlCreateButton("Delete", 335, 190, 57, 21, 0)
	GUICtrlCreateLabel("Drivers source folder", 16, 14, 101, 17)
	$BtnCancel = GUICtrlCreateButton("Cancel", 133, 281, 97, 25, 0)
	$EditTarget = GUICtrlCreateInput("", 13, 79, 305, 21)
	GUICtrlSetState(-1, $GUI_DISABLE)
	GUICtrlCreateLabel("Drivers target folder", 16, 61, 96, 17)
	GUISetState(@SW_SHOW)
	#EndRegion ### END Koda GUI section ###

	GUICtrlSetData($EditSource, $SourceFolder)
	GUICtrlSetData($EditTarget, $TargetFolder)

	;IniReadSection returns a 2 dimensional array of keywords and values; $ini[n][0] is key # n, $ini[n][1] is value # n; $ini[0][0] is the number of elements
	$items = 0
	$count = 0
	$ini = IniReadSection($iniFilename, "Models")
	If Not @error Then
		$items = $ini[0][0]
	EndIf

	$count1 = $items
	If ($items = 0) Then $count1 = 1

	Dim $item[$count1][4] ; we'll store the model in item[n][0], the folder in item[n][1], the listview controlid in item[n][2] and the state (0 = deleted, 1 = active) in item[n][3]
	If $items = 0 Then
		; if the section was empty or non-existing, we create one dummy item to avoid run-time errors
		$item[0][0] = ""
		$item[0][1] = ""
		$item[0][2] = 0
		$item[0][3] = 0
	EndIf

	For $n = 1 To $items
		$item[$n - 1][0] = $ini[$n][0]
		$item[$n - 1][1] = $ini[$n][1]
		$item[$n - 1][2] = GUICtrlCreateListViewItem($ini[$n][0] & "|" & $ini[$n][1], $ListView1)
		$item[$n - 1][3] = 1
		$count = $count + 1
	Next

	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE
				Exit
			Case $BtnOK
				IniWrite($iniFilename, "Config", "DriversSource", $SourceFolder)
				IniWrite($iniFilename, "Config", "DriversTarget", $TargetFolder)
				Dim $ini[$count][2]
				$i = 0
				For $n = 0 To UBound($item) - 1
					If $item[$n][3] = 1 Then
						$ini[$i][0] = $item[$n][0]
						$ini[$i][1] = $item[$n][1]
						$i = $i + 1
					EndIf
				Next
				IniWriteSection($iniFilename, "Models", $ini, 0)
				Exit
			Case $BtnCancel
				Exit
			Case $BtnConfig
				$newSource = $SourceFolder
				$newTarget = $TargetFolder
				If EditConfig($newSource, $newTarget) Then
					$SourceFolder = $newSource
					$TargetFolder = $newTarget
					GUICtrlSetData($EditSource, $SourceFolder)
					GUICtrlSetData($EditTarget, $TargetFolder)
				EndIf
			Case $BtnDelete
				$id = GUICtrlRead($ListView1)
				For $n = 0 To UBound($item) - 1
					If $item[$n][2] = $id And $item[$n][3] = 1 Then
						GUICtrlDelete($id)
						$item[$n][3] = 0
						$count = $count - 1
					EndIf
				Next
			Case $BtnAdd
				Dim $newModel = ""
				Dim $newFolder = ""
				If AddModel($newModel, $newFolder, 0) Then
					$n = UBound($item);
					ReDim $item[$n + 1][4]
					$item[$n][0] = $newModel
					$item[$n][1] = $newFolder
					$item[$n][2] = GUICtrlCreateListViewItem($newModel & "|" & $newFolder, $ListView1)
					$item[$n][3] = 1
					$count = $count + 1
				EndIf
			Case $BtnEdit
				$id = GUICtrlRead($ListView1)
				For $n = 0 To UBound($item) - 1
					If $item[$n][2] = $id And $item[$n][3] = 1 Then
						ExitLoop
					EndIf
				Next
				If $n >= UBound($item) Then ContinueLoop
				Dim $newModel = $item[$n][0]
				Dim $newFolder = $item[$n][1]
				If AddModel($newModel, $newFolder, 1) Then
					$item[$n][0] = $newModel
					$item[$n][1] = $newFolder
					GUICtrlSetData($id, $newModel & "|" & $newFolder)
				EndIf

		EndSwitch
	WEnd


EndFunc   ;==>DoGui

; ===========================================================================================
; GUI function called when the Add or Edit button is pressed.
Func AddModel(ByRef $Model, ByRef $folder, $flag) ; flag = 0: Add  flag = 1: Edit
; ===========================================================================================

	$title = "Add Model"
	If $flag = 1 Then $title = "Edit Model"

	#Region ### START Koda GUI section ### Form=Z:\install\AutoIt\koda_1.7.0.1\Forms\Form_AddItem.kxf
	$Form_AddModel = GUICreate($title, 429, 413)
	GUICtrlCreateGroup("", 8, 1, 297, 137)
	$InputModel = GUICtrlCreateInput("", 16, 40, 209, 21)
	$InputFolder = GUICtrlCreateInput("", 16, 97, 209, 21, BitOR($ES_AUTOHSCROLL, $ES_READONLY))
	GUICtrlCreateLabel("Model", 16, 16, 33, 17)
	GUICtrlCreateLabel("Subfolder", 17, 73, 49, 17)
	$Btn_WMI = GUICtrlCreateButton("WMI", 236, 40, 57, 21, 0)
	$BtnBrowse = GUICtrlCreateButton("..", 236, 97, 57, 21, 0)
	$CmdLines = GUICtrlCreateEdit("", 16, 192, 385, 81, BitOR($ES_AUTOVSCROLL, $ES_AUTOHSCROLL, $ES_WANTRETURN, $WS_HSCROLL, $WS_VSCROLL, $WS_BORDER), $ES_MULTILINE)
	$ButtonOK = GUICtrlCreateButton("&OK", 321, 11, 75, 25, 0)
	$ButtonCancel = GUICtrlCreateButton("&Cancel", 322, 43, 75, 25, 0)
	$RunOnce = GUICtrlCreateEdit("", 16, 304, 385, 81, BitOR($ES_AUTOVSCROLL, $ES_AUTOHSCROLL, $ES_WANTRETURN, $WS_HSCROLL, $WS_VSCROLL, $WS_BORDER), $ES_MULTILINE)
	GUICtrlCreateGroup("Command lines for drivers that require a setup program", 8, 152, 409, 249)
	GUICtrlCreateLabel("Before reboot (cmdlines.txt)", 16, 174, 200, 17)
	GUICtrlCreateLabel("After reboot (GuiRunonce section of sysprep.inf)", 16, 287, 300, 17)
	GUISetState(@SW_SHOW)
	#EndRegion ### END Koda GUI section ###

	GUICtrlSetData($InputModel, $Model)
	GUICtrlSetData($InputFolder, $folder)

	; If the model specific folder exists, read the cmdlines.txt and GuiRunOnce.txt files from it and display their contents in the $cmdlines and $RunOnce edit boxes.
	; If the model specific folder does not exist, disable the edit boxes

	$RunOnceText = ""
	$CmdLinesText = ""

	$ModelFolder = $SourceFolder & "\" & $folder
	If $SourceFolder = "" Or $folder = "" Then $ModelFolder = "---dummy---"
	If IsFolder($ModelFolder) Then
		$CmdLinesText = ReadCmdLines($ModelFolder)
		GUICtrlSetData($CmdLines, $CmdLinesText)
		$RunOnceText = ReadRunOnce($ModelFolder)
		GUICtrlSetData($RunOnce, $RunOnceText)
	Else
		GUICtrlSetState($CmdLines, $GUI_DISABLE)
		GUICtrlSetState($RunOnce, $GUI_DISABLE)
	EndIf

	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE
				Return False
			Case $ButtonCancel
				GUIDelete($Form_AddModel)
				Return False
			Case $BtnBrowse
				$flag = 1
				If IsWinPE() Then $flag = 0
				$old = $SourceFolder + "\" + GUICtrlRead($InputFolder)
				$new = FileSelectFolder("Select Folder", $SourceFolder, $flag, $ModelFolder) ; $flag = 1 : Show Create Folder Button (does not work in WinPE)
				If $new <> "" Then
					$new = StringMid($new, StringLen($SourceFolder) + 2)
					GUICtrlSetData($InputFolder, $new)
				EndIf
				$ModelFolder = $SourceFolder & "\" & $new
				If $SourceFolder = "" Or $new = "" Then $ModelFolder = "---dummy---"
				If IsFolder($ModelFolder) Then
					; User selected a new model specific folder - and the folder exists. Read the cmdlines.txt and GuiRunOnce.txt files from it and display their contents in the
					; $cmdlines and $RunOnce edit boxes.
					GUICtrlSetState($CmdLines, $GUI_ENABLE)
					GUICtrlSetState($RunOnce, $GUI_ENABLE)
					$CmdLinesText = ReadCmdLines($ModelFolder)
					$RunOnceText = ReadRunOnce($ModelFolder)
				Else
					; User selected a new model specific folder - and the folder does not exist. Disable the $cmdlines and $RunOnce edit boxes.
					GUICtrlSetState($CmdLines, $GUI_DISABLE)
					GUICtrlSetState($RunOnce, $GUI_DISABLE)
					$CmdLinesText = ""
					$RunOnceText = ""
				EndIf
				GUICtrlSetData($CmdLines, $CmdLinesText)
				GUICtrlSetData($RunOnce, $RunOnceText)
			Case $Btn_WMI
				ReadWmi($Manufacturer, $Model, $Version)
				GUICtrlSetData($InputModel, $Model)
			Case $ButtonOK
				$Model = StringStripWS(GUICtrlRead($InputModel), 3)
				$folder = StringStripWS(GUICtrlRead($InputFolder), 3)
				If $Model = "" Then
					MsgBox(0, $progname, "A model is required")
				ElseIf $folder = "" Then
					MsgBox(0, $progname, "A folder is required")
				Else
					$newCmdLinesText = GUICtrlRead($CmdLines)
					$newRunOnceText = GUICtrlRead($RunOnce)
					GUIDelete($Form_AddModel)
					If IsFolder($ModelFolder) Then
						If $newCmdLinesText <> $CmdLinesText Then SaveCmdLines($ModelFolder, $newCmdLinesText)
						If $newRunOnceText <> $RunOnceText Then SaveRunOnce($ModelFolder, $newRunOnceText)
					EndIf
					Return True
				EndIf
		EndSwitch
	WEnd

EndFunc   ;==>AddModel

; ===========================================================================================
; GUI function to edit the Source Folder and target folder settings
Func EditConfig(ByRef $source, ByRef $target)
; ===========================================================================================

	#Region ### START Koda GUI section ### Form=Z:\install\AutoIt\koda_1.7.0.1\Forms\Form_Config.kxf
	$Form_Config = GUICreate("Edit Config", 316, 197)
	; GUISetIcon("D:\003.ico")
	GUICtrlCreateGroup("", 8, 1, 297, 153)
	$EditSource = GUICtrlCreateInput("", 16, 38, 217, 21)
	$EditTarget = GUICtrlCreateInput("", 16, 110, 217, 21)
	GUICtrlCreateLabel("Drivers Source Folder (Specify UNC path)", 16, 16, 200, 17)
	GUICtrlCreateLabel("DriversTarget Folder", 16, 87, 100, 17)
	$BtnBrowse = GUICtrlCreateButton("..", 244, 38, 57, 21, 0)
	GUICtrlCreateGroup("", -99, -99, 1, 1)
	$BtnOK = GUICtrlCreateButton("&OK", 65, 163, 75, 25, 0)
	$BtnCancel = GUICtrlCreateButton("&Cancel", 162, 163, 75, 25, 0)
	GUISetState(@SW_SHOW)
	#EndRegion ### END Koda GUI section ###

	GUICtrlSetData($EditSource, $source)
	GUICtrlSetData($EditTarget, $target)

	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE
				Return False
			Case $BtnCancel
				GUIDelete($Form_Config)
				Return False
			Case $BtnOK
				$source = StringStripWS(GUICtrlRead($EditSource), 3)
				$target = StringStripWS(GUICtrlRead($EditTarget), 3)
				If $source = "" Then
					MsgBox(0, $progname, "A source folder is required")
				Else
					GUIDelete($Form_Config)
					Return True
				EndIf
			Case $BtnBrowse
				$new = FileSelectFolder("Select Source Folder", "", 0, $source)
				$new = StringStripWS($new, 3)
				GUICtrlSetData($EditSource, $new)
				If $new <> "" Then $source = $new
		EndSwitch
	WEnd

EndFunc   ;==>EditConfig

; ===========================================================================================
; Used by GUI to read cmdlines.txt from specified folder. The data is returned in a format ready to be fed into a GUI edit box
; (lines separated by CR-LF). The header line ([Commands]) is not included in the data.
Func ReadCmdLines($folder)
; ===========================================================================================
	Local $retstring = ""
	Local $filename = $folder & "\cmdlines.txt"
	Local $lineno = 0
	Local $file = FileOpen($filename, 0) ; 0 = read
	If $file = -1 Then Return ""

	Local $line
	While 1
		$line = FileReadLine($file)
		If @error Then ExitLoop
		If StringStripWS($line, 8) = "[commands]" Then ContinueLoop ; StringStripWS($line,8) strips all white space
		$line = StringStripWS($line, 3) ; 3 = strip leading & trailing while space
		If $line = "" Then ContinueLoop
		If StringLeft($line, 1) = '"' And StringRight($line, 1) = '"' Then
			$line = StringTrimLeft($line, 1)
			$line = StringTrimRight($line, 1)
		EndIf
		$line = StringStripWS($line, 3) ; 3 = strip leading & trailing while space
		If $line = "" Then ContinueLoop
		If $lineno > 0 Then $retstring = $retstring & @CRLF
		$lineno = $lineno + 1
		$retstring = $retstring & $line
	WEnd

	FileClose($file)
	Return $retstring

EndFunc   ;==>ReadCmdLines

; ===========================================================================================
; Used by GUI to read GuiRunOnce.ini from specified folder. The data is returned in a format ready to be fed into a GUI edit box
; (lines separated by CR-LF). The header line ([GuiRunOnce]) is not included in the data, nor are the CommandN= prefixes.
Func ReadRunOnce($folder)
; ===========================================================================================
	Local $retstring = ""
	Local $filename = $folder & "\GuiRunOnce.ini"
	Local $lineno = 0
	Local $lines = IniReadSection($filename, "GuiRunOnce")
	If @error Then Return ""

	Local $i, $line
	For $i = 1 To $lines[0][0]
		$line = $lines[$i][1]
		If $line = "" Then ContinueLoop
		If StringLeft($line, 1) = '"' And StringRight($line, 1) = '"' Then
			$line = StringTrimLeft($line, 1)
			$line = StringTrimRight($line, 1)
		EndIf
		$line = StringStripWS($line, 3) ; 3 = strip leading & trailing while space
		If $line = "" Then ContinueLoop
		If $lineno > 0 Then $retstring = $retstring & @CRLF
		$lineno = $lineno + 1
		$retstring = $retstring & $line
	Next

	Return $retstring

EndFunc   ;==>ReadRunOnce

; ===========================================================================================
; Used by GUI to save cmdlines.txt in specified folder. The input data ($text) is the raw data as read from the GUI edit control. The header line ([Commands]) is not expected to
; be included in the input data.
Func SaveCmdLines($folder, $text)
; ===========================================================================================
	If Not IsFolder($folder) Then DirCreate($folder)
	Local $filename = $folder & "\cmdlines.txt"
	Local $file = FileOpen($filename, 2) ; 2 = create
	$text = StringReplace($text, @LF, "")
	Local $lineno = 0
	Local $lines = StringSplit($text, @CR)
	Local $i, $line
	For $i = 1 To $lines[0]
		$line = StringReplace($lines[$i], @LF, "")
		$line = StringStripWS($line, 3) ; 3 = strip leading & trailing while space
		If $line = "" Then ContinueLoop
		If StringLeft($line, 1) = '"' And StringRight($line, 1) = '"' Then
			$line = StringTrimLeft($line, 1)
			$line = StringTrimRight($line, 1)
		EndIf
		$line = '"' & $line & '"'
		If $lineno = 0 Then FileWriteLine($file, "[Commands]")
		$lineno = $lineno + 1
		FileWriteLine($file, $line)
	Next
	FileClose($file)
	If $lineno = 0 Then FileDelete($filename)

EndFunc   ;==>SaveCmdLines

; ===========================================================================================
; Used by GUI to save GuiRunOnce.ini in specified folder. The input data ($text) is the raw data as read from the GUI edit control. The header line ([GuiRunOnce]) is not expected to
; be included in the input data, nor are the "CommandN=" prefixes.
Func SaveRunOnce($folder, $text)
; ===========================================================================================
	If Not IsFolder($folder) Then DirCreate($folder)
	Local $filename = $folder & "\GuiRunOnce.ini"
	Local $file = FileOpen($filename, 2) ; 2 = create
	$text = StringReplace($text, @LF, "")
	Local $lineno = 0
	Local $lines = StringSplit($text, @CR)
	Local $i, $line
	For $i = 1 To $lines[0]
		$line = StringReplace($lines[$i], @LF, "")
		$line = StringStripWS($line, 3) ; 3 = strip leading & trailing while space
		If $line = "" Then ContinueLoop
		If StringLeft($line, 1) = '"' And StringRight($line, 1) = '"' Then
			$line = StringTrimLeft($line, 1)
			$line = StringTrimRight($line, 1)
		EndIf
		$line = 'Command' & $lineno & '="' & $line & '"'
		If $lineno = 0 Then FileWriteLine($file, "[GuiRunOnce]")
		$lineno = $lineno + 1
		FileWriteLine($file, $line)
	Next
	FileClose($file)
	If $lineno = 0 Then FileDelete($filename)

EndFunc   ;==>SaveRunOnce

#cs -----------------------------------------------------------------------------------------
	End of GUI stuff.
#ce -----------------------------------------------------------------------------------------

#cs -----------------------------------------------------------------------------------------
	Copy with Progress Bar. Copied from from http://www.autoitscript.com/forum/index.php?showtopic=11313
#ce -----------------------------------------------------------------------------------------

; ===========================================================================================
Func _CopyDirWithProgress($sOriginalDir, $sDestDir)
; ===========================================================================================
	;$sOriginalDir and $sDestDir are quite selfexplanatory...
	;This func returns:
	; -1 in case of critical error, bad original or destination dir
	;  0 if everything went all right
	; >0 is the number of file not copied and it makes a log file
	;  if in the log appear as error message '0 file copied' it is a bug of some windows' copy command that does not redirect output...

	If StringRight($sOriginalDir, 1) <> '\' Then $sOriginalDir = $sOriginalDir & '\'
	If StringRight($sDestDir, 1) <> '\' Then $sDestDir = $sDestDir & '\'
	If $sOriginalDir = $sDestDir Then Return -1

	ProgressOn('Copying Drivers...', 'Building list of files...' & @LF & @LF, '', -1, -1, 18)
	Local $aFileList = _FileSearch($sOriginalDir)
	If $aFileList[0] = 0 Then
		ProgressOff()
		SetError(1)
		Return -1
	EndIf

	If FileExists($sDestDir) Then
		If Not StringInStr(FileGetAttrib($sDestDir), 'd') Then
			ProgressOff()
			SetError(2)
			Return -1
		EndIf
	Else
		DirCreate($sDestDir)
		If Not FileExists($sDestDir) Then
			ProgressOff()
			SetError(2)
			Return -1
		EndIf
	EndIf

	Local $iDirSize, $iCopiedSize = 0, $fProgress = 0
	Local $c, $filename, $iOutPut = 0, $sLost = '', $sError
	Local $Sl = StringLen($sOriginalDir)

	_Quick_Sort($aFileList, 1, $aFileList[0])

	$iDirSize = Int(DirGetSize($sOriginalDir) / 1024)

	ProgressSet(Int($fProgress * 100), $aFileList[$c], 'Copying file:')
	For $c = 1 To $aFileList[0]
		$filename = StringTrimLeft($aFileList[$c], $Sl)
		ProgressSet(Int($fProgress * 100), $aFileList[$c] & ' -> ' & $sDestDir & $filename & @LF & 'Total KB: ' & $iDirSize & @LF & 'Done KB: ' & $iCopiedSize, 'Coping file:  ' & Round($fProgress * 100, 2) & ' %   ' & $c & '/' & $aFileList[0])

		If StringInStr(FileGetAttrib($aFileList[$c]), 'd') Then
			DirCreate($sDestDir & $filename)
		Else
			If Not FileCopy($aFileList[$c], $sDestDir & $filename, 1) Then
				If Not FileCopy($aFileList[$c], $sDestDir & $filename, 1) Then ;Tries a second time
					If RunWait(@ComSpec & ' /c copy /y "' & $aFileList[$c] & '" "' & $sDestDir & $filename & '">' & @TempDir & '\o.tmp', '', @SW_HIDE) = 1 Then ; and a third time, but this time it takes the error message
						$sError = FileReadLine(@TempDir & '\o.tmp', 1)
						$iOutPut = $iOutPut + 1
						$sLost = $sLost & $aFileList[$c] & '  ' & $sError & @CRLF
					EndIf
					FileDelete(@TempDir & '\o.tmp')
				EndIf
			EndIf

			FileSetAttrib($sDestDir & $filename, "+A-RSH");<- Comment this line if you do not want attribs reset.

			$iCopiedSize = $iCopiedSize + Int(FileGetSize($aFileList[$c]) / 1024)
			$fProgress = $iCopiedSize / $iDirSize
		EndIf
	Next

	ProgressOff()

	If $sLost <> '' Then;tries to write the log somewhere.
		If FileWrite($sDestDir & 'notcopied.txt', $sLost) = 0 Then
			If FileWrite($sOriginalDir & 'notcopied.txt', $sLost) = 0 Then
				FileWrite(@WorkingDir & '\notcopied.txt', $sLost)
			EndIf
		EndIf
	EndIf

	Return $iOutPut
EndFunc   ;==>_CopyDirWithProgress

; ===========================================================================================
Func _FileSearch($sIstr, $bSF = 1)
; ===========================================================================================
	; $bSF = 1 means looking in subfolders
	; $sSF = 0 means looking only in the current folder.
	; An array is returned with the full path of all files found. The pos [0] keeps the number of elements.
	Local $sCriteria, $sBuffer, $iH, $iH2, $sCS, $sCF, $sCF2, $sCP, $sFP, $sOutPut = '', $aNull[1]
	$sCP = StringLeft($sIstr, StringInStr($sIstr, '\', 0, -1))
	If $sCP = '' Then $sCP = @WorkingDir & '\'
	$sCriteria = StringTrimLeft($sIstr, StringInStr($sIstr, '\', 0, -1))
	If $sCriteria = '' Then $sCriteria = '*.*'

	;To begin we seek in the starting path.
	$sCS = FileFindFirstFile($sCP & $sCriteria)
	If $sCS <> -1 Then
		Do
			$sCF = FileFindNextFile($sCS)
			If @error Then
				FileClose($sCS)
				ExitLoop
			EndIf
			If $sCF = '.' Or $sCF = '..' Then ContinueLoop
			$sOutPut = $sOutPut & $sCP & $sCF & @LF
		Until 0
	EndIf

	;And after, if needed, in the rest of the folders.
	If $bSF = 1 Then
		$sBuffer = @CR & $sCP & '*' & @LF;The buffer is set for keeping the given path plus a *.
		Do
			$sCS = StringTrimLeft(StringLeft($sBuffer, StringInStr($sBuffer, @LF, 0, 1) - 1), 1);current search.
			$sCP = StringLeft($sCS, StringInStr($sCS, '\', 0, -1));Current search path.
			$iH = FileFindFirstFile($sCS)
			If $iH <> -1 Then
				Do
					$sCF = FileFindNextFile($iH)
					If @error Then
						FileClose($iH)
						ExitLoop
					EndIf
					If $sCF = '.' Or $sCF = '..' Then ContinueLoop
					If StringInStr(FileGetAttrib($sCP & $sCF), 'd') Then
						$sBuffer = @CR & $sCP & $sCF & '\*' & @LF & $sBuffer;Every folder found is added in the begin of buffer
						$sFP = $sCP & $sCF & '\';                               for future searches
						$iH2 = FileFindFirstFile($sFP & $sCriteria);         and checked with the criteria.
						If $iH2 <> -1 Then
							Do
								$sCF2 = FileFindNextFile($iH2)
								If @error Then
									FileClose($iH2)
									ExitLoop
								EndIf
								If $sCF2 = '.' Or $sCF2 = '..' Then ContinueLoop
								$sOutPut = $sOutPut & $sFP & $sCF2 & @LF;Found items are put in the Output.
							Until 0
						EndIf
					EndIf
				Until 0
			EndIf
			$sBuffer = StringReplace($sBuffer, @CR & $sCS & @LF, '')
		Until $sBuffer = ''
	EndIf

	If $sOutPut = '' Then
		$aNull[0] = 0
		Return $aNull
	Else
		Return StringSplit(StringTrimRight($sOutPut, 1), @LF)
	EndIf
EndFunc   ;==>_FileSearch

; ===========================================================================================
Func _Quick_Sort(ByRef $SortArray, $First, $Last);Larry's code
; ===========================================================================================
	Local $Low, $High
	Local $Temp, $List_Separator

	$Low = $First
	$High = $Last
	$List_Separator = StringLen($SortArray[($First + $Last) / 2])
	Do
		While (StringLen($SortArray[$Low]) < $List_Separator)
			$Low = $Low + 1
		WEnd
		While (StringLen($SortArray[$High]) > $List_Separator)
			$High = $High - 1
		WEnd
		If ($Low <= $High) Then
			$Temp = $SortArray[$Low]
			$SortArray[$Low] = $SortArray[$High]
			$SortArray[$High] = $Temp
			$Low = $Low + 1
			$High = $High - 1
		EndIf
	Until $Low > $High
	If ($First < $High) Then _Quick_Sort($SortArray, $First, $High)
	If ($Low < $Last) Then _Quick_Sort($SortArray, $Low, $Last)
EndFunc   ;==>_Quick_Sort
