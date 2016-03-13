#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Change2CUI=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.2
 Author:         relascope

 Script Function:
	Automatically exports Cubase Projects (.cpr) to Music XML
	and PDF (with PDFCreator) 
	using Cubase 5 v5.1.1 (may work with other versions)
	Working Directory should be used before
	Problems with Network Drives. - Avoid!

#ce ----------------------------------------------------------------------------

#include <WindowsConstants.au3>
#include <MsgBoxConstants.au3>
#include <Array.au3>
#include <File.au3>
#include <GuiMenu.au3>

AutoItSetOption ("TrayIconDebug", 1);0-off

$timeComplete = TimerInit()

$iniFile = "Cubase5BatchExport.ini"

$cubaseProgramFile = IniRead($iniFile, "General", "cubaseProgramFile", "C:\Program Files\Steinberg\Cubase 5\Cubase5.exe")
$sourceFolder = IniRead($iniFile, "General", "sourceFolder", "")
$destinationFolder = IniRead($iniFile, "General", "destinationFolder", "")

$timoutCubaseStartup = Number(IniRead($iniFile, "Timeouts", "CubaseStartup", "30"))
$timeoutImport = Number(IniRead($iniFile, "Timeouts", "Import", "90"))
$timeoutRegeneration = Number(IniRead($iniFile, "Timeouts", "Regeneration", "300"))
$timeoutUiInteraction = Number(IniRead($iniFile, "Timeouts", "UiInteraction", "100"))
$timeoutExport = Number(IniRead($iniFile, "Timeoutes", "Export", "60"))

$cubaseProgramTitle = "Cubase 5"

Opt("WinTitleMatchMode", 1) ; Match Window Titles from the Beginning *

; TODO WHILE
If Not WinExists($cubaseProgramTitle) Then
	Run($cubaseProgramFile)
	WaitForCubaseToStart()
EndIf

If $sourceFolder == "" Then
	$sourceFolder = FileSelectFolder("Choose Directory containing your old ARR and ALL files", "")
EndIf

If $destinationFolder == "" Then
	$destinationFolder = FileSelectFolder("Choose Directory for Output", "")
EndIf

If $sourceFolder == "" Then
	MsgBox($MB_OK, "Folder Empty", "Source Folder cannot be empty")
	Exit(-1)
EndIf
If $destinationFolder == "" Then
	MsgBox($MB_OK, "Folder Empty", "Destination Folder cannot be empty")
	Exit(-1)
EndIf


$projectFiles = _RecFileListToArray($sourceFolder, "*.cpr", 1, 1, 0)
If @error Then

; Display the return value, @error and @extended value.
MsgBox($MB_SYSTEMMODAL, "", "Return value = " & $projectFiles & @CRLF & _
        "Value of @error is: " & @error & @CRLF & _
        "Value of @extended is: " & @extended)
EndIf

; Create Directory Structure
For $i = 1 to $projectFiles[0]
		$songFilePath = GetDir($projectFiles[$i])
		DirCreate($destinationFolder & "\" &  $songFilePath)
Next

WinActivate($cubaseProgramTitle)
Sleep($timeoutUiInteraction)
If WinExists("[CLASS:SteinbergWindowClass]") Then ; Assistant
	Send("{ESC}")
	Sleep($timeoutUiInteraction)
EndIf

If WinExists("[CLASS:SteinbergWindowClass]") Then ; registration
	Send("{ESC}")
	Sleep($timeoutUiInteraction)
EndIf

For $i = 1 to $projectFiles[0]
	$projectFile = $sourceFolder & "\" & $projectFiles[$i]
	$destinationFile = $destinationFolder & "\" & $projectFiles[$i]

	ConsoleWrite($projectFile & "...")

	Local $sDrive = "", $sDir = "", $sFileName = "", $sExtension = ""
	Local $aPathSplit = _PathSplit($destinationFile, $sDrive, $sDir, $sFileName, $sExtension)

	$destinationFile = $sDrive & $sDir & $sFileName

	if FileExists ( $destinationFile & ".xml") And FileExists ($destinationFile & ".pdf") Then
		ConsoleWrite("Exists!" & @CRLF)
		ContinueLoop
	EndIf

	If FileExists ( $destinationFile & ".xml") And PrinterPdfExists($destinationFile) Then
		ConsoleWrite("Exists!" & @CRLF)
		ContinueLoop
	EndIf

	CubaseXmlPdfExport($projectFile, $destinationFile)
	ConsoleWrite("Finished!" & @CRLF)
	Sleep($timeoutRegeneration) ; prevent from hanging around
Next

$timeInMinutes = TimerDiff($timeComplete)/60000

MsgBox($MB_OK, "Finished", "Finished processing " & @CRLF & $projectFiles[0] & " Project Files " & @CRLF & "Duration (in minutes): " & $timeInMinutes & @CRLF & @CRLF & "Errors (if any) reported to log-file." & @CRLF & @CRLF & "Donations to www.dojoy.at" )




Func CubaseXmlPdfExport($inputProjectFile, $outputFile)
	WinMenuSelectItem("[CLASS:CubaseFrameWindowClass]", "", "&File", "&Open...")

	WinWaitActive("Open Project")
	Sleep($timeoutUiInteraction)
	Send($inputProjectFile)
	Sleep($timeoutUiInteraction)
	Send("{ENTER}")

	If Not (WinWaitActive("Set Project Folder", "", 10) == 0)	Then
		Sleep($timeoutUiInteraction)
		Sleep($timeoutUiInteraction)
		Send("{ENTER}")
		Sleep($timeoutUiInteraction)
		Send("{ENTER}")
		Sleep($timeoutUiInteraction)
	EndIf

	; OK BUTTON
	; ControlClick("Select directory", "", 1)
	; Sleep($timeoutUiInteraction)

	; Wait until control (MDI-child) is here
	; Optional Click away "Pending Connections"
	Local $hTimer = TimerInit()
	Global $hCtrl = 0, $Waiting = True
	While ($Waiting)
		Local $fDiff = TimerDiff($hTimer)

		If $fDiff > ($timeoutImport * 1000) Then ; Timeout for Import
			; DRY Violation CubaseRestart
			ConsoleWriteError(@CRLF & "Error processing " & $inputProjectFile & @CRLF)
			ConsoleWriteError("quitting and restart" & @CRLF)

			$killCMD = "taskkill /im Cubase5.exe /t /f"
			RunWait(@ComSpec & " /c " & $killCMD)

			$echoError = "echo " & $inputProjectFile & " >> logbadfiles5.txt" ; too lazy to learn AutoIt FileWriting
			RunWait(@ComSpec & " /c " & $echoError)

			Run($cubaseProgramFile)
			WinWait($cubaseProgramTitle)

			WaitForCubaseToStart()

			Return
		EndIf

		If $Waiting And WinExists($cubaseProgramTitle) Then
			$hCtrl = ControlGetHandle($cubaseProgramTitle, "", "SteinbergMDIWindowClass1")
			If $hCtrl Then
				; we got the handle, so the MDI Window is there
				$Waiting = False
			EndIf

			; OPTIONAL STEP
			If WinActive("Missing Ports") Then
				Sleep($timeoutUiInteraction)
				Send("{ENTER}")
				Sleep($timeoutUiInteraction)
			Endif
		EndIf
		Sleep($timeoutUiInteraction)
	WEnd

	Sleep($timeoutUiInteraction)
	Send("^a") ; Select all tracks
	Sleep($timeoutUiInteraction)
	Send("^r") ; Open Score Editor
	Sleep($timeoutUiInteraction)
	Sleep($timeoutUiInteraction)
	Sleep($timeoutUiInteraction)

	ExportMusicXML($outputFile)

	PrintScore($outputFile)

	closeOpenMdiWindows()
EndFunc

Func ActualizeMenuStatus()
		; Menu status is only actualized, if the menu is invoked. so we just do "something harmless"
	WinMenuSelectItem("[CLASS:CubaseFrameWindowClass]", "", "&File", "Clea&nup...");
	WinWaitActive("[CLASS:SteinbergWindowClass]")
	Send("{Escape}")
EndFunc

Func IsPrintingAvailable()
	ActualizeMenuStatus()
	$hWnd = WinGetHandle("[CLASS:CubaseFrameWindowClass]")
	$hMain = _GUICtrlMenu_GetMenu($hWnd)
	$hFile = _GUICtrlMenu_GetItemSubMenu($hMain, 0)

	$printingEnabled = _GUICtrlMenu_GetItemEnabled($hFile, 14)

	Return $printingEnabled
EndFunc

Func IsMusicXmlExportAvailable()
	ActualizeMenuStatus()

	$hWnd = WinGetHandle("[CLASS:CubaseFrameWindowClass]")
	$hMain = _GUICtrlMenu_GetMenu($hWnd)
	$hFile = _GUICtrlMenu_GetItemSubMenu($hMain, 0)
	$hExport = _GUICtrlMenu_GetItemSubMenu($hFile, 17)

	$musicXmlEnabled = _GUICtrlMenu_GetItemEnabled($hExport, 6)

	$menuText = _GUICtrlMenu_GetItemText($hExport, 6)
	If StringCompare($menuText, "M&usicXML...") Then
		ConsoleWrite(@CRLF & "Cannot find MusicXML Export Menu" & @CRLF)
		MsgBox(0, "CubaseExport", "MusicXML Export Menu could not be found on the right position...")
	EndIf

	Return $musicXmlEnabled
EndFunc

Func ExportMusicXML($outputFile)
	If FileExists ( $destinationFile & ".xml") Then
		ConsoleWrite("..skipping XML...")
		Return
	EndIf

	If Not IsMusicXmlExportAvailable() Then
		ConsoleWrite("..XML Export not available...(1)...")
		Return
	EndIf

	While Not WinMenuSelectItem("[CLASS:CubaseFrameWindowClass]", "", "&File", "&Export", "M&usicXML...")
		If Not IsMusicXmlExportAvailable() Then
			ConsoleWrite("..XML Export not available...(2)...")
			Return
		EndIf
		Sleep($timeoutUiInteraction)
	WEnd

	Send($outputFile)

	Sleep($timeoutUiInteraction)
	Send("{ENTER}")

	$hWnd = WinWaitActive("Export MusicXML", "", 5)

	If Not IsMusicXmlExportAvailable() Then
		ConsoleWrite("..XML Export not available...(3)...")
		Return
	EndIf

	Sleep($timeoutUiInteraction)
	Sleep($timeoutUiInteraction)

	Local $hTimer = TimerInit()
	Local $Waiting = True

	While ($Waiting)
		 If WinExists ("Export MusicXML") Then
			Sleep(200)
		 Else
			ConsoleWrite("Export did not start!")
			Return
		 EndIf

		$fDiff = TimerDiff($hTimer)
		If $fDiff > ($timeoutExport * 1000) Then

			; DRY Violation CubaseRestart
			ConsoleWriteError(@CRLF & "Error exporting " & $outputFile & @CRLF)
			ConsoleWriteError("quitting and restart" & @CRLF)

			$killCMD = "taskkill /im Cubase5.exe /t /f"
			RunWait(@ComSpec & " /c " & $killCMD)

			$echoError = "echo " & $outputFile & " >> logbadfiles5.txt" ; too lazy to learn AutoIt FileWriting
			RunWait(@ComSpec & " /c " & $echoError)

			Run($cubaseProgramFile)
			WinWait($cubaseProgramTitle)

			WaitForCubaseToStart()

			Return
		EndIf

	WEnd
EndFunc

Func PageSetupA2()
	WinMenuSelectItem("[CLASS:CubaseFrameWindowClass]", "", "&File", "&Page Setup...")
	Local $hWndPageSetup = WinWaitActive("Page Setup", "", 5)
	If $hWndPageSetup = 0 Then
		ConsoleWrite("nopagesetup")
		Return
	EndIf

	ControlCommand($hWndPageSetup, "", "ComboBox1", "SelectString", "A2")
	Sleep($timeoutUiInteraction)
	Send("{ENTER}")
	Sleep($timeoutUiInteraction)
EndFunc

Func PrinterPdfExists($outputFile)
	$strippedDestinationFile = $outputFile
	Do
		_PathSplit($strippedDestinationFile, $sDrive, $sDir, $sFileName, $sExtension)
		$strippedDestinationFile = $sDrive & $sDir & $sFileName
	Until $sExtension == ""

	if FileExists ($strippedDestinationFile & ".pdf") Then
		Return True
	EndIf

	Return False
EndFunc

Func PrintScore($outputFile)
	; PDF Creator will strip extensions.....

	If PrinterPdfExists($outputFile) Then
		ConsoleWrite("pdf Exists!" & @CRLF)
		Return
	EndIf

	If Not IsPrintingAvailable() Then
		ConsoleWrite("Printing Not Available" & @CRLF)
		Return
	EndIf

	PageSetupA2()
	WinMenuSelectItem("[CLASS:CubaseFrameWindowClass]", "", "&File", "P&rint...")
	$hWndPrint = WinWaitActive("Print", "", 5)
	if $hWndPrint = 0 Then
		ConsoleWrite("noprintwindow")
		Return
	EndIf

	Send("{ENTER}")
	Sleep($timeoutUiInteraction)

	AutomatePrinter($outputFile)
EndFunc

Func WaitForPdfCreator()
	While (WinWaitActive("PDFCreator", "", 20) == 0)
		$printerMsg = "We are automating using PDFCreator, it looks like, you are using a different printer. " & @CRLF & @CRLF
		$printerMsg &= "Please set PDFCreator to your default printer and ==THEN== click Yes. " & @CRLF
		$printerMsg &= "If you want to use a different printer, please adjust the script. - Or find someone who can..." & @CRLF & @CRLF
		$printerMsg &= "Continue?"
		$answer = MsgBox($MB_YESNO, "Printer Config", $printerMsg)
		If ($answer == $IDYES) Then
			; Nothing to do
		ElseIf ($answer == $IDNO) Then
			MsgBox(0, "Bye", "Script terminating. Consult www.dojoy.at if you need help!")
			Exit(0)
		Else
			MsgBox(0, "Strange things...", "...can happen...")
			Exit(0)
		Endif
	WEnd
EndFunc

Func AutomatePrinter($outputFile)
	; HERE WE AUTOMATE PDFCREATOR: HAS TO BE ADJUSTED!!!
	Local $sDrive = "", $sDir = "", $sFileName = "", $sExtension = ""

	WaitForPdfCreator()

	Send("{TAB}")
	Sleep($timeoutUiInteraction)
	Send("+{END}")
	Send("{DEL}")

	Local $aPathSplit = _PathSplit($outputFile, $sDrive, $sDir, $sFileName, $sExtension)

	Sleep($timeoutUiInteraction)
	Send($sFileName)
	Sleep($timeoutUiInteraction)
	Send("{ENTER}")

	WinWaitActive("Select destination")
	Sleep($timeoutUiInteraction)

	Send($outputFile)
	Sleep($timeoutUiInteraction)
	Send("{ENTER}")

	; We do not know, how long it takes...
	Sleep($timeoutUiInteraction * 5)
EndFunc

Func closeOpenMdiWindows()
	Sleep($timeoutUiInteraction)
	While Not WinMenuSelectItem("[CLASS:CubaseFrameWindowClass]", "", "&Window", "&Close All")
		Sleep($timeoutUiInteraction)
	WEnd

	If Not (WinWaitActive("[CLASS:SteinbergWindowClass]", "", 5) == 0) Then
		; Button1 Save
		; Button2 Don't Save
		; Button3 Cancel
		Sleep($timeoutUiInteraction)
		ControlClick($cubaseProgramTitle, "", "Button2")
		Sleep($timeoutUiInteraction)
	EndIf

	Sleep($timeoutUiInteraction)
	Sleep($timeoutUiInteraction)
EndFunc

Func WaitForCubaseToStart($timeoutInSec = $timoutCubaseStartup, $displayMsg = True)

	$timer = TimerInit()

	While Not WinMenuSelectItem("[CLASS:CubaseFrameWindowClass]", "", "&Help", "&About Cubase")
		$timerDiff = TimerDiff($timer)

		If $timerDiff > $timeoutInSec * 1000 Then
			If $displayMsg Then
				MsgBox($MB_OK, "Start Cubase5", "Please Start Cubase5. When it has started, click OK. ")
			EndIf
			Return False
		EndIf

		Sleep($timeoutUiInteraction)
	WEnd
	Send ("{ESC}")
	Sleep($timeoutUiInteraction)
	Return True
EndFunc



;;;;;;;;;;;;;;;;;;========================
;;;;;;;;;;;;;;;;;;========================
;;;;;;;;;;;;;;;;;;========================THIRD PARTY FUNCTIONS
;;;;;;;;;;;;;;;;;;========================
;;;;;;;;;;;;;;;;;;========================


; #FUNCTION# ======================================================================================================
; Name...........: GetDir
; Description ...: Returns the directory of the given file path
; Syntax.........: GetDir($sFilePath)
; Parameters ....: $sFilePath - File path
; Return values .: Success - The file directory
;                  Failure - -1, sets @error to:
;                  |1 - $sFilePath is not a string
; Author ........: Renan Maronni <renanmaronni@hotmail.com>
; Modified.......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......:
; ==================================================================================================================

Func GetDir($sFilePath)

    Local $aFolders = StringSplit($sFilePath, "\")
    Local $iArrayFoldersSize = UBound($aFolders)
    Local $FileDir = ""

    If (Not IsString($sFilePath)) Then
        Return SetError(1, 0, -1)
    EndIf

    $aFolders = StringSplit($sFilePath, "\")
    $iArrayFoldersSize = UBound($aFolders)

    For $i = 1 To ($iArrayFoldersSize - 2)
        $FileDir &= $aFolders[$i] & "\"
    Next

    Return $FileDir

EndFunc   ;==>GetDir




; #FUNCTION# ====================================================================================================

; ================
; Name...........: _RecFileListToArray
; Description ...: Lists files and\or folders in a specified path (Similar to using Dir with the /B Switch)
; Syntax.........: _RecFileListToArray($sPath[, $sFilter = "*"[, $iFlag = 0[, $iRecur = 0[, $iFullPath = 0]]]])
; Parameters ....: $sPath   - Path to generate filelist for.
;                 $sFilter - Optional the filter to use, default is *. Search the Autoit3 helpfile for the word "WildCards" For details.
;                 $iFlag   - Optional: specifies whether to return files folders or both
;                 |$iFlag=0 (Default) Return both files and folders
;                 |$iFlag=1 Return files only
;                 |$iFlag=2 Return Folders only
;                 $iRecur  - Optional: specifies whether to search in subfolders
;                 |$iRecur=0 (Default) Do not search in subfolders
;                 |$iRecur=1 Search in subfolders
;                 $iFullPath  - Optional: specifies whether to include initial path in result string
;                 |$iFullPath=0 (Default) Do not include initial path
;                 |$iFullPath=1 Include initial path
; Return values .: @Error - 1 = Path not found or invalid
;                 |2 = Invalid $sFilter
;                 |3 = Invalid $iFlag
;                 |4 = Invalid $iRecur
;                 |5 = Invalid $iFullPath
;                 |6 = No File/Folder Found
; Author ........: SolidSnake <MetalGX91 at GMail dot com>
; Modified.......: 22 Jan 09 by Melba23 - added recursive search and full path options
; Remarks .......: The array returned is one-dimensional and is made up as follows:
;                               $array[0] = Number of Files\Folders returned
;                               $array[1] = 1st File\Folder
;                               $array[2] = 2nd File\Folder
;                               $array[3] = 3rd File\Folder
;                               $array[n] = nth File\Folder
; Related .......:
; Link ..........;
; Example .......; Yes
; ====================================================================================================

; ===========================
;Special Thanks to Helge and Layer for help with the $iFlag update
; speed optimization by code65536
;===============================================================================
Func _RecFileListToArray($sPath, $sFilter = "*", $iFlag = 0, $iRecur = 0, $iFullPath = 0)
    Local $asFileList[1], $sFullPath
    If Not FileExists($sPath) Then Return SetError(1, 1, "")
    If StringRight($sPath, 1) <> "\" Then $sPath = $sPath & "\"
    If (StringInStr($sFilter, "\")) Or (StringInStr($sFilter, "/")) Or (StringInStr($sFilter, ":")) Or (StringInStr($sFilter, ">")) Or (StringInStr($sFilter, "<")) Or (StringInStr($sFilter, "|")) Or (StringStripWS($sFilter, 8) = "") Then Return SetError(2, 2, "")
    If Not ($iFlag = 0 Or $iFlag = 1 Or $iFlag = 2) Then Return SetError(3, 3, "")
    If Not ($iRecur = 0 Or $iRecur = 1) Then Return SetError(4, 4, "")
    If $iFullPath = 0 Then
        $sFullPath = $sPath
    ElseIf $iFullPath = 1 Then
        $sFullPath = ""
    Else
        Return SetError(5, 5, "")
    EndIf
    _FLTA_Search($sPath, $sFilter, $iFlag, $iRecur, $sFullPath, $asFileList)
    If $asFileList[0] = 0 Then Return SetError(6, 6, "")
    Return $asFileList
EndFunc  ;==>_FileListToArray

; #INTERNAL_USE_ONLY#=================================================================================

; ===========================
; Name...........: _FLTA_Search
; Description ...: Searches folder for files and then recursively searches in subfolders
; Syntax.........: _FLTA_Search($sStartFolder, $sFilter, $iFlag, $iRecur, $sFullPath, ByRef $asFileList)
; Parameters ....: $sStartFolder - Value passed on from UBound($avArray)
;                 $sFilter - As set in _FileListToArray
;                 $iFlag - As set in _FileListToArray
;                 $iRecur - As set in _FileListToArray
;                 $sFullPath - $sPath as set in _FileListToArray
;                 $asFileList - Array containing found files/folders
; Return values .: None
; Author ........: Melba23 based on code from _FileListToArray by SolidSnake <MetalGX91 at GMail dot com>
; Modified.......:
; Remarks .......: This function is used internally by _FileListToArray.
; Related .......:
; Link ..........;
; Example .......;
; ====================================================================================================

; ===========================
Func _FLTA_Search($sStartFolder, $sFilter, $iFlag, $iRecur, $sFullPath, ByRef $asFileList)

    Local $hSearch, $sFile

    If StringRight($sStartFolder, 1) <> "\" Then $sStartFolder = $sStartFolder & "\"
; First look for filtered files/folders in folder
    $hSearch = FileFindFirstFile($sStartFolder & $sFilter)
    If $hSearch > 0 Then
        While 1
            $sFile = FileFindNextFile($hSearch)
            If @error Then ExitLoop
            Switch $iFlag
                Case 0; Both files and folders
                    If $iRecur And StringInStr(FileGetAttrib($sStartFolder & $sFile), "D") <> 0 Then ContinueLoop
                Case 1; Files Only
                    If StringInStr(FileGetAttrib($sStartFolder & $sFile), "D") <> 0 Then ContinueLoop
                Case 2; Folders only
                    If StringInStr(FileGetAttrib($sStartFolder & $sFile), "D") = 0 Then ContinueLoop
            EndSwitch
            If $iFlag = 1 And StringInStr(FileGetAttrib($sStartFolder & $sFile), "D") <> 0 Then ContinueLoop
            If $iFlag = 2 And StringInStr(FileGetAttrib($sStartFolder & $sFile), "D") = 0 Then ContinueLoop
            _FLTA_Add($asFileList, $sFullPath, $sStartFolder, $sFile)
        WEnd
        FileClose($hSearch)
        ReDim $asFileList[$asFileList[0] + 1]
    EndIf

    If $iRecur = 1 Then
    ; Now look for subfolders
        $hSearch = FileFindFirstFile($sStartFolder & "*.*")
        If $hSearch > 0 Then
            While 1
                $sFile = FileFindNextFile($hSearch)
                If @error Then ExitLoop
                If StringInStr(FileGetAttrib($sStartFolder & $sFile), "D") And ($sFile <> "." Or $sFile <> "..") Then
                ; If folders needed, add subfolder to array
                    If $iFlag <> 1 Then _FLTA_Add($asFileList, $sFullPath, $sStartFolder, $sFile)
                ; Recursive search of this subfolder
                    _FLTA_Search($sStartFolder & $sFile, $sFilter, $iFlag, $iRecur, $sFullPath, $asFileList)
                EndIf
            WEnd
            FileClose($hSearch)
        EndIf
    EndIf

EndFunc

; #INTERNAL_USE_ONLY#=================================================================================

; ===========================
; Name...........: _FLTA_Add
; Description ...: Searches folder for files and then recursively searches in subfolders
; Syntax.........: _FLTA_Add(ByRef $asFileList, $sFullPath, $sStartFolder, $sFile)
; Parameters ....: $asFileList - Array containing found files/folders
;                 $sFullPath - $sPath as set in _FileListToArray
;                 $sStartFolder - Value passed on from UBound($avArray)
;                 $sFile - Full path of file/folder to add to $asFileList
; Return values .: Function only changes $asFileList ByRef
; Author ........: Melba23 based on code from _FileListToArray by SolidSnake <MetalGX91 at GMail dot com>
; Modified.......:
; Remarks .......: This function is used internally by _FileListToArray.
; Related .......:
; Link ..........;
; Example .......;
; ====================================================================================================

; ===========================
Func _FLTA_Add(ByRef $asFileList, $sFullPath, $sStartFolder, $sFile)

    Local $sAddFolder

    $asFileList[0] += 1
    If UBound($asFileList) <= $asFileList[0] Then ReDim $asFileList[UBound($asFileList) * 2]
    If $sFullPath = "" Then
        $sAddFolder = $sStartFolder
    Else
        $sAddFolder = StringReplace($sStartFolder, $sFullPath, "")
    EndIf
    $asFileList[$asFileList[0]] = $sAddFolder & $sFile

EndFunc
