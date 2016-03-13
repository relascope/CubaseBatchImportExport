#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Change2CUI=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.2
 Author:         relascope

 Script Function:
	Automatically exports (really) old Cubase (.ALL, .ARR) to MIDI and CPR
	using Cubase SX v3.1.1 (may work with other SX versions)
	MIDI export options should be set in preferences!
	Cubase Working Directory should be used before
	Problems with Network Drives. - Avoid!

#ce ----------------------------------------------------------------------------

#include <WindowsConstants.au3>
#include <MsgBoxConstants.au3>


$timeComplete = TimerInit()

$iniFile = "CubaseSxBatchExport.ini"

$cubaseProgramFile = IniRead($iniFile, "General", "cubaseProgramFile", "C:\Program Files\Steinberg\Cubase SX 3\Cubasesx3.exe")
$sourceFolder = IniRead($iniFile, "General", "sourceFolder", "")
$destinationFolder = IniRead($iniFile, "General", "destinationFolder", "")

$timoutCubaseStartup = Number(IniRead($iniFile, "Timeouts", "CubaseStartup", "30"))
$timeoutImport = Number(IniRead($iniFile, "Timeouts", "Import", "90"))
$timeoutRegeneration = Number(IniRead($iniFile, "Timeouts", "Regeneration", "300"))
$timeoutUiInteraction = Number(IniRead($iniFile, "Timeouts", "UiInteraction", "100"))


Opt("WinTitleMatchMode", 1) ; Match Window Titles from the Beginning *

; TODO WHILE
If Not WinExists("Cubase SX") Then
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


$songFiles = _RecFileListToArray($sourceFolder, "*.all", 1, 1, 0)
If @error Then

; Display the return value, @error and @extended value.
MsgBox($MB_SYSTEMMODAL, "", "Return value = " & $songFiles & @CRLF & _
        "Value of @error is: " & @error & @CRLF & _
        "Value of @extended is: " & @extended)
EndIf

$arrFiles = _RecFileListToArray($sourceFolder, "*.arr", 1, 1, 0)
If @error Then

; Display the return value, @error and @extended value.
MsgBox($MB_SYSTEMMODAL, "", "Return value = " & $arrFiles & @CRLF & _
        "Value of @error is: " & @error & @CRLF & _
        "Value of @extended is: " & @extended)
EndIf

; Create Directory Structure
For $i = 1 to $songFiles[0]
		$songFilePath = GetDir($songFiles[$i])
		DirCreate($destinationFolder & "\" &  $songFilePath)
Next
For $i = 1 to $arrFiles[0]
	$arrFilePath = GetDir($arrFiles[$i])
	DirCreate($destinationFolder & "\" & $arrFilePath)
Next

WinActivate("Cubase SX")

For $i = 1 to $songFiles[0]
	$songFile = $sourceFolder & "\" & $songFiles[$i]
	$destinationFile = $destinationFolder & "\" & $songFiles[$i]

	ConsoleWrite($songFile & "...")

	if FileExists ( $destinationFile & ".mid") And FileExists ($destinationFile & ".cpr") Then
		ConsoleWrite("Exists!" & @CRLF)
		ContinueLoop
	EndIf

	CubaseSongImportExport($songFile, $destinationFile)
	ConsoleWrite("Finished!" & @CRLF)
	Sleep($timeoutRegeneration) ; prevent from hanging around
Next

For $i = 1 to $arrFiles[0]
	$arrFile = $sourceFolder & "\" & $arrFiles[$i]
	$destinationFile = $destinationFolder & "\" & $arrFiles[$i]

	ConsoleWrite($arrFile & "...")

	if FileExists ( $destinationFile & ".mid") And FileExists ($destinationFile & ".cpr") Then
		ConsoleWrite("Exists!" & @CRLF)
		ContinueLoop
	EndIf

	CubaseSongImportExport($arrFile, $destinationFile, True)
	ConsoleWrite("Finished!" & @CRLF)
	Sleep($timeoutRegeneration) ; prevent from hanging around
Next

$timeInMinutes = TimerDiff($timeComplete)/60000

MsgBox($MB_OK, "Finished", "Finished processing " & @CRLF & $songFiles[0] & " Songfiles and " & @CRLF & $arrFiles[0] & " Arrangements. " & @CRLF & "Duration: " & $timeInMinutes & @CRLF & @CRLF & "Errors (if any) reported to log-file." & @CRLF & @CRLF & "Donations to www.dojoy.at" )



Func CubaseSongImportExport($inputAllFile, $outputFile, $isArrangement = False)
	If $isArrangement Then
		WinMenuSelectItem("[CLASS:Cubase SXFrameWindowClass]", "", "&File", "&Import", "Cu&base Arrangement...")
	Else
		WinMenuSelectItem("[CLASS:Cubase SXFrameWindowClass]", "", "&File", "&Import", "&Cubase Song...")
	EndIf

	WinWait("Import Cubase")
	Sleep($timeoutUiInteraction)
	Send($inputAllFile)
	Sleep($timeoutUiInteraction)
	Send("{ENTER}")

	WinWaitActive("Select directory")
	Send("{ENTER}")
	Sleep($timeoutUiInteraction)
	Send("{ENTER}")
	Sleep($timeoutUiInteraction)

	; OK BUTTON
	ControlClick("Select directory", "", 1)
	Sleep($timeoutUiInteraction)

	; Wait until control (MDI-child) is here
	; Optional Click away "Pending Connections"
	Local $hTimer = TimerInit()
	Global $hCtrl = 0, $Waiting = True
	While ($Waiting)
		Local $fDiff = TimerDiff($hTimer)

		If $fDiff > ($timeoutImport * 1000) Then ; Timeout for Import
			; DRY Violation CubaseRestart
			ConsoleWriteError(@CRLF & "Error processing " & $inputAllFile & @CRLF)
			ConsoleWriteError("quitting and restart" & @CRLF)

			$killCMD = "taskkill /im Cubasesx3.exe /t /f"
			RunWait(@ComSpec & " /c " & $killCMD)

			$echoError = "echo " & $inputAllFile & " >> logbadfiles.txt" ; too lazy to learn AutoIt FileWriting
			RunWait(@ComSpec & " /c " & $echoError)

			Run($cubaseProgramFile)
			WinWait("Cubase SX")

			WaitForCubaseToStart()

			Return
		EndIf

		If $Waiting And WinExists("Cubase SX") Then
			$hCtrl = ControlGetHandle("Cubase SX", "", "SteinbergDocWindowClass1")
			If $hCtrl Then
				; we got the handle, so the MDI Window is there
				$Waiting = False
			EndIf

			; OPTIONAL STEP
			If WinActive("Pending Connections") Then
				Send("{ENTER}")
				Sleep($timeoutUiInteraction)
			Endif
		EndIf
		Sleep($timeoutUiInteraction)
	WEnd

	If Not FileExists($outputFile & ".mid") Then
		WinMenuSelectItem("[CLASS:Cubase SXFrameWindowClass]", "", "&File", "&Export", "&MIDI File...")

		WinWaitActive("Export MIDI File")
		; TODO EXPORT MIDI FILENAME
		Sleep($timeoutUiInteraction)
		Send($outputFile)
		Sleep($timeoutUiInteraction)
		Send("{ENTER}")
		Sleep($timeoutUiInteraction)
		; MIDI Options should be set in File=>Preferences!
		WinWaitActive("Export Options")
		Sleep($timeoutUiInteraction)
		Send("{ENTER}")
		Sleep($timeoutUiInteraction)
		WinWaitActive("Cubase SX")
		Sleep($timeoutUiInteraction)
	Else
		ConsoleWrite(@CRLF &  "Midi File Exists => SKIP!" & @CRLF)
	EndIf



	; CLOSING
	Sleep($timeoutUiInteraction)
	While Not WinMenuSelectItem("[CLASS:Cubase SXFrameWindowClass]", "", "&Window", "C&lose All")
	WEnd

	WinWaitActive("[CLASS:SteinbergModalWindowClass]")
	Sleep($timeoutUiInteraction)

	If Not FileExists($outputFile & ".cpr") Then
		; Button1 Save
		; Button2 Don't Save
		; Button3 Cancel
		Sleep($timeoutUiInteraction)
		ControlClick("Cubase SX", "", "Button1")
		Sleep($timeoutUiInteraction)
		WinWaitActive("Save As")
		Sleep($timeoutUiInteraction)
		Send($outputFile)
		Sleep($timeoutUiInteraction)
		Send("{ENTER}")
		Sleep($timeoutUiInteraction)

		$timer = TimerInit()
		; TODO Does not work: "Save As" window gets destroyed even if hanging
		While WinExists("Save As")
			$timerDiff = TimerDiff($timer)
			If $timerDiff > (120 * 1000) Then
				; DRY Violation CubaseRestart
				ConsoleWriteError(@CRLF & "Error processing " & $inputAllFile & @CRLF)
				ConsoleWriteError("quitting and restart" & @CRLF)

				$killCMD = "taskkill /im Cubasesx3.exe /t /f"
				RunWait(@ComSpec & " /c " & $killCMD)

				$echoError = "echo " & $inputAllFile & " >> logbadfiles.txt" ; too lazy to learn AutoIt FileWriting
				RunWait(@ComSpec & " /c " & $echoError)

				Run($cubaseProgramFile)
				WinWait("Cubase SX")

				WaitForCubaseToStart()

				Return

			EndIf
			Sleep($timeoutUiInteraction)
		WEnd

	Else
		ConsoleWrite(@CRLF &"Project File Exists => SKIP!" & @CRLF)
		ControlClick("Cubase SX", "", "Button2")
	EndIf

	If Not WaitForCubaseToStart(120, False) Then
		; DRY Violation CubaseRestart
		ConsoleWriteError(@CRLF & "Error processing " & $inputAllFile & @CRLF)
		ConsoleWriteError("quitting and restart" & @CRLF)

		$killCMD = "taskkill /im Cubasesx3.exe /t /f"
		RunWait(@ComSpec & " /c " & $killCMD)

		$echoError = "echo " & $inputAllFile & " >> logbadfiles.txt" ; too lazy to learn AutoIt FileWriting
		RunWait(@ComSpec & " /c " & $echoError)

		Run($cubaseProgramFile)
		WinWait("Cubase SX")

		WaitForCubaseToStart()
	EndIf

	Sleep($timeoutUiInteraction)
	Sleep($timeoutUiInteraction)

EndFunc

Func WaitForCubaseToStart($timeoutInSec = $timoutCubaseStartup, $displayMsg = True)

	$timer = TimerInit()

	While Not WinMenuSelectItem("[CLASS:Cubase SXFrameWindowClass]", "", "&Help", "&About Cubase SX")
		$timerDiff = TimerDiff($timer)

		If $timerDiff > $timeoutInSec * 1000 Then
			If $displayMsg Then
				MsgBox($MB_OK, "Start Cubase SX", "Please Start Cubase SX. When it has started, click OK. ")
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
