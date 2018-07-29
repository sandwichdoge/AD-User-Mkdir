;//Create directories based on username exported from ActiveDirectory.
;//sandwichdoge@gmail.com

#include <Array.au3>
#include <GuiConstantsEx.au3>

Global Const $DEBUG_MODE = False
Global Const $PROGRAM_NAME = "AD-User-MkDir"

Opt("TrayAutoPause", 0)
Opt("TrayIconHide", 1)


FileChangeDir(@ScriptDir)
$hGUI = GUICreate($PROGRAM_NAME, 350, 200, -1, -1, -1, 0x00000010)
$rCsv = GUICtrlCreateRadio("From csv:", 10, 10)
GUICtrlSetState(-1, $GUI_CHECKED)
$bExportCsv = GUICtrlCreateButton("Export users to CSV file", 80, 10, 125, 20)
GUICtrlSetTip(-1, "Export all users on this server's AD to a CSV file.")
$iCsvPath = GUICtrlCreateInput("Csv_file_path", 10, 36, 300, 22)
GUICtrlSetState(-1, $GUI_DROPACCEPTED)
$bCsvBrowse = GUICtrlCreateButton("...", 315, 35, 23, 23)
GUICtrlCreateGroup("Destination", 5, 115, 339, 46)
$iDest = GUICtrlCreateInput("", 10, 132, 300, 22)
$bDestBrowse = GUICtrlCreateButton("...", 315, 131, 23, 23)
$bStart = GUICtrlCreateButton("Make", 130, 170, 80, 25)
GUISetState()


$hGUIExportCsv = GUICreate("Export AD users to CSV", 300, 150)
GUICtrlCreateGroup("OU path", 10, 10, 280, 50)
$iOUPath = GUICtrlCreateInput("Acquiring domain name..", 20, 30, 260, 22)
GUICtrlCreateGroup("Destination", 10, 62, 280, 50)
$iExportPath = GUICtrlCreateInput("", 20, 82, 235, 22)
$bOUPathBrowse = GUICtrlCreateButton("...", 258, 81, 23, 23)
$bExport = GUICtrlCreateButton("Export", 120, 120, 50, 23)
;GUISetState()


While 1
	$msg = GUIGetMsg()
	Switch $msg
		Case $GUI_EVENT_DROPPED ;//this part removes the default csv_file_path field
			$sOldInputRead = GUICtrlRead($iCsvPath)
			$sOldInputRead = StringReplace($sOldInputRead, "Csv_file_path", "")
			GUICtrlSetData($iCsvPath, $sOldInputRead)
		Case $GUI_EVENT_CLOSE
			$hTemp = WinGetHandle("[ACTIVE]")
			Switch $hTemp
				Case $hGUI
					ExitLoop
				Case $hGUIExportCsv
					GUISetState(@SW_HIDE, $hTemp)
			EndSwitch
		Case $bExportCsv
			GUISetState(@SW_SHOW, $hGUIExportCsv)
			$sDomain = GetMachineDomainName()
			GUICtrlSetData($iOUPath, $sDomain & "/")
		Case $bCsvBrowse
			$sCsvPathRead = FileOpenDialog("Select exported CSV user file", @ScriptDir, "CSV files (*.csv)")
			If $sCsvPathRead Then GUICtrlSetData($iCsvPath, $sCsvPathRead)
		Case $bDestBrowse
			$sDestRead = FileSelectFolder("Select destination folder", @ScriptDir)
			If $sDestRead Then GUICtrlSetData($iDest, $sDestRead)
		Case $bStart
			GUICtrlSetState($bStart, $GUI_DISABLE)
			GUICtrlSetData($bStart, "Working")
			$sCsvPath = GUICtrlRead($iCsvPath)
			$sParentPath = GUICtrlRead($iDest)
			$nRet = MkDir($sCsvPath, $sParentPath) ;//nRet = count
			If $nRet > 0 Then
				MsgBox(64, $PROGRAM_NAME, "Finished!" & @CRLF & "Created " & $nRet & " folders in " & $sParentPath)
			Else
				Switch $nRet ;//now nRet is errno
					Case -1
						MsgBox(16, $PROGRAM_NAME, "Please select input paths.")
					Case -2
						MsgBox(16, $PROGRAM_NAME, "Could not parse CSV file.")
				EndSwitch
			EndIf
			GUICtrlSetState($bStart, $GUI_ENABLE)
			GUICtrlSetData($bStart, "Start")
		Case $bOUPathBrowse
			$sCsvExportPath = FileSaveDialog("Export AD users to CSV", @ScriptDir, "CSV files (*.csv)")
			If $sCsvExportPath Then GUICtrlSetData($iExportPath, $sCsvExportPath)
		Case $bExport
			$sCsvExportPath = GUICtrlRead($iExportPath)
			If Not $sCsvExportPath Then
				MsgBox(0, $PROGRAM_NAME, "Please select an output path.")
				ContinueLoop
			EndIf
			$sOUFilter = ConvertDirStringToADDir(GUICtrlRead($iOUPath))
			Switch $sOUFilter
				Case -1
					MsgBox(0, $PROGRAM_NAME, "Invalid OU input")
					ContinueLoop
				Case -2
					MsgBox(0, $PROGRAM_NAME, "Invalid domain name format")
					ContinueLoop
			EndSwitch
			$sExportCmd = 'powershell "Import-Module ActiveDirectory; Get-ADUser -Filter * ' & "-SearchBase '" & $sOUFilter & "'" & ' -Properties * | export-csv ' & $sCsvExportPath & '"'
			If $DEBUG_MODE Then MsgBox(0, 'cmd', $sExportCmd)
			$pid = Run(@ComSpec & " /c " & $sExportCmd, "", @SW_SHOW, 0x2)
			GUISetState(@SW_HIDE, $hGUIExportCsv)
			MsgBox(64, $PROGRAM_NAME, "Done!" & @CRLF & "CSV file saved to " & $sCsvExportPath)
	EndSwitch
WEnd



Func MkDir($sCsvPath, $sParentPath)
	Local $nRet = 0 ;//return code, is the count of total folders created
	Local $nUsernamePos = 15
	If Not $sCsvPath Or Not $sParentPath Then Return -1 ;//sanity check
	Local $sData = FileRead($sCsvPath)
	Local $aLines = StringSplit($sData, @CRLF, 3)
	If @error Then Return -2 ;
	
	If $DEBUG_MODE Then _ArrayDisplay($aLines)
	
	If MsgBox(4, $PROGRAM_NAME, "This will create " & UBound($aLines) - 3 & " items in " & $sParentPath & @CRLF & "Proceed?") = 7 Then Return 0
	;//start at 3 since first 2 lines are param template
	For $i = 2 To UBound($aLines) - 1
		If Not $aLines[$i] Then
			ContinueLoop
		EndIf
		$aParameters = StringSplit($aLines[$i], ",", 3)
		If @error Or UBound($aParameters) < $nUsernamePos Then ContinueLoop
		If $DEBUG_MODE Then _ArrayDisplay($aParameters)
		If $aParameters[$nUsernamePos] = '"0"' Then $nUsernamePos = 14 ;//sometimes it's the 14th but sometimes its 15th index
		$sUsername = StringReplace($aParameters[$nUsernamePos], '"', '')
		If DirCreate($sParentPath & "\" & $sUsername) = 0 Then
			MsgBox(0, "", "Error creating " & $sParentPath & "\" & $sUsername, 10)
		Else
			$nRet += 1
		EndIf
	Next
	
	Return $nRet
EndFunc   ;==>MkDir


Func ConvertDirStringToADDir($sDirString) ;//"child.robo17.local/Robo_HCM/Phong Ke Toan" -> OU=Phong Ke Toan,OU=Robo_HCM,DC=child,DC=robo17,DC=local
	$aItems = StringSplit($sDirString, "/", 3)
	If @error Then Return -1 ;//invalid input
	If Not StringInStr($aItems[0], ".") Then Return -2 ;//invalid domain name format
	
	Local $sRet, $sDC
	
	;//Append OU items
	If $aItems[1] Then ;//OU specified
		For $i = UBound($aItems) - 1 To 1 Step -1
			$sRet &= "OU=" & $aItems[$i] & ","
		Next
	EndIf
	
	;//Append DC items
	$aDomain = StringSplit($aItems[0], ".", 3) ;//Process domain string
	For $i = 0 To UBound($aDomain) - 1
		$sDC &= "DC=" & $aDomain[$i] & ","
	Next
	$sDC = StringTrimRight($sDC, 1) ;//Remove trailing comma
	
	$sRet &= $sDC ;//Append OU and DC
	
	If $DEBUG_MODE Then MsgBox(0, '', $sRet)
	Return $sRet
EndFunc   ;==>ConvertDirStringToADDir


Func GetMachineDomainName()
	$sRaw = ShellExecuteReturn('systeminfo | findstr /B /C:"Domain"')
	$aReg = StringRegExp($sRaw, "Domain:(?:\h*)(.+)", 3)
	If Not @error Then 
		Return $aReg[0]
	Else
		Return "Could not acquire domain name. Please enter manually."
	EndIf	
EndFunc   ;==>GetMachineDomainName


Func ShellExecuteReturn($sInput)
	$pid = Run(@ComSpec & " /c " & $sInput, "", @SW_HIDE, 0x2)
	$sOut = ProcessWaitClose($pid, 5)
	$sOut = StdoutRead($pid)
	Return $sOut
EndFunc   ;==>ShellExecuteReturn
