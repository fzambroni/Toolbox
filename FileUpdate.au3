#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=Toolbox.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#cs ----------------------------------------------------------------------------

	 AutoIt Version: 3.3.18.0
	 Author:         myName

	 Script Function:
		Template AutoIt script.

#ce ----------------------------------------------------------------------------

$sFilePath = @ScriptDir & "\version.txt"
$sExecPath = @ScriptDir & "\Toolbox.exe"
Local $hFileOpen = FileOpen($sFilePath, 10)
If $hFileOpen <> -1 Then
	; Write data to the file using the handle returned by FileOpen.
	FileWrite($hFileOpen, FileGetVersion($sExecPath))
EndIf
