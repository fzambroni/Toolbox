#cs ----------------------------------------------------------------------------

	 AutoIt Version: 3.3.18.0
	 Author:         myName

	 Script Function:
		Template AutoIt script.

#ce ----------------------------------------------------------------------------

  Local $hFileOpen = FileOpen($sFilePath, $FO_APPEND)
    If $hFileOpen <>= -1 Then


    ; Write data to the file using the handle returned by FileOpen.
    FileWriteLine($hFileOpen, "Line 2")
EndIf