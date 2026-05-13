#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=Toolbox.ico
#AutoIt3Wrapper_Res_Description=Updater
#AutoIt3Wrapper_Res_Fileversion=1.0.0.3
#AutoIt3Wrapper_Res_ProductName=Toolbox Updater
#AutoIt3Wrapper_Res_File_Add=E:\GitHub\Toolbox\splash.jpg
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.16.1
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------
#pragma compile(inputboxres, true)

Opt("TrayIconHide", 1)
Opt("TrayAutoPause", 0)
#include <WindowsStylesConstants.au3>
#include <StaticConstants.au3>

;####################################################
;####################################################
$AppName = "Toolbox"
;####################################################
;####################################################

$sSplashPath = @TempDir & "\splash.jpg"
FileInstall("splash.jpg", $sSplashPath, 1)
Sleep(1000)
If $CmdLine[0] >= 1 Then
	$Path = $CmdLine[1]
Else
	$Path = StringReplace(StringReplace($CmdLineRaw,"'", ""), '"', "")
EndIf

;~ MsgBox(262144,"",$Path & "\" & $AppName & ".tmp")
;~ _splash()
;~ Sleep(5000)

If Not FileExists($Path & "\" & $AppName & ".tmp") Then
	Exit
Else
	_splash()
	Sleep(3000)
	FileMove($Path & "\" & $AppName & ".tmp",$Path & "\" & $AppName & ".exe",9)
	Sleep(2000)
	Run('"' & $Path & "\" & $AppName & ".exe" & '"')
EndIf
Exit

Func _splash()

	$splashWin_X = 640
	$splashWin_Y = 360

	Global $Form_Splash = GUICreate("", $splashWin_X, $splashWin_Y, -1, -1, $WS_POPUP, BitOR($WS_EX_TOPMOST, $WS_EX_TOOLWINDOW, $WS_EX_LAYERED))


	Global $Pic_Splash = GUICtrlCreatePic($sSplashPath, 5, 5, 630, 350)

	Global $Label_Percentage = GUICtrlCreateLabel("Updating " & $AppName & " . . .", 5, 290, 630, 30, $SS_CENTER)
	GUICtrlSetFont($Label_Percentage, 17)
	GUICtrlSetColor($Label_Percentage, 0xFF0000)

	GUISetState(@SW_SHOW, $Form_Splash)


EndFunc   ;==>_splash
