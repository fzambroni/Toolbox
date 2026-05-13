#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_UseUpx=n
#AutoIt3Wrapper_Res_CompanyName=Fabricio Zambroni
#AutoIt3Wrapper_Res_Fileversion=1.1.5.1
#AutoIt3Wrapper_Res_ProductVersion=1.1.1.1
#AutoIt3Wrapper_Res_LegalCopyright=Copyright © 2026 Fabricio Zambroni
#AutoIt3Wrapper_Icon=Toolbox.ico
#AutoIt3Wrapper_Res_Description=Ortems SQL Toolbox
#AutoIt3Wrapper_Res_ProductName=Ortems SQL Toolbox
#AutoIt3Wrapper_Res_File_Add=E:\GitHub\Toolbox\Updater.exe
#AutoIt3Wrapper_Res_File_Add=E:\GitHub\Toolbox\Help.html


#NoTrayIcon
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <StructureConstants.au3>
#include <TabConstants.au3>
#include <GuiTab.au3>
#include <ListViewConstants.au3>
#include <GuiListView.au3>
#include <GuiComboBox.au3>
#include <ComboConstants.au3>
#include <Array.au3>
#include <File.au3>
#include <String.au3>
#include <EditConstants.au3>
#include <WindowsStylesConstants.au3>
#include <StaticConstants.au3>
#include <InetConstants.au3>


;=============================================================================
; ORTEMS TOOLBOX - AutoIt Interface
; Replaces the original Excel workbook Toolbox_v2.0.1.xlsm
; Target database: SQL Server
;=============================================================================

Opt("GUIOnEventMode", 1)
Opt("MustDeclareVars", 0)

Global $DevBy = "Developed by Fabricio Zambroni"

; ----------------------------------------------------------------------------------------------------------------------
; Updater - GitHub based
; ----------------------------------------------------------------------------------------------------------------------
; This updater no longer uses a shared network folder. It reads the latest version from GitHub version.txt,
; downloads the published Toolbox.exe only when a newer version is available, then lets Updater.exe replace the file.
Global Const $g_sGitHubDefaultRawBase = "https://raw.githubusercontent.com/fzambroni/Toolbox/main"
Global $g_sGitHubRawBase = IniRead(@ScriptDir & "\settings.ini", "Update", "github_raw_base", $g_sGitHubDefaultRawBase)
If StringStripWS($g_sGitHubRawBase, 3) = "" Then $g_sGitHubRawBase = $g_sGitHubDefaultRawBase

; Keep settings.ini explicit and self-documenting. The old [Update] path entry is intentionally ignored.
IniWrite(@ScriptDir & "\settings.ini", "Update", "source", "github")
IniWrite(@ScriptDir & "\settings.ini", "Update", "github_raw_base", $g_sGitHubRawBase)

; Skip automatic update checks when running the .au3 directly from SciTE/dev mode.
If Not StringInStr(StringLower(@ScriptName), ".au3") Then
    _CheckGitHubUpdate()
EndIf

; === Window sizing ===
Global Const $MIN_W = 1050
Global Const $MIN_H = 700
Global Const $BOTTOM_BAR_H = 70
Global Const $TAB_TOP_Y = 58

; === GLOBAL CONSTANTS ===
Global Const $TITLE       = "Ortems Toolbox"
Global Const $APP_WIDTH   = 1100
Global Const $APP_HEIGHT  = 720
Global Const $COL_W       = 140
Global Const $ORTEMS_DEFAULT_NOM_ENC = "ENCOURS" ; default Ortems en-cours/plan key used by E_OF.NOM_ENC and B_BT.NOM_ENC

; === GLOBAL STATE VARIABLES ===
Global $g_sServer   = ""
Global $g_sDatabase = ""
Global $g_sConnStr  = ""
Global $g_bConnected = False

; Selected modules
Global $g_bModPS    = True    ; Production Scheduling (PS)
Global $g_bModMP    = False   ; Master Planning (MP)
Global $g_bModSRP   = False
Global $g_bModWOL   = False   ; WO Links
Global $g_bModSR    = False   ; Secondary Resources
Global $g_bModINV   = False   ; Inventory Movements
Global $g_bModMRK   = False   ; Markers
Global $g_bModLR    = False   ; Limited Resources
Global $g_bModPRM   = False   ; Parameters
Global $g_bModCPRM  = False   ; Continuous Parameters
Global $g_bModBATCH = False   ; Batch Machines
Global $g_bModCROUT = False   ; Complex Routings

; In-memory data (array of arrays)
Global $g_aCals[0][7]      ; Calendars
Global $g_aMach[0][13]     ; Machines
Global $g_aOps[0][11]      ; Operations
Global $g_aRout[0][5]      ; Routings
Global $g_aMat[0][6]       ; Materials
Global $g_aBOM[0][7]       ; BOM
Global $g_aWO[0][6]        ; Work Orders
Global $g_aWOL[0][7]       ; WO Links
Global $g_aSR[0][5]        ; Secondary Resources
Global $g_aCap[0][7]       ; Capacity Calendars
Global $g_aStk[0][5]       ; Stock Movements
Global $g_aMPS[0][3]       ; MPS Buckets
Global $g_aFcst[0][4]      ; Forecasts
Global $g_aSO[0][6]        ; Sales Orders

; Main control handles
Global $g_hMain, $g_hTab
Global $g_hTabHandle = 0
Global $g_idBottomLine = 0
Global $g_lblFooter = 0
Global $g_hLV_Cal, $g_hLV_Mach, $g_hLV_Ops, $g_hLV_Rout
Global $g_hLV_Mat, $g_hLV_BOM, $g_hLV_WO, $g_hLV_WOL
Global $g_hLV_SR, $g_hLV_Cap, $g_hLV_Stk
Global $g_hLog
Global $g_sIntegrityReport = ""
Global $g_sLogDir = @ScriptDir & "\log"
Global $g_sLogFile = ""
Global $g_bVerbose = False
Global $g_bFullCleanImport = False
Global $g_iBusyDepth = 0
Global $g_aBusyControls[0]
Global $g_aBusyStates[0]
Global $g_sBusyFooterBackup = ""
Global $g_idTabSQLItem = 0
; Tab enable/disable state (Tab control doesn't truly disable items; we block selection via WM_NOTIFY)
Global $g_aTabEnabled[0]
Global $g_aTabBaseText[0]
Global $g_bAllowProgrammaticTabChange = False

; ListView sorting state - one click ascending, second click descending
Global Const $ORTEMS_LVN_COLUMNCLICK = -108
Global $g_aLVSortState[0][3] ; [ListView control ID, last column, descending flag]

; Settings persistence
Global Const $g_sIniFile = @ScriptDir & "\settings.ini"

; COM error handler (catches $oConn.Execute failures instead of crashing)
Global $g_oComErr = ObjEvent("AutoIt.Error", "_ComErrorHandler")
Global $g_sLastComError = ""

Func _ComErrorHandler()
    Local $oErr = $g_oComErr
    $g_sLastComError = "COM Error 0x" & Hex($oErr.number, 8) & ": " & $oErr.description
EndFunc

Func _CheckGitHubUpdate()
    Local $sCurrentVersion = FileGetVersion(@ScriptFullPath)
    If StringStripWS($sCurrentVersion, 3) = "" Then
        ConsoleWrite("GitHub update check skipped: local file version could not be read." & @CRLF)
        Return
    EndIf

    Local $sAppName = _FileNameWithoutExtension(@ScriptName)
    Local $sRemoteVersionUrl = _JoinUrl($g_sGitHubRawBase, "version.txt")
    Local $sRemoteExeUrl = _JoinUrl($g_sGitHubRawBase, @ScriptName)
    Local $sRemoteVersionTmp = @TempDir & "\" & $sAppName & "_github_version.txt"
    Local $sRemoteExeTmp = @TempDir & "\" & $sAppName & "_github_latest.exe"
    Local $sLocalTmp = @ScriptDir & "\" & $sAppName & ".tmp"
    Local $sUpdaterFile = @TempDir & "\Updater.exe"

    ConsoleWrite("Checking for updates from GitHub version file: " & $sRemoteVersionUrl & @CRLF)

    Local $sRemoteVersion = _GetGitHubVersionFromTextFile($sRemoteVersionUrl, $sRemoteVersionTmp)
    If StringStripWS($sRemoteVersion, 3) = "" Then
        ConsoleWrite("GitHub update check skipped: remote version.txt could not be read." & @CRLF)
        FileDelete($sRemoteVersionTmp)
        Return
    EndIf

    ConsoleWrite("Local version: " & $sCurrentVersion & " | GitHub version: " & $sRemoteVersion & @CRLF)

    If _CompareVersions($sRemoteVersion, $sCurrentVersion) <= 0 Then
        ConsoleWrite("No update required." & @CRLF)
        FileDelete($sRemoteVersionTmp)
        Return
    EndIf

    ConsoleWrite("Newer GitHub version found. Downloading: " & $sRemoteExeUrl & @CRLF)

    If Not _DownloadFile($sRemoteExeUrl, $sRemoteExeTmp) Then
        FileDelete($sRemoteVersionTmp)
        Return
    EndIf

    ; Validate that the executable we are about to install matches the version announced in version.txt.
    ; This prevents installing an older exe when version.txt was updated before Toolbox.exe was published.
    Local $sDownloadedExeVersion = FileGetVersion($sRemoteExeTmp)
    If StringStripWS($sDownloadedExeVersion, 3) = "" Then
        ConsoleWrite("Update aborted: downloaded executable version could not be read." & @CRLF)
        FileDelete($sRemoteVersionTmp)
        FileDelete($sRemoteExeTmp)
        Return
    EndIf

    If _CompareVersions($sDownloadedExeVersion, $sRemoteVersion) < 0 Then
        ConsoleWrite("Update aborted: downloaded executable version is older than version.txt. Downloaded=" & $sDownloadedExeVersion & ", version.txt=" & $sRemoteVersion & @CRLF)
        FileDelete($sRemoteVersionTmp)
        FileDelete($sRemoteExeTmp)
        Return
    EndIf

    If _CompareVersions($sDownloadedExeVersion, $sCurrentVersion) <= 0 Then
        ConsoleWrite("Update aborted: downloaded executable version is not newer. Downloaded=" & $sDownloadedExeVersion & ", Local=" & $sCurrentVersion & @CRLF)
        FileDelete($sRemoteVersionTmp)
        FileDelete($sRemoteExeTmp)
        Return
    EndIf

    FileDelete($sLocalTmp)
    If Not FileCopy($sRemoteExeTmp, $sLocalTmp, 9) Then
        ConsoleWrite("Update aborted: could not stage downloaded file at " & $sLocalTmp & @CRLF)
        FileDelete($sRemoteVersionTmp)
        FileDelete($sRemoteExeTmp)
        Return
    EndIf

    ConsoleWrite("Update staged at: " & $sLocalTmp & @CRLF)
    FileInstall("Updater.exe", $sUpdaterFile, 1)
    Sleep(500)
    Run($sUpdaterFile & " '" & @ScriptDir & "'")
    Sleep(100)
    Exit
EndFunc

Func _GetGitHubVersionFromTextFile($sUrl, $sDestination)
    FileDelete($sDestination)
    If Not _DownloadFile($sUrl, $sDestination) Then Return ""

    Local $sContent = FileRead($sDestination)
    If @error Then Return ""

    $sContent = StringStripWS($sContent, 3)

    ; version.txt must contain only the version number, for example: 1.1.5.0
    Local $aMatch = StringRegExp($sContent, "^([0-9]+(?:\.[0-9]+){1,3})$", 1)
    If @error Or UBound($aMatch) = 0 Then
        ConsoleWrite("Invalid version.txt content: " & $sContent & @CRLF)
        Return ""
    EndIf

    Return $aMatch[0]
EndFunc

Func _DownloadFile($sUrl, $sDestination)
    FileDelete($sDestination)
    Local $hDownload = InetGet($sUrl, $sDestination, $INET_FORCERELOAD, $INET_DOWNLOADWAIT)
    If @error Or $hDownload = 0 Or Not FileExists($sDestination) Or FileGetSize($sDestination) <= 0 Then
        ConsoleWrite("Download failed: " & $sUrl & @CRLF)
        Return False
    EndIf
    Return True
EndFunc

Func _JoinUrl($sBase, $sFile)
    Local $sCleanBase = StringStripWS($sBase, 3)
    While StringRight($sCleanBase, 1) = "/"
        $sCleanBase = StringTrimRight($sCleanBase, 1)
    WEnd
    Return $sCleanBase & "/" & $sFile
EndFunc

Func _FileNameWithoutExtension($sFileName)
    Local $iDot = StringInStr($sFileName, ".", 0, -1)
    If $iDot <= 1 Then Return $sFileName
    Return StringLeft($sFileName, $iDot - 1)
EndFunc

Func _CompareVersions($sLeft, $sRight)
    Local $aLeft = StringSplit(StringStripWS($sLeft, 3), ".")
    Local $aRight = StringSplit(StringStripWS($sRight, 3), ".")
    Local $iMax = $aLeft[0]
    If $aRight[0] > $iMax Then $iMax = $aRight[0]

    For $i = 1 To $iMax
        Local $nLeft = 0
        Local $nRight = 0
        If $i <= $aLeft[0] Then $nLeft = Number($aLeft[$i])
        If $i <= $aRight[0] Then $nRight = Number($aRight[$i])

        If $nLeft > $nRight Then Return 1
        If $nLeft < $nRight Then Return -1
    Next

    Return 0
EndFunc

Func _GetFooterText()
    Local $sVersion = FileGetVersion(@ScriptFullPath)
    If StringStripWS($sVersion, 3) = "" Then $sVersion = "N/A"
    Return "Version: " & $sVersion & " | " & $TITLE & " | " & @YEAR
EndFunc


; Tab indexes
Global $g_iTabDB = 0, $g_iTabMod = 1
Global $g_iTabCal = 2, $g_iTabMach = 3, $g_iTabOps = 4
Global $g_iTabRout = 5, $g_iTabMat = 6, $g_iTabBOM = 7
Global $g_iTabWO = 8, $g_iTabWOL = 9, $g_iTabSR = 10
Global $g_iTabCap = 11, $g_iTabStk = 12, $g_iTabSQL = 13

;=============================================================================
; ENTRY POINT
;=============================================================================
Main()

Func Main()
    CreateMainWindow()
    _LoadSettings()
    GUISetState(@SW_SHOW, $g_hMain)
    _LogSessionStart()
    _RefreshModuleFlags(True)
    _OnResize()

    While 1
        Sleep(100)
    WEnd
EndFunc

;=============================================================================
; MAIN WINDOW CREATION
;=============================================================================
Func CreateMainWindow()
    $g_hMain = GUICreate($TITLE, $APP_WIDTH, $APP_HEIGHT, -1, -1, _
        BitOR($GUI_SS_DEFAULT_GUI, $WS_SIZEBOX, $WS_MAXIMIZEBOX))
    GUISetOnEvent($GUI_EVENT_CLOSE, "_OnClose")
    GUISetOnEvent($GUI_EVENT_RESIZED, "_OnResize")
    GUIRegisterMsg($WM_GETMINMAXINFO, "_WM_GETMINMAXINFO")
    GUIRegisterMsg($WM_NOTIFY, "_WM_NOTIFY")

    ; Barra de titulo/logo
    GUICtrlCreateLabel("ORTEMS TOOLBOX", 10, 8, 170, 26)
    GUICtrlSetFont(-1, 14, 800, 0, "Segoe UI")
    GUICtrlSetColor(-1, 0x003366)

    GUICtrlCreateLabel("BETA VERSION", 188, 11, 130, 20)
    GUICtrlSetFont(-1, 10, 800, 0, "Segoe UI")
    GUICtrlSetColor(-1, 0xCC0000)

    GUICtrlCreateLabel("Validated Ortems demo-data builder for SQL Server", 10, 34, 500, 18)
    GUICtrlSetFont(-1, 9, 400, 2, "Segoe UI")
    GUICtrlSetColor(-1, 0x666666)

    ; Connection status
    GUICtrlCreateLabel("Status:", 650, 10, 50, 18)
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    Global $g_lblStatus = GUICtrlCreateLabel("Disconnected", 705, 10, 345, 18);,$SS_BLACKRECT)
    GUICtrlSetFont(-1, 9, 700, 0, "Segoe UI")
    GUICtrlSetColor($g_lblStatus, 0xCC0000)

    Global $g_btnHelp = GUICtrlCreateButton("?", $APP_WIDTH - 42, 8, 40, 40)
    GUICtrlSetOnEvent($g_btnHelp, "_Help_Click")
    GUICtrlSetFont(-1, 12, 800, 0, "Segoe UI")
    GUICtrlSetTip(-1,"Help")

    ; === ABAS PRINCIPAIS ===
    $g_hTab = GUICtrlCreateTab(5, $TAB_TOP_Y, $APP_WIDTH - 10, $APP_HEIGHT - 130)
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    $g_hTabHandle = GUICtrlGetHandle($g_hTab)

    ; ---- TAB 1: DATABASE ----
    GUICtrlCreateTabItem("1. Database")
    _CreateTabDatabase()

    ; ---- TAB 2: MODULES ----
    GUICtrlCreateTabItem("2. Modules")
    _CreateTabModules()

    ; ---- TAB 3: CALENDARS ----
    GUICtrlCreateTabItem("3. Calendars")
    _CreateTabCalendars()

    ; ---- TAB 4: MACHINES ----
    GUICtrlCreateTabItem("4. Machines")
    _CreateTabMachines()

    ; ---- TAB 5: OPERATIONS ----
    GUICtrlCreateTabItem("5. Operations")
    _CreateTabOperations()

    ; ---- TAB 6: ROUTINGS ----
    GUICtrlCreateTabItem("6. Routings")
    _CreateTabRoutings()

    ; ---- TAB 7: ITEMS ----
    GUICtrlCreateTabItem("7. Items")
    _CreateTabMaterials()

    ; ---- TAB 8: BOM ----
    GUICtrlCreateTabItem("8. BOM")
    _CreateTabBOM()

    ; ---- TAB 9: WORK ORDERS ----
    GUICtrlCreateTabItem("9. Work Orders (WO)")
    _CreateTabWO()

    ; ---- TAB 10: WO LINKS ----
    GUICtrlCreateTabItem("10. WO Links")
    _CreateTabWOLinks()

    ; ---- TAB 11: SECONDARY RESOURCES ----
    GUICtrlCreateTabItem("11. Secondary Resources")
    _CreateTabSecResources()

    ; ---- TAB 12: CAPACITY ----
    GUICtrlCreateTabItem("12. Capacity")
    _CreateTabCapacity()

    ; ---- TAB 13: INVENTORY ----
    GUICtrlCreateTabItem("13. Inventory Movements")
    _CreateTabStockMov()

    ; ---- TAB 14: GENERATE SQL ----
    $g_idTabSQLItem = GUICtrlCreateTabItem("14. Generate & Run SQL")
    _CreateTabSQL()

    GUICtrlCreateTabItem("")

    ; === BOTTOM ACTION BAR ===
    $g_idBottomLine = GUICtrlCreateLabel("", 0, $APP_HEIGHT - 65, $APP_WIDTH, 2)
    GUICtrlSetBkColor(-1, 0xCCCCCC)

    Global $g_btnImportXLS = GUICtrlCreateButton("Import Workbook...", 10, $APP_HEIGHT - 58, 132, 32)
    GUICtrlSetOnEvent($g_btnImportXLS, "_ImportExcel_Click")
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    GUICtrlSetTip($g_btnImportXLS, "Import a round-trip Excel workbook exported by this tool")
    _RegisterBusyControl($g_btnImportXLS)

    Global $g_btnExportXLS = GUICtrlCreateButton("Export Workbook...", 150, $APP_HEIGHT - 58, 132, 32)
    GUICtrlSetOnEvent($g_btnExportXLS, "_ExportExcel_Click")
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    GUICtrlSetTip($g_btnExportXLS, "Export all tabs to a structured Excel workbook that can be imported back")
    _RegisterBusyControl($g_btnExportXLS)

    Global $g_btnLoadDB = GUICtrlCreateButton("Load from DB", 290, $APP_HEIGHT - 58, 122, 32)
    GUICtrlSetOnEvent($g_btnLoadDB, "_LoadFromDB_Click")
    GUICtrlSetFont(-1, 9, 700, 0, "Segoe UI")
    GUICtrlSetBkColor($g_btnLoadDB, 0xE8F4FD)
    GUICtrlSetTip($g_btnLoadDB, "Read data from the connected Ortems database and populate all tabs")
    _RegisterBusyControl($g_btnLoadDB)

    Global $g_btnClearDB = GUICtrlCreateButton("Clear DB", 420, $APP_HEIGHT - 58, 92, 32)
    GUICtrlSetOnEvent($g_btnClearDB, "_ClearDatabase_Click")
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    GUICtrlSetBkColor($g_btnClearDB, 0xFFDDDD)
    _RegisterBusyControl($g_btnClearDB)

    Global $g_btnGenerate = GUICtrlCreateButton("Generate SQL", 520, $APP_HEIGHT - 58, 118, 32)
    GUICtrlSetOnEvent($g_btnGenerate, "_GenerateSQL_Click")
    GUICtrlSetFont(-1, 9, 700, 0, "Segoe UI")
    _RegisterBusyControl($g_btnGenerate)

    Global $g_btnExecute = GUICtrlCreateButton("Run on DB", 646, $APP_HEIGHT - 58, 115, 32)
    GUICtrlSetOnEvent($g_btnExecute, "_ExecuteSQL_Click")
    GUICtrlSetFont(-1, 9, 700, 0, "Segoe UI")
    GUICtrlSetBkColor($g_btnExecute, 0xDDFFDD)
    _RegisterBusyControl($g_btnExecute)

    Global $g_btnSaveSQL = GUICtrlCreateButton("Save SQL...", 769, $APP_HEIGHT - 58, 105, 32)
    GUICtrlSetOnEvent($g_btnSaveSQL, "_SaveSQL_Click")
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    _RegisterBusyControl($g_btnSaveSQL)

	Global $g_btnIntegrity = GUICtrlCreateButton("Integrity Check", 882, $APP_HEIGHT - 58, 125, 32)
    GUICtrlSetOnEvent($g_btnIntegrity, "_IntegrityCheck_Click")
    GUICtrlSetFont(-1, 9, 700, 0, "Segoe UI")
    GUICtrlSetBkColor($g_btnIntegrity, 0xFFF2CC)
    GUICtrlSetTip($g_btnIntegrity, "Validate cross-tab references before generating or running the SQL")
    _RegisterBusyControl($g_btnIntegrity)

    $g_lblFooter = GUICtrlCreateLabel($DevBy, 715, $APP_HEIGHT - 22, 390, 20);,$SS_GRAYFRAME)
	GUICtrlSetFont(-1, 7, 400, 0, "Segoe UI")
    GUICtrlSetColor(-1, 0x777777)
    $g_lblFooter2 = GUICtrlCreateLabel(_GetFooterText(), 15, $APP_HEIGHT - 22, 390, 20);,$SS_GRAYFRAME)
    GUICtrlSetFont(-1, 7, 400, 0, "Segoe UI")
    GUICtrlSetColor(-1, 0x777777)
    _InitTabState()
EndFunc

;=============================================================================
; TAB: DATABASE
;=============================================================================
Func _CreateTabDatabase()
    Local $y = 85, $xL = 20, $xV = 200

    GUICtrlCreateGroup("SQL Server Connection", $xL, $y, 620, 225)
    $y += 28

    GUICtrlCreateLabel("Server / Instance:", $xL + 10, $y, 170, 20)
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    Global $g_edtServer = GUICtrlCreateInput("localhost\SQLEXPRESS", $xV, $y, 220, 22)
    _RegisterBusyControl($g_edtServer)
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    $y += 32

    GUICtrlCreateLabel("Database name:", $xL + 10, $y, 170, 20)
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    Global $g_edtDatabase = GUICtrlCreateInput("ORTEMS_DEMO", $xV, $y, 220, 22)
    _RegisterBusyControl($g_edtDatabase)
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    $y += 32

    GUICtrlCreateLabel("Authentication:", $xL + 10, $y, 170, 20)
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    Global $g_cmbAuth = GUICtrlCreateCombo("Windows Authentication", $xV, $y, 220, 22)
    _RegisterBusyControl($g_cmbAuth)
    GUICtrlSetData($g_cmbAuth, "SQL Server Authentication")
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    GUICtrlSetOnEvent($g_cmbAuth, "_OnAuthChange")
    $y += 32

    GUICtrlCreateLabel("User:", $xL + 10, $y, 170, 20)
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    Global $g_edtUser = GUICtrlCreateInput("", $xV, $y, 220, 22)
    _RegisterBusyControl($g_edtUser)
    GUICtrlSetState($g_edtUser, $GUI_DISABLE)
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    $y += 32

    GUICtrlCreateLabel("Password:", $xL + 10, $y, 170, 20)
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    Global $g_edtPass = GUICtrlCreateInput("", $xV, $y, 220, 22, $ES_PASSWORD)
    _RegisterBusyControl($g_edtPass)
    GUICtrlSetState($g_edtPass, $GUI_DISABLE)
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    $y += 38

    Global $g_btnConnect = GUICtrlCreateButton("DB Connect", $xV, $y, 160, 30)
    GUICtrlSetOnEvent($g_btnConnect, "_TestConnection_Click")
    GUICtrlSetFont(-1, 9, 700, 0, "Segoe UI")

    Global $g_btnInspect = GUICtrlCreateButton("Inspect Table...", $xV + 170, $y, 140, 30)
    GUICtrlSetOnEvent($g_btnInspect, "_InspectTable_Click")
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    GUICtrlSetTip($g_btnInspect, "Show columns of any Ortems table (useful to verify column names)")
    _RegisterBusyControl($g_btnConnect)
    _RegisterBusyControl($g_btnInspect)

    ; Connection string display
    $y = 330
    GUICtrlCreateGroup("Active connection string", $xL, $y, 620, 80)
    Global $g_edtConnStr = GUICtrlCreateEdit("", $xL + 10, $y + 22, 590, 48, $ES_READONLY + $WS_VSCROLL)
    GUICtrlSetFont(-1, 8, 400, 0, "Courier New")
    GUICtrlSetBkColor($g_edtConnStr, 0xF5F5F5)

    ; Instructions
    $y = 430
    GUICtrlCreateGroup("How to use the Toolbox", $xL, $y, 720, 170)
    $y += 22
    Local $sInfo = "WORKFLOW:" & @CRLF & _
        "  1. Configure the connection to the Ortems database (above) and click 'DB Connect'" & @CRLF & _
        "  2. Go to the '2. Modules' tab and select the required Ortems modules for the demo" & @CRLF & _
        "  3. Fill in the data in the tabs (Calendars, Machines, Operations, Routings, Items, etc.)" & @CRLF & _
        "  4. Click 'Generate SQL' in the bottom toolbar to build the SQL script" & @CRLF & _
        "  5. Click 'Run on DB' to execute the SQL and insert data into the Ortems database" & @CRLF & _
        "  Tip: Settings (server, modules) are automatically saved to settings.ini"
    GUICtrlCreateEdit($sInfo, $xL + 10, $y, 700, 130, $ES_READONLY + $WS_VSCROLL)
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    GUICtrlSetBkColor(-1, 0xFFFFF0)
EndFunc

;=============================================================================
; TAB: MODULE SELECTION
;=============================================================================
Func _CreateTabModules()
    GUICtrlCreateLabel("Ortems Module Selection", 20, 80, 500, 22)
    GUICtrlSetFont(-1, 12, 700, 0, "Segoe UI")
    GUICtrlSetColor(-1, 0x003366)

    GUICtrlCreateLabel("Answer the questions below to configure the demo modules:", 20, 105, 700, 18)
    GUICtrlSetFont(-1, 9, 400, 2, "Segoe UI")

    Local $y = 135, $xL = 20
    GUICtrlCreateGroup("Module Configuration", $xL, $y, 900, 520)
    $y += 20

    ; Q1: PS vs MP
    _CreateModuleQuestion($y, "Q1:", "Does the customer need continuous operation scheduling (PS) or", "bucket planning with load leveling (MP)?")
    Global $g_rbPS = GUICtrlCreateRadio("PS - Production Scheduling (continuous scheduling)", $xL + 30, $y + 40, 420, 20)
    GUICtrlSetOnEvent($g_rbPS, "_OnModuleSelectionChanged")
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    GUICtrlSetState($g_rbPS, $GUI_CHECKED)
    _RegisterBusyControl($g_rbPS)
    Global $g_rbMP = GUICtrlCreateRadio("MP - Master Planning (bucket planning)", $xL + 30, $y + 62, 420, 20)
    GUICtrlSetOnEvent($g_rbMP, "_OnModuleSelectionChanged")
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    _RegisterBusyControl($g_rbMP)
    $y += 92

    ; Q2: SRP
    Global $g_chkSRP = _CreateModuleCheckbox($y, $xL, "Q2:", "Automatic synchronization across BOM levels? (SRP module)")
    $y += 38

    ; Q3: WO Links
    Global $g_chkWOL = _CreateModuleCheckbox($y, $xL, "Q3:", "Pre-defined precedence constraints between WOs/operations? (WO Links)")
    $y += 38

    ; Q4: Complex Routings
    Global $g_chkCROUT = _CreateModuleCheckbox($y, $xL, "Q4:", "Complex, non-linear routings, or start-start synchronization? (Complex routings)")
    $y += 38

    ; Q5: Secondary Resources
    Global $g_chkSR = _CreateModuleCheckbox($y, $xL, "Q5:", "Labor/tool constraints? Multi-resource optimization? (Secondary resources)")
    $y += 38

    ; Q6: Batch Machines
    Global $g_chkBATCH = _CreateModuleCheckbox($y, $xL, "Q6:", "Batch machines where operations can be grouped (e.g., oven)? (Batch machines)")
    $y += 38

    ; Q7: Inventory
    Global $g_chkINV = _CreateModuleCheckbox($y, $xL, "Q7:", "Show detailed inventory features (e.g., raw material replenishment)? (Inventory)")
    $y += 38

    ; Q8: Markers
    Global $g_chkMRK = _CreateModuleCheckbox($y, $xL, "Q8:", "Visual markers in planning for special events / collaboration? (Markers)")
    $y += 38

    ; Q9: Limited Resources
    Global $g_chkLR = _CreateModuleCheckbox($y, $xL, "Q9:", "Limited shared resources across multiple operations? (Limited resources)")
    $y += 38

    ; Q10: Parameters
    Global $g_chkPRM = _CreateModuleCheckbox($y, $xL, "Q10:", "Grouping by categories (color, temperature), changeovers between tools? (Parameters)")
    $y += 38

    ; Note: changes apply automatically on selection. Button below for confirmation summary.
    $y += 18
    Global $g_btnApplyMod = GUICtrlCreateButton("Apply / Show Summary", 20, $y, 200, 28)
    GUICtrlSetOnEvent($g_btnApplyMod, "_ApplyModules")
    GUICtrlSetFont(-1, 9, 700, 0, "Segoe UI")
    GUICtrlSetBkColor($g_btnApplyMod, 0xDDEEFF)
    _RegisterBusyControl($g_btnApplyMod)

    Global $g_lblModHint = GUICtrlCreateLabel("  OK  Tabs are updated automatically when you change a selection.", 230, $y + 6, 660, 18)
    GUICtrlSetFont(-1, 9, 400, 2, "Segoe UI")
    GUICtrlSetColor(-1, 0x007700)
EndFunc

Func _CreateModuleQuestion($y, $sNum, $sQ1, $sQ2)
    GUICtrlCreateLabel($sNum & " " & $sQ1, 30, $y, 700, 18)
    GUICtrlSetFont(-1, 9, 700, 0, "Segoe UI")
    If $sQ2 <> "" Then
        GUICtrlCreateLabel("    " & $sQ2, 30, $y + 18, 700, 18)
        GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    EndIf
EndFunc

Func _CreateModuleCheckbox($y, $xL, $sNum, $sText)
    GUICtrlCreateLabel($sNum, $xL + 10, $y + 2, 35, 18)
    GUICtrlSetFont(-1, 9, 700, 0, "Segoe UI")
    Local $hChk = GUICtrlCreateCheckbox($sText, $xL + 50, $y, 800, 20,-1);,$WS_EX_TOPMOST)
    GUICtrlSetOnEvent($hChk, "_OnModuleSelectionChanged")
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    Return _RegisterBusyControl($hChk)
EndFunc

;=============================================================================
; TAB: CALENDARS (SV_CALENDARS)
;=============================================================================
Func _CreateTabCalendars()
    Local $xL = 20
    GUICtrlCreateLabel("Work Calendars (SV_CALENDARS)", $xL, 80, 600, 20)
    GUICtrlSetFont(-1, 11, 700, 0, "Segoe UI")
    GUICtrlCreateLabel("A calendar defines working hours for machines or labor. Each row represents one shift slot.", $xL, 103, 900, 18)
    GUICtrlSetFont(-1, 8, 400, 2, "Segoe UI")
    GUICtrlCreateLabel("Days: 1=Mon, 2=Tue, 3=Wed, 4=Thu, 5=Fri, 6=Sat, 7=Sun   |   Times in HH:MM format", $xL, 123, 900, 18)
    GUICtrlSetFont(-1, 8, 400, 0, "Segoe UI")
    GUICtrlSetColor(-1, 0x555555)

    Local $y = 145
    _CreateCRUDButtons($y, "_Cal_Add", "_Cal_Edit", "_Cal_Del", "_Cal_DelAll", "_Cal_ImpCSV", "_Cal_ExpCSV", "_Cal_LoadEx")

    $g_hLV_Cal = GUICtrlCreateListView("Calendar ID|Calendar name|Start day|Start time|End day|End time|Line|", _
        $xL, $y + 35, $APP_WIDTH - 40, 400, _
        BitOR($LVS_REPORT, $LVS_SHOWSELALWAYS, $WS_BORDER))
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    _GUICtrlListView_SetColumnWidth($g_hLV_Cal, 0, 130)
    _GUICtrlListView_SetColumnWidth($g_hLV_Cal, 1, 150)
    _GUICtrlListView_SetColumnWidth($g_hLV_Cal, 2, 90)
    _GUICtrlListView_SetColumnWidth($g_hLV_Cal, 3, 90)
    _GUICtrlListView_SetColumnWidth($g_hLV_Cal, 4, 90)
    _GUICtrlListView_SetColumnWidth($g_hLV_Cal, 5, 90)
    _GUICtrlListView_SetColumnWidth($g_hLV_Cal, 6, 55)
EndFunc

Func _LoadExampleCalendars()
    ; Add standard examples
    _AddCalendar("Cal_1x8", "Cal 1x8", 1, "08:00", 1, "12:00")
    _AddCalendar("Cal_1x8", "Cal 1x8", 1, "14:00", 1, "17:00")
    _AddCalendar("Cal_1x8", "Cal 1x8", 2, "08:00", 2, "12:00")
    _AddCalendar("Cal_1x8", "Cal 1x8", 2, "14:00", 2, "17:00")
    _AddCalendar("Cal_2x8", "Cal 2x8", 1, "08:00", 2, "00:00")
EndFunc

Func _AddCalendar($sID, $sNome, $iDiaI, $sHoraI, $iDiaF, $sHoraF)
    Local $nRows = UBound($g_aCals, 1)
    ReDim $g_aCals[$nRows + 1][7]
    $g_aCals[$nRows][0] = $sID
    $g_aCals[$nRows][1] = $sNome
    $g_aCals[$nRows][2] = $iDiaI
    $g_aCals[$nRows][3] = $sHoraI
    $g_aCals[$nRows][4] = $iDiaF
    $g_aCals[$nRows][5] = $sHoraF
    _LV_AppendDataRow($g_hLV_Cal, $sID & "|" & $sNome & "|" & $iDiaI & "|" & $sHoraI & "|" & $iDiaF & "|" & $sHoraF)
EndFunc

;=============================================================================
; TAB: MACHINES (SV_MACHINE)
;=============================================================================
Func _CreateTabMachines()
    Local $xL = 20
    GUICtrlCreateLabel("Machines & Work Centers (SV_MACHINE)", $xL, 80, 600, 20)
    GUICtrlSetFont(-1, 11, 700, 0, "Segoe UI")
    GUICtrlCreateLabel("A machine is a resource that executes operations. Each machine belongs to a single work center (CT).", $xL, 103, 900, 18)
    GUICtrlSetFont(-1, 8, 400, 2, "Segoe UI")
    GUICtrlCreateLabel("CT type: 2=Finite capacity | 4=Infinite capacity   |   Machine type: NR=Standard, BA=Batch, RN=Run, CU=Tank", $xL, 123, 900, 18)
    GUICtrlSetFont(-1, 8, 400, 0, "Segoe UI")
    GUICtrlSetColor(-1, 0x555555)

    Local $y = 145
    _CreateCRUDButtons($y, "_Mach_Add", "_Mach_Edit", "_Mach_Del", "_Mach_DelAll", "_Mach_ImpCSV", "_Mach_ExpCSV", "_Mach_LoadEx")

    $g_hLV_Mach = GUICtrlCreateListView("Site ID|Site Name|CT ID|CT Name|CT Type|Section ID|Section Name|Machine ID|Machine Name|Mach Type|Cal ID|Cap Cal ID|Line|", _
        $xL, $y + 35, $APP_WIDTH - 40, 400, _
        BitOR($LVS_REPORT, $LVS_SHOWSELALWAYS, $WS_BORDER))
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    For $i = 0 To 11
        _GUICtrlListView_SetColumnWidth($g_hLV_Mach, $i, $COL_W - 30)
    Next
    _GUICtrlListView_SetColumnWidth($g_hLV_Mach, 12, 55)
EndFunc

Func _LoadExampleMachines()
    _AddMachine("LYON", "LYON", "Milling", "Milling", 2, "Sec_A", "Section A", "Milling_1", "Milling 1", "NR", "Cal_1x8", "Cal_1x8")
    _AddMachine("LYON", "LYON", "Milling", "Milling", 2, "Sec_A", "Section A", "Milling_2", "Milling 2", "NR", "Cal_1x8", "Cal_1x8")
    _AddMachine("LYON", "LYON", "Drilling", "Drilling", 2, "Sec_A", "Section A", "Drilling_1", "Drilling 1", "NR", "Cal_1x8", "Cal_1x8")
    _AddMachine("LYON", "LYON", "Assembly_Robot", "Robotics", 2, "Sec_B", "Section B", "Assy_Robot_1", "Assembly Robot 1", "NR", "Cal_2x8", "Cal_2x8")
EndFunc

Func _AddMachine($sSiteID, $sSiteNm, $sCTID, $sCTNm, $iCtTp, $sSecID, $sSecNm, $sMaqID, $sMaqNm, $sMaqTp, $sCalID, $sCalCap)
    Local $n = UBound($g_aMach, 1)
    ReDim $g_aMach[$n + 1][13]
    $g_aMach[$n][0] = $sSiteID ; ...store all fields
    _LV_AppendDataRow($g_hLV_Mach, $sSiteID & "|" & $sSiteNm & "|" & $sCTID & "|" & $sCTNm & "|" & $iCtTp & "|" & $sSecID & "|" & $sSecNm & "|" & $sMaqID & "|" & $sMaqNm & "|" & $sMaqTp & "|" & $sCalID & "|" & $sCalCap)
EndFunc

;=============================================================================
; TAB: OPERATIONS (SV_OPERATIONS)
;=============================================================================
Func _CreateTabOperations()
    Local $xL = 20
    GUICtrlCreateLabel("Operations (SV_OPERATIONS)", $xL, 80, 600, 20)
    GUICtrlSetFont(-1, 11, 700, 0, "Segoe UI")
    GUICtrlCreateLabel("An operation describes a process executed by a machine (e.g., Milling, Drilling, Assembly).", $xL, 103, 900, 18)
    GUICtrlSetFont(-1, 8, 400, 2, "Segoe UI")
    GUICtrlCreateLabel("Unit: D=Days | H=Hours | C=Hundredths of an hour   |   Interruptible: 1=Yes, 0=No", $xL, 123, 900, 18)
    GUICtrlSetFont(-1, 8, 400, 0, "Segoe UI")
    GUICtrlSetColor(-1, 0x555555)

    Local $y = 145
    _CreateCRUDButtons($y, "_Ops_Add", "_Ops_Edit", "_Ops_Del", "_Ops_DelAll", "_Ops_ImpCSV", "_Ops_ExpCSV", "_Ops_LoadEx")

    $g_hLV_Ops = GUICtrlCreateListView("Operation ID|Operation name|WC ID|Machine ID|Ref qty|Ref duration|Unit|Setup|Break|Interrupt|Line|", _
        $xL, $y + 35, $APP_WIDTH - 40, 400, _
        BitOR($LVS_REPORT, $LVS_SHOWSELALWAYS, $WS_BORDER))
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    For $i = 0 To 9
        _GUICtrlListView_SetColumnWidth($g_hLV_Ops, $i, $COL_W - 20)
    Next
    _GUICtrlListView_SetColumnWidth($g_hLV_Ops, 10, 55)
EndFunc

Func _LoadExampleOperations()
    _AddOperation("MILL", "Milling Axes", "Milling", "Milling_1", 26, 1, "H", 0, 0, 1)
    _AddOperation("MILL", "Milling Axes", "Milling", "Milling_2", 24, 1, "H", 0, 0, 1)
    _AddOperation("DRIL", "Drilling Housing", "Drilling", "Drilling_1", 49, 1, "H", 0, 0, 1)
    _AddOperation("ASSY", "Assembly Gearbox", "Assembly_Robot", "Assy_Robot_1", 19, 1, "H", 0, 0, 1)
EndFunc

Func _AddOperation($sID, $sNome, $sCT, $sMaq, $nQtdRef, $nDurRef, $sUni, $nPrep, $nTout, $bInterr)
    Local $n = UBound($g_aOps, 1)
    ReDim $g_aOps[$n + 1][11]
    _LV_AppendDataRow($g_hLV_Ops, $sID & "|" & $sNome & "|" & $sCT & "|" & $sMaq & "|" & $nQtdRef & "|" & $nDurRef & "|" & $sUni & "|" & $nPrep & "|" & $nTout & "|" & $bInterr)
EndFunc

;=============================================================================
; TAB: ROUTINGS (SV_ROUTINGS)
;=============================================================================
Func _CreateTabRoutings()
    Local $xL = 20
    GUICtrlCreateLabel("Production Routings (SV_ROUTINGS)", $xL, 80, 600, 20)
    GUICtrlSetFont(-1, 11, 700, 0, "Segoe UI")
    GUICtrlCreateLabel("A routing describes the steps of a manufacturing process (sequence of phases/operations).", $xL, 103, 900, 18)
    GUICtrlSetFont(-1, 8, 400, 2, "Segoe UI")
    GUICtrlCreateLabel("Phase code: number (10, 20, 30...) that defines the operation sequence within the routing.", $xL, 123, 900, 18)
    GUICtrlSetFont(-1, 8, 400, 0, "Segoe UI")
    GUICtrlSetColor(-1, 0x555555)

    Local $y = 145
    _CreateCRUDButtons($y, "_Rout_Add", "_Rout_Edit", "_Rout_Del", "_Rout_DelAll", "_Rout_ImpCSV", "_Rout_ExpCSV", "_Rout_LoadEx")

    $g_hLV_Rout = GUICtrlCreateListView("Routing ID|Routing name|Phase code|Operation ID|Phase name|Line|", _
        $xL, $y + 35, $APP_WIDTH - 40, 400, _
        BitOR($LVS_REPORT, $LVS_SHOWSELALWAYS, $WS_BORDER))
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    _GUICtrlListView_SetColumnWidth($g_hLV_Rout, 0, 160)
    _GUICtrlListView_SetColumnWidth($g_hLV_Rout, 1, 200)
    _GUICtrlListView_SetColumnWidth($g_hLV_Rout, 2, 110)
    _GUICtrlListView_SetColumnWidth($g_hLV_Rout, 3, 160)
    _GUICtrlListView_SetColumnWidth($g_hLV_Rout, 4, 250)
    _GUICtrlListView_SetColumnWidth($g_hLV_Rout, 5, 55)
EndFunc

Func _LoadExampleRoutings()
    _AddRouting("STD_AX", "Axes Standard", 10, "MILL", "Milling Axes Standard")
    _AddRouting("SPT_AX", "Axes Sport", 10, "MILL", "Milling Axes Sport")
    _AddRouting("SPT_AX", "Axes Sport", 20, "MILL", "Finishing Axes Sport")
    _AddRouting("CT", "Housing", 10, "DRIL", "Drilling Housing")
    _AddRouting("STD_GB", "Gearbox Standard", 10, "ASSY", "Assembly Gearbox Standard")
    _AddRouting("SPT_GB", "Gearbox Sport", 10, "ASSY", "Assembly Gearbox Sport")
EndFunc

Func _AddRouting($sID, $sNome, $nFase, $sOp, $sNomFase)
    Local $n = UBound($g_aRout, 1)
    ReDim $g_aRout[$n + 1][5]
    _LV_AppendDataRow($g_hLV_Rout, $sID & "|" & $sNome & "|" & $nFase & "|" & $sOp & "|" & $sNomFase)
EndFunc

;=============================================================================
; TAB: ITEMS (SV_MATERIALS)
;=============================================================================
Func _CreateTabMaterials()
    Local $xL = 20
    GUICtrlCreateLabel("Items (SV_MATERIALS)", $xL, 80, 600, 20)
    GUICtrlSetFont(-1, 11, 700, 0, "Segoe UI")
    GUICtrlCreateLabel("An item is a product that is consumed or produced during the production process.", $xL, 103, 900, 18)
    GUICtrlSetFont(-1, 8, 400, 2, "Segoe UI")
    GUICtrlCreateLabel("Type: MP=Raw material | SF=Semi-finished | PF=Finished good   |   Version is required for SF/PF only", $xL, 123, 900, 18)
    GUICtrlSetFont(-1, 8, 400, 0, "Segoe UI")
    GUICtrlSetColor(-1, 0x555555)

    Local $y = 145
    _CreateCRUDButtons($y, "_Mat_Add", "_Mat_Edit", "_Mat_Del", "_Mat_DelAll", "_Mat_ImpCSV", "_Mat_ExpCSV", "_Mat_LoadEx")

    $g_hLV_Mat = GUICtrlCreateListView("Item ID|Item name|Type|Version|Routing ID|On-hand qty|Line|", _
        $xL, $y + 35, $APP_WIDTH - 40, 400, _
        BitOR($LVS_REPORT, $LVS_SHOWSELALWAYS, $WS_BORDER))
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    _GUICtrlListView_SetColumnWidth($g_hLV_Mat, 0, 150)
    _GUICtrlListView_SetColumnWidth($g_hLV_Mat, 1, 200)
    _GUICtrlListView_SetColumnWidth($g_hLV_Mat, 2, 80)
    _GUICtrlListView_SetColumnWidth($g_hLV_Mat, 3, 80)
    _GUICtrlListView_SetColumnWidth($g_hLV_Mat, 4, 150)
    _GUICtrlListView_SetColumnWidth($g_hLV_Mat, 5, 110)
    _GUICtrlListView_SetColumnWidth($g_hLV_Mat, 6, 55)
EndFunc

Func _LoadExampleMaterials()
    _AddMaterial("R_Axes", "Axes Raw", "MP", "", "", 300)
    _AddMaterial("R_Housing", "Housing Raw", "MP", "", "", 72)
    _AddMaterial("STD_Axes", "Axes Standard", "SF", "STD", "STD_AX", 23)
    _AddMaterial("SPT_Axes", "Axes Sport", "SF", "SPT", "SPT_AX", 33)
    _AddMaterial("Housing", "Housing", "SF", "00", "CT", 36)
    _AddMaterial("STD_Gearbox", "Gearbox Standard", "PF", "STD", "STD_GB", 0)
    _AddMaterial("SPT_Gearbox", "Gearbox Sport", "PF", "SPT", "SPT_GB", 0)
EndFunc

Func _AddMaterial($sID, $sNome, $sTipo, $sVer, $sRot, $nStk)
    Local $n = UBound($g_aMat, 1)
    ReDim $g_aMat[$n + 1][6]
    _LV_AppendDataRow($g_hLV_Mat, $sID & "|" & $sNome & "|" & $sTipo & "|" & $sVer & "|" & $sRot & "|" & $nStk)
EndFunc

;=============================================================================
; TAB: BOM (SV_BOM)
;=============================================================================
Func _CreateTabBOM()
    Local $xL = 20
    GUICtrlCreateLabel("Bill of Materials - BOM (SV_BOM)", $xL, 80, 600, 20)
    GUICtrlSetFont(-1, 11, 700, 0, "Segoe UI")
    GUICtrlCreateLabel("The BOM defines which component items are required to produce a parent item.", $xL, 103, 900, 18)
    GUICtrlSetFont(-1, 8, 400, 2, "Segoe UI")
    GUICtrlCreateLabel("Example: STD_Box | STD | STD_Axes | STD_GB | 10 | 1 | 1   (1 axis to produce 1 box in phase 10)", $xL, 123, 900, 18)
    GUICtrlSetFont(-1, 8, 400, 0, "Segoe UI")
    GUICtrlSetColor(-1, 0x555555)

    Local $y = 145
    _CreateCRUDButtons($y, "_BOM_Add", "_BOM_Edit", "_BOM_Del", "_BOM_DelAll", "_BOM_ImpCSV", "_BOM_ExpCSV", "_BOM_LoadEx")

    $g_hLV_BOM = GUICtrlCreateListView("Parent item ID|Parent version|Component item ID|Routing ID|Phase|Ref qty|Required qty|Line|", _
        $xL, $y + 35, $APP_WIDTH - 40, 400, _
        BitOR($LVS_REPORT, $LVS_SHOWSELALWAYS, $WS_BORDER))
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    For $i = 0 To 6
        _GUICtrlListView_SetColumnWidth($g_hLV_BOM, $i, $COL_W)
    Next
    _GUICtrlListView_SetColumnWidth($g_hLV_BOM, 7, 55)
EndFunc

Func _LoadExampleBOM()
    _AddBOM("STD_Gearbox", "STD", "R_Axes",  "STD_GB", 10, 1, 2)
    _AddBOM("STD_Gearbox", "STD", "STD_Axes","STD_GB", 10, 1, 1)
    _AddBOM("STD_Gearbox", "STD", "Housing",   "STD_GB", 10, 1, 1)
    _AddBOM("SPT_Gearbox", "SPT", "R_Axes",  "SPT_GB", 10, 1, 2)
    _AddBOM("SPT_Gearbox", "SPT", "SPT_Axes","SPT_GB", 10, 1, 1)
EndFunc

Func _AddBOM($sPaiID, $sVPai, $sCompID, $sRotID, $nFase, $nQRef, $nQNec)
    Local $n = UBound($g_aBOM, 1)
    ReDim $g_aBOM[$n + 1][7]
    _LV_AppendDataRow($g_hLV_BOM, $sPaiID & "|" & $sVPai & "|" & $sCompID & "|" & $sRotID & "|" & $nFase & "|" & $nQRef & "|" & $nQNec)
EndFunc

;=============================================================================
; TAB: WORK ORDERS (SV_WO)
;=============================================================================
Func _CreateTabWO()
    Local $xL = 20
    GUICtrlCreateLabel("Work Orders (SV_WO)", $xL, 80, 600, 20)
    GUICtrlSetFont(-1, 11, 700, 0, "Segoe UI")
    GUICtrlCreateLabel("A work order (WO) is a production order with item, quantity, and planned due date.", $xL, 103, 900, 18)
    GUICtrlSetFont(-1, 8, 400, 2, "Segoe UI")
    GUICtrlCreateLabel("Date format: dd/mm/yyyy hh:mm   |   WO number must be unique", $xL, 123, 900, 18)
    GUICtrlSetFont(-1, 8, 400, 0, "Segoe UI")
    GUICtrlSetColor(-1, 0x555555)

    Local $y = 145
    _CreateCRUDButtons($y, "_WO_Add", "_WO_Edit", "_WO_Del", "_WO_DelAll", "_WO_ImpCSV", "_WO_ExpCSV", "_WO_LoadEx")

    $g_hLV_WO = GUICtrlCreateListView("WO ID|Item ID|Routing|Version|Qty|Start date|End date|Line|", _
        $xL, $y + 35, $APP_WIDTH - 40, 400, _
        BitOR($LVS_REPORT, $LVS_SHOWSELALWAYS, $WS_BORDER))
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    _GUICtrlListView_SetColumnWidth($g_hLV_WO, 0, 120)
    _GUICtrlListView_SetColumnWidth($g_hLV_WO, 1, 140)
    _GUICtrlListView_SetColumnWidth($g_hLV_WO, 2, 140)
    _GUICtrlListView_SetColumnWidth($g_hLV_WO, 3, 80)
    _GUICtrlListView_SetColumnWidth($g_hLV_WO, 4, 80)
    _GUICtrlListView_SetColumnWidth($g_hLV_WO, 5, 140)
    _GUICtrlListView_SetColumnWidth($g_hLV_WO, 6, 140)
    _GUICtrlListView_SetColumnWidth($g_hLV_WO, 7, 55)
EndFunc

Func _LoadExampleWOs()
    _AddWO("WO001", "STD_Gearbox", "STD_GB", "STD", 10, "01/02/2025 00:00", "28/02/2025 23:59")
    _AddWO("WO002", "STD_Gearbox", "STD_GB", "STD", 15, "01/02/2025 00:00", "28/02/2025 23:59")
    _AddWO("WO003", "SPT_Gearbox", "SPT_GB", "SPT", 8,  "01/03/2025 00:00", "31/03/2025 23:59")
EndFunc

Func _AddWO($sNum, $sMatID, $sRotID, $sVer, $nQtd, $sDtI, $sDtF)
    Local $n = UBound($g_aWO, 1)
    ReDim $g_aWO[$n + 1][7]
    _LV_AppendDataRow($g_hLV_WO, $sNum & "|" & $sMatID & "|" & $sRotID & "|" & $sVer & "|" & $nQtd & "|" & $sDtI & "|" & $sDtF)
EndFunc

;=============================================================================
; TAB: WO LINKS
;=============================================================================
Func _CreateTabWOLinks()
    Local $xL = 20
    GUICtrlCreateLabel("WO / Operation Links - Precedence Constraints (SV_WO_LINKS)", $xL, 80, 700, 20)
    GUICtrlSetFont(-1, 11, 700, 0, "Segoe UI")
    GUICtrlCreateLabel("A link enforces a precedence constraint between two WOs/operations (e.g., WO2 starts only after phase 10 of WO1 ends).", $xL, 103, 950, 18)
    GUICtrlSetFont(-1, 8, 400, 2, "Segoe UI")
    GUICtrlCreateLabel("Link types: FS=Finish-Start | SS=Start-Start | FF=Finish-Finish | SF=Start-Finish", $xL, 123, 900, 18)
    GUICtrlSetFont(-1, 8, 400, 0, "Segoe UI")
    GUICtrlSetColor(-1, 0x555555)

    Local $y = 145
    _CreateCRUDButtons($y, "_WOL_Add", "_WOL_Edit", "_WOL_Del", "_WOL_DelAll", "_WOL_ImpCSV", "_WOL_ExpCSV", "_WOL_LoadEx")

    $g_hLV_WOL = GUICtrlCreateListView("Predecessor WO|Pred routing|Pred phase|Successor WO|Succ routing|Succ phase|Link type|Line|", _
        $xL, $y + 35, $APP_WIDTH - 40, 400, _
        BitOR($LVS_REPORT, $LVS_SHOWSELALWAYS, $WS_BORDER))
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    For $i = 0 To 6
        _GUICtrlListView_SetColumnWidth($g_hLV_WOL, $i, $COL_W)
    Next
    _GUICtrlListView_SetColumnWidth($g_hLV_WOL, 7, 55)
EndFunc

;=============================================================================
; TAB: SECONDARY RESOURCES
;=============================================================================
Func _CreateTabSecResources()
    Local $xL = 20
    GUICtrlCreateLabel("Secondary Resources (SV_SEC_RESOURCES)", $xL, 80, 600, 20)
    GUICtrlSetFont(-1, 11, 700, 0, "Segoe UI")
    GUICtrlCreateLabel("A secondary resource is required during an operation in addition to the primary machine (e.g., labor, tooling).", $xL, 103, 900, 18)
    GUICtrlSetFont(-1, 8, 400, 2, "Segoe UI")
    GUICtrlCreateLabel("Qualification ID identifies the type of labor or tool required. Capacity calendar controls availability.", $xL, 123, 900, 18)
    GUICtrlSetFont(-1, 8, 400, 0, "Segoe UI")
    GUICtrlSetColor(-1, 0x555555)

    Local $y = 145
    _CreateCRUDButtons($y, "_SR_Add", "_SR_Edit", "_SR_Del", "_SR_DelAll", "_SR_ImpCSV", "_SR_ExpCSV", "_SR_LoadEx")

    $g_hLV_SR = GUICtrlCreateListView("Operation ID|WC ID|Machine ID|Qualification ID|Capacity calendar ID|Line|", _
        $xL, $y + 35, $APP_WIDTH - 40, 400, _
        BitOR($LVS_REPORT, $LVS_SHOWSELALWAYS, $WS_BORDER))
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    For $i = 0 To 4
        _GUICtrlListView_SetColumnWidth($g_hLV_SR, $i, $COL_W + 20)
    Next
    _GUICtrlListView_SetColumnWidth($g_hLV_SR, 5, 55)
EndFunc

;=============================================================================
; TAB: CAPACITY
;=============================================================================
Func _CreateTabCapacity()
    Local $xL = 20
    GUICtrlCreateLabel("Capacity Calendars (SV_CAPACITY)", $xL, 80, 600, 20)
    GUICtrlSetFont(-1, 11, 700, 0, "Segoe UI")
    GUICtrlCreateLabel("A capacity calendar defines shifts and the number of resources available for load calculation.", $xL, 103, 900, 18)
    GUICtrlSetFont(-1, 8, 400, 2, "Segoe UI")
    GUICtrlCreateLabel("Days: 1=Mon ... 7=Sun   |   #Resources = number of concurrent resources available during that shift", $xL, 123, 900, 18)
    GUICtrlSetFont(-1, 8, 400, 0, "Segoe UI")
    GUICtrlSetColor(-1, 0x555555)

    Local $y = 145
    _CreateCRUDButtons($y, "_Cap_Add", "_Cap_Edit", "_Cap_Del", "_Cap_DelAll", "_Cap_ImpCSV", "_Cap_ExpCSV", "_Cap_LoadEx")

    $g_hLV_Cap = GUICtrlCreateListView("Cap cal ID|Start day|Start time|End day|End time|#Resources|Line|", _
        $xL, $y + 35, $APP_WIDTH - 40, 400, _
        BitOR($LVS_REPORT, $LVS_SHOWSELALWAYS, $WS_BORDER))
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    For $i = 0 To 5
        _GUICtrlListView_SetColumnWidth($g_hLV_Cap, $i, $COL_W + 10)
    Next
    _GUICtrlListView_SetColumnWidth($g_hLV_Cap, 6, 55)
EndFunc

;=============================================================================
; TAB: INVENTORY MOVEMENTS
;=============================================================================
Func _CreateTabStockMov()
    Local $xL = 20
    GUICtrlCreateLabel("Inventory Movements (SV_STOCK_MOVEMENTS)", $xL, 80, 600, 20)
    GUICtrlSetFont(-1, 11, 700, 0, "Segoe UI")
    GUICtrlCreateLabel("An inventory movement is an in/out transaction of an item on a given date.", $xL, 103, 900, 18)
    GUICtrlSetFont(-1, 8, 400, 2, "Segoe UI")
    GUICtrlCreateLabel("Positive quantity = receipt; negative quantity = issue. Date format: dd/mm/yyyy", $xL, 123, 900, 18)
    GUICtrlSetFont(-1, 8, 400, 0, "Segoe UI")
    GUICtrlSetColor(-1, 0x555555)

    Local $y = 145
    _CreateCRUDButtons($y, "_Stk_Add", "_Stk_Edit", "_Stk_Del", "_Stk_DelAll", "_Stk_ImpCSV", "_Stk_ExpCSV", "_Stk_LoadEx")

    $g_hLV_Stk = GUICtrlCreateListView("Item ID|Routing ID|Version|Move date|Quantity|Line|", _
        $xL, $y + 35, $APP_WIDTH - 40, 400, _
        BitOR($LVS_REPORT, $LVS_SHOWSELALWAYS, $WS_BORDER))
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    For $i = 0 To 4
        _GUICtrlListView_SetColumnWidth($g_hLV_Stk, $i, $COL_W + 30)
    Next
    _GUICtrlListView_SetColumnWidth($g_hLV_Stk, 5, 55)
EndFunc

;=============================================================================
; TAB: GENERATE SQL
;=============================================================================
Func _CreateTabSQL()
    Local $xL = 20
    GUICtrlCreateLabel("SQL Generation & Execution", $xL, 80, 500, 20)
    GUICtrlSetFont(-1, 11, 700, 0, "Segoe UI")
    GUICtrlCreateLabel("Click 'GENERATE SQL' in the bottom toolbar to build the script. Then click 'RUN ON DB' to execute it.", $xL, 103, 950, 18)
    GUICtrlSetFont(-1, 8, 400, 2, "Segoe UI")

    ; Options row (inline, no group box needed)
    Local $y = 128
    GUICtrlCreateGroup("Options", $xL, $y, 520, 92)
    Global $g_chkClearFirst = GUICtrlCreateCheckbox("Clear existing data before insert (DELETE FROM ...)", $xL + 15, $y + 18, 470, 20)
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    GUICtrlSetState($g_chkClearFirst, $GUI_CHECKED)
    _RegisterBusyControl($g_chkClearFirst)
    Global $g_chkTransaction = GUICtrlCreateCheckbox("Wrap in a transaction (auto ROLLBACK on error)", $xL + 15, $y + 40, 470, 20)
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    GUICtrlSetState($g_chkTransaction, $GUI_CHECKED)
    _RegisterBusyControl($g_chkTransaction)
    Global $g_chkVerbose = GUICtrlCreateCheckbox("Verbose Mode (detailed troubleshooting log)", $xL + 15, $y + 62, 470, 20)
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    GUICtrlSetOnEvent($g_chkVerbose, "_OnVerboseModeChanged")
    _RegisterBusyControl($g_chkVerbose)

    ; SQL editor
    $y = 236
    GUICtrlCreateLabel("Generated SQL:", $xL, $y, 200, 18)
    GUICtrlSetFont(-1, 9, 700, 0, "Segoe UI")

    ; SQL box: height leaves room for exec log at the bottom (fixed 90px area)
    $g_hLog = GUICtrlCreateEdit("-- Click 'GENERATE SQL' to build the script..." & @CRLF & _
        "-- SQL is built from data entered in each tab.", _
        $xL, $y + 20, $APP_WIDTH - 40, 300, _
        BitOR($ES_MULTILINE, $WS_VSCROLL, $WS_HSCROLL, $ES_READONLY, $ES_AUTOVSCROLL))
    GUICtrlSetFont(-1, 9, 400, 0, "Courier New")
    GUICtrlSetBkColor($g_hLog, 0x1E1E1E)
    GUICtrlSetColor($g_hLog, 0x00FF00)

    ; Execution log - placed after SQL box with fixed offset, well within tab area
    $y = $y + 20 + 300 + 10
    GUICtrlCreateLabel("Execution log:", $xL, $y, 200, 18)
    GUICtrlSetFont(-1, 9, 700, 0, "Segoe UI")
    Global $g_hExecLog = GUICtrlCreateEdit("", $xL, $y + 20, $APP_WIDTH - 40, 55, _
        BitOR($ES_MULTILINE, $WS_VSCROLL, $ES_READONLY))
    GUICtrlSetFont(-1, 8, 400, 0, "Segoe UI")
    GUICtrlSetBkColor($g_hExecLog, 0xFFFFF0)
EndFunc

;=============================================================================
; HELPER: STANDARD CRUD BUTTONS
;=============================================================================
Func _CreateCRUDButtons($y, $sAdd, $sEdit, $sDel, $sDelAll, $sImp = "", $sExp = "", $sLoadEx = "")
    Local $btnAdd = GUICtrlCreateButton("+ Add", 20, $y - 2, 110, 28)
    GUICtrlSetOnEvent($btnAdd, $sAdd)
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    GUICtrlSetTip($btnAdd, "Add a new row to this tab")
    _RegisterBusyControl($btnAdd)

    Local $btnEdit = GUICtrlCreateButton("Edit", 140, $y - 2, 90, 28)
    GUICtrlSetOnEvent($btnEdit, $sEdit)
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    GUICtrlSetTip($btnEdit, "Edit the currently selected row")
    _RegisterBusyControl($btnEdit)

    Local $btnDel = GUICtrlCreateButton("- Remove", 240, $y - 2, 100, 28)
    GUICtrlSetOnEvent($btnDel, $sDel)
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    GUICtrlSetTip($btnDel, "Remove the currently selected row")
    _RegisterBusyControl($btnDel)

    Local $btnDelAll = GUICtrlCreateButton("Clear All", 350, $y - 2, 105, 28)
    GUICtrlSetOnEvent($btnDelAll, $sDelAll)
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    GUICtrlSetTip($btnDelAll, "Remove ALL rows in this tab")
    _RegisterBusyControl($btnDelAll)

    ; Load Example - highlighted so users see it as the easy way to populate a tab
    If $sLoadEx <> "" Then
        Local $btnLoadEx = GUICtrlCreateButton("Load Example", 465, $y - 2, 140, 28)
        GUICtrlSetOnEvent($btnLoadEx, $sLoadEx)
        GUICtrlSetFont(-1, 9, 700, 0, "Segoe UI")
        GUICtrlSetBkColor($btnLoadEx, 0xE8F4FD)
        GUICtrlSetTip($btnLoadEx, "Fill this tab with sample demo data (appends to any existing rows)")
        _RegisterBusyControl($btnLoadEx)
    EndIf

    Local $btnImp = GUICtrlCreateButton("Import CSV...", 615, $y - 2, 120, 28)
    If $sImp <> "" Then GUICtrlSetOnEvent($btnImp, $sImp)
    _RegisterBusyControl($btnImp)
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")

    Local $btnExp = GUICtrlCreateButton("Export CSV...", 745, $y - 2, 120, 28)
    If $sExp <> "" Then GUICtrlSetOnEvent($btnExp, $sExp)
    _RegisterBusyControl($btnExp)
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
EndFunc


;=============================================================================
; LOOKUP / CONTROLLED-SELECTION HELPERS
; These helpers feed dropdown fields from the current tab data instead of asking
; the user to manually type cross-tab references.
;=============================================================================
Func _PipeAppendUnique($sList, $sValue)
    Local $sVal = StringStripWS($sValue, 3)
    If $sVal = "" Then Return $sList
    If StringInStr("|" & $sList & "|", "|" & $sVal & "|") Then Return $sList
    If $sList <> "" Then $sList &= "|"
    Return $sList & $sVal
EndFunc

Func _PipeHasValue($sList, $sValue)
    Local $sVal = StringStripWS($sValue, 3)
    If $sVal = "" Then Return False
    Return (StringInStr("|" & $sList & "|", "|" & $sVal & "|") > 0)
EndFunc

Func _PipeEnsureValue($sList, $sValue)
    If StringStripWS($sValue, 3) = "" Then Return $sList
    Return _PipeAppendUnique($sList, $sValue)
EndFunc

Func _PipeUnion($sListA, $sListB)
    Local $sOut = $sListA
    If $sListB = "" Then Return $sOut
    Local $aItems = StringSplit($sListB, "|", 1)
    For $i = 1 To $aItems[0]
        $sOut = _PipeAppendUnique($sOut, $aItems[$i])
    Next
    Return $sOut
EndFunc

Func _LVColumnPipe($hLV, $iCol)
    Local $sList = ""
    If $hLV = 0 Then Return $sList
    Local $nRows = _GUICtrlListView_GetItemCount($hLV)
    For $i = 0 To $nRows - 1
        $sList = _PipeAppendUnique($sList, _GUICtrlListView_GetItemText($hLV, $i, $iCol))
    Next
    Return $sList
EndFunc

Func _LVColumnPipeFiltered1($hLV, $iReturnCol, $iFilterCol, $sFilterValue)
    Local $sList = ""
    If $hLV = 0 Then Return $sList
    Local $sFilter = StringStripWS($sFilterValue, 3)
    Local $nRows = _GUICtrlListView_GetItemCount($hLV)
    For $i = 0 To $nRows - 1
        If $sFilter <> "" And String(_GUICtrlListView_GetItemText($hLV, $i, $iFilterCol)) <> String($sFilter) Then ContinueLoop
        $sList = _PipeAppendUnique($sList, _GUICtrlListView_GetItemText($hLV, $i, $iReturnCol))
    Next
    Return $sList
EndFunc

Func _LVColumnPipeFiltered2($hLV, $iReturnCol, $iFilterCol1, $sFilterValue1, $iFilterCol2, $sFilterValue2)
    Local $sList = ""
    If $hLV = 0 Then Return $sList
    Local $sFilter1 = StringStripWS($sFilterValue1, 3)
    Local $sFilter2 = StringStripWS($sFilterValue2, 3)
    Local $nRows = _GUICtrlListView_GetItemCount($hLV)
    For $i = 0 To $nRows - 1
        If $sFilter1 <> "" And String(_GUICtrlListView_GetItemText($hLV, $i, $iFilterCol1)) <> String($sFilter1) Then ContinueLoop
        If $sFilter2 <> "" And String(_GUICtrlListView_GetItemText($hLV, $i, $iFilterCol2)) <> String($sFilter2) Then ContinueLoop
        $sList = _PipeAppendUnique($sList, _GUICtrlListView_GetItemText($hLV, $i, $iReturnCol))
    Next
    Return $sList
EndFunc

Func _LVHasValueInColumn($hLV, $iCol, $sValue)
    Local $sVal = StringStripWS($sValue, 3)
    If $sVal = "" Or $hLV = 0 Then Return False
    Local $nRows = _GUICtrlListView_GetItemCount($hLV)
    For $i = 0 To $nRows - 1
        If String(_GUICtrlListView_GetItemText($hLV, $i, $iCol)) = String($sVal) Then Return True
    Next
    Return False
EndFunc

Func _LVHasRowMatching2($hLV, $iCol1, $sVal1, $iCol2, $sVal2)
    If $hLV = 0 Then Return False
    Local $v1 = StringStripWS($sVal1, 3)
    Local $v2 = StringStripWS($sVal2, 3)
    Local $nRows = _GUICtrlListView_GetItemCount($hLV)
    For $i = 0 To $nRows - 1
        If String(_GUICtrlListView_GetItemText($hLV, $i, $iCol1)) = String($v1) And _
           String(_GUICtrlListView_GetItemText($hLV, $i, $iCol2)) = String($v2) Then Return True
    Next
    Return False
EndFunc

Func _LVHasRowMatching3($hLV, $iCol1, $sVal1, $iCol2, $sVal2, $iCol3, $sVal3)
    If $hLV = 0 Then Return False
    Local $v1 = StringStripWS($sVal1, 3)
    Local $v2 = StringStripWS($sVal2, 3)
    Local $v3 = StringStripWS($sVal3, 3)
    Local $nRows = _GUICtrlListView_GetItemCount($hLV)
    For $i = 0 To $nRows - 1
        If String(_GUICtrlListView_GetItemText($hLV, $i, $iCol1)) = String($v1) And _
           String(_GUICtrlListView_GetItemText($hLV, $i, $iCol2)) = String($v2) And _
           String(_GUICtrlListView_GetItemText($hLV, $i, $iCol3)) = String($v3) Then Return True
    Next
    Return False
EndFunc

Func _DialogGetValue(ByRef $aFields, ByRef $aValues, $sFieldName)
    Local $sTarget = StringLower(StringStripWS($sFieldName, 3))
    For $i = 0 To UBound($aFields) - 1
        If StringLower(StringStripWS($aFields[$i], 3)) = $sTarget Then
            If IsArray($aValues) And $i < UBound($aValues) Then Return StringStripWS($aValues[$i], 3)
            Return ""
        EndIf
    Next
    Return ""
EndFunc

Func _Lookup_RoutingIDsForItem($sItemID)
    If _ItemIsRawMaterial($sItemID) Then Return ""
    Local $sList = _LVColumnPipeFiltered1($g_hLV_Mat, 4, 0, $sItemID)
    If $sList = "" Then $sList = _LVColumnPipe($g_hLV_Rout, 0)
    Return $sList
EndFunc

Func _Lookup_RoutingIDsForWO($sWOID)
    Local $sList = _LVColumnPipeFiltered1($g_hLV_WO, 2, 0, $sWOID)
    If $sList = "" Then $sList = _LVColumnPipe($g_hLV_Rout, 0)
    Return $sList
EndFunc

Func _Lookup_ItemVersions($sItemID = "", $sRoutingID = "")
    If $sItemID <> "" And _ItemIsRawMaterial($sItemID) Then Return ""
    Local $sList = ""
    If $sItemID <> "" And $sRoutingID <> "" Then $sList = _LVColumnPipeFiltered2($g_hLV_Mat, 3, 0, $sItemID, 4, $sRoutingID)
    If $sList = "" And $sItemID <> "" Then $sList = _LVColumnPipeFiltered1($g_hLV_Mat, 3, 0, $sItemID)
    If $sList = "" And $sRoutingID <> "" Then $sList = _LVColumnPipeFiltered1($g_hLV_Mat, 3, 4, $sRoutingID)
    If $sList = "" Then $sList = _LVColumnPipe($g_hLV_Mat, 3)
    Return $sList
EndFunc

Func _DialogLookupKind($sTitle, $sFieldName)
    Local $t = StringLower($sTitle)
    Local $f = StringLower(StringStripWS($sFieldName, 3))

    If $f = "start day" Or $f = "end day" Then Return "DAY"

    If StringInStr($t, "machine") Then
        Switch $f
            Case "ct type"
                Return "CTTYPE"
            Case "machine type", "mach type"
                Return "MACHTYPE"
            Case "calendar id"
                Return "CALENDAR"
            Case "capacity calendar id"
                Return "ANYCALENDAR"
        EndSwitch
    EndIf

    If StringInStr($t, "operation") Then
        Switch $f
            Case "wc / ct id", "wc id", "ct id"
                Return "CT"
            Case "machine id"
                Return "MACHINE_BY_CT"
            Case "unit"
                Return "UNIT"
            Case "interruptible"
                Return "BOOL"
        EndSwitch
    EndIf

    If StringInStr($t, "routing") And Not StringInStr($t, "work order") Then
        If $f = "operation id" Then Return "OPERATION"
    EndIf

    If StringInStr($t, "item") Then
        Switch $f
            Case "type", "item type"
                Return "ITEMTYPE"
            Case "routing id", "routing"
                Return "ROUTING"
        EndSwitch
    EndIf

    If StringInStr($t, "bom") Then
        Switch $f
            Case "parent item id", "component item id"
                Return "ITEM"
            Case "parent version"
                Return "PARENTVERSION"
            Case "routing id", "routing"
                Return "ROUTING"
            Case "phase", "phase code"
                Return "PHASE_BY_ROUTING"
        EndSwitch
    EndIf

    If StringInStr($t, "work order") Then
        Switch $f
            Case "item id"
                Return "ITEM"
            Case "routing id", "routing"
                Return "ROUTING_BY_ITEM"
            Case "version", "routing version"
                Return "VERSION_BY_ITEM_ROUTING"
        EndSwitch
    EndIf

    If StringInStr($t, "wo link") Then
        Switch $f
            Case "predecessor wo", "successor wo"
                Return "WO"
            Case "pred routing"
                Return "ROUTING_BY_PRED_WO"
            Case "succ routing"
                Return "ROUTING_BY_SUCC_WO"
            Case "pred phase"
                Return "PHASE_BY_PRED_ROUTING"
            Case "succ phase"
                Return "PHASE_BY_SUCC_ROUTING"
            Case "link type"
                Return "LINKTYPE"
        EndSwitch
    EndIf

    If StringInStr($t, "secondary resource") Then
        Switch $f
            Case "operation id"
                Return "OPERATION"
            Case "wc / ct id", "wc id", "ct id"
                Return "CT"
            Case "machine id"
                Return "MACHINE_BY_CT"
            Case "capacity calendar id"
                Return "ANYCALENDAR"
        EndSwitch
    EndIf

    If StringInStr($t, "inventory movement") Then
        Switch $f
            Case "item id"
                Return "ITEM"
            Case "routing id", "routing"
                Return "ROUTING_BY_ITEM"
            Case "version", "routing version"
                Return "VERSION_BY_ITEM_ROUTING"
        EndSwitch
    EndIf

    Return ""
EndFunc

Func _DialogLookupAllowsBlank($sTitle, $sFieldName, ByRef $aFields, ByRef $aValues)
    Local $t = StringLower($sTitle)
    Local $f = StringLower(StringStripWS($sFieldName, 3))
    If StringInStr($t, "item") And ($f = "routing id" Or $f = "routing") Then Return True
    If StringInStr($t, "machine") And ($f = "calendar id" Or $f = "capacity calendar id") Then Return True
    If StringInStr($t, "secondary resource") And $f = "capacity calendar id" Then Return True
    If StringInStr($t, "inventory movement") And ($f = "routing id" Or $f = "routing") Then Return True
    If StringInStr($t, "inventory movement") And ($f = "version" Or $f = "routing version") Then
        Local $sItem = _DialogGetValue($aFields, $aValues, "Item ID")
        If _ItemIsRawMaterial($sItem) Then Return True
    EndIf
    Return False
EndFunc

Func _DialogLookupOptions($sKind, ByRef $aFields, ByRef $aValues)
    Switch $sKind
        Case "DAY"
            Return "1|2|3|4|5|6|7"
        Case "CTTYPE"
            Return "2|4"
        Case "MACHTYPE"
            Return "NR|BA|RN|CU"
        Case "ITEMTYPE"
            Return "MP|SF|PF"
        Case "UNIT"
            Return "D|H|C"
        Case "BOOL"
            Return "0|1"
        Case "LINKTYPE"
            Return "FS|SS|FF|SF"
        Case "ITEM"
            Return _LVColumnPipe($g_hLV_Mat, 0)
        Case "ROUTING"
            Return _LVColumnPipe($g_hLV_Rout, 0)
        Case "ROUTING_BY_ITEM"
            Return _Lookup_RoutingIDsForItem(_DialogGetValue($aFields, $aValues, "Item ID"))
        Case "ROUTING_BY_PRED_WO"
            Return _Lookup_RoutingIDsForWO(_DialogGetValue($aFields, $aValues, "Predecessor WO"))
        Case "ROUTING_BY_SUCC_WO"
            Return _Lookup_RoutingIDsForWO(_DialogGetValue($aFields, $aValues, "Successor WO"))
        Case "PARENTVERSION"
            Return _Lookup_ItemVersions(_DialogGetValue($aFields, $aValues, "Parent item ID"), "")
        Case "VERSION_BY_ITEM_ROUTING"
            Return _Lookup_ItemVersions(_DialogGetValue($aFields, $aValues, "Item ID"), _DialogGetValue($aFields, $aValues, "Routing ID"))
        Case "PHASE_BY_ROUTING"
            Return _LVColumnPipeFiltered1($g_hLV_Rout, 2, 0, _DialogGetValue($aFields, $aValues, "Routing ID"))
        Case "PHASE_BY_PRED_ROUTING"
            Return _LVColumnPipeFiltered1($g_hLV_Rout, 2, 0, _DialogGetValue($aFields, $aValues, "Pred routing"))
        Case "PHASE_BY_SUCC_ROUTING"
            Return _LVColumnPipeFiltered1($g_hLV_Rout, 2, 0, _DialogGetValue($aFields, $aValues, "Succ routing"))
        Case "OPERATION"
            Return _LVColumnPipe($g_hLV_Ops, 0)
        Case "CT"
            Return _LVColumnPipe($g_hLV_Mach, 2)
        Case "MACHINE_BY_CT"
            Local $sList = _LVColumnPipeFiltered1($g_hLV_Mach, 7, 2, _DialogGetValue($aFields, $aValues, "WC / CT ID"))
            If $sList = "" Then $sList = _LVColumnPipe($g_hLV_Mach, 7)
            Return $sList
        Case "CALENDAR"
            Return _LVColumnPipe($g_hLV_Cal, 0)
        Case "ANYCALENDAR"
            Return _PipeUnion(_LVColumnPipe($g_hLV_Cal, 0), _LVColumnPipe($g_hLV_Cap, 0))
        Case "WO"
            Return _LVColumnPipe($g_hLV_WO, 0)
    EndSwitch
    Return ""
EndFunc

Func _ComboSetPipeList($idCombo, $sOptions, $sSelected, $bAllowBlank = False, $bPreserveInvalid = False)
    Local $hCombo = GUICtrlGetHandle($idCombo)
    Local $sSel = StringStripWS($sSelected, 3)
    If $bPreserveInvalid Then $sOptions = _PipeEnsureValue($sOptions, $sSel)

    _GUICtrlComboBox_ResetContent($hCombo)
    If $bAllowBlank Then _GUICtrlComboBox_AddString($hCombo, "")

    If $sOptions <> "" Then
        Local $aItems = StringSplit($sOptions, "|", 1)
        For $i = 1 To $aItems[0]
            If StringStripWS($aItems[$i], 3) <> "" Then _GUICtrlComboBox_AddString($hCombo, $aItems[$i])
        Next
    EndIf

    Local $iSel = -1
    If $sSel <> "" Or $bAllowBlank Then $iSel = _GUICtrlComboBox_FindStringExact($hCombo, $sSel)
    If $iSel >= 0 Then
        _GUICtrlComboBox_SetCurSel($hCombo, $iSel)
    ElseIf _GUICtrlComboBox_GetCount($hCombo) > 0 Then
        _GUICtrlComboBox_SetCurSel($hCombo, 0)
    EndIf
EndFunc

Func _DialogReadControlValues(ByRef $aInputs)
    Local $n = UBound($aInputs)
    Local $aValues[$n]
    For $i = 0 To $n - 1
        $aValues[$i] = StringStripWS(GUICtrlRead($aInputs[$i]), 3)
    Next
    Return $aValues
EndFunc

Func _DialogRefreshLookupCombos($sTitle, ByRef $aFields, ByRef $aInputs, ByRef $aIsCombo)
    Local $aValues = _DialogReadControlValues($aInputs)
    For $i = 0 To UBound($aFields) - 1
        If Not $aIsCombo[$i] Then ContinueLoop
        Local $sKind = _DialogLookupKind($sTitle, $aFields[$i])
        If $sKind = "" Then ContinueLoop
        Local $sOptions = _DialogLookupOptions($sKind, $aFields, $aValues)
        Local $bAllowBlank = _DialogLookupAllowsBlank($sTitle, $aFields[$i], $aFields, $aValues)
        _ComboSetPipeList($aInputs[$i], $sOptions, $aValues[$i], $bAllowBlank, False)
        $aValues = _DialogReadControlValues($aInputs)
    Next
EndFunc

Func _ValidateDialogLookupValues($sTitle, ByRef $aFields, ByRef $aValues)
    Local $sErrors = ""
    Local $t = StringLower($sTitle)

    For $i = 0 To UBound($aFields) - 1
        Local $sKind = _DialogLookupKind($sTitle, $aFields[$i])
        If $sKind = "" Then ContinueLoop
        Local $sVal = StringStripWS($aValues[$i], 3)
        If $sVal = "" Then ContinueLoop
        Local $sOptions = _DialogLookupOptions($sKind, $aFields, $aValues)
        If $sOptions = "" Then
            _AppendDialogError($sErrors, $aFields[$i] & " has no available master-data option yet. Create or load the source tab first.")
        ElseIf Not _PipeHasValue($sOptions, $sVal) Then
            _AppendDialogError($sErrors, $aFields[$i] & " must be selected from the available list.")
        EndIf
    Next

    If StringInStr($t, "bom") Then
        Local $sParent = _DialogGetValue($aFields, $aValues, "Parent item ID")
        Local $sPVer = _DialogGetValue($aFields, $aValues, "Parent version")
        Local $sComp = _DialogGetValue($aFields, $aValues, "Component item ID")
        Local $sRout = _DialogGetValue($aFields, $aValues, "Routing ID")
        Local $sPhase = _DialogGetValue($aFields, $aValues, "Phase")
        If $sParent <> "" And Not _LVHasValueInColumn($g_hLV_Mat, 0, $sParent) Then _AppendDialogError($sErrors, "Parent item ID must exist in the Items tab.")
        If $sParent <> "" And _ItemIsRawMaterial($sParent) Then _AppendDialogError($sErrors, "Parent item cannot be MP. Raw materials do not have produced version/routing.")
        If $sComp <> "" And Not _LVHasValueInColumn($g_hLV_Mat, 0, $sComp) Then _AppendDialogError($sErrors, "Component item ID must exist in the Items tab.")
        If $sParent <> "" And Not _ItemIsRawMaterial($sParent) And $sPVer = "" Then _AppendDialogError($sErrors, "Parent version is required for manufactured parent items.")
        If $sParent <> "" And $sPVer <> "" And Not _LVHasRowMatching2($g_hLV_Mat, 0, $sParent, 3, $sPVer) Then _AppendDialogError($sErrors, "Parent version must exist for the selected parent item.")
        If $sRout <> "" And Not _LVHasValueInColumn($g_hLV_Rout, 0, $sRout) Then _AppendDialogError($sErrors, "Routing ID must exist in the Routings tab.")
        If $sRout <> "" And $sPhase <> "" And Not _LVHasRowMatching2($g_hLV_Rout, 0, $sRout, 2, $sPhase) Then _AppendDialogError($sErrors, "Phase must exist for the selected Routing ID.")
    EndIf

    If StringInStr($t, "work order") Then
        Local $sItem = _DialogGetValue($aFields, $aValues, "Item ID")
        Local $sRoutWO = _DialogGetValue($aFields, $aValues, "Routing ID")
        Local $sVerWO = _DialogGetValue($aFields, $aValues, "Version")
        If $sItem <> "" And Not _LVHasValueInColumn($g_hLV_Mat, 0, $sItem) Then _AppendDialogError($sErrors, "Item ID must exist in the Items tab.")
        If $sItem <> "" And _ItemIsRawMaterial($sItem) Then _AppendDialogError($sErrors, "Work Orders must produce manufactured items. MP raw materials do not have routing/version.")
        If $sRoutWO <> "" And Not _LVHasValueInColumn($g_hLV_Rout, 0, $sRoutWO) Then _AppendDialogError($sErrors, "Routing ID must exist in the Routings tab.")
        If $sItem <> "" And $sRoutWO <> "" And $sVerWO <> "" And Not _LVHasRowMatching3($g_hLV_Mat, 0, $sItem, 4, $sRoutWO, 3, $sVerWO) Then _AppendDialogError($sErrors, "The Item ID + Routing ID + Version combination must exist in the Items tab.")
    EndIf

    If StringInStr($t, "inventory movement") Then
        Local $sInvItem = _DialogGetValue($aFields, $aValues, "Item ID")
        Local $sInvRout = _DialogGetValue($aFields, $aValues, "Routing ID")
        Local $sInvVer = _DialogGetValue($aFields, $aValues, "Version")
        If $sInvItem <> "" And Not _LVHasValueInColumn($g_hLV_Mat, 0, $sInvItem) Then _AppendDialogError($sErrors, "Item ID must exist in the Items tab.")
        If $sInvRout <> "" And Not _LVHasValueInColumn($g_hLV_Rout, 0, $sInvRout) Then _AppendDialogError($sErrors, "Routing ID must exist in the Routings tab.")
        If $sInvItem <> "" And _ItemIsRawMaterial($sInvItem) Then
            If $sInvRout <> "" Then _AppendDialogError($sErrors, "MP raw materials cannot have Routing ID. Leave Routing ID blank.")
            If $sInvVer <> "" Then _AppendDialogError($sErrors, "MP raw materials cannot have Version. Leave Version blank.")
        ElseIf $sInvItem <> "" Then
            If $sInvVer = "" Then _AppendDialogError($sErrors, "Version is required for manufactured inventory movements.")
            If $sInvRout = "" Or $sInvVer = "" Then _AppendDialogError($sErrors, "Routing ID and Version are required for manufactured inventory movements.")
        EndIf
        If $sInvItem <> "" And Not _ItemIsRawMaterial($sInvItem) And $sInvRout <> "" And $sInvVer <> "" And Not _LVHasRowMatching3($g_hLV_Mat, 0, $sInvItem, 4, $sInvRout, 3, $sInvVer) Then _AppendDialogError($sErrors, "The Item ID + Routing ID + Version combination must exist in the Items tab.")
    EndIf

    If StringInStr($t, "operation") Or StringInStr($t, "secondary resource") Then
        Local $sCT = _DialogGetValue($aFields, $aValues, "WC / CT ID")
        Local $sMach = _DialogGetValue($aFields, $aValues, "Machine ID")
        If $sCT <> "" And Not _LVHasValueInColumn($g_hLV_Mach, 2, $sCT) Then _AppendDialogError($sErrors, "WC / CT ID must exist in the Machines tab.")
        If $sMach <> "" And Not _IsStandByMachine($sMach) And Not _LVHasValueInColumn($g_hLV_Mach, 7, $sMach) Then _AppendDialogError($sErrors, "Machine ID must exist in the Machines tab.")
        If $sCT <> "" And $sMach <> "" And Not _IsStandByMachine($sMach) And Not _LVHasRowMatching2($g_hLV_Mach, 2, $sCT, 7, $sMach) Then _AppendDialogError($sErrors, "Machine ID must belong to the selected WC / CT ID.")
    EndIf

    If StringInStr($t, "routing") And Not StringInStr($t, "work order") Then
        Local $sOp = _DialogGetValue($aFields, $aValues, "Operation ID")
        If $sOp <> "" And Not _LVHasValueInColumn($g_hLV_Ops, 0, $sOp) Then _AppendDialogError($sErrors, "Operation ID must exist in the Operations tab.")
    EndIf

    Return $sErrors
EndFunc

;=============================================================================
; GENERIC DIALOG TO ADD/EDIT A ROW
;   $aFields  = array of field labels
;   $aValues  = array of current values (for edit) OR 0 for empty
;   $aHints   = OPTIONAL array of hint strings shown under each input
;   $sSubtitle= OPTIONAL extra text under the dialog title (e.g. "Row 3 of 5")
;
; IMPORTANT - the main GUI runs in OnEvent mode, but this popup is handled as a
; local modal dialog with GUIGetMsg. That avoids the deadlock/freeze seen when
; Save/Cancel tried to signal back through a busy wait loop.
;=============================================================================
Func _ShowRowDialog($sTitle, $aFields, $aValues, $aHints = "", $sSubtitle = "")
    Local $nF = UBound($aFields)
    Local $bHasHints = (IsArray($aHints) And UBound($aHints) >= $nF)

    Local $iRowH = $bHasHints ? 50 : 34
    Local $iDlgW = 640
    Local $iHeaderH = ($sSubtitle = "" ? 62 : 80)
    Local $iDlgH = $iHeaderH + 10 + $nF * $iRowH + 70

    ; Center the row dialog over the current position of the main window.
    ; Using -1/-1 here centers on the desktop/default position, which feels wrong
    ; when the main Toolbox window has been moved to another screen area.
    Local $iDlgX = -1, $iDlgY = -1
    If $g_hMain <> 0 Then
        Local $aMainPos = WinGetPos($g_hMain)
        If Not @error And IsArray($aMainPos) Then
            $iDlgX = $aMainPos[0] + Int(($aMainPos[2] - $iDlgW) / 2)
            $iDlgY = $aMainPos[1] + Int(($aMainPos[3] - $iDlgH) / 2)
        EndIf
    EndIf

    Local $iPrevOnEvent = Opt("GUIOnEventMode", 0)
    Local $hDlg = GUICreate($sTitle, $iDlgW, $iDlgH, $iDlgX, $iDlgY, _
        BitOR($WS_DLGFRAME, $WS_POPUP, $WS_CAPTION), -1, $g_hMain)
    GUISetBkColor(0xF7F7F7, $hDlg)

    ; ---- Header ----
    GUICtrlCreateLabel($sTitle, 15, 10, $iDlgW - 30, 22)
    GUICtrlSetFont(-1, 12, 700, 0, "Segoe UI")
    GUICtrlSetColor(-1, 0x003366)

    Local $sHeaderLine2 = "Review each field below. Labels in bold are the column being edited."
    If $sSubtitle <> "" Then $sHeaderLine2 = $sSubtitle & "  -  " & $sHeaderLine2
    GUICtrlCreateLabel($sHeaderLine2, 15, 34, $iDlgW - 30, 18)
    GUICtrlSetFont(-1, 8, 400, 2, "Segoe UI")
    GUICtrlSetColor(-1, 0x555555)

    GUICtrlCreateLabel("", 10, $iHeaderH - 4, $iDlgW - 20, 1)
    GUICtrlSetBkColor(-1, 0xCCCCCC)

    Local $iLabelX = 15, $iNumW = 22, $iLabelW = 205
    Local $iInputX = 245, $iInputW = $iDlgW - $iInputX - 20
    Local $y = $iHeaderH + 6

    Local $aInputs[$nF]
    Local $aIsCombo[$nF]
    For $i = 0 To $nF - 1
        If Mod($i, 2) = 1 Then
            Local $hBg = GUICtrlCreateLabel("", 10, $y - 3, $iDlgW - 20, $iRowH)
            GUICtrlSetBkColor($hBg, 0xECECEC)
            GUICtrlSetState($hBg, $GUI_DISABLE)
        EndIf

        GUICtrlCreateLabel(($i + 1) & ".", $iLabelX, $y + 4, $iNumW, 18)
        GUICtrlSetFont(-1, 9, 700, 0, "Segoe UI")
        GUICtrlSetColor(-1, 0x888888)
        GUICtrlSetBkColor(-1, -2)

        GUICtrlCreateLabel($aFields[$i] & ":", $iLabelX + $iNumW, $y + 4, $iLabelW, 18)
        GUICtrlSetFont(-1, 9, 700, 0, "Segoe UI")
        GUICtrlSetColor(-1, 0x222222)
        GUICtrlSetBkColor(-1, -2)

        Local $sVal = ""
        If IsArray($aValues) And $i < UBound($aValues) Then $sVal = $aValues[$i]

        Local $sLookupKind = _DialogLookupKind($sTitle, $aFields[$i])
        If $sLookupKind <> "" Then
            $aIsCombo[$i] = True
            $aInputs[$i] = GUICtrlCreateCombo("", $iInputX, $y + 2, $iInputW, 22, BitOR($CBS_DROPDOWNLIST, $WS_VSCROLL))
            GUICtrlSetFont(-1, 10, 400, 0, "Segoe UI")
            GUICtrlSetBkColor(-1, 0xFFFFFF)
            Local $sOptions = _DialogLookupOptions($sLookupKind, $aFields, $aValues)
            Local $bAllowBlank = _DialogLookupAllowsBlank($sTitle, $aFields[$i], $aFields, $aValues)
            _ComboSetPipeList($aInputs[$i], $sOptions, $sVal, $bAllowBlank, True)
        Else
            $aIsCombo[$i] = False
            $aInputs[$i] = GUICtrlCreateInput($sVal, $iInputX, $y + 2, $iInputW, 22)
            GUICtrlSetFont(-1, 10, 400, 0, "Segoe UI")
            GUICtrlSetBkColor(-1, 0xFFFFFF)
        EndIf

        If $bHasHints And $aHints[$i] <> "" Then
            GUICtrlCreateLabel($aHints[$i], $iInputX, $y + 26, $iInputW, 18)
            GUICtrlSetFont(-1, 8, 400, 2, "Segoe UI")
            GUICtrlSetColor(-1, 0x0066AA)
            GUICtrlSetBkColor(-1, -2)
            GUICtrlSetTip($aInputs[$i], $aHints[$i])
        EndIf

        $y += $iRowH
    Next

    $y += 6
    GUICtrlCreateLabel("", 10, $y, $iDlgW - 20, 1)
    GUICtrlSetBkColor(-1, 0xCCCCCC)
    $y += 10

    Local $btnOK = GUICtrlCreateButton("Save", $iDlgW - 230, $y, 100, 32)
    GUICtrlSetFont(-1, 10, 700, 0, "Segoe UI")
    GUICtrlSetBkColor($btnOK, 0xDDFFDD)

    Local $btnCancel = GUICtrlCreateButton("Cancel", $iDlgW - 120, $y, 100, 32)
    GUICtrlSetFont(-1, 10, 400, 0, "Segoe UI")

    If $g_hMain <> 0 Then GUISetState(@SW_DISABLE, $g_hMain)
    GUISetState(@SW_SHOW, $hDlg)
    If $nF > 0 Then ControlFocus($hDlg, "", $aInputs[0])

    Local $aResult[$nF]
    Local $bOK = False
    While 1
        Local $nMsg = GUIGetMsg()
        Switch $nMsg
            Case $GUI_EVENT_CLOSE, $btnCancel
                ExitLoop
            Case $btnOK
                For $i = 0 To $nF - 1
                    $aResult[$i] = StringStripWS(GUICtrlRead($aInputs[$i]), 3)
                Next
                Local $sValidation = _ValidateDialogInput($sTitle, $aFields, $aResult)
                If $sValidation <> "" Then
                    MsgBox(262144+48, "Review required fields", $sValidation,0,$g_hMain)
                    ContinueLoop
                EndIf
                $bOK = True
                ExitLoop
            Case Else
                For $i = 0 To $nF - 1
                    If $aIsCombo[$i] And $nMsg = $aInputs[$i] Then
                        _DialogRefreshLookupCombos($sTitle, $aFields, $aInputs, $aIsCombo)
                        ExitLoop
                    EndIf
                Next
        EndSwitch
        Sleep(15)
    WEnd

    If $g_hMain <> 0 Then GUISetState(@SW_ENABLE, $g_hMain)
    GUIDelete($hDlg)
    If $g_hMain <> 0 Then WinActivate($g_hMain)
    Opt("GUIOnEventMode", $iPrevOnEvent)

    If $bOK Then Return $aResult
    Return 0
EndFunc

Func _IsNumericText($sVal)
    Local $s = StringStripWS($sVal, 3)
    If $s = "" Then Return False
    Return StringRegExp($s, "^-?\d+([\.,]\d+)?$")
EndFunc

Func _IsIntegerText($sVal)
    Local $s = StringStripWS($sVal, 3)
    If $s = "" Then Return False
    Return StringRegExp($s, "^-?\d+$")
EndFunc

Func _LooksLikeDate($sVal)
    Return StringRegExp(StringStripWS($sVal, 3), "^\d{2}/\d{2}/\d{4}$")
EndFunc

Func _LooksLikeDateTime($sVal)
    Return StringRegExp(StringStripWS($sVal, 3), "^\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2}$")
EndFunc

Func _AppendDialogError(ByRef $sErrors, $sMsg)
    $sErrors &= "- " & $sMsg & @CRLF
EndFunc

Func _ValidateDialogInput($sTitle, ByRef $aFields, ByRef $aValues)
    Local $sErrors = ""
    Local $sTitleNorm = StringLower($sTitle)

    For $i = 0 To UBound($aFields) - 1
        Local $sField = $aFields[$i]
        Local $sFieldNorm = StringLower($sField)
        Local $sVal = StringStripWS($aValues[$i], 3)

        ; Do not reject spaces in identifiers.
        ; Existing Ortems databases may contain Item IDs, Routing IDs, WO IDs,
        ; machine IDs, calendar IDs, and other master-data keys with spaces.
        ; Integrity is enforced through lookup validation, not through a "no spaces" rule.

        If StringInStr($sFieldNorm, "day") And Not StringInStr($sFieldNorm, "name") Then
            If Not _IsIntegerText($sVal) Or Number($sVal) < 1 Or Number($sVal) > 7 Then _AppendDialogError($sErrors, $sField & " must be a number from 1 to 7.")
        EndIf

        If StringInStr($sFieldNorm, "time") Then
            If _NormalizeTimeText($sVal) = "" Then _AppendDialogError($sErrors, $sField & " must use HH:MM format, for example 08:00.")
        EndIf

        If StringInStr($sFieldNorm, "date") Then
            If StringInStr($sFieldNorm, "movement") Then
                If Not _LooksLikeDate($sVal) Then _AppendDialogError($sErrors, $sField & " must use dd/mm/yyyy format.")
            Else
                If Not _LooksLikeDateTime($sVal) Then _AppendDialogError($sErrors, $sField & " must use dd/mm/yyyy hh:mm format.")
            EndIf
        EndIf

        If StringInStr($sFieldNorm, "qty") Or StringInStr($sFieldNorm, "duration") Or StringInStr($sFieldNorm, "setup") Or StringInStr($sFieldNorm, "break") Or StringInStr($sFieldNorm, "resource count") Or StringInStr($sFieldNorm, "phase") Then
            If $sVal <> "" And Not _IsNumericText($sVal) Then _AppendDialogError($sErrors, $sField & " must be numeric.")
        EndIf
    Next

    If StringInStr($sTitleNorm, "machine") Then
        If UBound($aValues) >= 12 Then
            If StringStripWS($aValues[0], 3) = "" Then _AppendDialogError($sErrors, "Site ID is required.")
            If StringStripWS($aValues[2], 3) = "" Then _AppendDialogError($sErrors, "CT ID is required.")
            If StringStripWS($aValues[7], 3) = "" Then _AppendDialogError($sErrors, "Machine ID is required.")
            If Not _IsMachineCalendarOptional($aValues[4]) And StringStripWS($aValues[10], 3) = "" Then _AppendDialogError($sErrors, "Calendar ID is required unless CT type is 4.")
        EndIf
    ElseIf StringInStr($sTitleNorm, "item") Then
        If UBound($aValues) >= 5 Then
            If StringStripWS($aValues[0], 3) = "" Then _AppendDialogError($sErrors, "Item ID is required.")
            If StringStripWS($aValues[2], 3) = "" Then _AppendDialogError($sErrors, "Type is required.")
            If _IsRawMaterialType($aValues[2]) Then
                If StringStripWS($aValues[3], 3) <> "" Then _AppendDialogError($sErrors, "Raw materials (MP) cannot have Version. Leave Version blank.")
                If StringStripWS($aValues[4], 3) <> "" Then _AppendDialogError($sErrors, "Raw materials (MP) cannot have Routing ID. Leave Routing ID blank.")
            Else
                If StringStripWS($aValues[3], 3) = "" Then _AppendDialogError($sErrors, "Version is required for manufactured items.")
                If StringStripWS($aValues[4], 3) = "" Then _AppendDialogError($sErrors, "Routing ID is required for manufactured items.")
            EndIf
        EndIf
    ElseIf StringInStr($sTitleNorm, "calendar") Then
        If UBound($aValues) > 0 And StringStripWS($aValues[0], 3) = "" Then _AppendDialogError($sErrors, "Calendar ID is required.")
    ElseIf StringInStr($sTitleNorm, "operation") Then
        If UBound($aValues) >= 4 Then
            If StringStripWS($aValues[0], 3) = "" Then _AppendDialogError($sErrors, "Operation ID is required.")
            If StringStripWS($aValues[2], 3) = "" Then _AppendDialogError($sErrors, "WC ID is required.")
            If StringStripWS($aValues[3], 3) = "" And Not _IsStandByMachine($aValues[3]) Then _AppendDialogError($sErrors, "Machine ID is required unless the value is STAND-BY.")
        EndIf
    ElseIf StringInStr($sTitleNorm, "routing") Then
        If UBound($aValues) >= 4 Then
            If StringStripWS($aValues[0], 3) = "" Then _AppendDialogError($sErrors, "Routing ID is required.")
            If StringStripWS($aValues[2], 3) = "" Then _AppendDialogError($sErrors, "Phase code is required.")
            If StringStripWS($aValues[3], 3) = "" Then _AppendDialogError($sErrors, "Operation ID is required.")
        EndIf
    ElseIf StringInStr($sTitleNorm, "work order") Then
        If UBound($aValues) >= 4 Then
            If StringStripWS($aValues[0], 3) = "" Then _AppendDialogError($sErrors, "WO ID is required.")
            If StringStripWS($aValues[1], 3) = "" Then _AppendDialogError($sErrors, "Item ID is required.")
            If StringStripWS($aValues[2], 3) = "" Then _AppendDialogError($sErrors, "Routing ID is required.")
            If StringStripWS($aValues[3], 3) = "" Then _AppendDialogError($sErrors, "Version is required.")
        EndIf
    EndIf

    $sErrors &= _ValidateDialogLookupValues($sTitle, $aFields, $aValues)

    If $sErrors <> "" Then
        Return "Some values need attention before this row can be saved:" & @CRLF & @CRLF & $sErrors
    EndIf
    Return ""
EndFunc

;=============================================================================
; LISTVIEW HELPER - visible line number column (last column)
;=============================================================================
Func _LV_AppendDataRow($hLV, $sData)
    GUICtrlCreateListViewItem($sData & "|" & (_GUICtrlListView_GetItemCount($hLV) + 1), $hLV)
EndFunc

Func _LV_Renumber($hLV)
    If $hLV = 0 Then Return
    Local $nCols = _GUICtrlListView_GetColumnCount($hLV)
    If $nCols < 1 Then Return
    Local $iLineCol = $nCols - 1
    Local $nRows = _GUICtrlListView_GetItemCount($hLV)
    For $i = 0 To $nRows - 1
        _GUICtrlListView_SetItemText($hLV, $i, $i + 1, $iLineCol)
    Next
EndFunc

Func _LV_GetControlFromNotifyHandle($hFrom)
    Local $aListViews[11] = [$g_hLV_Cal, $g_hLV_Mach, $g_hLV_Ops, $g_hLV_Rout, $g_hLV_Mat, $g_hLV_BOM, $g_hLV_WO, $g_hLV_WOL, $g_hLV_SR, $g_hLV_Cap, $g_hLV_Stk]
    For $i = 0 To UBound($aListViews) - 1
        If $aListViews[$i] = 0 Then ContinueLoop
        If HWnd(GUICtrlGetHandle($aListViews[$i])) = HWnd($hFrom) Then Return $aListViews[$i]
    Next
    Return 0
EndFunc

Func _LV_GetSortStateIndex($hLV)
    For $i = 0 To UBound($g_aLVSortState, 1) - 1
        If $g_aLVSortState[$i][0] = $hLV Then Return $i
    Next

    Local $n = UBound($g_aLVSortState, 1)
    ReDim $g_aLVSortState[$n + 1][3]
    $g_aLVSortState[$n][0] = $hLV
    $g_aLVSortState[$n][1] = -1
    $g_aLVSortState[$n][2] = False
    Return $n
EndFunc

Func _LV_GetRowTextPipe($hLV, $iRow, $nCols)
    Local $sRow = ""
    For $c = 0 To $nCols - 1
        If $c > 0 Then $sRow &= "|"
        $sRow &= _GUICtrlListView_GetItemText($hLV, $iRow, $c)
    Next
    Return $sRow
EndFunc

Func _LV_MakeSortKey($sValue)
    Local $s = StringStripWS(String($sValue), 3)
    If $s = "" Then Return "9|"

    ; Date formats used by the toolbox: dd/mm/yyyy and yyyy-mm-dd
    Local $aDate = StringRegExp($s, "^([0-9]{1,2})/([0-9]{1,2})/([0-9]{4})$", 1)
    If IsArray($aDate) Then Return "1|" & StringFormat("%04d%02d%02d", Number($aDate[2]), Number($aDate[1]), Number($aDate[0]))

    $aDate = StringRegExp($s, "^([0-9]{4})-([0-9]{1,2})-([0-9]{1,2})$", 1)
    If IsArray($aDate) Then Return "1|" & StringFormat("%04d%02d%02d", Number($aDate[0]), Number($aDate[1]), Number($aDate[2]))

    ; Time format: HH:MM or HH:MM:SS
    Local $aTime = StringRegExp($s, "^([0-9]{1,2}):([0-9]{2})(:([0-9]{2}))?$", 1)
    If IsArray($aTime) Then
        Local $iSeconds = (Number($aTime[0]) * 3600) + (Number($aTime[1]) * 60)
        If UBound($aTime) >= 4 Then $iSeconds += Number($aTime[3])
        Return "2|" & StringFormat("%08d", $iSeconds)
    EndIf

    ; Numeric values - avoids lexical problems like 10 being sorted before 2
    Local $sNum = StringReplace($s, ",", ".")
    If StringRegExp($sNum, "^-?[0-9]+(\.[0-9]+)?$") Then Return "3|" & StringFormat("%020.6f", Number($sNum) + 1000000000000)

    Return "4|" & StringLower($s)
EndFunc

Func _LV_SortByColumn($hLV, $iCol)
    If $hLV = 0 Then Return False

    Local $nRows = _GUICtrlListView_GetItemCount($hLV)
    Local $nCols = _GUICtrlListView_GetColumnCount($hLV)
    If $nRows < 2 Or $nCols < 1 Then Return False
    If $iCol < 0 Or $iCol >= $nCols Then Return False

    Local $iState = _LV_GetSortStateIndex($hLV)
    Local $bDescending = False
    If $g_aLVSortState[$iState][1] = $iCol Then
        $bDescending = Not $g_aLVSortState[$iState][2]
    Else
        $bDescending = False
    EndIf

    $g_aLVSortState[$iState][1] = $iCol
    $g_aLVSortState[$iState][2] = $bDescending

    Local $aRows[$nRows][2]
    For $i = 0 To $nRows - 1
        $aRows[$i][0] = _LV_MakeSortKey(_GUICtrlListView_GetItemText($hLV, $i, $iCol)) & "|" & StringFormat("%010d", $i)
        $aRows[$i][1] = _LV_GetRowTextPipe($hLV, $i, $nCols)
    Next

    _ArraySort($aRows, Number($bDescending), 0, 0, 0)

    _GUICtrlListView_BeginUpdate($hLV)
    _GUICtrlListView_DeleteAllItems($hLV)
    For $i = 0 To $nRows - 1
        GUICtrlCreateListViewItem($aRows[$i][1], $hLV)
    Next
    _LV_Renumber($hLV)
    _GUICtrlListView_EndUpdate($hLV)

    If $g_lblFooter <> 0 Then
        Local $sDir = "ascending"
        If $bDescending Then $sDir = "descending"
        GUICtrlSetData($g_lblFooter, "Table sorted by column " & ($iCol + 1) & " (" & $sDir & ").")
    EndIf

    Return True
EndFunc

;=============================================================================
; GENERIC EDIT HELPER - reads the selected row of a ListView, opens the edit
; dialog pre-filled with current values, and writes the result back in place.
; Returns True on save, False on cancel / no selection.
;=============================================================================
Func _LV_EditSelectedRow($hLV, $sTitle, $aFields, $aHints = "")
    Local $iSel = _GUICtrlListView_GetSelectionMark($hLV)
    If $iSel < 0 Then
        Local $sSel = _GUICtrlListView_GetSelectedIndices($hLV)
        If $sSel <> "" Then $iSel = Number($sSel)
    EndIf
    If $iSel < 0 Then
        MsgBox(262144+48, "No row selected", "Please click on a row in the list first, then press Edit.",0,$g_hMain)
        Return False
    EndIf

    Local $nCols  = _GUICtrlListView_GetColumnCount($hLV)
    Local $nDataCols = $nCols - 1
    Local $nTotal = _GUICtrlListView_GetItemCount($hLV)
    Local $aValues[$nDataCols]
    For $c = 0 To $nDataCols - 1
        $aValues[$c] = _GUICtrlListView_GetItemText($hLV, $iSel, $c)
    Next

    Local $sSubtitle = "Editing row " & ($iSel + 1) & " of " & $nTotal
    Local $aResult = _ShowRowDialog($sTitle, $aFields, $aValues, $aHints, $sSubtitle)
    If Not IsArray($aResult) Then Return False

    ; Write new values back into the same row (preserves position and selection)
    For $c = 0 To $nDataCols - 1
        If $c < UBound($aResult) Then
            _GUICtrlListView_SetItemText($hLV, $iSel, $aResult[$c], $c)
        EndIf
    Next
    Return True
EndFunc

;=============================================================================
; EVENT HANDLERS - CALENDARS CRUD
;=============================================================================
Func _Cal_Fields()
    Local $a[] = ["Calendar ID", "Calendar name", "Start day", "Start time", "End day", "End time"]
    Return $a
EndFunc
Func _Cal_Hints()
    Local $a[] = [ _
        "Unique calendar ID exactly as used in Ortems. Spaces are allowed. Example: Cal 1x8", _
        "Human-friendly label. Example: Cal 1x8", _
        "Day of week where the shift starts: 1=Mon ... 7=Sun", _
        "Shift start time in HH:MM (24h). Example: 08:00", _
        "Day of week where the shift ends (same as start for intra-day shifts)", _
        "Shift end time in HH:MM (24h). Example: 17:00"]
    Return $a
EndFunc
Func _Cal_Add()
    Local $aFields = _Cal_Fields()
    Local $aHints  = _Cal_Hints()
    Local $aVals[] = ["Cal 1x8", "Cal 1x8", "1", "08:00", "1", "17:00"]
    Local $aResult = _ShowRowDialog("Add calendar", $aFields, $aVals, $aHints)
    If IsArray($aResult) Then _AddCalendar($aResult[0], $aResult[1], $aResult[2], $aResult[3], $aResult[4], $aResult[5])
EndFunc
Func _Cal_Edit()
    _LV_EditSelectedRow($g_hLV_Cal, "Edit calendar", _Cal_Fields(), _Cal_Hints())
EndFunc
Func _Cal_Del()
    Local $sSel = _GUICtrlListView_GetSelectedIndices($g_hLV_Cal)
    If $sSel = "" Then
        MsgBox(262144+48, "No row selected", "Please select a row first.",0,$g_hMain)
        Return
    EndIf
    _GUICtrlListView_DeleteItem($g_hLV_Cal, Number($sSel))
    _LV_Renumber($g_hLV_Cal)
EndFunc
Func _Cal_DelAll()
    _GUICtrlListView_DeleteAllItems($g_hLV_Cal)
    _LV_Renumber($g_hLV_Cal)
EndFunc
Func _Cal_LoadEx()
    _LoadExampleCalendars()
EndFunc

;=============================================================================
; HANDLERS - MACHINES
;=============================================================================
Func _Mach_Fields()
    Local $a[] = ["Site ID", "Site name", "CT ID", "CT name", "CT type", "Section ID", "Section name", "Machine ID", "Machine name", "Machine type", "Calendar ID", "Capacity calendar ID"]
    Return $a
EndFunc
Func _Mach_Hints()
    Local $a[] = [ _
        "Plant / site code. Example: LYON", _
        "Plant / site label", _
        "Work center / CT ID exactly as used in Ortems. Spaces are allowed. Example: Milling Center", _
        "Work center human label", _
        "Capacity: 2 = Finite (bottleneck), 4 = Infinite", _
        "Section code (organisational sub-unit)", _
        "Section label", _
        "Machine ID exactly as used in Ortems. Spaces are allowed. Example: Milling 1", _
        "Machine human label", _
        "NR=Standard, BA=Batch, RN=Run, CU=Tank", _
        "Work-hours calendar (must exist in the Calendars tab)", _
        "Capacity calendar for load calculation (often same as Calendar ID)"]
    Return $a
EndFunc
Func _Mach_Add()
    Local $aFields = _Mach_Fields()
    Local $aHints  = _Mach_Hints()
    Local $aVals[] = ["LYON","LYON","Milling Center","Milling","2","Section A","Section A","Milling 1","Milling 1","NR","Cal 1x8","Cal 1x8"]
    Local $aResult = _ShowRowDialog("Add machine", $aFields, $aVals, $aHints)
    If IsArray($aResult) Then
        _AddMachine($aResult[0],$aResult[1],$aResult[2],$aResult[3],$aResult[4],$aResult[5],$aResult[6],$aResult[7],$aResult[8],$aResult[9],$aResult[10],$aResult[11])
    EndIf
EndFunc
Func _Mach_Edit()
    _LV_EditSelectedRow($g_hLV_Mach, "Edit machine", _Mach_Fields(), _Mach_Hints())
EndFunc
Func _Mach_Del()
    Local $sSel = _GUICtrlListView_GetSelectedIndices($g_hLV_Mach)
    If $sSel = "" Then
        MsgBox(262144+48, "No row selected", "Please select a row first.",0,$g_hMain)
        Return
    EndIf
    _GUICtrlListView_DeleteItem($g_hLV_Mach, Number($sSel))
    _LV_Renumber($g_hLV_Mach)
EndFunc
Func _Mach_DelAll()
    _GUICtrlListView_DeleteAllItems($g_hLV_Mach)
    _LV_Renumber($g_hLV_Mach)
EndFunc
Func _Mach_LoadEx()
    _LoadExampleMachines()
EndFunc

;=============================================================================
; HANDLERS - OPERATIONS
;=============================================================================
Func _Ops_Fields()
    Local $a[] = ["Operation ID","Operation name","WC / CT ID","Machine ID","Ref qty","Ref duration","Unit","Setup time","Break time","Interruptible"]
    Return $a
EndFunc
Func _Ops_Hints()
    Local $a[] = [ _
        "Unique operation code. Example: MILL", _
        "Human label. Example: Milling Axes", _
        "Work center / grouping code where this op runs (must exist in Machines tab)", _
        "Machine that executes this op (must exist in Machines tab)", _
        "Reference quantity the duration is measured for (integer)", _
        "Reference duration (in Unit below) to produce Ref qty", _
        "D=Days, H=Hours, C=Hundredths of an hour", _
        "Setup time before the op starts (hours)", _
        "Break / teardown time after op ends (hours)", _
        "1 = op can be interrupted and resumed, 0 = must run continuously"]
    Return $a
EndFunc
Func _Ops_Add()
    Local $aFields = _Ops_Fields()
    Local $aHints  = _Ops_Hints()
    Local $aVals[] = ["MILL","Milling","Milling Center","Milling 1","26","1","H","0","0","1"]
    Local $aResult = _ShowRowDialog("Add operation", $aFields, $aVals, $aHints)
    If IsArray($aResult) Then _AddOperation($aResult[0],$aResult[1],$aResult[2],$aResult[3],$aResult[4],$aResult[5],$aResult[6],$aResult[7],$aResult[8],$aResult[9])
EndFunc
Func _Ops_Edit()
    _LV_EditSelectedRow($g_hLV_Ops, "Edit operation", _Ops_Fields(), _Ops_Hints())
EndFunc
Func _Ops_Del()
    Local $sSel = _GUICtrlListView_GetSelectedIndices($g_hLV_Ops)
    If $sSel = "" Then
        MsgBox(262144+48, "No row selected", "Please select a row first.",0,$g_hMain)
        Return
    EndIf
    _GUICtrlListView_DeleteItem($g_hLV_Ops, Number($sSel))
    _LV_Renumber($g_hLV_Ops)
EndFunc
Func _Ops_DelAll()
    _GUICtrlListView_DeleteAllItems($g_hLV_Ops)
    _LV_Renumber($g_hLV_Ops)
EndFunc
Func _Ops_LoadEx()
    _LoadExampleOperations()
EndFunc

;=============================================================================
; HANDLERS - ROUTINGS
;=============================================================================
Func _Rout_Fields()
    Local $a[] = ["Routing ID","Routing name","Phase code","Operation ID","Phase name"]
    Return $a
EndFunc
Func _Rout_Hints()
    Local $a[] = [ _
        "Routing ID exactly as used in Ortems. Spaces are allowed. Example: STD AX", _
        "Human label. Example: Axes Standard", _
        "Phase sequence number (10, 20, 30, ...). Lower = earlier in the routing", _
        "Operation to run at this phase (must exist in Operations tab)", _
        "Free-text label for this specific phase"]
    Return $a
EndFunc
Func _Rout_Add()
    Local $aFields = _Rout_Fields()
    Local $aHints  = _Rout_Hints()
    Local $aVals[] = ["STD AX","Axes Standard","10","MILL","Milling Axes Standard"]
    Local $aResult = _ShowRowDialog("Add routing", $aFields, $aVals, $aHints)
    If IsArray($aResult) Then _AddRouting($aResult[0],$aResult[1],$aResult[2],$aResult[3],$aResult[4])
EndFunc
Func _Rout_Edit()
    _LV_EditSelectedRow($g_hLV_Rout, "Edit routing", _Rout_Fields(), _Rout_Hints())
EndFunc
Func _Rout_Del()
    Local $sSel = _GUICtrlListView_GetSelectedIndices($g_hLV_Rout)
    If $sSel = "" Then
        MsgBox(262144+48, "No row selected", "Please select a row first.",0,$g_hMain)
        Return
    EndIf
    _GUICtrlListView_DeleteItem($g_hLV_Rout, Number($sSel))
    _LV_Renumber($g_hLV_Rout)
EndFunc
Func _Rout_DelAll()
    _GUICtrlListView_DeleteAllItems($g_hLV_Rout)
    _LV_Renumber($g_hLV_Rout)
EndFunc
Func _Rout_LoadEx()
    _LoadExampleRoutings()
EndFunc

;=============================================================================
; HANDLERS - ITEMS
;=============================================================================
Func _Mat_Fields()
    Local $a[] = ["Item ID","Item name","Type","Version","Routing ID","On-hand qty"]
    Return $a
EndFunc
Func _Mat_Hints()
    Local $a[] = [ _
        "Item ID exactly as used in Ortems. Spaces are allowed. Example: STD Gearbox", _
        "Human label. Example: Gearbox Standard", _
        "MP = Raw material, SF = Semi-finished, PF = Finished good", _
        "Version code for manufactured items (SF/PF). Leave blank for MP raw materials.", _
        "Routing used to produce this item (leave blank for raw materials)", _
        "Current stock on hand (numeric). Starting inventory for this item/version"]
    Return $a
EndFunc
Func _Mat_Add()
    Local $aFields = _Mat_Fields()
    Local $aHints  = _Mat_Hints()
    Local $aVals[] = ["STD Gearbox","Gearbox Standard","PF","STD","STD GB","0"]
    Local $aResult = _ShowRowDialog("Add item", $aFields, $aVals, $aHints)
    If IsArray($aResult) Then _AddMaterial($aResult[0],$aResult[1],$aResult[2],$aResult[3],$aResult[4],$aResult[5])
EndFunc
Func _Mat_Edit()
    _LV_EditSelectedRow($g_hLV_Mat, "Edit item", _Mat_Fields(), _Mat_Hints())
EndFunc
Func _Mat_Del()
    Local $sSel = _GUICtrlListView_GetSelectedIndices($g_hLV_Mat)
    If $sSel = "" Then
        MsgBox(262144+48, "No row selected", "Please select a row first.",0,$g_hMain)
        Return
    EndIf
    _GUICtrlListView_DeleteItem($g_hLV_Mat, Number($sSel))
    _LV_Renumber($g_hLV_Mat)
EndFunc
Func _Mat_DelAll()
    _GUICtrlListView_DeleteAllItems($g_hLV_Mat)
    _LV_Renumber($g_hLV_Mat)
EndFunc
Func _Mat_LoadEx()
    _LoadExampleMaterials()
EndFunc

;=============================================================================
; HANDLERS - BOM
;=============================================================================
Func _BOM_Fields()
    Local $a[] = ["Parent item ID","Parent version","Component item ID","Routing ID","Phase","Ref qty","Required qty"]
    Return $a
EndFunc
Func _BOM_Hints()
    Local $a[] = [ _
        "Item that will be produced (parent). Must exist in Items tab", _
        "Version of the parent (must match a version in Items tab)", _
        "Component to consume (must exist in Items tab)", _
        "Routing of the parent (consumption happens within this routing)", _
        "Phase number where the component is consumed (e.g. 10)", _
        "Reference batch size (typically 1)", _
        "Qty of component needed to produce 'Ref qty' of parent"]
    Return $a
EndFunc
Func _BOM_Add()
    Local $aFields = _BOM_Fields()
    Local $aHints  = _BOM_Hints()
    Local $aVals[] = ["STD_Gearbox","STD","R_Axes","STD_GB","10","1","2"]
    Local $aResult = _ShowRowDialog("Add BOM row", $aFields, $aVals, $aHints)
    If IsArray($aResult) Then _AddBOM($aResult[0],$aResult[1],$aResult[2],$aResult[3],$aResult[4],$aResult[5],$aResult[6])
EndFunc
Func _BOM_Edit()
    _LV_EditSelectedRow($g_hLV_BOM, "Edit BOM row", _BOM_Fields(), _BOM_Hints())
EndFunc
Func _BOM_Del()
    Local $sSel = _GUICtrlListView_GetSelectedIndices($g_hLV_BOM)
    If $sSel = "" Then
        MsgBox(262144+48, "No row selected", "Please select a row first.",0,$g_hMain)
        Return
    EndIf
    _GUICtrlListView_DeleteItem($g_hLV_BOM, Number($sSel))
    _LV_Renumber($g_hLV_BOM)
EndFunc
Func _BOM_DelAll()
    _GUICtrlListView_DeleteAllItems($g_hLV_BOM)
    _LV_Renumber($g_hLV_BOM)
EndFunc
Func _BOM_LoadEx()
    _LoadExampleBOM()
EndFunc

;=============================================================================
; HANDLERS - WO
;=============================================================================
Func _WO_Fields()
    Local $a[] = ["WO number","Item ID","Routing ID","Version","Quantity","Start date","End date"]
    Return $a
EndFunc
Func _WO_Hints()
    Local $a[] = [ _
        "Unique work-order number. Example: WO001", _
        "Item to produce (must exist in Items tab)", _
        "Routing used for this WO (must exist in Routings tab)", _
        "Version of the item (STD, SPT, 00, ...)", _
        "Quantity to produce (numeric)", _
        "Start date in dd/mm/yyyy hh:mm. Example: 01/02/2025 00:00", _
        "End / due date in dd/mm/yyyy hh:mm. Example: 28/02/2025 23:59"]
    Return $a
EndFunc
Func _WO_Add()
    Local $aFields = _WO_Fields()
    Local $aHints  = _WO_Hints()
    Local $aVals[] = ["WO001","STD_Gearbox","STD_GB","STD","10","01/02/2025 00:00","28/02/2025 23:59"]
    Local $aResult = _ShowRowDialog("Add work order", $aFields, $aVals, $aHints)
    If IsArray($aResult) Then _AddWO($aResult[0],$aResult[1],$aResult[2],$aResult[3],$aResult[4],$aResult[5],$aResult[6])
EndFunc
Func _WO_Edit()
    _LV_EditSelectedRow($g_hLV_WO, "Edit work order", _WO_Fields(), _WO_Hints())
EndFunc
Func _WO_Del()
    Local $sSel = _GUICtrlListView_GetSelectedIndices($g_hLV_WO)
    If $sSel = "" Then
        MsgBox(262144+48, "No row selected", "Please select a row first.",0,$g_hMain)
        Return
    EndIf
    _GUICtrlListView_DeleteItem($g_hLV_WO, Number($sSel))
    _LV_Renumber($g_hLV_WO)
EndFunc
Func _WO_DelAll()
    _GUICtrlListView_DeleteAllItems($g_hLV_WO)
    _LV_Renumber($g_hLV_WO)
EndFunc
Func _WO_LoadEx()
    _LoadExampleWOs()
EndFunc

;=============================================================================
; HANDLERS - WO LINKS
;=============================================================================
Func _WOL_Fields()
    Local $a[] = ["Predecessor WO","Pred routing","Pred phase","Successor WO","Succ routing","Succ phase","Link type"]
    Return $a
EndFunc
Func _WOL_Hints()
    Local $a[] = [ _
        "The WO that must complete first (must exist in Work Orders tab)", _
        "Routing of the predecessor WO", _
        "Phase number in the predecessor routing (e.g. 10)", _
        "The WO that depends on the predecessor", _
        "Routing of the successor WO", _
        "Phase number in the successor routing", _
        "FS=Finish-Start, SS=Start-Start, FF=Finish-Finish, SF=Start-Finish"]
    Return $a
EndFunc
Func _WOL_Add()
    Local $aFields = _WOL_Fields()
    Local $aHints  = _WOL_Hints()
    Local $aVals[] = ["WO001","STD_GB","10","WO002","STD_GB","10","FS"]
    Local $aResult = _ShowRowDialog("Add WO link", $aFields, $aVals, $aHints)
    If IsArray($aResult) Then _AddWOLink($aResult[0],$aResult[1],$aResult[2],$aResult[3],$aResult[4],$aResult[5],$aResult[6])
EndFunc
Func _WOL_Edit()
    _LV_EditSelectedRow($g_hLV_WOL, "Edit WO link", _WOL_Fields(), _WOL_Hints())
EndFunc
Func _WOL_Del()
    Local $sSel = _GUICtrlListView_GetSelectedIndices($g_hLV_WOL)
    If $sSel = "" Then
        MsgBox(262144+48, "No row selected", "Please select a row first.",0,$g_hMain)
        Return
    EndIf
    _GUICtrlListView_DeleteItem($g_hLV_WOL, Number($sSel))
    _LV_Renumber($g_hLV_WOL)
EndFunc
Func _WOL_DelAll()
    _GUICtrlListView_DeleteAllItems($g_hLV_WOL)
    _LV_Renumber($g_hLV_WOL)
EndFunc
Func _WOL_LoadEx()
    _LoadExampleWOLinks()
EndFunc

;=============================================================================
; HANDLERS - SECONDARY RESOURCES
;=============================================================================
Func _SR_Fields()
    Local $a[] = ["Operation ID","WC / CT ID","Machine ID","Qualification ID","Capacity calendar ID"]
    Return $a
EndFunc
Func _SR_Hints()
    Local $a[] = [ _
        "Operation that requires this secondary resource (must exist in Operations tab)", _
        "Work center where the op runs", _
        "Machine where the op runs (primary resource)", _
        "Labor/tool qualification required. Example: OPERATOR, WELDER", _
        "Capacity calendar controlling availability of the secondary resource"]
    Return $a
EndFunc
Func _SR_Add()
    Local $aFields = _SR_Fields()
    Local $aHints  = _SR_Hints()
    Local $aVals[] = ["MILL","Milling","Milling_1","OPERATOR","Cal_1x8"]
    Local $aResult = _ShowRowDialog("Add secondary resource", $aFields, $aVals, $aHints)
    If IsArray($aResult) Then _AddSR($aResult[0],$aResult[1],$aResult[2],$aResult[3],$aResult[4])
EndFunc
Func _SR_Edit()
    _LV_EditSelectedRow($g_hLV_SR, "Edit secondary resource", _SR_Fields(), _SR_Hints())
EndFunc
Func _SR_Del()
    Local $sSel = _GUICtrlListView_GetSelectedIndices($g_hLV_SR)
    If $sSel = "" Then
        MsgBox(262144+48, "No row selected", "Please select a row first.",0,$g_hMain)
        Return
    EndIf
    _GUICtrlListView_DeleteItem($g_hLV_SR, Number($sSel))
    _LV_Renumber($g_hLV_SR)
EndFunc
Func _SR_DelAll()
    _GUICtrlListView_DeleteAllItems($g_hLV_SR)
    _LV_Renumber($g_hLV_SR)
EndFunc
Func _SR_LoadEx()
    _LoadExampleSR()
EndFunc

;=============================================================================
; HANDLERS - CAPACITY
;=============================================================================
Func _Cap_Fields()
    Local $a[] = ["Capacity calendar ID","Start day","Start time","End day","End time","Resource count"]
    Return $a
EndFunc
Func _Cap_Hints()
    Local $a[] = [ _
        "Unique capacity calendar code. Example: Cap_1x8", _
        "Day of week (1=Mon ... 7=Sun)", _
        "Start time HH:MM. Example: 08:00", _
        "Day of week (1=Mon ... 7=Sun)", _
        "End time HH:MM. Example: 17:00", _
        "Number of concurrent resources available during this shift"]
    Return $a
EndFunc
Func _Cap_Add()
    Local $aFields = _Cap_Fields()
    Local $aHints  = _Cap_Hints()
    Local $aVals[] = ["Cap_1x8","1","08:00","1","17:00","2"]
    Local $aResult = _ShowRowDialog("Add capacity calendar", $aFields, $aVals, $aHints)
    If IsArray($aResult) Then _AddCap($aResult[0],$aResult[1],$aResult[2],$aResult[3],$aResult[4],$aResult[5])
EndFunc
Func _Cap_Edit()
    _LV_EditSelectedRow($g_hLV_Cap, "Edit capacity calendar", _Cap_Fields(), _Cap_Hints())
EndFunc
Func _Cap_Del()
    Local $sSel = _GUICtrlListView_GetSelectedIndices($g_hLV_Cap)
    If $sSel = "" Then
        MsgBox(262144+48, "No row selected", "Please select a row first.",0,$g_hMain)
        Return
    EndIf
    _GUICtrlListView_DeleteItem($g_hLV_Cap, Number($sSel))
    _LV_Renumber($g_hLV_Cap)
EndFunc
Func _Cap_DelAll()
    _GUICtrlListView_DeleteAllItems($g_hLV_Cap)
    _LV_Renumber($g_hLV_Cap)
EndFunc
Func _Cap_LoadEx()
    _LoadExampleCap()
EndFunc

;=============================================================================
; HANDLERS - INVENTORY
;=============================================================================
Func _Stk_Fields()
    Local $a[] = ["Item ID","Routing ID","Version","Movement date","Quantity"]
    Return $a
EndFunc
Func _Stk_Hints()
    Local $a[] = [ _
        "Item being moved (must exist in Items tab)", _
        "Routing associated with the item (for receipts linked to production)", _
        "Item version (STD, SPT, 00, ...)", _
        "Movement date in dd/mm/yyyy. Example: 15/02/2025", _
        "Positive = receipt (+), negative = issue (-)"]
    Return $a
EndFunc
Func _Stk_Add()
    Local $aFields = _Stk_Fields()
    Local $aHints  = _Stk_Hints()
    Local $aVals[] = ["R_Axes","","00","15/02/2025","500"]
    Local $aResult = _ShowRowDialog("Add inventory movement", $aFields, $aVals, $aHints)
    If IsArray($aResult) Then _AddStk($aResult[0],$aResult[1],$aResult[2],$aResult[3],$aResult[4])
EndFunc
Func _Stk_Edit()
    _LV_EditSelectedRow($g_hLV_Stk, "Edit inventory movement", _Stk_Fields(), _Stk_Hints())
EndFunc
Func _Stk_Del()
    Local $sSel = _GUICtrlListView_GetSelectedIndices($g_hLV_Stk)
    If $sSel = "" Then
        MsgBox(262144+48, "No row selected", "Please select a row first.",0,$g_hMain)
        Return
    EndIf
    _GUICtrlListView_DeleteItem($g_hLV_Stk, Number($sSel))
    _LV_Renumber($g_hLV_Stk)
EndFunc
Func _Stk_DelAll()
    _GUICtrlListView_DeleteAllItems($g_hLV_Stk)
    _LV_Renumber($g_hLV_Stk)
EndFunc
Func _Stk_LoadEx()
    _LoadExampleStk()
EndFunc

;=============================================================================
; ROW APPEND HELPERS for WOL / SR / Cap / Stk (kept here so Add and
; LoadExample functions share the same column layout)
;=============================================================================
Func _AddWOLink($sPredWO, $sPredRot, $nPredPh, $sSuccWO, $sSuccRot, $nSuccPh, $sType)
    _LV_AppendDataRow($g_hLV_WOL, $sPredWO & "|" & $sPredRot & "|" & $nPredPh & "|" & _
        $sSuccWO & "|" & $sSuccRot & "|" & $nSuccPh & "|" & $sType)
EndFunc

Func _AddSR($sOp, $sCT, $sMach, $sQualif, $sCapCal)
    _LV_AppendDataRow($g_hLV_SR, $sOp & "|" & $sCT & "|" & $sMach & "|" & $sQualif & "|" & $sCapCal)
EndFunc

Func _AddCap($sID, $nDiaI, $sHoraI, $nDiaF, $sHoraF, $nRes)
    _LV_AppendDataRow($g_hLV_Cap, $sID & "|" & $nDiaI & "|" & $sHoraI & "|" & $nDiaF & "|" & $sHoraF & "|" & $nRes)
EndFunc

Func _AddStk($sItem, $sRot, $sVer, $sDate, $nQty)
    _LV_AppendDataRow($g_hLV_Stk, $sItem & "|" & $sRot & "|" & $sVer & "|" & $sDate & "|" & $nQty)
EndFunc

;=============================================================================
; SAMPLE DATA for tabs that didn't have an example loader before
;=============================================================================
Func _LoadExampleWOLinks()
    ; Example: WO002 must wait for WO001 to finish (Finish-Start on phase 10)
    _AddWOLink("WO001", "STD_GB", 10, "WO002", "STD_GB", 10, "FS")
    _AddWOLink("WO002", "STD_GB", 10, "WO003", "SPT_GB", 10, "FS")
EndFunc

Func _LoadExampleSR()
    ; Example: milling and drilling operations require an operator;
    ; assembly requires a qualified technician
    _AddSR("MILL", "Milling",        "Milling_1",    "OPERATOR",  "Cal_1x8")
    _AddSR("MILL", "Milling",        "Milling_2",    "OPERATOR",  "Cal_1x8")
    _AddSR("DRIL", "Drilling",       "Drilling_1",   "OPERATOR",  "Cal_1x8")
    _AddSR("ASSY", "Assembly_Robot", "Assy_Robot_1", "TECHNICIAN","Cal_2x8")
EndFunc

Func _LoadExampleCap()
    ; Example: 1x8 shift Mon-Fri with 2 resources
    _AddCap("Cap_1x8", 1, "08:00", 1, "17:00", 2)
    _AddCap("Cap_1x8", 2, "08:00", 2, "17:00", 2)
    _AddCap("Cap_1x8", 3, "08:00", 3, "17:00", 2)
    _AddCap("Cap_1x8", 4, "08:00", 4, "17:00", 2)
    _AddCap("Cap_1x8", 5, "08:00", 5, "17:00", 2)
    ; 2x8 shift
    _AddCap("Cap_2x8", 1, "08:00", 2, "00:00", 4)
EndFunc

Func _LoadExampleStk()
    ; Example: raw material receipts and a semi-finished issue
    _AddStk("R_Axes",  "", "",  "15/02/2025",  500)
    _AddStk("R_Housing", "", "",  "15/02/2025",  200)
    _AddStk("STD_Axes","STD_AX", "STD", "20/02/2025", -10)
EndFunc

;=============================================================================
; DATABASE CONNECTION
;=============================================================================
Func _TestConnection()
    $g_sServer   = GUICtrlRead($g_edtServer)
    $g_sDatabase = GUICtrlRead($g_edtDatabase)
    Local $sAuth  = GUICtrlRead($g_cmbAuth)

    If $g_sServer = "" Or $g_sDatabase = "" Then
        MsgBox(262144+48, "Error", "Server and database are required.",0,$g_hMain)
        Return
    EndIf

    If $sAuth = "Windows Authentication" Then
        $g_sConnStr = "Driver={SQL Server};Server=" & $g_sServer & ";Database=" & $g_sDatabase & ";Trusted_Connection=yes;"
    Else
        $g_sConnStr = "Driver={SQL Server};Server=" & $g_sServer & ";Database=" & $g_sDatabase & ";UID=" & GUICtrlRead($g_edtUser) & ";PWD=" & GUICtrlRead($g_edtPass) & ";"
    EndIf

    GUICtrlSetData($g_edtConnStr, $g_sConnStr)
    _Log("Connection test started for database " & $g_sDatabase & " on server " & $g_sServer)
    _LogVerbose("Connection string: " & _MaskConnectionString($g_sConnStr))

    ; Test the connection via ADO
    Local $oConn = ObjCreate("ADODB.Connection")
    If Not IsObj($oConn) Then
        GUICtrlSetData($g_lblStatus, "Error: ADO not available")
        GUICtrlSetColor($g_lblStatus, 0xCC0000)
        MsgBox(262144+16, "Error", "Could not create ADODB.Connection." & @CRLF & "Check that the SQL Server ODBC / OLE DB provider is installed.",0,$g_hMain)
        Return
    EndIf

    $oConn.ConnectionString = $g_sConnStr
    $oConn.ConnectionTimeout = 5

    Local $bOK = False
    Local $sErr = ""
    ; Try to open
    ; The COM error handler captures connection errors without crashing the app
    $oConn.Open($g_sConnStr)
    If $oConn.State = 1 Then ; adStateOpen = 1
        $bOK = True
        $oConn.Close()
    Else
        $sErr = "Could not open the connection."
    EndIf

    If $bOK Then
        $g_bConnected = True
		MsgBox(262144+262144,"","Connected: " & $g_sDatabase & "@" & $g_sServer,0,$g_hMain)
        GUICtrlSetData($g_lblStatus, "Connected: " & $g_sDatabase & "@" & $g_sServer)
        GUICtrlSetColor($g_lblStatus, 0x007700)
        _SaveSettings()   ; persist server + database immediately
        MsgBox(262144+64, "Connection", "Successfully connected to the database!" & @CRLF & @CRLF & "Database: " & $g_sDatabase & @CRLF & "Server: " & $g_sServer,0,$g_hMain)
    Else
        $g_bConnected = False
        GUICtrlSetData($g_lblStatus, "Connection error")
        GUICtrlSetColor($g_lblStatus, 0xCC0000)
        MsgBox(262144+16, "Connection error", "Could not connect to the database." & @CRLF & @CRLF & $sErr & @CRLF & @CRLF & "Connection string used:" & @CRLF & $g_sConnStr,0,$g_hMain)
    EndIf
EndFunc

Func _InspectTable()
    If Not $g_bConnected Then
        MsgBox(262144+48, "Not connected", "Test the connection first (tab 1. Database).",0,$g_hMain)
        Return
    EndIf

    Local $sTable = InputBox("Inspect Table", "Enter the table name to inspect (e.g. B_NOME, B_MACH, B_OF):", "B_NOME", "", 320, 130)
    If $sTable = "" Or @error Then Return
    $sTable = StringStripWS(StringUpper($sTable), 3)

    Local $oConn = ObjCreate("ADODB.Connection")
    If Not IsObj($oConn) Then Return
    $g_sLastComError = ""
    $oConn.Open($g_sConnStr)
    If $g_sLastComError <> "" Or $oConn.State <> 1 Then Return

    Local $sQuery = "SELECT COLUMN_NAME, DATA_TYPE, CHARACTER_MAXIMUM_LENGTH, IS_NULLABLE " & _
        "FROM INFORMATION_SCHEMA.COLUMNS " & _
        "WHERE TABLE_NAME='" & $sTable & "' ORDER BY ORDINAL_POSITION"
    _LogVerboseSQL("Inspect table " & $sTable, $sQuery)
    $g_sLastComError = ""
    Local $oRS = $oConn.Execute($sQuery)

    If $g_sLastComError <> "" Or Not IsObj($oRS) Then
        $oConn.Close()
        MsgBox(262144+16, "Error", "Could not query schema: " & $g_sLastComError,0,$g_hMain)
        Return
    EndIf

    Local $sResult = "Columns of table [" & $sTable & "]:" & @CRLF & @CRLF
    Local $nCols = 0
    While Not $oRS.EOF
        $nCols += 1
        Local $sCol  = $oRS.Fields(0).Value
        Local $sType = $oRS.Fields(1).Value
        Local $sLen  = $oRS.Fields(2).Value
        Local $sNull = $oRS.Fields(3).Value
        If $sLen <> "" And $sLen <> "Null" Then $sType &= "(" & $sLen & ")"
        $sResult &= "  " & $sCol & "  [" & $sType & "]" & ($sNull = "YES" ? " NULL" : " NOT NULL") & @CRLF
        $oRS.MoveNext()
    WEnd
    $oRS.Close()
    $oConn.Close()

    If $nCols = 0 Then
        MsgBox(262144+48, "Not found", "Table '" & $sTable & "' not found in the database.",0,$g_hMain)
    Else
        MsgBox(262144+64, "Schema: " & $sTable & " (" & $nCols & " columns)", $sResult,0,$g_hMain)
    EndIf
EndFunc

Func _OnAuthChange()
    Local $sAuth = GUICtrlRead($g_cmbAuth)
    If $sAuth = "SQL Server Authentication" Then
        GUICtrlSetState($g_edtUser, $GUI_ENABLE)
        GUICtrlSetState($g_edtPass, $GUI_ENABLE)
    Else
        GUICtrlSetState($g_edtUser, $GUI_DISABLE)
        GUICtrlSetState($g_edtPass, $GUI_DISABLE)
    EndIf
EndFunc

Func _Help_Click()
    FileInstall("Help.html", @ScriptDir & "\Help.html", 1)
    ShellExecute(@ScriptDir & "\Help.html")
EndFunc

;=============================================================================
; UI BUSY MODE / SAFE COMMAND WRAPPERS
;=============================================================================
Func _RegisterBusyControl($idCtrl)
    If $idCtrl = 0 Then Return $idCtrl
    Local $n = UBound($g_aBusyControls)
    For $i = 0 To $n - 1
        If $g_aBusyControls[$i] = $idCtrl Then Return $idCtrl
    Next
    ReDim $g_aBusyControls[$n + 1]
    $g_aBusyControls[$n] = $idCtrl
    Return $idCtrl
EndFunc

Func _ForceMainRedraw()
    If $g_hMain <> 0 Then DllCall("user32.dll", "bool", "RedrawWindow", "hwnd", $g_hMain, "ptr", 0, "ptr", 0, "uint", 0x0185)
EndFunc

Func _SetBusyControlsEnabled($bEnable)
    Local $n = UBound($g_aBusyControls)
    If $n <= 0 Then Return

    If Not $bEnable Then
        ReDim $g_aBusyStates[$n]
        For $i = 0 To $n - 1
            $g_aBusyStates[$i] = GUICtrlGetState($g_aBusyControls[$i])
            GUICtrlSetState($g_aBusyControls[$i], $GUI_DISABLE)
        Next
    Else
        For $i = 0 To $n - 1
            If $i < UBound($g_aBusyStates) And BitAND($g_aBusyStates[$i], $GUI_DISABLE) Then
                GUICtrlSetState($g_aBusyControls[$i], $GUI_DISABLE)
            Else
                GUICtrlSetState($g_aBusyControls[$i], $GUI_ENABLE)
            EndIf
        Next
        ReDim $g_aBusyStates[0]
        _OnAuthChange()
    EndIf
EndFunc

Func _SetUIBusy($bBusy, $sLabel = "")
    If $bBusy Then
        $g_iBusyDepth += 1
        If $g_iBusyDepth = 1 Then
            _LogVerbose("UI controls disabled while executing: " & $sLabel)
            If $g_hMain <> 0 Then GUISetCursor(15, 1, $g_hMain)
            If $g_lblFooter <> 0 Then
                $g_sBusyFooterBackup = GUICtrlRead($g_lblFooter)
                GUICtrlSetData($g_lblFooter, "Executing: " & $sLabel & " ...")
            EndIf
            _SetBusyControlsEnabled(False)
            _ForceMainRedraw()
        EndIf
    Else
        If $g_iBusyDepth > 0 Then $g_iBusyDepth -= 1
        If $g_iBusyDepth = 0 Then
            _SetBusyControlsEnabled(True)
            If $g_lblFooter <> 0 Then GUICtrlSetData($g_lblFooter, $g_sBusyFooterBackup)
            If $g_hMain <> 0 Then GUISetCursor(2, 0, $g_hMain)
            _ForceMainRedraw()
        EndIf
    EndIf
EndFunc

Func _RunWithBusy($sLabel, $sFunction)
    If $g_iBusyDepth > 0 Then Return
    _SetUIBusy(True, $sLabel)
    Local $vResult = Call($sFunction)
    _SetUIBusy(False, $sLabel)
    Return $vResult
EndFunc

Func _TestConnection_Click()
    _RunWithBusy("Test DB connection", "_TestConnection")
EndFunc

Func _InspectTable_Click()
    _RunWithBusy("Inspect table", "_InspectTable")
EndFunc

Func _ImportExcel_Click()
    _RunWithBusy("Import workbook", "_ImportExcel")
EndFunc

Func _ExportExcel_Click()
    _RunWithBusy("Export workbook", "_ExportExcel")
EndFunc

Func _LoadFromDB_Click()
    _RunWithBusy("Load from DB", "_LoadFromDB")
EndFunc

Func _ClearDatabase_Click()
    _RunWithBusy("Clear DB", "_ClearDatabase")
EndFunc

Func _GenerateSQL_Click()
    _RunWithBusy("Generate SQL", "_GenerateSQL")
EndFunc

Func _ExecuteSQL_Click()
    _RunWithBusy("Run SQL on DB", "_ExecuteSQL")
EndFunc

Func _SaveSQL_Click()
    _RunWithBusy("Save SQL", "_SaveSQL")
EndFunc

Func _IntegrityCheck_Click()
    _RunWithBusy("Integrity validation", "_IntegrityCheck")
EndFunc

Func _ShowSQLTab()
    If $g_idTabSQLItem <> 0 Then GUICtrlSetState($g_idTabSQLItem, $GUI_SHOW)
    If $g_hTabHandle <> 0 Then
        $g_bAllowProgrammaticTabChange = True
        _GUICtrlTab_SetCurSel($g_hTabHandle, $g_iTabSQL)
        $g_bAllowProgrammaticTabChange = False
    EndIf
    If $g_hLog <> 0 Then GUICtrlSetState($g_hLog, $GUI_FOCUS)
    If $g_hMain <> 0 Then DllCall("user32.dll", "bool", "RedrawWindow", "hwnd", $g_hMain, "ptr", 0, "ptr", 0, "uint", 0x0185)
EndFunc

Func _OnVerboseModeChanged()
    If IsDeclared("g_chkVerbose") Then $g_bVerbose = (GUICtrlRead($g_chkVerbose) = $GUI_CHECKED)
    IniWrite($g_sIniFile, "SQL", "Verbose", ($g_bVerbose ? "1" : "0"))
    _Log("Verbose Mode " & ($g_bVerbose ? "enabled" : "disabled"))
EndFunc

;=============================================================================
; APPLY MODULES
;=============================================================================
Func _ApplyModules()
    _RefreshModuleFlags(False)
EndFunc

;=============================================================================
; SQL GENERATION
;=============================================================================
Func _GenerateSQL()
    Local $nIntegrity = _IntegrityCheck(False, True, "generate the SQL")
    If $nIntegrity = -1 Then Return

    Local $sSQL = ""
    ; Force full clean-import mode. The Ortems SQL schema can contain many
    ; deep child tables beyond the core Toolbox list, so every import now
    ; starts by clearing all user tables dynamically before inserting data.
    Local $bClear = True
    $g_bFullCleanImport = True
    If IsDeclared("g_chkClearFirst") Then GUICtrlSetState($g_chkClearFirst, $GUI_CHECKED)
    Local $bTrans = (GUICtrlRead($g_chkTransaction) = $GUI_CHECKED)

    $sSQL &= "-- =============================================" & @CRLF
    $sSQL &= "-- ORTEMS TOOLBOX - Import SQL" & @CRLF
    $sSQL &= "-- Generated on: " & @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & @CRLF
    $sSQL &= "-- Database: " & GUICtrlRead($g_edtDatabase) & @CRLF
    $sSQL &= "-- =============================================" & @CRLF & @CRLF

    If $bTrans Then $sSQL &= "BEGIN TRANSACTION;" & @CRLF & @CRLF

    ; Clear all user tables dynamically before importing data.
    ; This is now mandatory for reliable Ortems SQL imports because customer/demo
    ; databases may contain deep child tables not represented in the workbook.
    If $bClear Then
        _Log("Global clean-import mode enabled: all non-system user tables will be cleared before import; constraints/triggers remain disabled until the import finishes.")
        $sSQL &= "-- ===== GLOBAL TABLE CLEANUP (ALL USER TABLES - FORCED) =====" & @CRLF
        $sSQL &= _GetClearSQLScript() & @CRLF
    EndIf

    ; Calendars
    $sSQL &= "-- ===== CALENDARS (B_CAL / B_PERI) =====" & @CRLF
    $sSQL &= _GenerateCalSQL() & @CRLF

    ; Machines
    $sSQL &= "-- ===== MACHINES AND WORK CENTERS =====" & @CRLF
    $sSQL &= _GenerateMachSQL() & @CRLF

    ; Operations
    $sSQL &= "-- ===== OPERATIONS (B_OPE / B_CADE) =====" & @CRLF
    $sSQL &= _GenerateOpsSQL() & @CRLF

    ; Routings
    $sSQL &= "-- ===== ROUTINGS (B_GAMM / B_PHAS) =====" & @CRLF
    $sSQL &= _GenerateRoutSQL() & @CRLF

    ; Materials
    $sSQL &= "-- ===== ITEMS (B_ART / B_VER_ART) =====" & @CRLF
    $sSQL &= _GenerateMatSQL() & @CRLF

    ; Ortems item-version-routing bridge
    $sSQL &= "-- ===== ITEM VERSION ROUTING BRIDGE (E_GAMME_NOME) =====" & @CRLF
    $sSQL &= _GenerateEGammeNomeSQL() & @CRLF

    ; BOM
    $sSQL &= "-- ===== BOM (B_NOME) =====" & @CRLF
    $sSQL &= _GenerateBOMSQL() & @CRLF

    ; Required Ortems reference/master rows used by generated Work Orders.
    ; Global cleanup clears reference tables such as B_SER, so they must be
    ; recreated before E_OF/B_OF/B_BT are inserted.
    $sSQL &= "-- ===== ORTEMS REFERENCE DATA REQUIRED BY WORK ORDERS =====" & @CRLF
    $sSQL &= _GenerateWOReferenceDataSQL() & @CRLF

    ; Targeted work-order cleanup
    ; When refreshing an existing demo database without full Clear first, old B_BT/B_OF rows
    ; can keep FK references against E_OF and block the parent update. Clean only the WOs
    ; present in the Work Orders tab, then recreate them below.
    If Not $bClear Then
        $sSQL &= "-- ===== TARGETED WORK ORDER CLEANUP =====" & @CRLF
        $sSQL &= _GenerateTargetedWOCleanupSQL() & @CRLF
    EndIf

    ; Work order header / parent rows
    ; Some Ortems SQL Server schemas validate B_BT through E_OF, not only B_OF.
    $sSQL &= "-- ===== WORK ORDER HEADERS (E_OF) =====" & @CRLF
    $sSQL &= _GenerateWOHeaderSQL() & @CRLF

    ; Ordens de Producao
    $sSQL &= "-- ===== WORK ORDERS (B_OF) =====" & @CRLF
    $sSQL &= _GenerateWOSQL() & @CRLF

    ; Ortems requires phase rows for each WO. Without B_BT rows, the application can raise PL_0170.
    $sSQL &= "-- ===== WORK ORDER PHASES (B_BT) =====" & @CRLF
    $sSQL &= _GenerateWOBTSQL() & @CRLF

    ; WO Links (se modulo ativo)
    If $g_bModWOL Then
        $sSQL &= "-- ===== LINKS WO (B_PROF) =====" & @CRLF
        $sSQL &= _GenerateWOLSQL() & @CRLF
    EndIf

    ; Re-enable triggers and FK constraints only after all workbook rows were inserted.
    ; Keeping them disabled during the import avoids intermediate FK conflicts while
    ; E_OF/B_OF/B_BT and related Ortems tables are being rebuilt.
    If $bClear Then
        $sSQL &= "-- ===== POST-IMPORT CONSTRAINT/TRIGGER RE-ENABLE =====" & @CRLF
        $sSQL &= _GetPostImportEnableSQLScript() & @CRLF
    EndIf

    If $bTrans Then
        $sSQL &= @CRLF & "COMMIT TRANSACTION;" & @CRLF
        $sSQL &= "-- If an error occurs, execute ROLLBACK TRANSACTION" & @CRLF
    EndIf

    $sSQL &= @CRLF & "-- ===== END OF SCRIPT =====" & @CRLF

    GUICtrlSetData($g_hLog, $sSQL)
    _Log("SQL generated successfully. " & StringLen($sSQL) & " characters.")

    ; Switch to the SQL tab and force the tab page to repaint.
    _ShowSQLTab()
EndFunc

;=============================================================================
; RUN SQL ON THE DATABASE
;=============================================================================
;=============================================================================
; SQL GENERATION FUNCTIONS
;=============================================================================

Func _NormalizeTimeText($vTime)
    Local $s = StringStripWS(String($vTime), 3)
    If $s = "" Then Return ""

    If StringInStr($s, ":") Then
        Local $a = StringSplit($s, ":", 1)
        If $a[0] >= 2 Then
            Local $h = Number($a[1])
            Local $m = Number($a[2])
            If $h >= 0 And $h <= 23 And $m >= 0 And $m <= 59 Then Return StringFormat("%02d:%02d", $h, $m)
        EndIf
    EndIf

    If StringRegExp($s, "^\d+[\.,]\d+$") Then
        Local $f = Number(StringReplace($s, ",", "."))
        If $f >= 0 And $f < 1 Then
            Local $nMin = Round($f * 24 * 60, 0)
            Local $h2 = Int($nMin / 60)
            Local $m2 = Mod($nMin, 60)
            If $h2 >= 0 And $h2 <= 23 Then Return StringFormat("%02d:%02d", $h2, $m2)
        EndIf
        If $f >= 0 And $f <= 23.999 Then
            Local $h3 = Int($f)
            Local $m3 = Round(($f - $h3) * 60, 0)
            If $m3 = 60 Then
                $h3 += 1
                $m3 = 0
            EndIf
            If $h3 >= 0 And $h3 <= 23 And $m3 >= 0 And $m3 <= 59 Then Return StringFormat("%02d:%02d", $h3, $m3)
        EndIf
    EndIf

    Local $sDigits = StringRegExpReplace($s, "\D", "")
    If $sDigits = "" Then Return ""

    Switch StringLen($sDigits)
        Case 1, 2
            Local $h4 = Number($sDigits)
            If $h4 >= 0 And $h4 <= 23 Then Return StringFormat("%02d:00", $h4)
        Case 3
            Local $h5 = Number(StringLeft($sDigits, 1))
            Local $m5 = Number(StringRight($sDigits, 2))
            If $h5 >= 0 And $h5 <= 23 And $m5 >= 0 And $m5 <= 59 Then Return StringFormat("%02d:%02d", $h5, $m5)
        Case Else
            $sDigits = StringRight($sDigits, 4)
            Local $h6 = Number(StringLeft($sDigits, 2))
            Local $m6 = Number(StringRight($sDigits, 2))
            If $h6 >= 0 And $h6 <= 23 And $m6 >= 0 And $m6 <= 59 Then Return StringFormat("%02d:%02d", $h6, $m6)
    EndSwitch

    Return ""
EndFunc

Func _InferBPeriTimeMode()
    Local $sMode = "text"
    If Not $g_bConnected Or $g_sConnStr = "" Then Return $sMode

    Local $oConn = ObjCreate("ADODB.Connection")
    If Not IsObj($oConn) Then Return $sMode
    $g_sLastComError = ""
    $oConn.Open($g_sConnStr)
    If $g_sLastComError <> "" Or $oConn.State <> 1 Then Return $sMode

    Local $sQuery = "SELECT TOP 1 DEB_PERIO FROM B_PERI WHERE DEB_PERIO IS NOT NULL"
    _LogVerboseSQL("Infer B_PERI time mode", $sQuery)
    $g_sLastComError = ""
    Local $oRS = $oConn.Execute($sQuery)
    If $g_sLastComError = "" And IsObj($oRS) And Not $oRS.EOF Then
        Local $sSample = StringStripWS(String($oRS.Fields(0).Value), 3)
        If StringInStr($sSample, ":") Then
            $sMode = "text"
        ElseIf StringRegExp($sSample, "^\d+$") Then
            Local $n = Number($sSample)
            If $n >= 0 And $n <= 23 Then
                $sMode = "hour"
            Else
                $sMode = "hhmm"
            EndIf
        EndIf
        $oRS.Close()
    EndIf
    $oConn.Close()
    _Log("B_PERI time mode detected: " & $sMode)
    Return $sMode
EndFunc

Func _TimeToSQLLiteral($vTime, $sMode)
    Local $sNorm = _NormalizeTimeText($vTime)
    If $sNorm = "" Then Return "''"

    Switch StringLower($sMode)
        Case "hour"
            Return String(Number(StringLeft($sNorm, 2)))
        Case "hhmm"
            Return String(Number(StringLeft($sNorm, 2) & StringRight($sNorm, 2)))
        Case Else
            Return "'" & $sNorm & "'"
    EndSwitch
EndFunc

Func _SQLQ($sVal)
    Return "'" & StringReplace(String($sVal), "'", "''") & "'"
EndFunc

Func _SQLIdent($sName)
    Return "[" & StringReplace(String($sName), "]", "]]") & "]"
EndFunc

Func _SQLNum($sVal, $sDefault = "0")
    Local $s = StringStripWS(String($sVal), 3)
    If $s = "" Then Return $sDefault
    $s = StringReplace($s, ",", ".")
    If Not StringRegExp($s, "^-?\d+(\.\d+)?$") Then Return $sDefault
    Return $s
EndFunc

Func _SQLDateTime103($sVal, $sDefault = "01/01/1995 00:00")
    Local $s = StringStripWS(String($sVal), 3)
    If $s = "" Then $s = $sDefault
    Return "CONVERT(datetime," & _SQLQ($s) & ",103)"
EndFunc

Func _DetectOptionalColumn($sTable, $sRole, $sCandidates)
    ; For optional fields, do not guess when we are offline. Guessing would generate
    ; columns that may not exist in the target Ortems schema.
    If Not $g_bConnected Or $g_sConnStr = "" Then Return ""
    Return _DetectColumn($sTable, $sRole, $sCandidates & ";__NONE__")
EndFunc

Func _DetectFKParentColumnByReferenced($sParentTable, $sReferencedTable, $sReferencedColumn, $sRole)
    If Not $g_bConnected Or $g_sConnStr = "" Then Return ""

    Local $oConn = ObjCreate("ADODB.Connection")
    If Not IsObj($oConn) Then Return ""

    $g_sLastComError = ""
    $oConn.Open($g_sConnStr)
    If $g_sLastComError <> "" Or $oConn.State <> 1 Then Return ""

    Local $sParentSafe = StringReplace($sParentTable, "'", "''")
    Local $sReferencedSafe = StringReplace($sReferencedTable, "'", "''")
    Local $sRefColSafe = StringReplace($sReferencedColumn, "'", "''")

    Local $sQuery = "SELECT TOP 1 pc.name " & _
        "FROM sys.foreign_key_columns fkc " & _
        "INNER JOIN sys.columns pc ON pc.object_id=fkc.parent_object_id AND pc.column_id=fkc.parent_column_id " & _
        "INNER JOIN sys.columns rc ON rc.object_id=fkc.referenced_object_id AND rc.column_id=fkc.referenced_column_id " & _
        "INNER JOIN sys.foreign_keys fk ON fk.object_id=fkc.constraint_object_id " & _
        "WHERE OBJECT_NAME(fkc.parent_object_id)='" & $sParentSafe & "' " & _
        "AND OBJECT_NAME(fkc.referenced_object_id)='" & $sReferencedSafe & "' " & _
        "AND UPPER(rc.name)=UPPER('" & $sRefColSafe & "') " & _
        "ORDER BY CASE WHEN fk.name='FK2_B_PART' THEN 0 ELSE 1 END, pc.column_id"
    _LogVerboseSQL("Detect FK parent column " & $sParentTable & " -> " & $sReferencedTable & "." & $sReferencedColumn, $sQuery)
    $g_sLastComError = ""
    Local $oRS = $oConn.Execute($sQuery)

    Local $sFound = ""
    If $g_sLastComError = "" And IsObj($oRS) Then
        If Not $oRS.EOF Then $sFound = $oRS.Fields(0).Value
        $oRS.Close()
    EndIf

    $oConn.Close()

    If $sFound <> "" Then
        _Log($sParentTable & " schema: FK to " & $sReferencedTable & " maps referenced column '" & $sReferencedColumn & "' to parent column '" & $sFound & "' for " & $sRole)
    Else
        _Log($sParentTable & " schema: no FK parent column detected for " & $sRole & " through " & $sReferencedTable & "." & $sReferencedColumn)
    EndIf

    Return $sFound
EndFunc

Func _ColumnIfPresent(ByRef $sColList, ByRef $sValList, $sCol, $sVal)
    If $sCol = "" Then Return
    $sColList &= ", " & $sCol
    $sValList &= ", " & $sVal
EndFunc

Func _BaseTableExists($sTable)
    If Not $g_bConnected Or $g_sConnStr = "" Then Return False

    Local $oConn = ObjCreate("ADODB.Connection")
    If Not IsObj($oConn) Then Return False

    $g_sLastComError = ""
    $oConn.Open($g_sConnStr)
    If $g_sLastComError <> "" Or $oConn.State <> 1 Then Return False

    Local $sSafeTable = StringReplace($sTable, "'", "''")
    Local $sQuery = "SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME='" & $sSafeTable & "' AND TABLE_TYPE='BASE TABLE'"
    _LogVerboseSQL("Check base table exists: " & $sTable, $sQuery)
    Local $oRS = $oConn.Execute($sQuery)
    Local $bExists = False
    If $g_sLastComError = "" And IsObj($oRS) Then
        $bExists = (Number($oRS.Fields(0).Value) > 0)
        $oRS.Close()
    EndIf

    $oConn.Close()
    Return $bExists
EndFunc


Func _GetTableColumnsForLog($sTable)
    ; Returns the actual columns for a table in a compact string used by verbose logs.
    ; This is useful when an Ortems schema uses version-specific column names.
    If Not $g_bConnected Or $g_sConnStr = "" Then Return ""

    Local $oConn = ObjCreate("ADODB.Connection")
    If Not IsObj($oConn) Then Return ""

    $g_sLastComError = ""
    $oConn.Open($g_sConnStr)
    If $g_sLastComError <> "" Or $oConn.State <> 1 Then Return ""

    Local $sSafeTable = StringReplace($sTable, "'", "''")
    Local $sQuery = "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS " & _
        "WHERE TABLE_NAME='" & $sSafeTable & "' ORDER BY ORDINAL_POSITION"
    _LogVerboseSQL("Get table columns for " & $sTable, $sQuery)
    $g_sLastComError = ""
    Local $oRS = $oConn.Execute($sQuery)

    Local $sCols = ""
    If $g_sLastComError = "" And IsObj($oRS) Then
        While Not $oRS.EOF
            If $sCols <> "" Then $sCols &= ", "
            $sCols &= $oRS.Fields(0).Value
            $oRS.MoveNext()
        WEnd
        $oRS.Close()
    EndIf

    $oConn.Close()
    Return $sCols
EndFunc

Func _DetectRequiredColumn($sTable, $sRole, $sCandidates)
    ; Same detection logic as optional columns, but with a clearer warning when the
    ; field is required to build a safe INSERT/WHERE condition.
    Local $sFound = _DetectOptionalColumn($sTable, $sRole, $sCandidates)
    If $sFound = "" Then
        Local $sActualCols = _GetTableColumnsForLog($sTable)
        _Log("WARNING: Required column for " & $sTable & "." & $sRole & " was not detected. Candidates tried: " & $sCandidates)
        If $sActualCols <> "" Then _Log($sTable & " actual columns: " & $sActualCols)
    EndIf
    Return $sFound
EndFunc

Func _AppendColumnValue(ByRef $sColList, ByRef $sValList, $sCol, $sVal)
    ; Uses SQL Server quoted identifiers so schema-specific column names remain safe.
    If $sCol = "" Then Return
    $sColList &= ", " & _SQLIdent($sCol)
    $sValList &= ", " & $sVal
EndFunc

Func _AppendWhereEquals(ByRef $sWhere, $sCol, $sVal)
    If $sCol = "" Then Return
    If $sWhere <> "" Then $sWhere &= " AND "
    $sWhere &= _SQLIdent($sCol) & "=" & $sVal
EndFunc

Func _AppendUpdateSet(ByRef $sSetList, $sCol, $sVal)
    If $sCol = "" Then Return
    If $sSetList <> "" Then $sSetList &= ", "
    $sSetList &= _SQLIdent($sCol) & "=" & $sVal
EndFunc

Func _InsertListHasColumn($sColList, $sCol)
    If $sCol = "" Then Return True
    If StringInStr($sColList, _SQLIdent($sCol)) Then Return True
    If StringInStr(", " & $sColList & ",", ", " & $sCol & ",") Then Return True
    Return False
EndFunc

Func _AppendColumnValueIfMissing(ByRef $sColList, ByRef $sValList, $sCol, $sVal)
    If $sCol = "" Then Return
    If _InsertListHasColumn($sColList, $sCol) Then Return
    _AppendColumnValue($sColList, $sValList, $sCol, $sVal)
EndFunc

Func _GetRequiredNoDefaultColumns($sTable)
    ; Returns required columns that SQL Server will not auto-populate.
    ; Format: column|datatype;column|datatype
    If Not $g_bConnected Or $g_sConnStr = "" Then Return ""

    Local $oConn = ObjCreate("ADODB.Connection")
    If Not IsObj($oConn) Then Return ""

    $g_sLastComError = ""
    $oConn.Open($g_sConnStr)
    If $g_sLastComError <> "" Or $oConn.State <> 1 Then Return ""

    Local $sSafeTable = StringReplace($sTable, "'", "''")
    Local $sQuery = "SELECT COLUMN_NAME, DATA_TYPE " & _
        "FROM INFORMATION_SCHEMA.COLUMNS " & _
        "WHERE TABLE_NAME='" & $sSafeTable & "' " & _
        "AND IS_NULLABLE='NO' " & _
        "AND COLUMN_DEFAULT IS NULL " & _
        "AND COLUMNPROPERTY(OBJECT_ID(TABLE_SCHEMA + '.' + TABLE_NAME), COLUMN_NAME, 'IsIdentity') = 0 " & _
        "ORDER BY ORDINAL_POSITION"
    _LogVerboseSQL("Get required no-default columns for " & $sTable, $sQuery)
    $g_sLastComError = ""
    Local $oRS = $oConn.Execute($sQuery)

    Local $sOut = ""
    If $g_sLastComError = "" And IsObj($oRS) Then
        While Not $oRS.EOF
            If $sOut <> "" Then $sOut &= ";"
            $sOut &= $oRS.Fields("COLUMN_NAME").Value & "|" & $oRS.Fields("DATA_TYPE").Value
            $oRS.MoveNext()
        WEnd
        $oRS.Close()
    EndIf

    $oConn.Close()
    Return $sOut
EndFunc

Func _LookupWorkCenterByOperation($sOperation)
    Local $sOpNorm = StringLower(StringStripWS($sOperation, 3))
    If $sOpNorm = "" Then Return ""

    Local $nRows = _GUICtrlListView_GetItemCount($g_hLV_Ops)
    For $i = 0 To $nRows - 1
        Local $sRowOp = StringLower(StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Ops, $i, 0), 3))
        If $sRowOp = $sOpNorm Then Return StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Ops, $i, 2), 3)
    Next
    Return ""
EndFunc

Func _LookupOperationRefQty($sOperation)
    Local $sOpNorm = StringLower(StringStripWS($sOperation, 3))
    If $sOpNorm = "" Then Return "1"

    Local $nRows = _GUICtrlListView_GetItemCount($g_hLV_Ops)
    For $i = 0 To $nRows - 1
        Local $sRowOp = StringLower(StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Ops, $i, 0), 3))
        If $sRowOp = $sOpNorm Then Return _SQLNum(_GUICtrlListView_GetItemText($g_hLV_Ops, $i, 4), "1")
    Next
    Return "1"
EndFunc

Func _LookupOperationRefDuration($sOperation)
    Local $sOpNorm = StringLower(StringStripWS($sOperation, 3))
    If $sOpNorm = "" Then Return "0"

    Local $nRows = _GUICtrlListView_GetItemCount($g_hLV_Ops)
    For $i = 0 To $nRows - 1
        Local $sRowOp = StringLower(StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Ops, $i, 0), 3))
        If $sRowOp = $sOpNorm Then Return _SQLNum(_GUICtrlListView_GetItemText($g_hLV_Ops, $i, 5), "0")
    Next
    Return "0"
EndFunc

Func _LookupPhaseName($sRouting, $sPhase)
    Local $sRoutingNorm = StringLower(StringStripWS($sRouting, 3))
    Local $nPhase = Number(_SQLNum($sPhase, "0"))
    If $sRoutingNorm = "" Then Return ""

    Local $nRows = _GUICtrlListView_GetItemCount($g_hLV_Rout)
    For $i = 0 To $nRows - 1
        Local $sRowRouting = StringLower(StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Rout, $i, 0), 3))
        Local $nRowPhase = Number(_SQLNum(_GUICtrlListView_GetItemText($g_hLV_Rout, $i, 2), "0"))
        If $sRowRouting = $sRoutingNorm And $nRowPhase = $nPhase Then Return StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Rout, $i, 4), 3)
    Next
    Return ""
EndFunc

Func _BTRequiredDefaultValue($sCol, $sDataType, $sWO, $sRouting, $sPhase, $sOpe, $sMach, $sWC, $sItem, $sVersion, $nQty, $sStartDate, $sEndDate, $sPhaseName)
    Local $sU = StringUpper(StringStripWS($sCol, 3))
    Local $sType = StringLower(StringStripWS($sDataType, 3))
    Local $nOpRefQty = _LookupOperationRefQty($sOpe)
    Local $nOpDuration = _LookupOperationRefDuration($sOpe)
    Local $sDescr = $sPhaseName
    If $sDescr = "" Then $sDescr = $sOpe

    Switch $sU
        Case "E_NOF", "NOF", "BT_NOF"
            Return _SQLQ($sWO)
        Case "NOM_ENC", "E_NOM_ENC", "B_B_NOM_ENC"
            Return _SQLQ($ORTEMS_DEFAULT_NOM_ENC)
        Case "NOMG", "BT_NOMG", "B_NOMG", "B_P_NOMG", "NOM_GAMME", "GAMME", "ROUTING"
            Return _SQLQ($sRouting)
        Case "BT_NOPHASE", "NOPHASE"
            Return _SQLNum($sPhase, "0")
        Case "BT_FRACTION", "FRACTION"
            Return "0"
        Case "BT_OPE", "OPE", "CODEOPE"
            Return _SQLQ($sOpe)
        Case "MACHINE", "B_M_MACHINE", "BT_MACHINE"
            Return _SQLQ($sMach)
        Case "ILOT", "B_I_ILOT", "B_M_ILOT"
            Return _SQLQ($sWC)
        Case "CODEARTIC", "B_CODEARTIC", "BT_CODEARTIC"
            Return _SQLQ($sItem)
        Case "VER_ART", "B_VER_ART", "BT_VER_ART"
            Return _SQLQ($sVersion)
        Case "SERIE", "BT_SERIE"
            Return "'0'"
        Case "BT_QTE", "QTE", "QTEBT", "QTEOF", "QTE_ORDER", "BT_QTEREST", "BT_QTE_FCT"
            Return $nQty
        Case "BT_QTEREAL", "BT_TEMPSPASS", "BT_DURPREP", "BT_DURCHANG", "BT_DURREAL_CHARGE", "BT_CST_CHARGE", "BT_STATION_COUNT", "BT_STATION_FROM", "BT_STATION_TO", "BT_TAKT_TIME"
            Return "0"
        Case "BT_DURREAL", "BT_TEMPSREST", "BT_TEMPSCYCL", "BT_CODEBASET_CHARGE"
            Return $nOpDuration
        Case "BT_CODEBASET"
            Return $nOpRefQty
        Case "BT_COEF_T", "BT_COEF_QTE", "BT_COEF_FRAC", "BT_INTENSITE", "BT_INTENSITE_QTE", "BT_EFFECTSTAN", "BT_EFFECTMIN", "BT_EFFECTMAX"
            Return "1"
        Case "BT_CODEGEST", "CODEGEST"
            Return "'F'"
        Case "BT_ETATOP", "ETATBT", "ETAT", "STATUS_POTENTIEL", "LBE_STATUT", "BT_STATUT_MATIERE", "BT_PUN_STATUT"
            Return "'S'"
        Case "BT_LIBPHASE", "BT_OPEDESCR"
            Return _SQLQ($sDescr)
        Case "BT_DATEFIN", "BT_FINPLUSTARD", "BT_DATE_FIN_OBJ", "BT_DATE_FIN_PLA_REF", "BT_PUN_DATE_FIN"
            Return _SQLDateTime103($sEndDate, "01/01/2050 23:59")
        Case "BT_DATE_CAL", "BT_DATE_CAL_SF", "BT_DATE_CAL_MP", "BT_DATEG", "BT_DATE_DEB_PROD", "BT_DATE_DEB_OBJ", "BT_DATE_DEB_PLUSTOT", "BT_DATE_DEB_PLA_REF", "BT_DATE_LIS", "BT_DATE_DATE_CAL", "DATE_DEB_ENCOURS", "BT_DATE_CAL_SAUVE", "DATE_PLA", "BT_PUN_DATE_DEB"
            Return _SQLDateTime103($sStartDate, "01/01/1995 00:00")
        Case "BT_AFFECTFORCE", "BT_AFFECT", "BT_INTERUPT", "BT_MONTEE_CADENCE", "PHASE_LOGIQUE", "BT_PROP_DATE_CAL", "CHGT_PARAM", "RANG_GPE_RESS", "BT_PRIORITE", "BT_MEM", "BT_DISPO", "CALAGETOT", "POTENTIEL", "FORCER_ACTIF_POTENTIEL", "LBE_PARAM_LOCAUX", "LBE_MARGE_AVANCE", "LBE_CTRL_AVANCE", "LBE_MARGE_RETARD", "LBE_CTRL_RETARD", "LBE_FIGE", "LBE_SPECIF", "BT_PREMSUIVI", "BT_NUMORDRE", "BT_CSMANAGEMENT", "BT_CSUNITMODE", "BT_CSMINIMAL", "BT_CSINITIAL", "BT_CSFREEZE", "BT_DELAI", "BT_LIEN_SUPPR", "BT_PUN", "BT_PUN_DURCHANG", "DURLCREST", "DURLCRESTI"
            Return "0"
        Case "PHASE_PROD"
            Return "1"
        Case "BT_MODE_CHARGE"
            Return "'N'"
    EndSwitch

    If StringInStr($sType, "date") Or StringInStr($sType, "time") Then Return _SQLDateTime103($sStartDate, "01/01/1995 00:00")
    If StringInStr($sType, "char") Or StringInStr($sType, "text") Or StringInStr($sType, "uniqueidentifier") Then Return "''"
    If StringInStr($sType, "bit") Then Return "0"

    Return "0"
EndFunc

Func _AppendBTTRequiredDefaults(ByRef $sColList, ByRef $sValList, $sWO, $sRouting, $sPhase, $sOpe, $sMach, $sWC, $sItem, $sVersion, $nQty, $sStartDate, $sEndDate, $sPhaseName)
    Static $sRequiredCache = "__UNSET__"
    Static $sLoggedCols = ";"
    If $sRequiredCache = "__UNSET__" Or ($sRequiredCache = "" And $g_bConnected) Then $sRequiredCache = _GetRequiredNoDefaultColumns("B_BT")

    Local $sRequired = $sRequiredCache
    If $sRequired = "" Then Return

    Local $aReq = StringSplit($sRequired, ";", 1)
    For $i = 1 To $aReq[0]
        Local $aParts = StringSplit($aReq[$i], "|", 1)
        If $aParts[0] < 1 Then ContinueLoop
        Local $sCol = $aParts[1]
        Local $sType = ""
        If $aParts[0] >= 2 Then $sType = $aParts[2]
        If _InsertListHasColumn($sColList, $sCol) Then ContinueLoop

        Local $sVal = _BTRequiredDefaultValue($sCol, $sType, $sWO, $sRouting, $sPhase, $sOpe, $sMach, $sWC, $sItem, $sVersion, $nQty, $sStartDate, $sEndDate, $sPhaseName)
        _AppendColumnValueIfMissing($sColList, $sValList, $sCol, $sVal)
        If Not StringInStr($sLoggedCols, ";" & $sCol & ";") Then
            _Log("B_BT schema: required column '" & $sCol & "' auto-filled with a safe default.")
            $sLoggedCols &= $sCol & ";"
        EndIf
    Next
EndFunc


Func _LookupFirstOperationByRouting($sRouting)
    Local $sRoutingNorm = StringLower(StringStripWS($sRouting, 3))
    If $sRoutingNorm = "" Then Return ""

    Local $sBestOp = ""
    Local $nBestPhase = 999999999
    Local $nRows = _GUICtrlListView_GetItemCount($g_hLV_Rout)
    For $i = 0 To $nRows - 1
        Local $sRowRouting = StringLower(StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Rout, $i, 0), 3))
        If $sRowRouting <> $sRoutingNorm Then ContinueLoop
        Local $nPhase = Number(_SQLNum(_GUICtrlListView_GetItemText($g_hLV_Rout, $i, 2), "0"))
        If $sBestOp = "" Or $nPhase < $nBestPhase Then
            $nBestPhase = $nPhase
            $sBestOp = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Rout, $i, 3), 3)
        EndIf
    Next
    Return $sBestOp
EndFunc

Func _WOHeaderRequiredDefaultValue($sCol, $sDataType, $sWO, $sItem, $sRouting, $sVersion, $nQty, $sStartDate, $sEndDate, $sFirstPhase, $sFirstOp, $sMach, $sWC)
    Local $sU = StringUpper(StringStripWS($sCol, 3))
    Local $sType = StringLower(StringStripWS($sDataType, 3))

    Switch $sU
        Case "NOF", "E_NOF", "BT_NOF", "NUMOF", "NUM_OF", "NO_OF", "N_OF", "OF", "ID_OF"
            Return _SQLQ($sWO)
        Case "CODEARTIC", "B_CODEARTIC", "E_CODEARTIC", "ARTICLE", "ITEM", "B_V_CODEARTIC"
            Return _SQLQ($sItem)
        Case "NOM_ENC", "E_NOM_ENC", "PLAN_CAP"
            Return _SQLQ($ORTEMS_DEFAULT_NOM_ENC)
        Case "NOMG", "E_NOMG", "NOM_GAMME", "GAMME", "ROUTING"
            Return _SQLQ($sRouting)
        Case "VER_ART", "E_VER_ART", "VERSION", "B_VER_ART"
            Return _SQLQ($sVersion)
        Case "SERIE", "E_SERIE"
            Return "'0'"
        Case "QTE", "QTEOF", "QTE_ORDER", "QTEOFORDER", "QTECMD", "QTE_LANCEE", "QTE_LANC"
            Return $nQty
        Case "DPLUSTOT", "DATEDEB", "DATDEB", "DATE_DEBUT", "DEBUT", "DATE_PLA", "DATE_LANCEMENT", "DATE_DEB", "VER_EFFET_DEBUT"
            Return _SQLDateTime103($sStartDate, "01/01/1995 00:00")
        Case "FPLUSTARD", "DATEFIN", "DATFIN", "DATE_FIN", "FIN", "DATE_BESOIN", "DATE_ECH", "VER_EFFET_FIN"
            Return _SQLDateTime103($sEndDate, "01/01/2050 23:59")
        Case "CODEGEST", "E_CODEGEST"
            Return "'F'"
        Case "ETATOF", "ETAT", "STATUS", "STATUT", "OF_STATUS"
            Return "'S'"
        Case "TYPEOF", "TYPE_OF", "OF_TYPE", "TYPE"
            Return "'F'"
        Case "NOPHASE", "BT_NOPHASE", "PHASE", "NO_PHASE"
            Return _SQLNum($sFirstPhase, "0")
        Case "OPE", "CODEOPE", "BT_OPE"
            Return _SQLQ($sFirstOp)
        Case "MACHINE", "BT_MACHINE"
            Return _SQLQ($sMach)
        Case "ILOT", "B_I_ILOT", "B_M_ILOT"
            Return _SQLQ($sWC)
        Case "MODE_UTIL", "MODE_GESTION", "MODUTIL"
            Return "'C'"
        Case "PRIORITE", "PRIORITY", "RANG", "NUMORDRE"
            Return "0"
    EndSwitch

    If StringInStr($sType, "date") Or StringInStr($sType, "time") Then Return _SQLDateTime103($sStartDate, "01/01/1995 00:00")
    If StringInStr($sType, "char") Or StringInStr($sType, "text") Or StringInStr($sType, "uniqueidentifier") Then Return "''"
    If StringInStr($sType, "bit") Then Return "0"

    Return "0"
EndFunc

Func _AppendWOHeaderRequiredDefaults(ByRef $sColList, ByRef $sValList, $sWO, $sItem, $sRouting, $sVersion, $nQty, $sStartDate, $sEndDate, $sFirstPhase, $sFirstOp, $sMach, $sWC)
    Static $sRequiredCache = "__UNSET__"
    Static $sLoggedCols = ";"
    If $sRequiredCache = "__UNSET__" Or ($sRequiredCache = "" And $g_bConnected) Then $sRequiredCache = _GetRequiredNoDefaultColumns("E_OF")

    Local $sRequired = $sRequiredCache
    If $sRequired = "" Then Return

    Local $aReq = StringSplit($sRequired, ";", 1)
    For $i = 1 To $aReq[0]
        Local $aParts = StringSplit($aReq[$i], "|", 1)
        If $aParts[0] < 1 Then ContinueLoop
        Local $sCol = $aParts[1]
        Local $sType = ""
        If $aParts[0] >= 2 Then $sType = $aParts[2]
        If _InsertListHasColumn($sColList, $sCol) Then ContinueLoop

        Local $sVal = _WOHeaderRequiredDefaultValue($sCol, $sType, $sWO, $sItem, $sRouting, $sVersion, $nQty, $sStartDate, $sEndDate, $sFirstPhase, $sFirstOp, $sMach, $sWC)
        _AppendColumnValueIfMissing($sColList, $sValList, $sCol, $sVal)
        If Not StringInStr($sLoggedCols, ";" & $sCol & ";") Then
            _Log("E_OF schema: required column '" & $sCol & "' auto-filled with a safe default.")
            $sLoggedCols &= $sCol & ";"
        EndIf
    Next
EndFunc

Func _LookupItemType($sItemID)
    Local $sNeedle = StringLower(StringStripWS($sItemID, 3))
    If $sNeedle = "" Then Return ""
    Local $nRows = _GUICtrlListView_GetItemCount($g_hLV_Mat)
    For $i = 0 To $nRows - 1
        Local $sRowItem = StringLower(StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Mat, $i, 0), 3))
        If $sRowItem = $sNeedle Then Return StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Mat, $i, 2), 3)
    Next
    Return ""
EndFunc

Func _ItemIsRawMaterial($sItemID)
    Return _IsRawMaterialType(_LookupItemType($sItemID))
EndFunc

Func _LookupUniqueMachineByOperation($sOperation)
    Local $sOpNorm = StringLower(StringStripWS($sOperation, 3))
    If $sOpNorm = "" Then Return ""

    Local $sFound = ""
    Local $sSeen = ";"
    Local $nUnique = 0
    Local $nRows = _GUICtrlListView_GetItemCount($g_hLV_Ops)

    For $i = 0 To $nRows - 1
        Local $sRowOp = StringLower(StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Ops, $i, 0), 3))
        If $sRowOp <> $sOpNorm Then ContinueLoop
        Local $sMachine = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Ops, $i, 3), 3)
        If $sMachine = "" Or _IsStandByMachine($sMachine) Then ContinueLoop
        If StringInStr($sSeen, ";" & $sMachine & ";") Then ContinueLoop
        $sSeen &= $sMachine & ";"
        $nUnique += 1
        $sFound = $sMachine
        If $nUnique > 1 Then Return ""
    Next
    Return $sFound
EndFunc

Func _GenerateCalSQL()
    Local $s = ""
    Local $sTimeMode = _InferBPeriTimeMode()
    Local $nRows = _GUICtrlListView_GetItemCount($g_hLV_Cal)
    For $i = 0 To $nRows - 1
        Local $sID    = _GUICtrlListView_GetItemText($g_hLV_Cal, $i, 0)
        Local $sNome  = _GUICtrlListView_GetItemText($g_hLV_Cal, $i, 1)
        Local $sDiaI  = _GUICtrlListView_GetItemText($g_hLV_Cal, $i, 2)
        Local $sHoraI = _GUICtrlListView_GetItemText($g_hLV_Cal, $i, 3)
        Local $sDiaF  = _GUICtrlListView_GetItemText($g_hLV_Cal, $i, 4)
        Local $sHoraF = _GUICtrlListView_GetItemText($g_hLV_Cal, $i, 5)

        Local $sHoraISQL = _TimeToSQLLiteral($sHoraI, $sTimeMode)
        Local $sHoraFSQL = _TimeToSQLLiteral($sHoraF, $sTimeMode)

        If $sHoraISQL = "''" Or $sHoraFSQL = "''" Then
            _Log("WARNING: Calendar row " & ($i + 1) & " has an invalid time format. Start='" & $sHoraI & "' End='" & $sHoraF & "'")
        EndIf

        $s &= "IF NOT EXISTS (SELECT 1 FROM B_CAL WHERE NOCALHEBD='" & $sID & "')" & @CRLF
        $s &= "    INSERT INTO B_CAL (NOCALHEBD, NOMCAL) VALUES ('" & $sID & "', '" & $sNome & "');" & @CRLF

        $s &= "IF NOT EXISTS (SELECT 1 FROM B_PERI WHERE NOCALHEBD='" & $sID & "' AND NOJOUR_DEB=" & $sDiaI & " AND DEB_PERIO=" & $sHoraISQL & ")" & @CRLF
        $s &= "    INSERT INTO B_PERI (NOCALHEBD, NOJOUR_DEB, DEB_PERIO, NOJOUR_FIN, FIN_PERIO) " & @CRLF
        $s &= "    VALUES ('" & $sID & "', " & $sDiaI & ", " & $sHoraISQL & ", " & $sDiaF & ", " & $sHoraFSQL & ");" & @CRLF
    Next
    Return $s
EndFunc

Func _GenerateMachSQL()
    Local $s = ""
    Local $nRows = _GUICtrlListView_GetItemCount($g_hLV_Mach)
    For $i = 0 To $nRows - 1
        Local $sSiteID = _GUICtrlListView_GetItemText($g_hLV_Mach, $i, 0)
        Local $sSiteNm = _GUICtrlListView_GetItemText($g_hLV_Mach, $i, 1)
        Local $sCTID   = _GUICtrlListView_GetItemText($g_hLV_Mach, $i, 2)
        Local $sCTNm   = _GUICtrlListView_GetItemText($g_hLV_Mach, $i, 3)
        Local $sCtTp   = _GUICtrlListView_GetItemText($g_hLV_Mach, $i, 4)
        Local $sSecID  = _GUICtrlListView_GetItemText($g_hLV_Mach, $i, 5)
        Local $sSecNm  = _GUICtrlListView_GetItemText($g_hLV_Mach, $i, 6)
        Local $sMaqID  = _GUICtrlListView_GetItemText($g_hLV_Mach, $i, 7)
        Local $sMaqNm  = _GUICtrlListView_GetItemText($g_hLV_Mach, $i, 8)
        Local $sMaqTp  = _GUICtrlListView_GetItemText($g_hLV_Mach, $i, 9)
        Local $sCalID  = _GUICtrlListView_GetItemText($g_hLV_Mach, $i, 10)
        Local $sCalCap = _GUICtrlListView_GetItemText($g_hLV_Mach, $i, 11)

        $s &= "IF NOT EXISTS (SELECT 1 FROM B_ZONE WHERE NOZONE='" & $sSiteID & "')" & @CRLF
        $s &= "    INSERT INTO B_ZONE (NOZONE, LIBZONE) VALUES ('" & $sSiteID & "', '" & $sSiteNm & "');" & @CRLF
        $s &= "IF NOT EXISTS (SELECT 1 FROM B_ILOT WHERE ILOT='" & $sCTID & "')" & @CRLF
        $s &= "    INSERT INTO B_ILOT (ILOT, LIBILOT, TYPEILOT) VALUES ('" & $sCTID & "', '" & $sCTNm & "', " & $sCtTp & ");" & @CRLF
        $s &= "IF NOT EXISTS (SELECT 1 FROM B_SECT WHERE CODESECTI='" & $sSecID & "')" & @CRLF
        $s &= "    INSERT INTO B_SECT (CODESECTI, DESIGSECT) VALUES ('" & $sSecID & "', '" & $sSecNm & "');" & @CRLF
        $s &= "IF NOT EXISTS (SELECT 1 FROM B_MACH WHERE MACHINE='" & $sMaqID & "')" & @CRLF
        $s &= "    INSERT INTO B_MACH (MACHINE, LIBMACH, ILOT, CODESECTI, NOZONE, MACH_MODEMACH, NOCALHEBD) " & @CRLF
        $s &= "    VALUES ('" & $sMaqID & "', '" & $sMaqNm & "', '" & $sCTID & "', '" & $sSecID & "', '" & $sSiteID & "', '" & $sMaqTp & "', '" & $sCalID & "');" & @CRLF
    Next
    Return $s
EndFunc

Func _GenerateOpsSQL()
    Local $s = ""
    Local $nRows = _GUICtrlListView_GetItemCount($g_hLV_Ops)
    For $i = 0 To $nRows - 1
        Local $sID    = _GUICtrlListView_GetItemText($g_hLV_Ops, $i, 0)
        Local $sNome  = _GUICtrlListView_GetItemText($g_hLV_Ops, $i, 1)
        Local $sCT    = _GUICtrlListView_GetItemText($g_hLV_Ops, $i, 2)
        Local $sMaq   = _GUICtrlListView_GetItemText($g_hLV_Ops, $i, 3)
        Local $nQRef  = _GUICtrlListView_GetItemText($g_hLV_Ops, $i, 4)
        Local $nDRef  = _GUICtrlListView_GetItemText($g_hLV_Ops, $i, 5)
        Local $sUni   = _GUICtrlListView_GetItemText($g_hLV_Ops, $i, 6)
        Local $nPrep  = _GUICtrlListView_GetItemText($g_hLV_Ops, $i, 7)
        Local $nTOut  = _GUICtrlListView_GetItemText($g_hLV_Ops, $i, 8)
        Local $bIntr  = _GUICtrlListView_GetItemText($g_hLV_Ops, $i, 9)

        $s &= "IF NOT EXISTS (SELECT 1 FROM B_OPE WHERE OPE='" & $sID & "')" & @CRLF
        $s &= "    INSERT INTO B_OPE (OPE, LIBOP, ILOT, CODEBASET, UNITE, DURPREP, THM, INTERUPT)" & @CRLF
        $s &= "    VALUES ('" & $sID & "', '" & $sNome & "', '" & $sCT & "', " & $nQRef & ", '" & $sUni & "', " & $nPrep & ", " & $nTOut & ", " & $bIntr & ");" & @CRLF
        $s &= "IF NOT EXISTS (SELECT 1 FROM B_CADE WHERE OPE='" & $sID & "' AND MACHINE='" & $sMaq & "')" & @CRLF
        $s &= "    INSERT INTO B_CADE (OPE, MACHINE, ILOT, CADE_DURREAL, CADE_CODEBASET, CADE_UNITE)" & @CRLF
        $s &= "    VALUES ('" & $sID & "', '" & $sMaq & "', '" & $sCT & "', " & $nDRef & ", " & $nQRef & ", '" & $sUni & "');" & @CRLF
    Next
    Return $s
EndFunc

Func _GenerateRoutSQL()
    Local $s = ""
    Local $nRows = _GUICtrlListView_GetItemCount($g_hLV_Rout)
    For $i = 0 To $nRows - 1
        Local $sID   = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Rout, $i, 0), 3)
        Local $sNome = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Rout, $i, 1), 3)
        Local $nFase = _SQLNum(_GUICtrlListView_GetItemText($g_hLV_Rout, $i, 2), "0")
        Local $sOp   = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Rout, $i, 3), 3)
        Local $sNomF = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Rout, $i, 4), 3)

        ; Upsert routing headers and phases. Existing demo databases may already
        ; contain a routing with different phase details, so the Toolbox must
        ; refresh the phase rows before B_VER_ART points an item version to them.
        $s &= "IF EXISTS (SELECT 1 FROM B_GAMM WHERE NOMG=" & _SQLQ($sID) & ")" & @CRLF
        $s &= "    UPDATE B_GAMM SET LIBGAM=" & _SQLQ($sNome) & " WHERE NOMG=" & _SQLQ($sID) & @CRLF
        $s &= "ELSE" & @CRLF
        $s &= "    INSERT INTO B_GAMM (NOMG, LIBGAM) VALUES (" & _SQLQ($sID) & ", " & _SQLQ($sNome) & ");" & @CRLF

        $s &= "IF EXISTS (SELECT 1 FROM B_PHAS WHERE NOMG=" & _SQLQ($sID) & " AND NOPHASE=" & $nFase & ")" & @CRLF
        $s &= "    UPDATE B_PHAS SET OPE=" & _SQLQ($sOp) & ", LIBPHASE=" & _SQLQ($sNomF) & " WHERE NOMG=" & _SQLQ($sID) & " AND NOPHASE=" & $nFase & @CRLF
        $s &= "ELSE" & @CRLF
        $s &= "    INSERT INTO B_PHAS (NOMG, NOPHASE, OPE, LIBPHASE) VALUES (" & _SQLQ($sID) & ", " & $nFase & ", " & _SQLQ($sOp) & ", " & _SQLQ($sNomF) & ");" & @CRLF
    Next
    Return $s
EndFunc

Func _GetFirstRoutingPhase($sRouting)
    Local $sTarget = StringStripWS(String($sRouting), 3)
    Local $sFirst = ""
    Local $nFirst = 999999999
    Local $nRows = _GUICtrlListView_GetItemCount($g_hLV_Rout)
    For $i = 0 To $nRows - 1
        Local $sID = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Rout, $i, 0), 3)
        If $sID <> $sTarget Then ContinueLoop
        Local $sPhase = _SQLNum(_GUICtrlListView_GetItemText($g_hLV_Rout, $i, 2), "0")
        Local $nPhase = Number($sPhase)
        If $sFirst = "" Or $nPhase < $nFirst Then
            $sFirst = $sPhase
            $nFirst = $nPhase
        EndIf
    Next
    Return $sFirst
EndFunc

Func _GetRoutingOperationForPhase($sRouting, $sPhase)
    Local $sTargetRouting = StringStripWS(String($sRouting), 3)
    Local $sTargetPhase = String(Number(_SQLNum($sPhase, "0")))
    If $sTargetRouting = "" Or $sTargetPhase = "" Then Return ""

    Local $nRows = _GUICtrlListView_GetItemCount($g_hLV_Rout)
    For $i = 0 To $nRows - 1
        Local $sID = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Rout, $i, 0), 3)
        If $sID <> $sTargetRouting Then ContinueLoop

        Local $sRowPhase = String(Number(_SQLNum(_GUICtrlListView_GetItemText($g_hLV_Rout, $i, 2), "0")))
        If $sRowPhase <> $sTargetPhase Then ContinueLoop

        Return StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Rout, $i, 3), 3)
    Next
    Return ""
EndFunc

Func _GenerateItemVersionDependencyCleanupSQL($sID, $sVer)
    ; The original Ortems Toolbox Excel clears B_PART/B_PARTC/I_PART/I_PARTC before B_VER_ART.
    ; Those parameter tables can keep old routing/phase references for an existing item version.
    ; If B_VER_ART.NOMG is changed while those rows still point to an old phase, SQL Server can
    ; raise FK2_B_PART against B_PHAS during the B_VER_ART update. Clean only the impacted item
    ; version, and only when the tables/columns exist in the target schema.
    If StringStripWS(String($sID), 3) = "" Or StringStripWS(String($sVer), 3) = "" Then Return ""

    Local $s = ""
    $s &= "IF OBJECT_ID(N'[dbo].[I_PARTC]', N'U') IS NOT NULL AND COL_LENGTH('dbo.I_PARTC','CODEARTIC') IS NOT NULL AND COL_LENGTH('dbo.I_PARTC','VER_ART') IS NOT NULL" & @CRLF
    $s &= "    DELETE FROM I_PARTC WHERE CODEARTIC=" & _SQLQ($sID) & " AND VER_ART=" & _SQLQ($sVer) & ";" & @CRLF
    $s &= "IF OBJECT_ID(N'[dbo].[I_PART]', N'U') IS NOT NULL AND COL_LENGTH('dbo.I_PART','CODEARTIC') IS NOT NULL AND COL_LENGTH('dbo.I_PART','VER_ART') IS NOT NULL" & @CRLF
    $s &= "    DELETE FROM I_PART WHERE CODEARTIC=" & _SQLQ($sID) & " AND VER_ART=" & _SQLQ($sVer) & ";" & @CRLF
    $s &= "IF OBJECT_ID(N'[dbo].[B_PARTC]', N'U') IS NOT NULL AND COL_LENGTH('dbo.B_PARTC','CODEARTIC') IS NOT NULL AND COL_LENGTH('dbo.B_PARTC','VER_ART') IS NOT NULL" & @CRLF
    $s &= "    DELETE FROM B_PARTC WHERE CODEARTIC=" & _SQLQ($sID) & " AND VER_ART=" & _SQLQ($sVer) & ";" & @CRLF
    $s &= "IF OBJECT_ID(N'[dbo].[B_PART]', N'U') IS NOT NULL AND COL_LENGTH('dbo.B_PART','CODEARTIC') IS NOT NULL AND COL_LENGTH('dbo.B_PART','VER_ART') IS NOT NULL" & @CRLF
    $s &= "    DELETE FROM B_PART WHERE CODEARTIC=" & _SQLQ($sID) & " AND VER_ART=" & _SQLQ($sVer) & ";" & @CRLF
    Return $s
EndFunc

Func _GenerateMatSQL()
    Local $s = ""
    Local $sVerArtPhaseCol = _DetectOptionalColumn("B_VER_ART", "routing_phase", "NOPHASE;NO_PHASE;PHASE;NO_PHAS;B_V_NOPHASE;B_NOPHASE;NPHASE;NUMPHASE;PHASE_DEFAUT")
    If $sVerArtPhaseCol = "" Then $sVerArtPhaseCol = _DetectFKParentColumnByReferenced("B_VER_ART", "B_PHAS", "NOPHASE", "routing_phase")

    ; Some Ortems schemas do not use an obvious phase column name in B_VER_ART.
    ; The reliable source is the SQL Server FK mapping to B_PHAS. If FK2_B_PART
    ; also maps an operation/code column, update it together with the routing and
    ; phase to avoid carrying an old routing-phase combination.
    Local $sVerArtRoutCol = _DetectFKParentColumnByReferenced("B_VER_ART", "B_PHAS", "NOMG", "routing_id")
    Local $sVerArtOpCol = _DetectFKParentColumnByReferenced("B_VER_ART", "B_PHAS", "OPE", "routing_operation")
    Local $nRows = _GUICtrlListView_GetItemCount($g_hLV_Mat)
    For $i = 0 To $nRows - 1
        Local $sID   = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Mat, $i, 0), 3)
        Local $sNome = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Mat, $i, 1), 3)
        Local $sTipo = StringUpper(StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Mat, $i, 2), 3))
        Local $sVer  = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Mat, $i, 3), 3)
        Local $sRot  = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Mat, $i, 4), 3)
        Local $nStk  = _SQLNum(_GUICtrlListView_GetItemText($g_hLV_Mat, $i, 5), "0")
        Local $sFirstPhase = _GetFirstRoutingPhase($sRot)
        Local $sFirstOp = _GetRoutingOperationForPhase($sRot, $sFirstPhase)

        ; Upsert the item header. The Toolbox is frequently used to refresh/customize
        ; an existing demo database, so an existing row must be updated, not ignored.
        $s &= "IF EXISTS (SELECT 1 FROM B_ART WHERE CODEARTIC=" & _SQLQ($sID) & ")" & @CRLF
        $s &= "    UPDATE B_ART SET LIBARTIC=" & _SQLQ($sNome) & ", TYPEMATI=" & _SQLQ($sTipo) & ", QTE_STOCK=" & $nStk & " WHERE CODEARTIC=" & _SQLQ($sID) & @CRLF
        $s &= "ELSE" & @CRLF
        $s &= "    INSERT INTO B_ART (CODEARTIC, LIBARTIC, TYPEMATI, QTE_STOCK) VALUES (" & _SQLQ($sID) & ", " & _SQLQ($sNome) & ", " & _SQLQ($sTipo) & ", " & $nStk & ");" & @CRLF

        If _IsRawMaterialType($sTipo) Then
            $s &= "-- MP raw material " & _SQLQ($sID) & ": B_VER_ART intentionally skipped because MP items do not require version or routing" & @CRLF
            ContinueLoop
        EndIf

        ; Remove stale parameter/phase links for this item version before touching B_VER_ART.
        ; This mirrors the original Excel Toolbox cleanup order for B_PART/B_PARTC and prevents
        ; FK2_B_PART from keeping an old routing/phase combination alive during the update.
        $s &= _GenerateItemVersionDependencyCleanupSQL($sID, $sVer)

        ; Upsert the item version/routing. This is critical when an existing item
        ; version is being reassigned to another routing, for example PINION 4G / 00
        ; from PINION to PINION ALTERNATE. Some Ortems SQL Server schemas also keep
        ; a phase reference in B_VER_ART, so update that phase together with NOMG.
        ; Otherwise SQL Server may reject the update because NOMG + old phase no
        ; longer exists in B_PHAS for the new routing.
        Local $sRoutUpdate = ""
        Local $sRoutInsertCol = ""
        Local $sRoutInsertVal = ""
        If $sVerArtRoutCol <> "" And StringUpper($sVerArtRoutCol) <> "NOMG" Then
            $sRoutUpdate = ", " & _SQLIdent($sVerArtRoutCol) & "=" & _SQLQ($sRot)
            $sRoutInsertCol = ", " & _SQLIdent($sVerArtRoutCol)
            $sRoutInsertVal = ", " & _SQLQ($sRot)
        EndIf

        Local $sPhaseUpdate = ""
        Local $sPhaseInsertCol = ""
        Local $sPhaseInsertVal = ""
        If $sVerArtPhaseCol <> "" And $sFirstPhase <> "" And StringUpper($sVerArtPhaseCol) <> "NOMG" And StringUpper($sVerArtPhaseCol) <> StringUpper($sVerArtRoutCol) Then
            $sPhaseUpdate = ", " & _SQLIdent($sVerArtPhaseCol) & "=" & $sFirstPhase
            $sPhaseInsertCol = ", " & _SQLIdent($sVerArtPhaseCol)
            $sPhaseInsertVal = ", " & $sFirstPhase
        EndIf

        Local $sOpUpdate = ""
        Local $sOpInsertCol = ""
        Local $sOpInsertVal = ""
        If $sVerArtOpCol <> "" And $sFirstOp <> "" And StringUpper($sVerArtOpCol) <> "NOMG" And StringUpper($sVerArtOpCol) <> StringUpper($sVerArtRoutCol) And StringUpper($sVerArtOpCol) <> StringUpper($sVerArtPhaseCol) Then
            $sOpUpdate = ", " & _SQLIdent($sVerArtOpCol) & "=" & _SQLQ($sFirstOp)
            $sOpInsertCol = ", " & _SQLIdent($sVerArtOpCol)
            $sOpInsertVal = ", " & _SQLQ($sFirstOp)
        EndIf

        $s &= "IF EXISTS (SELECT 1 FROM B_VER_ART WHERE CODEARTIC=" & _SQLQ($sID) & " AND VER_ART=" & _SQLQ($sVer) & ")" & @CRLF
        $s &= "    UPDATE B_VER_ART SET VER_DESC=" & _SQLQ($sNome) & ", NOMG=" & _SQLQ($sRot) & $sRoutUpdate & $sPhaseUpdate & $sOpUpdate & ", VER_EFFET_DEBUT=CONVERT(datetime,'01/01/1995',103), VER_EFFET_FIN=CONVERT(datetime,'01/01/2050',103) WHERE CODEARTIC=" & _SQLQ($sID) & " AND VER_ART=" & _SQLQ($sVer) & @CRLF
        $s &= "ELSE" & @CRLF
        $s &= "    INSERT INTO B_VER_ART (CODEARTIC, VER_ART, VER_DESC, NOMG" & $sRoutInsertCol & $sPhaseInsertCol & $sOpInsertCol & ", VER_EFFET_DEBUT, VER_EFFET_FIN)" & @CRLF
        $s &= "    VALUES (" & _SQLQ($sID) & ", " & _SQLQ($sVer) & ", " & _SQLQ($sNome) & ", " & _SQLQ($sRot) & $sRoutInsertVal & $sPhaseInsertVal & $sOpInsertVal & ", CONVERT(datetime,'01/01/1995',103), CONVERT(datetime,'01/01/2050',103));" & @CRLF
    Next
    Return $s
EndFunc

Func _GenerateEGammeNomeSQL()
    ; Some Ortems SQL Server schemas validate B_NOME against E_GAMME_NOME
    ; using the parent item + version + routing. Populate this bridge explicitly
    ; before inserting BOM rows.
    If Not _BaseTableExists("E_GAMME_NOME") Then
        _Log("E_GAMME_NOME table not found in this schema - explicit bridge generation skipped.")
        Return ""
    EndIf

    Local $sItemCol = _DetectOptionalColumn("E_GAMME_NOME", "item",    "CODEARTIC;B_V_CODEARTIC;B_CODEARTIC;ARTICLE")
    Local $sVerCol  = _DetectOptionalColumn("E_GAMME_NOME", "version", "VER_ART;B_VER_ART;VERSION")
    Local $sRoutCol = _DetectOptionalColumn("E_GAMME_NOME", "routing", "NOMG;NOM_GAMME;GAMME")

    If $sItemCol = "" Or $sVerCol = "" Or $sRoutCol = "" Then
        _Log("WARNING: E_GAMME_NOME exists, but required item/version/routing columns were not detected. Bridge generation skipped.")
        Return "-- WARNING: E_GAMME_NOME bridge generation skipped because required columns were not detected" & @CRLF
    EndIf

    Local $s = ""
    Local $sSeen = ";"
    Local $nRows = _GUICtrlListView_GetItemCount($g_hLV_Mat)
    For $i = 0 To $nRows - 1
        Local $sID   = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Mat, $i, 0), 3)
        Local $sTipo = StringUpper(StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Mat, $i, 2), 3))
        Local $sVer  = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Mat, $i, 3), 3)
        Local $sRot  = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Mat, $i, 4), 3)

        If $sID = "" Or $sVer = "" Or $sRot = "" Then ContinueLoop
        If _IsRawMaterialType($sTipo) Then ContinueLoop

        Local $sKey = StringLower($sID & "|" & $sVer & "|" & $sRot)
        If StringInStr($sSeen, ";" & $sKey & ";") Then ContinueLoop
        $sSeen &= $sKey & ";"

        $s &= "IF NOT EXISTS (SELECT 1 FROM E_GAMME_NOME WHERE " & _SQLIdent($sItemCol) & "=" & _SQLQ($sID) & " AND " & _SQLIdent($sVerCol) & "=" & _SQLQ($sVer) & " AND " & _SQLIdent($sRoutCol) & "=" & _SQLQ($sRot) & ")" & @CRLF
        $s &= "    INSERT INTO E_GAMME_NOME (" & _SQLIdent($sItemCol) & ", " & _SQLIdent($sVerCol) & ", " & _SQLIdent($sRoutCol) & ") VALUES (" & _SQLQ($sID) & ", " & _SQLQ($sVer) & ", " & _SQLQ($sRot) & ");" & @CRLF
    Next

    Return $s
EndFunc

Func _GenerateBOMSQL()
    ; B_NOME quantity column names vary between Ortems versions.
    ; Query the live schema to pick the right ones rather than hardcoding.
    Local $sQtyBase = _DetectBNOMEColumn("qty_base", "QTEREF;QTEBASET;CODEBASET;QTE_BASE;QTBASE;QTE1")
    Local $sQtyNec  = _DetectBNOMEColumn("qty_nec",  "QTENEC;QTEBESO;BES_CODEARTIC;QTE_BES;QTBESO;QTE2")

    Local $s = ""
    Local $nRows = _GUICtrlListView_GetItemCount($g_hLV_BOM)
    For $i = 0 To $nRows - 1
        Local $sPaiID  = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_BOM, $i, 0), 3)
        Local $sVPai   = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_BOM, $i, 1), 3)
        Local $sCompID = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_BOM, $i, 2), 3)
        Local $sRotID  = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_BOM, $i, 3), 3)
        Local $nFase   = _SQLNum(_GUICtrlListView_GetItemText($g_hLV_BOM, $i, 4), "0")
        Local $nQRef   = _SQLNum(_GUICtrlListView_GetItemText($g_hLV_BOM, $i, 5), "1")
        Local $nQNec   = _SQLNum(_GUICtrlListView_GetItemText($g_hLV_BOM, $i, 6), "1")

        $s &= "IF NOT EXISTS (SELECT 1 FROM B_NOME WHERE B_V_CODEARTIC=" & _SQLQ($sPaiID) & " AND VER_ART=" & _SQLQ($sVPai) & " AND CODEARTIC=" & _SQLQ($sCompID) & " AND NOMG=" & _SQLQ($sRotID) & " AND NOPHASE=" & $nFase & ")" & @CRLF
        $s &= "    INSERT INTO B_NOME (B_V_CODEARTIC, VER_ART, CODEARTIC, NOMG, NOPHASE, " & $sQtyBase & ", " & $sQtyNec & ")" & @CRLF
        $s &= "    VALUES (" & _SQLQ($sPaiID) & ", " & _SQLQ($sVPai) & ", " & _SQLQ($sCompID) & ", " & _SQLQ($sRotID) & ", " & $nFase & ", " & $nQRef & ", " & $nQNec & ");" & @CRLF
    Next
    Return $s
EndFunc

; Detects which of the candidate column names actually exists in B_NOME.
; $sCandidates = semicolon-separated list tried in order.
; Returns the first match, or the first candidate as fallback.
; Generic: query INFORMATION_SCHEMA.COLUMNS to find the first matching column
; $sTable = table name, $sRole = description for log, $sCandidates = semicolon list
Func _DetectColumn($sTable, $sRole, $sCandidates)
    Local $aCands = StringSplit($sCandidates, ";", 1)
    If Not $g_bConnected Or $g_sConnStr = "" Then Return $aCands[1]

    Local $oConn = ObjCreate("ADODB.Connection")
    If Not IsObj($oConn) Then Return $aCands[1]
    $g_sLastComError = ""
    $oConn.Open($g_sConnStr)
    If $g_sLastComError <> "" Or $oConn.State <> 1 Then Return $aCands[1]

    Local $sQuery = "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS " & _
        "WHERE TABLE_NAME='" & $sTable & "' ORDER BY ORDINAL_POSITION"
    _LogVerboseSQL("Detect column for " & $sTable & "." & $sRole, $sQuery)
    $g_sLastComError = ""
    Local $oRS = $oConn.Execute($sQuery)

    Local $sFound = ""
    If $g_sLastComError = "" And IsObj($oRS) Then
        Local $sActual = ""
        While Not $oRS.EOF
            $sActual &= ";" & $oRS.Fields(0).Value
            $oRS.MoveNext()
        WEnd
        $sActual &= ";"
        $oRS.Close()
        For $i = 1 To $aCands[0]
            If $aCands[$i] = "__NONE__" Then ContinueLoop   ; sentinel - skip
            If StringInStr($sActual, ";" & $aCands[$i] & ";") Then
                $sFound = $aCands[$i]
                _Log($sTable & " schema: column '" & $sFound & "' used for " & $sRole)
                ExitLoop
            EndIf
        Next
        If $sFound = "" Then
            ; Check if __NONE__ sentinel was used - means caller intentionally allows "not found"
            If StringInStr($sCandidates, "__NONE__") Then
                _Log($sTable & " schema: column for " & $sRole & " not present in this version (optional)")
                $sFound = ""   ; caller will check for empty string
            Else
                _Log("WARNING: No match in " & $sTable & " for " & $sRole & ". Actual:" & $sActual)
                $sFound = $aCands[1]   ; fallback to first candidate
            EndIf
        EndIf
    Else
        $sFound = $aCands[1]
    EndIf

    $oConn.Close()
    Return $sFound
EndFunc

; Backward-compat wrapper
Func _DetectBNOMEColumn($sRole, $sCandidates)
    Return _DetectColumn("B_NOME", $sRole, $sCandidates)
EndFunc


Func _BuildWOInList()
    ; Build a SQL IN-list with the WOs currently present in the Work Orders tab.
    ; Duplicates are harmless in SQL, but removing them keeps the generated script smaller.
    Local $sList = ""
    Local $sSeen = Chr(1)
    Local $nRows = _GUICtrlListView_GetItemCount($g_hLV_WO)
    For $i = 0 To $nRows - 1
        Local $sWO = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_WO, $i, 0), 3)
        If $sWO = "" Then ContinueLoop
        Local $sKey = StringReplace($sWO, Chr(1), "")
        If StringInStr($sSeen, Chr(1) & $sKey & Chr(1)) Then ContinueLoop
        $sSeen &= $sKey & Chr(1)
        If $sList <> "" Then $sList &= ", "
        $sList &= _SQLQ($sWO)
    Next
    Return $sList
EndFunc

Func _AppendTargetedDeleteByColumn(ByRef $sSQL, $sTable, $sColumn, $sWOInList)
    ; Guard with OBJECT_ID/COL_LENGTH so the same generated SQL can run across Ortems schemas.
    ; IMPORTANT: the DELETE itself must be executed as dynamic SQL. SQL Server validates
    ; column names at compile time even when the DELETE is protected by IF/COL_LENGTH,
    ; so a static DELETE can still fail with "Invalid column name" for columns that do
    ; not exist in the current Ortems schema.
    If $sWOInList = "" Or $sTable = "" Or $sColumn = "" Then Return
    Local $sTableSafe = StringReplace($sTable, "'", "''")
    Local $sColumnSafe = StringReplace($sColumn, "'", "''")
    Local $sDelete = "DELETE FROM [dbo].[" & $sTable & "] WHERE " & _SQLIdent($sColumn) & " IN (" & $sWOInList & ")"
    Local $sDeleteExec = StringReplace($sDelete, "'", "''")
    $sSQL &= "IF OBJECT_ID(N'[dbo].[" & $sTable & "]', N'U') IS NOT NULL AND COL_LENGTH('dbo." & $sTableSafe & "','" & $sColumnSafe & "') IS NOT NULL" & @CRLF
    $sSQL &= "    EXEC(N'" & $sDeleteExec & "');" & @CRLF
EndFunc

Func _AppendTargetedDeleteForColumns(ByRef $sSQL, $sTable, $sColumns, $sWOInList)
    Local $aCols = StringSplit($sColumns, ";", 2)
    For $i = 0 To UBound($aCols) - 1
        Local $sCol = StringStripWS($aCols[$i], 3)
        If $sCol = "" Then ContinueLoop
        _AppendTargetedDeleteByColumn($sSQL, $sTable, $sCol, $sWOInList)
    Next
EndFunc

Func _CountSQLInListValues($sInList)
    Local $s = StringStripWS(String($sInList), 3)
    If $s = "" Then Return 0
    Local $a = StringSplit($s, ",", 1)
    Return $a[0]
EndFunc

Func _GenerateTargetedWOCleanupSQL()
    ; Refresh only the WOs being imported. This avoids full database cleanup while still
    ; removing stale phase/link rows that can block E_OF updates through FK1_B_BT.
    Local $sWOInList = _BuildWOInList()
    If $sWOInList = "" Then Return "-- No Work Orders found in the workbook; targeted WO cleanup skipped." & @CRLF

    Local $s = ""
    Local $sWOCols = "E_NOF;NOF;BT_NOF;B_NOF;B_O_NOF;O_NOF;NO_OF;N_OF;NUMOF;NUM_OF;OF;ID_OF;BT_OF;NOF_PERE;NOF_FILS;NOF_PREC;NOF_SUIV;B_O_NOF_PERE;B_O_NOF_FILS"

    ; Child/link tables first. These may reference B_BT, B_OF or E_OF depending on schema/version.
    _AppendTargetedDeleteForColumns($s, "B_BT2",  $sWOCols, $sWOInList)
    _AppendTargetedDeleteForColumns($s, "B_PREN2", $sWOCols, $sWOInList)
    _AppendTargetedDeleteForColumns($s, "B_PROF",  $sWOCols, $sWOInList)
    _AppendTargetedDeleteForColumns($s, "B_PHM",   $sWOCols, $sWOInList)
    _AppendTargetedDeleteForColumns($s, "E_OF2",   $sWOCols, $sWOInList)
    _AppendTargetedDeleteForColumns($s, "ENC_CAD", $sWOCols, $sWOInList)

    ; ENC_CAD can reference B_BT in some Ortems schemas (for example through FK1_ENC_CAD).
    ; It must be removed before B_BT, otherwise the B_BT cleanup can be blocked.

    ; Main WO phase/header rows. B_BT must be removed before E_OF can be safely updated.
    _AppendTargetedDeleteForColumns($s, "B_BT",    $sWOCols, $sWOInList)
    _AppendTargetedDeleteForColumns($s, "B_OF",    "NOF;E_NOF;NUMOF;NUM_OF;NO_OF;N_OF;OF;ID_OF", $sWOInList)

    ; Do not delete E_OF here. After B_BT/B_OF are cleaned, _GenerateWOHeaderSQL() can update
    ; the existing E_OF row safely, preserving any columns outside the Toolbox mapping.
    _Log("Targeted WO cleanup generated for imported Work Orders. Existing ENC_CAD/B_BT/B_OF rows for those WOs will be refreshed before E_OF is updated.")
    _LogVerbose("Targeted WO cleanup IN-list contains " & _CountSQLInListValues($sWOInList) & " Work Order(s).")
    _LogVerboseSQL("Generated targeted WO cleanup SQL", $s)
    Return $s
EndFunc

Func _GenerateSingleWOCleanupSQL($sWO)
    ; Defensive cleanup for a single WO immediately before the E_OF upsert.
    ; This handles existing demo databases where stale B_BT rows still reference E_OF
    ; and block an UPDATE of the E_OF parent row through FK1_B_BT.
    Local $sWOValue = StringStripWS(String($sWO), 3)
    If $sWOValue = "" Then Return ""

    Local $sOneWOInList = _SQLQ($sWOValue)
    Local $s = ""
    Local $sWOCols = "E_NOF;NOF;BT_NOF;B_NOF;B_O_NOF;O_NOF;NO_OF;N_OF;NUMOF;NUM_OF;OF;ID_OF;BT_OF;NOF_PERE;NOF_FILS;NOF_PREC;NOF_SUIV;B_O_NOF_PERE;B_O_NOF_FILS"

    $s &= "-- Defensive cleanup for WO " & _SQLQ($sWOValue) & " before E_OF upsert" & @CRLF

    ; Child/link tables first. These are the common tables that can reference B_BT/B_OF/E_OF.
    _AppendTargetedDeleteForColumns($s, "B_BT2",  $sWOCols, $sOneWOInList)
    _AppendTargetedDeleteForColumns($s, "B_PREN2", $sWOCols, $sOneWOInList)
    _AppendTargetedDeleteForColumns($s, "B_PROF",  $sWOCols, $sOneWOInList)
    _AppendTargetedDeleteForColumns($s, "B_PHM",   $sWOCols, $sOneWOInList)
    _AppendTargetedDeleteForColumns($s, "E_OF2",   $sWOCols, $sOneWOInList)
    _AppendTargetedDeleteForColumns($s, "ENC_CAD", $sWOCols, $sOneWOInList)

    ; ENC_CAD can reference B_BT in some Ortems schemas (for example through FK1_ENC_CAD).
    ; It must be removed before B_BT, otherwise the B_BT cleanup can be blocked.

    ; Main WO rows. B_BT must be removed before updating E_OF, otherwise FK1_B_BT can block it.
    _AppendTargetedDeleteForColumns($s, "B_BT",    $sWOCols, $sOneWOInList)
    _AppendTargetedDeleteForColumns($s, "B_OF",    "NOF;E_NOF;NUMOF;NUM_OF;NO_OF;N_OF;OF;ID_OF", $sOneWOInList)

    _LogVerboseSQL("Generated defensive single-WO cleanup SQL for " & $sWOValue, $s)
    Return $s
EndFunc


Func _GenerateWOReferenceDataSQL()
    Local $s = ""

    ; In global clean-import mode all user tables are emptied, including Ortems
    ; reference/master tables. E_OF.SERIE is constrained to B_SER.SERIE in the
    ; SQL Server schema seen in the logs, so the default series used below must
    ; be recreated before E_OF is inserted.
    $s &= _GenerateSeriesReferenceSQL("0") & @CRLF
    $s &= _GenerateSimpleReferenceRowSQL("PLAN_CAP", "PLAN_CAP", $ORTEMS_DEFAULT_NOM_ENC) & @CRLF
    $s &= _GenerateSimpleReferenceRowSQL("ENCOURS", "NOM_ENC", $ORTEMS_DEFAULT_NOM_ENC) & @CRLF

    Return $s
EndFunc

Func _GenerateSimpleReferenceRowSQL($sTable, $sKeyCol, $sKeyValue)
    Local $s = ""
    Local $sTableSafe = StringReplace($sTable, "'", "''")
    Local $sKeyColSafe = StringReplace($sKeyCol, "'", "''")
    Local $sTableBracket = "[dbo].[" & StringReplace($sTable, "]", "]]" ) & "]"

    $s &= "IF OBJECT_ID(N'[dbo].[" & $sTableSafe & "]', N'U') IS NOT NULL AND COL_LENGTH('dbo." & $sTableSafe & "','" & $sKeyColSafe & "') IS NOT NULL" & @CRLF
    $s &= "BEGIN" & @CRLF
    $s &= "    IF NOT EXISTS (SELECT 1 FROM " & $sTableBracket & " WHERE [" & StringReplace($sKeyCol, "]", "]]" ) & "] = " & _SQLQ($sKeyValue) & ")" & @CRLF
    $s &= "    BEGIN" & @CRLF
    $s &= "        DECLARE @cols NVARCHAR(MAX), @vals NVARCHAR(MAX), @sql NVARCHAR(MAX)" & @CRLF
    $s &= "        SET @cols = N''" & @CRLF
    $s &= "        SET @vals = N''" & @CRLF
    $s &= "        SELECT" & @CRLF
    $s &= "            @cols = @cols + CASE WHEN LEN(@cols) > 0 THEN N', ' ELSE N'' END + QUOTENAME(COLUMN_NAME)," & @CRLF
    $s &= "            @vals = @vals + CASE WHEN LEN(@vals) > 0 THEN N', ' ELSE N'' END +" & @CRLF
    $s &= "                CASE" & @CRLF
    $s &= "                    WHEN UPPER(COLUMN_NAME) = UPPER(N'" & $sKeyColSafe & "') THEN N" & _SQLQ(_SQLQ($sKeyValue)) & "" & @CRLF
    $s &= "                    WHEN DATA_TYPE IN (N'date') THEN N'CONVERT(date,''19950101'',112)'" & @CRLF
    $s &= "                    WHEN DATA_TYPE IN (N'datetime', N'datetime2', N'smalldatetime') THEN N'CONVERT(datetime,''19950101'',112)'" & @CRLF
    $s &= "                    WHEN DATA_TYPE IN (N'time') THEN N'CONVERT(time,''00:00'')'" & @CRLF
    $s &= "                    WHEN DATA_TYPE IN (N'bit') THEN N'0'" & @CRLF
    $s &= "                    WHEN DATA_TYPE IN (N'uniqueidentifier') THEN N'NEWID()'" & @CRLF
    $s &= "                    WHEN DATA_TYPE IN (N'char', N'nchar', N'varchar', N'nvarchar', N'text', N'ntext') THEN N" & _SQLQ(_SQLQ($sKeyValue)) & "" & @CRLF
    $s &= "                    ELSE N'0'" & @CRLF
    $s &= "                END" & @CRLF
    $s &= "        FROM INFORMATION_SCHEMA.COLUMNS" & @CRLF
    $s &= "        WHERE TABLE_SCHEMA = 'dbo'" & @CRLF
    $s &= "          AND TABLE_NAME = '" & $sTableSafe & "'" & @CRLF
    $s &= "          AND (" & @CRLF
    $s &= "                UPPER(COLUMN_NAME) = UPPER(N'" & $sKeyColSafe & "')" & @CRLF
    $s &= "                OR (IS_NULLABLE = 'NO'" & @CRLF
    $s &= "                    AND COLUMN_DEFAULT IS NULL" & @CRLF
    $s &= "                    AND COLUMNPROPERTY(OBJECT_ID(TABLE_SCHEMA + '.' + TABLE_NAME), COLUMN_NAME, 'IsIdentity') = 0)" & @CRLF
    $s &= "              )" & @CRLF
    $s &= "        ORDER BY CASE WHEN UPPER(COLUMN_NAME) = UPPER(N'" & $sKeyColSafe & "') THEN 0 ELSE ORDINAL_POSITION END" & @CRLF
    $s &= "        IF LEN(@cols) > 0" & @CRLF
    $s &= "        BEGIN" & @CRLF
    $s &= "            SET @sql = N'INSERT INTO " & $sTableBracket & " (' + @cols + N') VALUES (' + @vals + N')'" & @CRLF
    $s &= "            EXEC sp_executesql @sql" & @CRLF
    $s &= "        END" & @CRLF
    $s &= "    END" & @CRLF
    $s &= "END;" & @CRLF

    _LogVerboseSQL("Generated required Ortems reference seed SQL for " & $sTable & "." & $sKeyCol & " = " & $sKeyValue, $s)
    Return $s
EndFunc

Func _GenerateSeriesReferenceSQL($sSeries)
    Local $s = ""
    $s &= "IF OBJECT_ID(N'[dbo].[B_SER]', N'U') IS NOT NULL AND COL_LENGTH('dbo.B_SER','SERIE') IS NOT NULL" & @CRLF
    $s &= "BEGIN" & @CRLF
    $s &= "    IF NOT EXISTS (SELECT 1 FROM [dbo].[B_SER] WHERE [SERIE] = " & _SQLQ($sSeries) & ")" & @CRLF
    $s &= "    BEGIN" & @CRLF
    $s &= "        DECLARE @cols NVARCHAR(MAX), @vals NVARCHAR(MAX), @sql NVARCHAR(MAX)" & @CRLF
    $s &= "        SET @cols = N''" & @CRLF
    $s &= "        SET @vals = N''" & @CRLF
    $s &= "        SELECT" & @CRLF
    $s &= "            @cols = @cols + CASE WHEN LEN(@cols) > 0 THEN N', ' ELSE N'' END + QUOTENAME(COLUMN_NAME)," & @CRLF
    $s &= "            @vals = @vals + CASE WHEN LEN(@vals) > 0 THEN N', ' ELSE N'' END +" & @CRLF
    $s &= "                CASE" & @CRLF
    $s &= "                    WHEN UPPER(COLUMN_NAME) IN (N'SERIE', N'E_SERIE', N'NO_SERIE', N'NUMSERIE') THEN N" & _SQLQ(_SQLQ($sSeries)) & "" & @CRLF
    $s &= "                    WHEN DATA_TYPE IN (N'date') THEN N'CONVERT(date,''19950101'',112)'" & @CRLF
    $s &= "                    WHEN DATA_TYPE IN (N'datetime', N'datetime2', N'smalldatetime') THEN N'CONVERT(datetime,''19950101'',112)'" & @CRLF
    $s &= "                    WHEN DATA_TYPE IN (N'time') THEN N'CONVERT(time,''00:00'')'" & @CRLF
    $s &= "                    WHEN DATA_TYPE IN (N'bit') THEN N'0'" & @CRLF
    $s &= "                    WHEN DATA_TYPE IN (N'uniqueidentifier') THEN N'NEWID()'" & @CRLF
    $s &= "                    WHEN DATA_TYPE IN (N'char', N'nchar', N'varchar', N'nvarchar', N'text', N'ntext') THEN N" & _SQLQ(_SQLQ($sSeries)) & "" & @CRLF
    $s &= "                    ELSE N'0'" & @CRLF
    $s &= "                END" & @CRLF
    $s &= "        FROM INFORMATION_SCHEMA.COLUMNS" & @CRLF
    $s &= "        WHERE TABLE_SCHEMA = 'dbo'" & @CRLF
    $s &= "          AND TABLE_NAME = 'B_SER'" & @CRLF
    $s &= "          AND (" & @CRLF
    $s &= "                UPPER(COLUMN_NAME) IN (N'SERIE', N'E_SERIE', N'NO_SERIE', N'NUMSERIE')" & @CRLF
    $s &= "                OR (IS_NULLABLE = 'NO'" & @CRLF
    $s &= "                    AND COLUMN_DEFAULT IS NULL" & @CRLF
    $s &= "                    AND COLUMNPROPERTY(OBJECT_ID(TABLE_SCHEMA + '.' + TABLE_NAME), COLUMN_NAME, 'IsIdentity') = 0)" & @CRLF
    $s &= "              )" & @CRLF
    $s &= "        ORDER BY CASE WHEN UPPER(COLUMN_NAME) = 'SERIE' THEN 0 ELSE ORDINAL_POSITION END" & @CRLF
    $s &= "        IF LEN(@cols) > 0" & @CRLF
    $s &= "        BEGIN" & @CRLF
    $s &= "            SET @sql = N'INSERT INTO [dbo].[B_SER] (' + @cols + N') VALUES (' + @vals + N')'" & @CRLF
    $s &= "            EXEC sp_executesql @sql" & @CRLF
    $s &= "        END" & @CRLF
    $s &= "    END" & @CRLF
    $s &= "END;" & @CRLF

    _LogVerboseSQL("Generated B_SER reference seed SQL for series " & $sSeries, $s)
    Return $s
EndFunc

Func _GenerateWOHeaderSQL()
    ; Some Ortems SQL Server schemas use E_OF as the WO parent table for B_BT.
    ; If this parent row is missing, SQL Server rejects B_BT with FK1_B_BT.
    If Not _BaseTableExists("E_OF") Then
        _Log("E_OF table not found in this schema - WO header generation skipped.")
        Return ""
    EndIf

    Local $sWOCol = _DetectOptionalColumn("E_OF", "work_order", "NOF;E_NOF;BT_NOF;NUMOF;NUM_OF;NO_OF;N_OF;OF;ID_OF")
    If $sWOCol = "" Then
        Local $sActualCols = _GetTableColumnsForLog("E_OF")
        _Log("WARNING: E_OF exists, but the work-order key column was not detected. WO header generation skipped.")
        If $sActualCols <> "" Then _Log("E_OF actual columns: " & $sActualCols)
        Return "-- WARNING: E_OF WO header generation skipped because the WO key column was not detected" & @CRLF
    EndIf

    Local $sItemCol   = _DetectOptionalColumn("E_OF", "item",       "CODEARTIC;B_CODEARTIC;E_CODEARTIC;ARTICLE;ITEM;B_V_CODEARTIC")
    ; NOM_ENC is the Ortems en-cours/plan key and usually references ENCOURS.NOM_ENC.
    ; It must not receive the routing/gamme value (for example FINAL ASSEMBLY).
    Local $sEncCol    = _DetectFKParentColumnByReferenced("E_OF", "ENCOURS", "NOM_ENC", "en_cours")
    If $sEncCol = "" Then $sEncCol = _DetectOptionalColumn("E_OF", "en_cours", "NOM_ENC;E_NOM_ENC;PLAN_CAP")
    Local $sRoutCol   = _DetectOptionalColumn("E_OF", "routing",    "NOMG;E_NOMG;NOM_GAMME;GAMME;ROUTING")
    If $sRoutCol = $sEncCol Then $sRoutCol = ""
    Local $sVerCol    = _DetectOptionalColumn("E_OF", "version",    "VER_ART;E_VER_ART;VERSION;B_VER_ART")
    Local $sSerieCol  = _DetectOptionalColumn("E_OF", "serie",      "SERIE;E_SERIE")
    Local $sQtyCol    = _DetectOptionalColumn("E_OF", "qty_order",  "QTE;QTEOF;QTE_ORDER;QTEOFORDER;QTECMD;QTE_LANCEE;QTE_LANC")
    Local $sStartCol  = _DetectOptionalColumn("E_OF", "start_date", "DPLUSTOT;DATEDEB;DATDEB;DATE_DEBUT;DEBUT;DATE_PLA;DATE_LANCEMENT;DATE_DEB")
    Local $sEndCol    = _DetectOptionalColumn("E_OF", "end_date",   "FPLUSTARD;DATEFIN;DATFIN;DATE_FIN;FIN;DATE_BESOIN;DATE_ECH")
    Local $sGestCol   = _DetectOptionalColumn("E_OF", "codegest",   "CODEGEST;E_CODEGEST")
    Local $sStatusCol = _DetectOptionalColumn("E_OF", "status",     "ETATOF;ETAT;STATUS;STATUT;OF_STATUS")
    Local $sModeCol   = _DetectOptionalColumn("E_OF", "mode_util",  "MODE_UTIL;MODE_GESTION;MODUTIL")

    Local $s = ""
    Local $nRows = _GUICtrlListView_GetItemCount($g_hLV_WO)
    For $i = 0 To $nRows - 1
        Local $sNum  = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_WO, $i, 0), 3)
        Local $sMat  = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_WO, $i, 1), 3)
        Local $sRot  = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_WO, $i, 2), 3)
        Local $sVer  = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_WO, $i, 3), 3)
        Local $nQtd  = _SQLNum(_GUICtrlListView_GetItemText($g_hLV_WO, $i, 4), "0")
        Local $sDtI  = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_WO, $i, 5), 3)
        Local $sDtF  = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_WO, $i, 6), 3)
        Local $sFirstPhase = _GetFirstRoutingPhase($sRot)
        Local $sFirstOp = _LookupFirstOperationByRouting($sRot)
        Local $sMach = _LookupUniqueMachineByOperation($sFirstOp)
        Local $sWC = _LookupWorkCenterByOperation($sFirstOp)

        Local $sCols = _SQLIdent($sWOCol)
        Local $sVals = _SQLQ($sNum)

        _AppendColumnValue($sCols, $sVals, $sItemCol,   _SQLQ($sMat))
        _AppendColumnValue($sCols, $sVals, $sEncCol,    _SQLQ($ORTEMS_DEFAULT_NOM_ENC))
        _AppendColumnValue($sCols, $sVals, $sRoutCol,   _SQLQ($sRot))
        _AppendColumnValue($sCols, $sVals, $sVerCol,    _SQLQ($sVer))
        _AppendColumnValue($sCols, $sVals, $sSerieCol,  "'0'")
        _AppendColumnValue($sCols, $sVals, $sQtyCol,    $nQtd)
        _AppendColumnValue($sCols, $sVals, $sStartCol,  _SQLDateTime103($sDtI, "01/01/1995 00:00"))
        _AppendColumnValue($sCols, $sVals, $sEndCol,    _SQLDateTime103($sDtF, "01/01/2050 23:59"))
        _AppendColumnValue($sCols, $sVals, $sGestCol,   "'F'")
        _AppendColumnValue($sCols, $sVals, $sStatusCol, "'S'")
        _AppendColumnValue($sCols, $sVals, $sModeCol,   "'C'")

        _AppendWOHeaderRequiredDefaults($sCols, $sVals, $sNum, $sMat, $sRot, $sVer, $nQtd, $sDtI, $sDtF, $sFirstPhase, $sFirstOp, $sMach, $sWC)

        ; Keep the E_OF parent row aligned with B_BT. In some customer databases,
        ; B_BT is constrained to E_OF through E_NOF + NOM_ENC. NOM_ENC is the
        ; en-cours/plan key, not the routing/gamme. Update by WO first.
        Local $sUpdateSet = ""
        _AppendUpdateSet($sUpdateSet, $sItemCol,   _SQLQ($sMat))
        _AppendUpdateSet($sUpdateSet, $sEncCol,    _SQLQ($ORTEMS_DEFAULT_NOM_ENC))
        _AppendUpdateSet($sUpdateSet, $sRoutCol,   _SQLQ($sRot))
        _AppendUpdateSet($sUpdateSet, $sVerCol,    _SQLQ($sVer))
        _AppendUpdateSet($sUpdateSet, $sSerieCol,  "'0'")
        _AppendUpdateSet($sUpdateSet, $sQtyCol,    $nQtd)
        _AppendUpdateSet($sUpdateSet, $sStartCol,  _SQLDateTime103($sDtI, "01/01/1995 00:00"))
        _AppendUpdateSet($sUpdateSet, $sEndCol,    _SQLDateTime103($sDtF, "01/01/2050 23:59"))
        _AppendUpdateSet($sUpdateSet, $sGestCol,   "'F'")
        _AppendUpdateSet($sUpdateSet, $sStatusCol, "'S'")
        _AppendUpdateSet($sUpdateSet, $sModeCol,   "'C'")

        ; E_OF is keyed by WO + en-cours in many Ortems SQL Server schemas (for example E_NOF + NOM_ENC).
        ; Even after a global clean, Ortems triggers or duplicate workbook rows can make a plain INSERT collide.
        ; Always guard/upsert by the same composite key that B_BT uses to reference E_OF.
        Local $sHeaderWhere = ""
        _AppendWhereEquals($sHeaderWhere, $sWOCol, _SQLQ($sNum))
        _AppendWhereEquals($sHeaderWhere, $sEncCol, _SQLQ($ORTEMS_DEFAULT_NOM_ENC))
        If $sHeaderWhere = "" Then $sHeaderWhere = _SQLIdent($sWOCol) & "=" & _SQLQ($sNum)

        If $sUpdateSet <> "" Then
            $s &= "IF EXISTS (SELECT 1 FROM E_OF WHERE " & $sHeaderWhere & ")" & @CRLF
            $s &= "    UPDATE E_OF SET " & $sUpdateSet & " WHERE " & $sHeaderWhere & @CRLF
            $s &= "ELSE" & @CRLF
            $s &= "    INSERT INTO E_OF (" & $sCols & ")" & @CRLF
            $s &= "    VALUES (" & $sVals & ");" & @CRLF
        Else
            $s &= "IF NOT EXISTS (SELECT 1 FROM E_OF WHERE " & $sHeaderWhere & ")" & @CRLF
            $s &= "    INSERT INTO E_OF (" & $sCols & ")" & @CRLF
            $s &= "    VALUES (" & $sVals & ");" & @CRLF
        EndIf
    Next

    Return $s
EndFunc

Func _GenerateWOSQL()
    ; Detect version-dependent B_OF column names at runtime
    Local $sQteCol = _DetectColumn("B_OF", "qty_order", "QTE;QTEORDER;QTEOF;QTE_ORDER;QTEOFORDER;QTECMD")

    ; MODE_UTIL exists in some Ortems versions but not all
    Local $sModeUtil = _DetectOptionalColumn("B_OF", "mode_util", "MODE_UTIL;MODE_GESTION;MODUTIL")
    Local $bHasModeUtil = ($sModeUtil <> "")

    Local $s = ""
    Local $nRows = _GUICtrlListView_GetItemCount($g_hLV_WO)
    For $i = 0 To $nRows - 1
        Local $sNum  = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_WO, $i, 0), 3)
        Local $sMat  = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_WO, $i, 1), 3)
        Local $sRot  = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_WO, $i, 2), 3)
        Local $sVer  = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_WO, $i, 3), 3)
        Local $nQtd  = _SQLNum(_GUICtrlListView_GetItemText($g_hLV_WO, $i, 4), "0")
        Local $sDtI  = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_WO, $i, 5), 3)
        Local $sDtF  = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_WO, $i, 6), 3)

        Local $sColList = "NOF, CODEARTIC, NOMG, SERIE, VER_ART, " & $sQteCol & ", DPLUSTOT, FPLUSTARD, CODEGEST, ETATOF, VER_EFFET_DEBUT, VER_EFFET_FIN"
        Local $sValList = _SQLQ($sNum) & ", " & _SQLQ($sMat) & ", " & _SQLQ($sRot) & ", '0', " & _SQLQ($sVer) & ", " & $nQtd & ", " & _
            _SQLDateTime103($sDtI, "01/01/1995 00:00") & ", " & _SQLDateTime103($sDtF, "01/01/2050 23:59") & ", " & _
            "'F', 'S', CONVERT(datetime,'01/01/1995',103), CONVERT(datetime,'01/01/2050',103)"

        If $bHasModeUtil Then
            $sColList &= ", " & $sModeUtil
            $sValList &= ", 'C'"
        EndIf

        $s &= "IF NOT EXISTS (SELECT 1 FROM B_OF WHERE NOF=" & _SQLQ($sNum) & ")" & @CRLF
        $s &= "    INSERT INTO B_OF (" & $sColList & ")" & @CRLF
        $s &= "    VALUES (" & $sValList & ");" & @CRLF
    Next
    Return $s
EndFunc

Func _GenerateWOBTSQL()
    ; PL_0170 in Ortems is commonly raised when a WO exists without phase rows.
    ; The B_BT schema is not identical across Ortems versions, so do not assume
    ; NOF/NOMG/NOPHASE are the physical column names. Detect them before generating SQL.
    ; Some SQL schemas use E_NOF for the WO key and NOM_ENC for the en-cours/plan key.
    ; NOM_ENC must match E_OF.NOM_ENC (= ENCOURS), while the routing/gamme stays on NOMG/BT_NOMG when present.
    Local $sWOCol      = _DetectOptionalColumn("B_BT", "work_order", "E_NOF;NOF;BT_NOF;B_NOF;B_O_NOF;O_NOF;NO_OF;N_OF;NUMOF;NUM_OF;OF;ID_OF;BT_OF")
    If $sWOCol = "" Then $sWOCol = _DetectFKParentColumnByReferenced("B_BT", "B_OF", "NOF", "work_order")

    Local $sEncCol     = _DetectFKParentColumnByReferenced("B_BT", "E_OF", "NOM_ENC", "en_cours")
    If $sEncCol = "" Then $sEncCol = _DetectOptionalColumn("B_BT", "en_cours", "NOM_ENC;B_B_NOM_ENC;E_NOM_ENC")

    Local $sRoutCol    = _DetectOptionalColumn("B_BT", "routing",    "NOMG;BT_NOMG;B_NOMG;B_P_NOMG;NOM_GAMME;NOGAMME;NO_GAMME;GAMME;ROUTING")
    If $sRoutCol = "" Then $sRoutCol = _DetectFKParentColumnByReferenced("B_BT", "B_PHAS", "NOMG", "routing")
    If $sRoutCol = $sEncCol Then $sRoutCol = ""

    Local $sPhaseCol   = _DetectOptionalColumn("B_BT", "phase",      "NOPHASE;BT_NOPHASE;B_NOPHASE;B_P_NOPHASE;NO_PHASE;NO_PHAS;PHASE;NPHASE;NUMPHASE;NUM_PHASE")
    If $sPhaseCol = "" Then $sPhaseCol = _DetectFKParentColumnByReferenced("B_BT", "B_PHAS", "NOPHASE", "phase")

    Local $sOpeCol     = _DetectOptionalColumn("B_BT", "operation",  "OPE;CODEOPE;BT_OPE")
    Local $sMachCol    = _DetectOptionalColumn("B_BT", "machine",    "MACHINE;MACH;BT_MACHINE")
    Local $sItemCol    = _DetectOptionalColumn("B_BT", "item",       "CODEARTIC;B_CODEARTIC;BT_CODEARTIC")
    Local $sVerCol     = _DetectOptionalColumn("B_BT", "version",    "VER_ART;B_VER_ART;BT_VER_ART")
    Local $sSerieCol   = _DetectOptionalColumn("B_BT", "serie",      "SERIE;BT_SERIE")
    Local $sQtyCol     = _DetectOptionalColumn("B_BT", "qty",        "QTE;QTEOF;QTEBT;BT_QTE;QTE_ORDER")
    Local $sStartCol   = _DetectOptionalColumn("B_BT", "start_date", "DPLUSTOT;DATEDEB;DATDEB;BT_DEBUT;DEBUT")
    Local $sEndCol     = _DetectOptionalColumn("B_BT", "end_date",   "FPLUSTARD;DATEFIN;DATFIN;BT_FIN;FIN")
    Local $sGestCol    = _DetectOptionalColumn("B_BT", "codegest",   "CODEGEST;BT_CODEGEST")
    Local $sStatusCol  = _DetectOptionalColumn("B_BT", "status",     "ETATBT;ETAT;BT_ETAT")
    Local $sModeCol    = _DetectOptionalColumn("B_BT", "mode_util",  "MODE_UTIL;MODE_GESTION;MODUTIL")

    Local $s = ""

    If $sWOCol = "" Or $sEncCol = "" Or $sPhaseCol = "" Then
        Local $sActualCols = _GetTableColumnsForLog("B_BT")
        $s &= "-- WARNING: B_BT phase generation skipped because required WO/en-cours/phase columns were not detected in this Ortems schema." & @CRLF
        $s &= "-- Check the verbose log for the actual B_BT column list and add the proper column aliases to _GenerateWOBTSQL()." & @CRLF
        _Log("WARNING: B_BT generation skipped. Missing required key mapping: WO='" & $sWOCol & "', en_cours='" & $sEncCol & "', phase='" & $sPhaseCol & "'. Optional routing='" & $sRoutCol & "'.")
        If $sActualCols <> "" Then _Log("B_BT actual columns: " & $sActualCols)
        Return $s
    EndIf

    Local $nWO = _GUICtrlListView_GetItemCount($g_hLV_WO)
    Local $nRout = _GUICtrlListView_GetItemCount($g_hLV_Rout)

    For $i = 0 To $nWO - 1
        Local $sNum = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_WO, $i, 0), 3)
        Local $sMat = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_WO, $i, 1), 3)
        Local $sRot = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_WO, $i, 2), 3)
        Local $sVer = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_WO, $i, 3), 3)
        Local $nQtd = _SQLNum(_GUICtrlListView_GetItemText($g_hLV_WO, $i, 4), "0")
        Local $sDtI = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_WO, $i, 5), 3)
        Local $sDtF = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_WO, $i, 6), 3)

        For $r = 0 To $nRout - 1
            Local $sRoutID = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Rout, $r, 0), 3)
            If $sRoutID <> $sRot Then ContinueLoop

            Local $nFase = _SQLNum(_GUICtrlListView_GetItemText($g_hLV_Rout, $r, 2), "0")
            Local $sOpe  = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Rout, $r, 3), 3)
            Local $sMach = _LookupUniqueMachineByOperation($sOpe)
            Local $sWC = _LookupWorkCenterByOperation($sOpe)
            Local $sPhaseName = _LookupPhaseName($sRot, $nFase)

            Local $sCols = _SQLIdent($sWOCol) & ", " & _SQLIdent($sEncCol) & ", " & _SQLIdent($sPhaseCol)
            Local $sVals = _SQLQ($sNum) & ", " & _SQLQ($ORTEMS_DEFAULT_NOM_ENC) & ", " & $nFase
            Local $sWhere = ""
            _AppendWhereEquals($sWhere, $sWOCol, _SQLQ($sNum))
            _AppendWhereEquals($sWhere, $sEncCol, _SQLQ($ORTEMS_DEFAULT_NOM_ENC))
            _AppendWhereEquals($sWhere, $sPhaseCol, $nFase)

            _AppendColumnValue($sCols, $sVals, $sRoutCol,  _SQLQ($sRot))

            _AppendColumnValue($sCols, $sVals, $sOpeCol,    _SQLQ($sOpe))
            _AppendColumnValue($sCols, $sVals, $sMachCol,   _SQLQ($sMach))
            _AppendColumnValue($sCols, $sVals, $sItemCol,   _SQLQ($sMat))
            _AppendColumnValue($sCols, $sVals, $sVerCol,    _SQLQ($sVer))
            _AppendColumnValue($sCols, $sVals, $sSerieCol,  "'0'")
            _AppendColumnValue($sCols, $sVals, $sQtyCol,    $nQtd)
            _AppendColumnValue($sCols, $sVals, $sStartCol,  _SQLDateTime103($sDtI, "01/01/1995 00:00"))
            _AppendColumnValue($sCols, $sVals, $sEndCol,    _SQLDateTime103($sDtF, "01/01/2050 23:59"))
            _AppendColumnValue($sCols, $sVals, $sGestCol,   "'F'")
            _AppendColumnValue($sCols, $sVals, $sStatusCol, "'S'")
            _AppendColumnValue($sCols, $sVals, $sModeCol,   "'C'")

            ; Some B_BT schemas define additional NOT NULL columns without SQL defaults
            ; (for example BT_FRACTION). Populate those dynamically to avoid rollback.
            _AppendBTTRequiredDefaults($sCols, $sVals, $sNum, $sRot, $nFase, $sOpe, $sMach, $sWC, $sMat, $sVer, $nQtd, $sDtI, $sDtF, $sPhaseName)

            $s &= "IF NOT EXISTS (SELECT 1 FROM B_BT WHERE " & $sWhere & ")" & @CRLF
            $s &= "    INSERT INTO B_BT (" & $sCols & ")" & @CRLF
            $s &= "    VALUES (" & $sVals & ");" & @CRLF
        Next
    Next

    Return $s
EndFunc

Func _GenerateWOLSQL()
    Local $s = ""
    Local $nRows = _GUICtrlListView_GetItemCount($g_hLV_WOL)
    For $i = 0 To $nRows - 1
        Local $sNOF_P = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_WOL, $i, 0), 3)
        Local $sNOMG_P= StringStripWS(_GUICtrlListView_GetItemText($g_hLV_WOL, $i, 1), 3)
        Local $sFas_P = _SQLNum(_GUICtrlListView_GetItemText($g_hLV_WOL, $i, 2), "0")
        Local $sNOF_S = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_WOL, $i, 3), 3)
        Local $sNOMG_S= StringStripWS(_GUICtrlListView_GetItemText($g_hLV_WOL, $i, 4), 3)
        Local $sFas_S = _SQLNum(_GUICtrlListView_GetItemText($g_hLV_WOL, $i, 5), "0")
        Local $sTipo  = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_WOL, $i, 6), 3)

        $s &= "IF NOT EXISTS (SELECT 1 FROM B_PROF WHERE NOF=" & _SQLQ($sNOF_S) & " AND NOMG=" & _SQLQ($sNOMG_S) & " AND NOPHASE=" & $sFas_S & " AND B_O_NOF=" & _SQLQ($sNOF_P) & " AND B_P_NOMG=" & _SQLQ($sNOMG_P) & " AND B_P_NOPHASE=" & $sFas_P & ")" & @CRLF
        $s &= "    INSERT INTO B_PROF (NOF, NOMG, NOPHASE, B_O_NOF, B_P_NOMG, B_P_NOPHASE, PROF_TYPEPREC)" & @CRLF
        $s &= "    VALUES (" & _SQLQ($sNOF_S) & ", " & _SQLQ($sNOMG_S) & ", " & $sFas_S & ", " & _SQLQ($sNOF_P) & ", " & _SQLQ($sNOMG_P) & ", " & $sFas_P & ", " & _SQLQ($sTipo) & ");" & @CRLF
    Next
    Return $s
EndFunc

; Extract the first table name referenced by an INSERT/UPDATE/DELETE/IF NOT EXISTS statement.
; Returns "?" if it cannot be determined.
Func _ExtractSQLTableName($sStmt)
    Local $s = StringStripWS($sStmt, 3)
    ; Try INSERT INTO <table>
    Local $aMatch = StringRegExp($s, "(?i)INSERT\s+INTO\s+([A-Za-z_][A-Za-z0-9_]*)", 1)
    If IsArray($aMatch) Then Return $aMatch[0]
    ; Try UPDATE <table>
    $aMatch = StringRegExp($s, "(?i)UPDATE\s+([A-Za-z_][A-Za-z0-9_]*)", 1)
    If IsArray($aMatch) Then Return $aMatch[0]
    ; Try DELETE FROM <table>
    $aMatch = StringRegExp($s, "(?i)DELETE\s+FROM\s+([A-Za-z_][A-Za-z0-9_]*)", 1)
    If IsArray($aMatch) Then Return $aMatch[0]
    ; Try IF NOT EXISTS (SELECT ... FROM <table>
    $aMatch = StringRegExp($s, "(?i)FROM\s+([A-Za-z_][A-Za-z0-9_]*)", 1)
    If IsArray($aMatch) Then Return $aMatch[0]
    Return "?"
EndFunc

; Interpret common SQL Server error messages and return a friendly hint string, or "".
Func _InterpretSQLError($sErr)
    Local $sU = StringUpper($sErr)
    If StringInStr($sU, "E_GAMME_NOME") Then
        Return "BOM insertion failed because Ortems could not validate the parent item/version/routing bridge. This usually means B_VER_ART or E_GAMME_NOME does not contain the parent item + version + routing used by the BOM row."
    EndIf
    If StringInStr($sU, "VIOLATION OF PRIMARY KEY") Or StringInStr($sU, "DUPLICATE KEY") Then
        Return "Primary key already exists in the target table. The row was previously loaded - this usually happens when you re-run the SQL on a database that already contains the data. The generator adds IF NOT EXISTS guards for known tables; if the error persists, the composite key may not match the guard columns."
    EndIf
    If StringInStr($sU, "FK2_B_PART") Or (StringInStr($sU, "B_PHAS") And StringInStr($sU, "B_VER_ART")) Then
        Return "B_VER_ART / B_PART is pointing an item version to a routing/phase combination that does not exist in B_PHAS. The import now cleans B_PART/B_PARTC/I_PART/I_PARTC links for the impacted item version before changing B_VER_ART.NOMG. Also make sure Clear existing data before insert is enabled for full dataset refreshes."
    EndIf
    If StringInStr($sU, "FOREIGN KEY") Or StringInStr($sU, "REFERENCE CONSTRAINT") Then
        Return "A foreign key reference is missing. Check that the parent row (item/routing/calendar) exists before inserting the child row. Run FK Pre-validation for more details."
    EndIf
    If StringInStr($sU, "CANNOT INSERT THE VALUE NULL") Or StringInStr($sU, "NULL INTO COLUMN") Then
        Return "A NOT NULL column received an empty value. Check the source row in the corresponding tab."
    EndIf
    If StringInStr($sU, "STRING OR BINARY DATA WOULD BE TRUNCATED") Then
        Return "A value is longer than the target column allows. Shorten the offending field in the source tab."
    EndIf
    If StringInStr($sU, "CONVERSION FAILED") Or StringInStr($sU, "ARITHMETIC OVERFLOW") Then
        Return "A value could not be converted to the target column type. Check numeric/date fields in the source row."
    EndIf
    If StringInStr($sU, "INVALID COLUMN NAME") Then
        Return "The SQL references a column that does not exist in this Ortems schema. Regenerate the SQL after testing the DB connection so the Toolbox can detect the version-specific column names."
    EndIf
    If StringInStr($sU, "INVALID OBJECT NAME") Then
        Return "The target table does not exist in this database. Check that the Ortems schema is installed."
    EndIf
    If StringInStr($sU, "PERMISSION") Or StringInStr($sU, "DENIED") Then
        Return "The connected user does not have permission to modify this table."
    EndIf
    Return ""
EndFunc

Func _SQLStripLineComments($sStmt)
    Local $sNormalized = StringReplace($sStmt, @CRLF, @LF)
    $sNormalized = StringReplace($sNormalized, @CR, @LF)
    Local $aLines = StringSplit($sNormalized, @LF, 1)
    Local $sOut = ""
    For $i = 1 To $aLines[0]
        Local $sLine = $aLines[$i]
        Local $sTrim = StringStripWS($sLine, 3)
        If $sTrim = "" Then ContinueLoop
        If StringLeft($sTrim, 2) = "--" Then ContinueLoop
        $sOut &= $sLine & @CRLF
    Next
    Return StringStripWS($sOut, 3)
EndFunc

Func _SQLSplitAppend(ByRef $aOut, ByRef $nCount, $sStmt)
    Local $sClean = _SQLStripLineComments($sStmt)
    If $sClean = "" Then Return
    $nCount += 1
    ReDim $aOut[$nCount + 1]
    $aOut[$nCount] = $sClean
EndFunc

Func _SQLSplitStatements($sSQL)
    Local $aOut[1]
    Local $nCount = 0
    Local $sCur = ""
    Local $bInQuote = False
    Local $bInLineComment = False
    Local $nLen = StringLen($sSQL)

    For $i = 1 To $nLen
        Local $ch = StringMid($sSQL, $i, 1)

        ; Keep line comments attached to the current chunk, but ignore any semicolon inside them.
        ; The append step removes comment-only lines before execution.
        If $bInLineComment Then
            $sCur &= $ch
            If $ch = @CR Or $ch = @LF Then $bInLineComment = False
            ContinueLoop
        EndIf

        If Not $bInQuote Then
            If $ch = "-" And $i < $nLen And StringMid($sSQL, $i + 1, 1) = "-" Then
                $sCur &= "--"
                $i += 1
                $bInLineComment = True
                ContinueLoop
            EndIf
        EndIf

        If $ch = "'" Then
            $sCur &= $ch
            If $bInQuote And $i < $nLen And StringMid($sSQL, $i + 1, 1) = "'" Then
                $i += 1
                $sCur &= "'"
            Else
                $bInQuote = Not $bInQuote
            EndIf
            ContinueLoop
        EndIf

        If $ch = ";" And Not $bInQuote Then
            _SQLSplitAppend($aOut, $nCount, $sCur)
            $sCur = ""
        Else
            $sCur &= $ch
        EndIf
    Next

    _SQLSplitAppend($aOut, $nCount, $sCur)
    $aOut[0] = $nCount
    Return $aOut
EndFunc

Func _ExecuteSQL()
    If Not $g_bConnected Then
        MsgBox(262144+48, "Warning", "You are not connected to the database." & @CRLF & "Go to the '1. Database' tab and test the connection first.",0,$g_hMain)
        Return
    EndIf

    Local $sSQL = GUICtrlRead($g_hLog)
    If $sSQL = "" Or StringInStr($sSQL, "INSERT") = 0 Then
        MsgBox(262144+48, "Warning", "No SQL has been generated yet. Click 'GENERATE SQL' first.",0,$g_hMain)
        Return
    EndIf

    Local $nRet = MsgBox(262144+4 + 48, "Confirm execution", _
        "You are about to run SQL on the database:" & @CRLF & @CRLF & _
        "Database: " & $g_sDatabase & @CRLF & "Server: " & $g_sServer & @CRLF & @CRLF & _
        "This action can DELETE and RECREATE existing data!" & @CRLF & @CRLF & _
        "Do you want to continue?",0,$g_hMain)

    If $nRet <> 6 Then Return

    ; Pre-flight integrity validation before touching the database
    Local $nFKIssues = _IntegrityCheck(False, True, "run the SQL on the database")
    If $nFKIssues = -1 Then Return

    Local $oConn = ObjCreate("ADODB.Connection")
    If Not IsObj($oConn) Then
        _Log("ERROR: Could not create ADODB.Connection object.")
        Return
    EndIf

    $g_sLastComError = ""
    $oConn.Open($g_sConnStr)
    If $g_sLastComError <> "" Then
        _Log("ERROR opening connection: " & $g_sLastComError)
        MsgBox(262144+16, "Connection error", "Could not open connection:" & @CRLF & $g_sLastComError,0,$g_hMain)
        Return
    EndIf
    If $oConn.State <> 1 Then
        _Log("ERROR: Connection state is not Open.")
        Return
    EndIf

    Local $bUseTrans = (GUICtrlRead($g_chkTransaction) = $GUI_CHECKED)
    If StringInStr($sSQL, "GLOBAL TABLE CLEANUP") > 0 And Not $bUseTrans Then
        _Log("Transaction forced ON for global clean-import mode. This prevents leaving the database half-loaded or constraints disabled if an error occurs.")
        $bUseTrans = True
    EndIf
    Local $nOK = 0, $nErr = 0
    Local $bTransStarted = False

    If $bUseTrans Then
        $g_sLastComError = ""
        $oConn.BeginTrans()
        If $g_sLastComError <> "" Then
            _Log("WARNING: BeginTrans failed (" & $g_sLastComError & ") - running without transaction.")
            $bUseTrans = False
        Else
            $bTransStarted = True
        EndIf
    EndIf

    Local $aStmts = _SQLSplitStatements($sSQL)

    For $i = 1 To $aStmts[0]
        Local $sStmt = StringStripWS($aStmts[$i], 3)
        If $sStmt = "" Then ContinueLoop

        ; Remove SQL comment-only lines before deciding whether this chunk is executable.
        ; This avoids skipping a valid statement just because it starts after a section header comment.
        Local $sNoComments = StringRegExpReplace($sStmt, "(?m)^\s*--[^\r\n]*", "")
        $sStmt = StringStripWS($sNoComments, 3)
        If $sStmt = "" Then ContinueLoop

        If StringUpper($sStmt) = "BEGIN TRANSACTION" Then ContinueLoop
        If StringUpper($sStmt) = "COMMIT TRANSACTION" Then ContinueLoop
        If StringUpper($sStmt) = "ROLLBACK TRANSACTION" Then ContinueLoop

        Local $sTblExec = _ExtractSQLTableName($sStmt)
        _LogVerboseSQL("Run SQL stmt " & $i & " on table [" & $sTblExec & "]", $sStmt)
        $g_sLastComError = ""
        $oConn.Execute($sStmt)

        If $g_sLastComError <> "" Then
            $nErr += 1
            Local $sTbl = _ExtractSQLTableName($sStmt)
            Local $sHint = _InterpretSQLError($g_sLastComError)
            _Log("ERROR (stmt " & $i & ") on table [" & $sTbl & "]: " & $g_sLastComError)
            If $sHint <> "" Then _Log("  HINT: " & $sHint)
            ; Show up to 400 chars of the offending statement (multiple lines preserved)
            Local $sStmtClean = StringStripWS($sStmt, 3)
            If StringLen($sStmtClean) > 400 Then $sStmtClean = StringLeft($sStmtClean, 400) & " ..."
            _Log("  >> " & StringReplace($sStmtClean, @CRLF, " | "))
            If $bTransStarted Then
                $g_sLastComError = ""
                $oConn.RollbackTrans()
                $oConn.Close()
                _Log("Transaction ROLLED BACK due to error.")
                Local $sMsg = "A SQL error occurred on table [" & $sTbl & "] - the transaction was rolled back." & @CRLF & @CRLF
                If $sHint <> "" Then $sMsg &= $sHint & @CRLF & @CRLF
                $sMsg &= "Check the execution log and the session log file for details."
                MsgBox(262144+16, "Execution error", $sMsg,0,$g_hMain)
                Return
            EndIf
        Else
            $nOK += 1
            _LogVerbose("SQL stmt " & $i & " OK on table [" & $sTblExec & "]")
        EndIf
    Next

    If $bTransStarted Then
        $g_sLastComError = ""
        $oConn.CommitTrans()
        If $g_sLastComError <> "" Then
            _Log("ERROR committing transaction: " & $g_sLastComError)
            $oConn.RollbackTrans()
        Else
            _Log("Transaction committed successfully.")
        EndIf
    EndIf

    $oConn.Close()
    _Log("Execution finished: " & $nOK & " statements OK, " & $nErr & " errors.")

    If $nErr = 0 Then
        MsgBox(262144+64, "Success", "All SQL executed successfully!" & @CRLF & @CRLF & $nOK & " statements executed.",0,$g_hMain)
    Else
        MsgBox(262144+48, "Completed with errors", $nOK & " statements OK, " & $nErr & " errors." & @CRLF & "Check the execution log and the session log file for details.",0,$g_hMain)
    EndIf
EndFunc

Func _GetClearSQL()
    ; Definitive global clean-import cleanup.
    ;
    ; Previous versions used a static Ortems table list or one dynamic DELETE batch.
    ; That made troubleshooting difficult: if one table was not cleared, the later
    ; E_OF upsert could collide with old B_BT rows and look like an import-order issue.
    ;
    ; This version is intentionally strict:
    ;   1. Disable constraints on every user table with NOCHECK CONSTRAINT ALL.
    ;   2. Verify no FK remains enabled.
    ;   3. Disable triggers on every user table.
    ;   4. Delete every non-system user table using a cursor and per-table dynamic SQL.
    ;   5. Verify all non-system user tables are empty before the import continues.
    ;
    ; If cleanup is incomplete, it raises a SQL error immediately. We do not continue
    ; with stale Ortems data because that creates misleading FK errors later.

    Local $sDisableFK = _
        "DECLARE @schema sysname, @table sysname, @sql NVARCHAR(MAX)" & @CRLF & _
        "DECLARE c CURSOR LOCAL FAST_FORWARD FOR" & @CRLF & _
        "SELECT SCHEMA_NAME(schema_id), name" & @CRLF & _
        "FROM sys.tables" & @CRLF & _
        "WHERE is_ms_shipped = 0 AND name <> 'sysdiagrams'" & @CRLF & _
        "OPEN c" & @CRLF & _
        "FETCH NEXT FROM c INTO @schema, @table" & @CRLF & _
        "WHILE @@FETCH_STATUS = 0" & @CRLF & _
        "BEGIN" & @CRLF & _
        "    SET @sql = N'ALTER TABLE ' + QUOTENAME(@schema) + N'.' + QUOTENAME(@table) + N' NOCHECK CONSTRAINT ALL'" & @CRLF & _
        "    EXEC sp_executesql @sql" & @CRLF & _
        "    FETCH NEXT FROM c INTO @schema, @table" & @CRLF & _
        "END" & @CRLF & _
        "CLOSE c" & @CRLF & _
        "DEALLOCATE c" & @CRLF & _
        "IF EXISTS (SELECT 1 FROM sys.foreign_keys WHERE is_disabled = 0)" & @CRLF & _
        "BEGIN" & @CRLF & _
        "    RAISERROR('Global cleanup failed: at least one FK constraint remains enabled after NOCHECK CONSTRAINT ALL.', 16, 1)" & @CRLF & _
        "END"

    Local $sDisableTrig = _
        "DECLARE @schema sysname, @table sysname, @sql NVARCHAR(MAX)" & @CRLF & _
        "DECLARE c CURSOR LOCAL FAST_FORWARD FOR" & @CRLF & _
        "SELECT SCHEMA_NAME(schema_id), name" & @CRLF & _
        "FROM sys.tables" & @CRLF & _
        "WHERE is_ms_shipped = 0 AND name <> 'sysdiagrams'" & @CRLF & _
        "OPEN c" & @CRLF & _
        "FETCH NEXT FROM c INTO @schema, @table" & @CRLF & _
        "WHILE @@FETCH_STATUS = 0" & @CRLF & _
        "BEGIN" & @CRLF & _
        "    SET @sql = N'DISABLE TRIGGER ALL ON ' + QUOTENAME(@schema) + N'.' + QUOTENAME(@table)" & @CRLF & _
        "    EXEC sp_executesql @sql" & @CRLF & _
        "    FETCH NEXT FROM c INTO @schema, @table" & @CRLF & _
        "END" & @CRLF & _
        "CLOSE c" & @CRLF & _
        "DEALLOCATE c"

    Local $sDeleteAllTables = _
        "DECLARE @schema sysname, @table sysname, @sql NVARCHAR(MAX), @msg NVARCHAR(4000)" & @CRLF & _
        "DECLARE c CURSOR LOCAL FAST_FORWARD FOR" & @CRLF & _
        "SELECT SCHEMA_NAME(schema_id), name" & @CRLF & _
        "FROM sys.tables" & @CRLF & _
        "WHERE is_ms_shipped = 0 AND name <> 'sysdiagrams'" & @CRLF & _
        "ORDER BY name DESC" & @CRLF & _
        "OPEN c" & @CRLF & _
        "FETCH NEXT FROM c INTO @schema, @table" & @CRLF & _
        "WHILE @@FETCH_STATUS = 0" & @CRLF & _
        "BEGIN" & @CRLF & _
        "    BEGIN TRY" & @CRLF & _
        "        SET @sql = N'DELETE FROM ' + QUOTENAME(@schema) + N'.' + QUOTENAME(@table)" & @CRLF & _
        "        EXEC sp_executesql @sql" & @CRLF & _
        "    END TRY" & @CRLF & _
        "    BEGIN CATCH" & @CRLF & _
        "        SET @msg = N'Global cleanup failed while deleting [' + @schema + N'].[' + @table + N']: ' + ERROR_MESSAGE()" & @CRLF & _
        "        RAISERROR(@msg, 16, 1)" & @CRLF & _
        "    END CATCH" & @CRLF & _
        "    FETCH NEXT FROM c INTO @schema, @table" & @CRLF & _
        "END" & @CRLF & _
        "CLOSE c" & @CRLF & _
        "DEALLOCATE c"

    Local $sValidateEmpty = _
        "DECLARE @schema sysname, @table sysname, @sql NVARCHAR(MAX), @cnt INT, @bad NVARCHAR(MAX)" & @CRLF & _
        "SET @bad = N''" & @CRLF & _
        "DECLARE c CURSOR LOCAL FAST_FORWARD FOR" & @CRLF & _
        "SELECT SCHEMA_NAME(schema_id), name" & @CRLF & _
        "FROM sys.tables" & @CRLF & _
        "WHERE is_ms_shipped = 0 AND name <> 'sysdiagrams'" & @CRLF & _
        "ORDER BY name" & @CRLF & _
        "OPEN c" & @CRLF & _
        "FETCH NEXT FROM c INTO @schema, @table" & @CRLF & _
        "WHILE @@FETCH_STATUS = 0" & @CRLF & _
        "BEGIN" & @CRLF & _
        "    SET @cnt = 0" & @CRLF & _
        "    SET @sql = N'SELECT @cnt_out = COUNT(1) FROM ' + QUOTENAME(@schema) + N'.' + QUOTENAME(@table)" & @CRLF & _
        "    EXEC sp_executesql @sql, N'@cnt_out INT OUTPUT', @cnt_out = @cnt OUTPUT" & @CRLF & _
        "    IF @cnt <> 0" & @CRLF & _
        "        SET @bad = @bad + QUOTENAME(@schema) + N'.' + QUOTENAME(@table) + N'=' + CONVERT(NVARCHAR(30), @cnt) + N'; '" & @CRLF & _
        "    FETCH NEXT FROM c INTO @schema, @table" & @CRLF & _
        "END" & @CRLF & _
        "CLOSE c" & @CRLF & _
        "DEALLOCATE c" & @CRLF & _
        "IF LEN(@bad) > 0" & @CRLF & _
        "BEGIN" & @CRLF & _
        "    RAISERROR('Global cleanup incomplete. Tables still contain rows: %s', 16, 1, @bad)" & @CRLF & _
        "END"

    Local $aResult[5]
    $aResult[0] = 4
    $aResult[1] = $sDisableFK
    $aResult[2] = $sDisableTrig
    $aResult[3] = $sDeleteAllTables
    $aResult[4] = $sValidateEmpty

    Return $aResult
EndFunc

Func _GetPostImportEnableSQL()
    ; Re-enable triggers and validate constraints table-by-table.
    ; The cursor approach makes the failing table/constraint easier to diagnose than
    ; one large dynamic batch, and avoids silent partial execution.

    Local $sEnableTrig = _
        "DECLARE @schema sysname, @table sysname, @sql NVARCHAR(MAX)" & @CRLF & _
        "DECLARE c CURSOR LOCAL FAST_FORWARD FOR" & @CRLF & _
        "SELECT SCHEMA_NAME(schema_id), name" & @CRLF & _
        "FROM sys.tables" & @CRLF & _
        "WHERE is_ms_shipped = 0 AND name <> 'sysdiagrams'" & @CRLF & _
        "OPEN c" & @CRLF & _
        "FETCH NEXT FROM c INTO @schema, @table" & @CRLF & _
        "WHILE @@FETCH_STATUS = 0" & @CRLF & _
        "BEGIN" & @CRLF & _
        "    SET @sql = N'ENABLE TRIGGER ALL ON ' + QUOTENAME(@schema) + N'.' + QUOTENAME(@table)" & @CRLF & _
        "    EXEC sp_executesql @sql" & @CRLF & _
        "    FETCH NEXT FROM c INTO @schema, @table" & @CRLF & _
        "END" & @CRLF & _
        "CLOSE c" & @CRLF & _
        "DEALLOCATE c"

    Local $sEnableFK = _
        "DECLARE @schema sysname, @table sysname, @sql NVARCHAR(MAX), @msg NVARCHAR(4000)" & @CRLF & _
        "DECLARE c CURSOR LOCAL FAST_FORWARD FOR" & @CRLF & _
        "SELECT SCHEMA_NAME(schema_id), name" & @CRLF & _
        "FROM sys.tables" & @CRLF & _
        "WHERE is_ms_shipped = 0 AND name <> 'sysdiagrams'" & @CRLF & _
        "ORDER BY name" & @CRLF & _
        "OPEN c" & @CRLF & _
        "FETCH NEXT FROM c INTO @schema, @table" & @CRLF & _
        "WHILE @@FETCH_STATUS = 0" & @CRLF & _
        "BEGIN" & @CRLF & _
        "    BEGIN TRY" & @CRLF & _
        "        SET @sql = N'ALTER TABLE ' + QUOTENAME(@schema) + N'.' + QUOTENAME(@table) + N' WITH CHECK CHECK CONSTRAINT ALL'" & @CRLF & _
        "        EXEC sp_executesql @sql" & @CRLF & _
        "    END TRY" & @CRLF & _
        "    BEGIN CATCH" & @CRLF & _
        "        SET @msg = N'Post-import FK/CHECK validation failed on [' + @schema + N'].[' + @table + N']: ' + ERROR_MESSAGE()" & @CRLF & _
        "        RAISERROR(@msg, 16, 1)" & @CRLF & _
        "    END CATCH" & @CRLF & _
        "    FETCH NEXT FROM c INTO @schema, @table" & @CRLF & _
        "END" & @CRLF & _
        "CLOSE c" & @CRLF & _
        "DEALLOCATE c" & @CRLF & _
        "IF EXISTS (SELECT 1 FROM sys.foreign_keys WHERE is_disabled = 1 OR is_not_trusted = 1)" & @CRLF & _
        "BEGIN" & @CRLF & _
        "    RAISERROR('Post-import FK validation incomplete: at least one FK is still disabled or not trusted.', 16, 1)" & @CRLF & _
        "END"

    Local $aResult[3]
    $aResult[0] = 2
    $aResult[1] = $sEnableTrig
    $aResult[2] = $sEnableFK

    Return $aResult
EndFunc

Func _GetPostImportEnableSQLScript()
    Local $aStmts = _GetPostImportEnableSQL()
    Local $s = ""
    For $i = 1 To $aStmts[0]
        $s &= $aStmts[$i]
        If StringRight(StringStripWS($aStmts[$i], 3), 1) <> ";" Then $s &= ";"
        $s &= @CRLF
    Next
    Return $s
EndFunc

Func _GetClearSQLScript()
    Local $aStmts = _GetClearSQL()
    Local $s = ""
    For $i = 1 To $aStmts[0]
        $s &= $aStmts[$i]
        If StringRight(StringStripWS($aStmts[$i], 3), 1) <> ";" Then $s &= ";"
        $s &= @CRLF
    Next
    Return $s
EndFunc

; Helper: turns a string array into a SQL IN-list  e.g. 'B_GAMM','B_OF',...
;=============================================================================
; FK PRE-VALIDATION: check referential integrity before attempting import
; Queries data in the ListViews against what already exists in the DB.
; Logs each violation to the execution log and session log file and returns the count of issues found.
;=============================================================================
;=============================================================================
; FK PRE-VALIDATION: cross-check references WITHIN the import data
; The DB was just cleared, so checking against the DB would always warn.
; Instead we verify that IDs referenced in one tab exist in another tab.
;=============================================================================
Func _IntegrityCheck($p1 = "", $p2 = "", $p3 = "")
    Local $bShowSuccess = True
    Local $bAskToContinue = False
    Local $sAction = ""

    If @NumParams >= 1 Then $bShowSuccess = $p1
    If @NumParams >= 2 Then $bAskToContinue = $p2
    If @NumParams >= 3 Then $sAction = $p3

    Local $nIssues = _ValidateFKBeforeImport()
    If $nIssues = 0 Then
        If $bShowSuccess Then
            MsgBox(262144+64, "Integrity Check", _
                "No cross-tab reference issue was found." & @CRLF & @CRLF & _
                "The data is consistent and ready for SQL generation.",0,$g_hMain)
        EndIf
        Return 0
    EndIf

    Local $sMsg = _IntegrityBuildMessage($nIssues)
    Local $nFixChoice = MsgBox(262144+3 + 48, "Integrity Check", _
        $sMsg & @CRLF & @CRLF & _
        "Do you want to try automatic correction for simple references now?" & @CRLF & @CRLF & _
        "Yes = try auto-fix" & @CRLF & _
        "No = keep current values" & @CRLF & _
        "Cancel = stop",0,$g_hMain)

    If $nFixChoice = 2 Then Return -1

    If $nFixChoice = 6 Then
        Local $nFixed = _AutoFixSimpleReferences()
        If $nFixed > 0 Then
            Local $nRemaining = _ValidateFKBeforeImport()
            If $nRemaining = 0 Then
                MsgBox(262144+64, "Integrity Check", _
                    "Automatic correction updated " & $nFixed & " field(s)." & @CRLF & @CRLF & _
                    "All cross-tab references are now consistent.",0,$g_hMain)
                Return 0
            EndIf

            $nIssues = $nRemaining
            $sMsg = "Automatic correction updated " & $nFixed & " field(s), but some issues still remain." & @CRLF & @CRLF & _
                _IntegrityBuildMessage($nIssues)
        Else
            $sMsg = "No simple reference could be corrected automatically." & @CRLF & @CRLF & _
                _IntegrityBuildMessage($nIssues)
        EndIf
    EndIf

    If $bAskToContinue Then
        Local $sQuestion = $sMsg
        If $sAction <> "" Then $sQuestion &= @CRLF & @CRLF & "Do you still want to " & $sAction & "?"
        Local $nRet = MsgBox(262144+4 + 48, "Integrity Check", $sQuestion,0,$g_hMain)
        If $nRet <> 6 Then Return -1
    ElseIf $nFixChoice = 6 Then
        MsgBox(262144+48, "Integrity Check", $sMsg,0,$g_hMain)
    EndIf

    Return $nIssues
EndFunc

Func _ValidateFKBeforeImport()
    Local $nIssues = 0
    $g_sIntegrityReport = ""
    _Log("=== Integrity Check started (cross-tab reference validation) ===")

    ; Reference sets
    Local $setCalIDs         = _LVColSet($g_hLV_Cal, 0)
    Local $setCapCalIDs      = _LVColSet($g_hLV_Cap, 0)
    Local $setAnyCalIDs      = $setCalIDs & StringTrimLeft($setCapCalIDs, 1)
    Local $setCTIDs          = _LVColSet($g_hLV_Mach, 2)
    Local $setMachIDs        = _LVColSet($g_hLV_Mach, 7)
    Local $setCTMach         = _LVCompositeSet($g_hLV_Mach, "2,7")
    Local $setOpeIDs         = _LVColSet($g_hLV_Ops, 0)
    Local $setOpeCTMach      = _LVCompositeSet($g_hLV_Ops, "0,2,3")
    Local $setRoutIDs        = _LVColSet($g_hLV_Rout, 0)
    Local $setRoutPhase      = _LVCompositeSet($g_hLV_Rout, "0,2")
    Local $setMatIDs         = _LVColSet($g_hLV_Mat, 0)
    Local $setMatVer         = _LVCompositeSet($g_hLV_Mat, "0,3")
    Local $setMatRoutVer     = _LVCompositeSet($g_hLV_Mat, "0,4,3")
    Local $setMatVerRout     = _LVCompositeSet($g_hLV_Mat, "0,3,4")
    Local $setWOIDs          = _LVColSet($g_hLV_WO, 0)
    Local $setWORout         = _LVCompositeSet($g_hLV_WO, "0,2")

    ; Data-shape checks based on the original Excel Toolbox rules
    $nIssues += _ValidateRequiredFields()
    $nIssues += _ValidateDuplicateKeys()

    ; Cross-tab checks
    $nIssues += _CrossCheckMachineCalendarRef($g_hLV_Mach, 4, 10, $setCalIDs, "Machines tab / Calendar ID")
    $nIssues += _CrossCheck($g_hLV_Mach, 11, $setAnyCalIDs,  "Machines tab / Capacity calendar ID", True)

    $nIssues += _CrossCheck($g_hLV_Ops,  2, $setCTIDs,       "Operations tab / WC ID")
    $nIssues += _CrossCheckMachineRef($g_hLV_Ops,  3, $setMachIDs, "Operations tab / Machine ID")
    $nIssues += _CrossCheckCompositeMachineRef($g_hLV_Ops, "2,3", 3, $setCTMach, "Operations tab / WC ID + Machine ID")

    $nIssues += _CrossCheck($g_hLV_Rout, 3, $setOpeIDs,      "Routings tab / Operation ID")

    $nIssues += _CrossCheckItemRoutingRef($g_hLV_Mat, 2, 4, $setRoutIDs, "Items tab / Routing ID", False)

    $nIssues += _CrossCheck($g_hLV_BOM,          0,       $setMatIDs,      "BOM tab / Parent item ID")
    $nIssues += _CrossCheck($g_hLV_BOM,          2,       $setMatIDs,      "BOM tab / Component item ID")
    $nIssues += _CrossCheck($g_hLV_BOM,          3,       $setRoutIDs,     "BOM tab / Routing ID")
    $nIssues += _CrossCheckComposite($g_hLV_BOM, "0,1",   $setMatVer,      "BOM tab / Parent item + version")
    $nIssues += _CrossCheckComposite($g_hLV_BOM, "0,3,1", $setMatRoutVer,  "BOM tab / Parent item + routing + version")
    $nIssues += _CrossCheckComposite($g_hLV_BOM, "3,4",   $setRoutPhase,   "BOM tab / Routing + phase")
    $nIssues += _ValidateBOMParentRules()

    $nIssues += _CrossCheckComposite($g_hLV_WO,  "1,2,3", $setMatRoutVer,  "Work Orders tab / Item + routing + version")
    $nIssues += _ValidateWorkOrderRules()

    $nIssues += _CrossCheck($g_hLV_WOL,          0,       $setWOIDs,       "WO Links tab / Predecessor WO", True)
    $nIssues += _CrossCheck($g_hLV_WOL,          3,       $setWOIDs,       "WO Links tab / Successor WO", True)
    $nIssues += _CrossCheckComposite($g_hLV_WOL, "0,1",   $setWORout,      "WO Links tab / Predecessor WO + routing", True)
    $nIssues += _CrossCheckComposite($g_hLV_WOL, "3,4",   $setWORout,      "WO Links tab / Successor WO + routing", True)
    $nIssues += _CrossCheckComposite($g_hLV_WOL, "1,2",   $setRoutPhase,   "WO Links tab / Predecessor routing + phase", True)
    $nIssues += _CrossCheckComposite($g_hLV_WOL, "4,5",   $setRoutPhase,   "WO Links tab / Successor routing + phase", True)

    $nIssues += _CrossCheck($g_hLV_SR,           0,       $setOpeIDs,      "Secondary Resources tab / Operation ID")
    $nIssues += _CrossCheck($g_hLV_SR,           1,       $setCTIDs,       "Secondary Resources tab / WC ID")
    $nIssues += _CrossCheckMachineRef($g_hLV_SR, 2,       $setMachIDs,     "Secondary Resources tab / Machine ID")
    $nIssues += _CrossCheckCompositeMachineRef($g_hLV_SR, "0,1,2", 2, $setOpeCTMach, "Secondary Resources tab / Operation + WC + Machine")
    $nIssues += _CrossCheck($g_hLV_SR,           4,       $setAnyCalIDs,   "Secondary Resources tab / Capacity calendar ID", True)

    $nIssues += _CrossCheckInventoryItemVersion($g_hLV_Stk, 0, 2, $setMatVer, "Inventory Movements tab / Item + version")
    $nIssues += _CrossCheck($g_hLV_Stk,          1,       $setRoutIDs,     "Inventory Movements tab / Routing ID", True)
    $nIssues += _CrossCheckInventoryItemRoutingVersion($g_hLV_Stk, 0, 1, 2, $setMatRoutVer, "Inventory Movements tab / Item + routing + version")

    ; Format checks
    $nIssues += _ValidateCalendarFormats()
    $nIssues += _ValidateCapacityFormats()
    $nIssues += _ValidateWODates()
    $nIssues += _ValidateStockDates()

    If $nIssues = 0 Then
        _Log("=== Integrity Check OK - all cross-tab references are consistent ===")
    Else
        _Log("=== Integrity Check finished with " & $nIssues & " issue(s) ===")
    EndIf
    Return $nIssues
EndFunc

Func _LVColSet($hLV, $iCol)
    Local $sSet = ";"
    Local $nRows = _GUICtrlListView_GetItemCount($hLV)
    For $i = 0 To $nRows - 1
        Local $sVal = StringStripWS(_GUICtrlListView_GetItemText($hLV, $i, $iCol), 3)
        If $sVal <> "" And Not StringInStr($sSet, ";" & $sVal & ";") Then $sSet &= $sVal & ";"
    Next
    Return $sSet
EndFunc

Func _LVCompositeSet($hLV, $sCols)
    Local $sSet = ";"
    Local $nRows = _GUICtrlListView_GetItemCount($hLV)
    For $i = 0 To $nRows - 1
        Local $sKey = _LVComposeKey($hLV, $i, $sCols)
        If $sKey <> "" And Not StringInStr($sSet, ";" & $sKey & ";") Then $sSet &= $sKey & ";"
    Next
    Return $sSet
EndFunc

Func _LVComposeKey($hLV, $iRow, $sCols)
    Local $aCols = StringSplit($sCols, ",", 1)
    Local $sKey = ""
    For $j = 1 To $aCols[0]
        Local $iCol = Number($aCols[$j])
        Local $sVal = StringStripWS(_GUICtrlListView_GetItemText($hLV, $iRow, $iCol), 3)
        If $sVal = "" Then Return ""
        If $j > 1 Then $sKey &= "|"
        $sKey &= $sVal
    Next
    Return $sKey
EndFunc

Func _IsStandByMachine($sVal)
    Local $sNorm = StringUpper(StringStripWS($sVal, 3))
    $sNorm = StringReplace($sNorm, "-", "")
    $sNorm = StringReplace($sNorm, "_", "")
    $sNorm = StringReplace($sNorm, " ", "")
    Return ($sNorm = "STANDBY")
EndFunc

Func _IsRawMaterialType($sVal)
    Local $sNorm = StringUpper(StringStripWS($sVal, 3))
    Return ($sNorm = "MP")
EndFunc

Func _IsMachineCalendarOptional($sCTType)
    Local $sNorm = StringStripWS($sCTType, 3)
    Return ($sNorm = "4")
EndFunc

Func _CrossCheckMachineCalendarRef($hLV, $iCTTypeCol, $iCalCol, $sSet, $sDesc, $bAllowEmpty = False)
    Local $nIssues = 0
    Local $nRows = _GUICtrlListView_GetItemCount($hLV)

    For $i = 0 To $nRows - 1
        Local $sCTType = StringStripWS(_GUICtrlListView_GetItemText($hLV, $i, $iCTTypeCol), 3)
        If _IsMachineCalendarOptional($sCTType) Then ContinueLoop

        Local $sVal = StringStripWS(_GUICtrlListView_GetItemText($hLV, $i, $iCalCol), 3)
        If $sVal = "" Then
            If $bAllowEmpty Then ContinueLoop
            $nIssues += 1
            _IntegrityIssue($sDesc & " - row " & ($i + 1) & ": empty value.")
            ContinueLoop
        EndIf

        If Not StringInStr($sSet, ";" & $sVal & ";") Then
            $nIssues += 1
            _IntegrityIssue($sDesc & " - row " & ($i + 1) & ": '" & $sVal & "' does not exist in the referenced tab.")
        EndIf
    Next
    Return $nIssues
EndFunc

Func _CrossCheckItemRoutingRef($hLV, $iTypeCol, $iRoutingCol, $sSet, $sDesc, $bAllowEmpty = False)
    Local $nIssues = 0
    Local $nRows = _GUICtrlListView_GetItemCount($hLV)

    For $i = 0 To $nRows - 1
        Local $sType = StringStripWS(_GUICtrlListView_GetItemText($hLV, $i, $iTypeCol), 3)
        If _IsRawMaterialType($sType) Then ContinueLoop

        Local $sVal = StringStripWS(_GUICtrlListView_GetItemText($hLV, $i, $iRoutingCol), 3)
        If $sVal = "" Then
            If $bAllowEmpty Then ContinueLoop
            $nIssues += 1
            _IntegrityIssue($sDesc & " - row " & ($i + 1) & ": empty value.")
            ContinueLoop
        EndIf

        If Not StringInStr($sSet, ";" & $sVal & ";") Then
            $nIssues += 1
            _IntegrityIssue($sDesc & " - row " & ($i + 1) & ": '" & $sVal & "' does not exist in the referenced tab.")
        EndIf
    Next
    Return $nIssues
EndFunc

Func _CrossCheckMachineRef($hLV, $iCol, $sSet, $sDesc, $bAllowEmpty = False)
    Local $nIssues = 0
    Local $nRows = _GUICtrlListView_GetItemCount($hLV)
    For $i = 0 To $nRows - 1
        Local $sVal = StringStripWS(_GUICtrlListView_GetItemText($hLV, $i, $iCol), 3)
        If _IsStandByMachine($sVal) Then ContinueLoop
        If $sVal = "" Then
            If $bAllowEmpty Then ContinueLoop
            $nIssues += 1
            _IntegrityIssue($sDesc & " - row " & ($i + 1) & ": empty value.")
            ContinueLoop
        EndIf
        If Not StringInStr($sSet, ";" & $sVal & ";") Then
            $nIssues += 1
            _IntegrityIssue($sDesc & " - row " & ($i + 1) & ": '" & $sVal & "' does not exist in the referenced tab.")
        EndIf
    Next
    Return $nIssues
EndFunc

Func _CrossCheckCompositeMachineRef($hLV, $sCols, $iMachineCol, $sSet, $sDesc, $bAllowEmpty = False)
    Local $nIssues = 0
    Local $nRows = _GUICtrlListView_GetItemCount($hLV)

    For $i = 0 To $nRows - 1
        Local $sMachine = StringStripWS(_GUICtrlListView_GetItemText($hLV, $i, $iMachineCol), 3)
        If _IsStandByMachine($sMachine) Then ContinueLoop

        Local $sKey = _LVComposeKey($hLV, $i, $sCols)
        If $sKey = "" Then
            If $bAllowEmpty Then ContinueLoop
            $nIssues += 1
            _IntegrityIssue($sDesc & " - row " & ($i + 1) & ": one or more key fields are empty.")
            ContinueLoop
        EndIf
        If Not StringInStr($sSet, ";" & $sKey & ";") Then
            $nIssues += 1
            _IntegrityIssue($sDesc & " - row " & ($i + 1) & ": '" & $sKey & "' does not match any valid combination in the referenced tab.")
        EndIf
    Next
    Return $nIssues
EndFunc

Func _CrossCheck($hLV, $iCol, $sSet, $sDesc, $bAllowEmpty = False)
    Local $nIssues = 0
    Local $nRows = _GUICtrlListView_GetItemCount($hLV)
    For $i = 0 To $nRows - 1
        Local $sVal = StringStripWS(_GUICtrlListView_GetItemText($hLV, $i, $iCol), 3)
        If $sVal = "" Then
            If $bAllowEmpty Then ContinueLoop
            $nIssues += 1
            _IntegrityIssue($sDesc & " - row " & ($i + 1) & ": empty value.")
            ContinueLoop
        EndIf
        If Not StringInStr($sSet, ";" & $sVal & ";") Then
            $nIssues += 1
            _IntegrityIssue($sDesc & " - row " & ($i + 1) & ": '" & $sVal & "' does not exist in the referenced tab.")
        EndIf
    Next
    Return $nIssues
EndFunc

Func _CrossCheckComposite($hLV, $sCols, $sSet, $sDesc, $bAllowEmpty = False)
    Local $nIssues = 0
    Local $nRows = _GUICtrlListView_GetItemCount($hLV)
    For $i = 0 To $nRows - 1
        Local $sKey = _LVComposeKey($hLV, $i, $sCols)
        If $sKey = "" Then
            If $bAllowEmpty Then ContinueLoop
            $nIssues += 1
            _IntegrityIssue($sDesc & " - row " & ($i + 1) & ": one or more key fields are empty.")
            ContinueLoop
        EndIf
        If Not StringInStr($sSet, ";" & $sKey & ";") Then
            $nIssues += 1
            _IntegrityIssue($sDesc & " - row " & ($i + 1) & ": '" & $sKey & "' does not match any valid combination in the referenced tab.")
        EndIf
    Next
    Return $nIssues
EndFunc

Func _CrossCheckInventoryItemVersion($hLV, $iItemCol, $iVerCol, $sSet, $sDesc)
    Local $nIssues = 0
    Local $nRows = _GUICtrlListView_GetItemCount($hLV)
    For $i = 0 To $nRows - 1
        Local $sItem = StringStripWS(_GUICtrlListView_GetItemText($hLV, $i, $iItemCol), 3)
        Local $sVer  = StringStripWS(_GUICtrlListView_GetItemText($hLV, $i, $iVerCol), 3)
        If $sItem = "" Then ContinueLoop

        If _ItemIsRawMaterial($sItem) Then
            If $sVer <> "" Then
                $nIssues += 1
                _IntegrityIssue($sDesc & " - row " & ($i + 1) & ": MP raw material items cannot have version '" & $sVer & "'. Leave Version blank.")
            EndIf
            ContinueLoop
        EndIf

        If $sVer = "" Then
            $nIssues += 1
            _IntegrityIssue($sDesc & " - row " & ($i + 1) & ": version is required for manufactured items, but not for MP raw materials.")
            ContinueLoop
        EndIf

        Local $sKey = $sItem & "|" & $sVer
        If Not StringInStr($sSet, ";" & $sKey & ";") Then
            $nIssues += 1
            _IntegrityIssue($sDesc & " - row " & ($i + 1) & ": '" & $sKey & "' does not match any manufactured item/version in the Items tab.")
        EndIf
    Next
    Return $nIssues
EndFunc

Func _CrossCheckInventoryItemRoutingVersion($hLV, $iItemCol, $iRoutingCol, $iVerCol, $sSet, $sDesc)
    Local $nIssues = 0
    Local $nRows = _GUICtrlListView_GetItemCount($hLV)
    For $i = 0 To $nRows - 1
        Local $sItem = StringStripWS(_GUICtrlListView_GetItemText($hLV, $i, $iItemCol), 3)
        Local $sRout = StringStripWS(_GUICtrlListView_GetItemText($hLV, $i, $iRoutingCol), 3)
        Local $sVer  = StringStripWS(_GUICtrlListView_GetItemText($hLV, $i, $iVerCol), 3)
        If $sItem = "" Then ContinueLoop

        If _ItemIsRawMaterial($sItem) Then
            If $sRout <> "" Then
                $nIssues += 1
                _IntegrityIssue($sDesc & " - row " & ($i + 1) & ": MP raw material items cannot have routing '" & $sRout & "'. Leave Routing ID blank.")
            EndIf
            If $sVer <> "" Then
                $nIssues += 1
                _IntegrityIssue($sDesc & " - row " & ($i + 1) & ": MP raw material items cannot have version '" & $sVer & "'. Leave Version blank.")
            EndIf
            ContinueLoop
        EndIf

        If $sRout = "" Or $sVer = "" Then
            $nIssues += 1
            _IntegrityIssue($sDesc & " - row " & ($i + 1) & ": routing and version are required for manufactured inventory movements. MP raw materials may leave both blank.")
            ContinueLoop
        EndIf

        Local $sKey = $sItem & "|" & $sRout & "|" & $sVer
        If Not StringInStr($sSet, ";" & $sKey & ";") Then
            $nIssues += 1
            _IntegrityIssue($sDesc & " - row " & ($i + 1) & ": '" & $sKey & "' does not match any item/routing/version in the Items tab.")
        EndIf
    Next
    Return $nIssues
EndFunc

Func _ValidateBOMParentRules()
    Local $nIssues = 0
    Local $nRows = _GUICtrlListView_GetItemCount($g_hLV_BOM)
    For $i = 0 To $nRows - 1
        Local $sParent = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_BOM, $i, 0), 3)
        Local $sPVer   = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_BOM, $i, 1), 3)
        If $sParent = "" Then ContinueLoop
        If _ItemIsRawMaterial($sParent) Then
            $nIssues += 1
            _IntegrityIssue("BOM tab - row " & ($i + 1) & ": parent item '" & $sParent & "' is MP. Raw materials cannot be BOM parents because they do not have a produced version/routing.")
        ElseIf $sPVer = "" Then
            $nIssues += 1
            _IntegrityIssue("BOM tab - row " & ($i + 1) & ": parent version is required for manufactured parent items.")
        EndIf
    Next
    Return $nIssues
EndFunc

Func _ValidateWorkOrderRules()
    Local $nIssues = 0
    Local $setRoutPhase = _LVCompositeSet($g_hLV_Rout, "0,2")
    Local $nRows = _GUICtrlListView_GetItemCount($g_hLV_WO)
    For $i = 0 To $nRows - 1
        Local $sWO   = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_WO, $i, 0), 3)
        Local $sItem = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_WO, $i, 1), 3)
        Local $sRout = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_WO, $i, 2), 3)
        If $sItem <> "" And _ItemIsRawMaterial($sItem) Then
            $nIssues += 1
            _IntegrityIssue("Work Orders tab - row " & ($i + 1) & ": WO '" & $sWO & "' points to MP item '" & $sItem & "'. WOs must produce manufactured items with routing/version.")
        EndIf
        If $sRout <> "" And Not _RoutingHasAtLeastOnePhase($sRout, $setRoutPhase) Then
            $nIssues += 1
            _IntegrityIssue("Work Orders tab - row " & ($i + 1) & ": routing '" & $sRout & "' has no phase. Ortems requires B_BT phase rows for each WO.")
        EndIf
    Next
    Return $nIssues
EndFunc

Func _RoutingHasAtLeastOnePhase($sRouting, $setRoutPhase)
    Local $sR = StringStripWS($sRouting, 3)
    If $sR = "" Then Return False
    Local $aParts = StringSplit($setRoutPhase, ";", 1)
    For $i = 1 To $aParts[0]
        If StringLeft($aParts[$i], StringLen($sR) + 1) = $sR & "|" Then Return True
    Next
    Return False
EndFunc

Func _ValidateManufacturedItemRequiredFields()
    Local $nIssues = 0
    Local $nRows = _GUICtrlListView_GetItemCount($g_hLV_Mat)
    For $i = 0 To $nRows - 1
        Local $sType = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Mat, $i, 2), 3)
        Local $sVer = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Mat, $i, 3), 3)
        Local $sRout = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Mat, $i, 4), 3)
        If _IsRawMaterialType($sType) Then
            If $sVer <> "" Then
                $nIssues += 1
                _IntegrityIssue("Items tab - row " & ($i + 1) & ": MP raw material items cannot have a version. Clear the Version field.")
            EndIf
            If $sRout <> "" Then
                $nIssues += 1
                _IntegrityIssue("Items tab - row " & ($i + 1) & ": MP raw material items cannot have a routing. Clear the Routing ID field.")
            EndIf
            ContinueLoop
        EndIf
        If $sVer = "" Then
            $nIssues += 1
            _IntegrityIssue("Items tab - row " & ($i + 1) & ": version is required for SF/PF manufactured items.")
        EndIf
        If $sRout = "" Then
            $nIssues += 1
            _IntegrityIssue("Items tab - row " & ($i + 1) & ": routing is required for SF/PF manufactured items.")
        EndIf
    Next
    Return $nIssues
EndFunc

Func _ValidateStockRequiredFieldsByItemType()
    Local $nIssues = 0
    Local $nRows = _GUICtrlListView_GetItemCount($g_hLV_Stk)
    For $i = 0 To $nRows - 1
        Local $sItem = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Stk, $i, 0), 3)
        Local $sRout = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Stk, $i, 1), 3)
        Local $sVer  = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Stk, $i, 2), 3)
        If $sItem = "" Then ContinueLoop
        If _ItemIsRawMaterial($sItem) Then
            If $sRout <> "" Then
                $nIssues += 1
                _IntegrityIssue("Inventory Movements tab - row " & ($i + 1) & ": MP raw material items cannot have a routing. Clear the Routing ID field.")
            EndIf
            If $sVer <> "" Then
                $nIssues += 1
                _IntegrityIssue("Inventory Movements tab - row " & ($i + 1) & ": MP raw material items cannot have a version. Clear the Version field.")
            EndIf
            ContinueLoop
        EndIf
        If $sRout = "" Or $sVer = "" Then
            $nIssues += 1
            _IntegrityIssue("Inventory Movements tab - row " & ($i + 1) & ": routing and version are required for manufactured inventory movements.")
        EndIf
    Next
    Return $nIssues
EndFunc

Func _IntegrityIssue($sMsg)
    $g_sIntegrityReport &= "- " & $sMsg & @CRLF
    _Log($sMsg)
EndFunc

Func _IntegrityPreview($iMaxLines = 12)
    Local $sText = StringStripCR($g_sIntegrityReport)
    If $sText = "" Then Return ""

    Local $aLines = StringSplit($sText, @LF, 1)
    Local $sOut = ""
    Local $nShown = 0
    Local $nTotal = 0

    For $i = 1 To $aLines[0]
        Local $sLine = StringStripWS($aLines[$i], 3)
        If $sLine = "" Then ContinueLoop
        $nTotal += 1
        If $nShown < $iMaxLines Then
            If $sOut <> "" Then $sOut &= @CRLF
            $sOut &= $sLine
            $nShown += 1
        EndIf
    Next

    If $nTotal > $nShown Then $sOut &= @CRLF & "..."
    Return $sOut
EndFunc

Func _IntegrityBuildMessage($nIssues)
    Local $sPreview = _IntegrityPreview(12)
    Local $sMsg = $nIssues & " integrity issue(s) were found." & @CRLF & @CRLF
    If $sPreview <> "" Then $sMsg &= $sPreview & @CRLF & @CRLF
    $sMsg &= "Full details were written to the execution log and to the session log file."
    Return $sMsg
EndFunc

Func _AutoFixSimpleReferences()
    Local $nFixed = 0
    _Log("=== Integrity Auto-Fix started ===")

    Local $setCalIDs         = _LVColSet($g_hLV_Cal, 0)
    Local $setCapCalIDs      = _LVColSet($g_hLV_Cap, 0)
    Local $setAnyCalIDs      = $setCalIDs & StringTrimLeft($setCapCalIDs, 1)
    Local $setCTIDs          = _LVColSet($g_hLV_Mach, 2)
    Local $setMachIDs        = _LVColSet($g_hLV_Mach, 7)
    Local $setOpeIDs         = _LVColSet($g_hLV_Ops, 0)
    Local $setRoutIDs        = _LVColSet($g_hLV_Rout, 0)
    Local $setMatIDs         = _LVColSet($g_hLV_Mat, 0)
    Local $setWOIDs          = _LVColSet($g_hLV_WO, 0)

    ; Case/spacing normalization against existing IDs
    $nFixed += _AutoFixMachineCalendarRef($g_hLV_Mach, 4, 10, $setCalIDs, "Machines tab / Calendar ID")
    $nFixed += _AutoFixExactRef($g_hLV_Mach, 11, $setAnyCalIDs, "Machines tab / Capacity calendar ID", True)

    $nFixed += _AutoFixExactRef($g_hLV_Ops, 2, $setCTIDs,       "Operations tab / WC ID")
    $nFixed += _AutoFixMachineExactRef($g_hLV_Ops, 3, $setMachIDs, "Operations tab / Machine ID")

    $nFixed += _AutoFixExactRef($g_hLV_Rout, 3, $setOpeIDs,     "Routings tab / Operation ID")

    $nFixed += _AutoFixItemRoutingRef($g_hLV_Mat, 2, 4, $setRoutIDs,     "Items tab / Routing ID", True)

    $nFixed += _AutoFixExactRef($g_hLV_BOM, 3, $setRoutIDs,     "BOM tab / Routing ID")

    $nFixed += _AutoFixExactRef($g_hLV_WO,  0, $setWOIDs,       "Work Orders tab / WO ID")
    $nFixed += _AutoFixExactRef($g_hLV_WO,  1, $setMatIDs,      "Work Orders tab / Item ID")
    $nFixed += _AutoFixExactRef($g_hLV_WO,  2, $setRoutIDs,     "Work Orders tab / Routing ID")

    $nFixed += _AutoFixExactRef($g_hLV_WOL, 0, $setWOIDs,       "WO Links tab / Predecessor WO", True)
    $nFixed += _AutoFixExactRef($g_hLV_WOL, 1, $setRoutIDs,     "WO Links tab / Predecessor routing", True)
    $nFixed += _AutoFixExactRef($g_hLV_WOL, 3, $setWOIDs,       "WO Links tab / Successor WO", True)
    $nFixed += _AutoFixExactRef($g_hLV_WOL, 4, $setRoutIDs,     "WO Links tab / Successor routing", True)

    $nFixed += _AutoFixExactRef($g_hLV_SR,  0, $setOpeIDs,      "Secondary Resources tab / Operation ID")
    $nFixed += _AutoFixExactRef($g_hLV_SR,  1, $setCTIDs,       "Secondary Resources tab / WC ID")
    $nFixed += _AutoFixMachineExactRef($g_hLV_SR, 2, $setMachIDs, "Secondary Resources tab / Machine ID")
    $nFixed += _AutoFixExactRef($g_hLV_SR,  4, $setAnyCalIDs,   "Secondary Resources tab / Capacity calendar ID", True)

    $nFixed += _AutoFixExactRef($g_hLV_Stk, 0, $setMatIDs,      "Inventory Movements tab / Item ID")
    $nFixed += _AutoFixExactRef($g_hLV_Stk, 1, $setRoutIDs,     "Inventory Movements tab / Routing ID", True)

    ; MP items cannot keep produced-item metadata.
    $nFixed += _AutoFixRawMaterialItemVersionRouting()
    $nFixed += _AutoFixRawMaterialInventoryVersionRouting()

    ; Safe contextual auto-fixes (only when the lookup resolves to a single possible value)
    $nFixed += _AutoFixMachineByWC($g_hLV_Ops, 2, 3, "Operations tab / Machine ID")
    $nFixed += _AutoFixMachineByWC($g_hLV_SR,  1, 2, "Secondary Resources tab / Machine ID")

    $nFixed += _AutoFixRoutingByItemVersion($g_hLV_BOM, 0, 1, 3, "BOM tab / Routing ID")
    $nFixed += _AutoFixRoutingByItemVersion($g_hLV_WO,  1, 3, 2, "Work Orders tab / Routing ID")
    $nFixed += _AutoFixRoutingByItemVersion($g_hLV_Stk, 0, 2, 1, "Inventory Movements tab / Routing ID")

    $nFixed += _AutoFixRoutingByWO($g_hLV_WOL, 0, 1, "WO Links tab / Predecessor routing")
    $nFixed += _AutoFixRoutingByWO($g_hLV_WOL, 3, 4, "WO Links tab / Successor routing")

    If $nFixed = 0 Then
        _Log("=== Integrity Auto-Fix finished - no automatic change was applied ===")
    Else
        _Log("=== Integrity Auto-Fix finished - " & $nFixed & " field(s) updated ===")
    EndIf
    Return $nFixed
EndFunc


Func _AutoFixRawMaterialItemVersionRouting()
    Local $nFixed = 0
    Local $nRows = _GUICtrlListView_GetItemCount($g_hLV_Mat)
    For $i = 0 To $nRows - 1
        Local $sType = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Mat, $i, 2), 3)
        If Not _IsRawMaterialType($sType) Then ContinueLoop
        Local $sVer = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Mat, $i, 3), 3)
        Local $sRout = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Mat, $i, 4), 3)
        If $sVer <> "" Then
            _GUICtrlListView_SetItemText($g_hLV_Mat, $i, "", 3)
            $nFixed += 1
            _Log("AUTO-FIX: Items tab row " & ($i + 1) & " MP version cleared ('" & $sVer & "').")
        EndIf
        If $sRout <> "" Then
            _GUICtrlListView_SetItemText($g_hLV_Mat, $i, "", 4)
            $nFixed += 1
            _Log("AUTO-FIX: Items tab row " & ($i + 1) & " MP routing cleared ('" & $sRout & "').")
        EndIf
    Next
    Return $nFixed
EndFunc

Func _AutoFixRawMaterialInventoryVersionRouting()
    Local $nFixed = 0
    Local $nRows = _GUICtrlListView_GetItemCount($g_hLV_Stk)
    For $i = 0 To $nRows - 1
        Local $sItem = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Stk, $i, 0), 3)
        If $sItem = "" Or Not _ItemIsRawMaterial($sItem) Then ContinueLoop
        Local $sRout = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Stk, $i, 1), 3)
        Local $sVer = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Stk, $i, 2), 3)
        If $sRout <> "" Then
            _GUICtrlListView_SetItemText($g_hLV_Stk, $i, "", 1)
            $nFixed += 1
            _Log("AUTO-FIX: Inventory Movements tab row " & ($i + 1) & " MP routing cleared ('" & $sRout & "').")
        EndIf
        If $sVer <> "" Then
            _GUICtrlListView_SetItemText($g_hLV_Stk, $i, "", 2)
            $nFixed += 1
            _Log("AUTO-FIX: Inventory Movements tab row " & ($i + 1) & " MP version cleared ('" & $sVer & "').")
        EndIf
    Next
    Return $nFixed
EndFunc

Func _AutoFixExactRef($hLV, $iCol, $sSet, $sDesc, $bAllowEmpty = False)
    Local $nFixed = 0
    Local $nRows = _GUICtrlListView_GetItemCount($hLV)
    For $i = 0 To $nRows - 1
        Local $sRaw = _GUICtrlListView_GetItemText($hLV, $i, $iCol)
        Local $sVal = StringStripWS($sRaw, 3)
        If $sVal = "" Then
            If $bAllowEmpty Then ContinueLoop
            ContinueLoop
        EndIf
        If StringInStr($sSet, ";" & $sVal & ";") Then ContinueLoop

        Local $sCanonical = _LookupCanonicalInSet($sSet, $sVal)
        If $sCanonical = "" Then ContinueLoop

        _GUICtrlListView_SetItemText($hLV, $i, $sCanonical, $iCol)
        $nFixed += 1
        _Log("AUTO-FIX: " & $sDesc & " - row " & ($i + 1) & ": '" & $sRaw & "' -> '" & $sCanonical & "'")
    Next
    Return $nFixed
EndFunc

Func _AutoFixMachineCalendarRef($hLV, $iCTTypeCol, $iCalCol, $sSet, $sDesc, $bAllowEmpty = False)
    Local $nFixed = 0
    Local $nRows = _GUICtrlListView_GetItemCount($hLV)
    For $i = 0 To $nRows - 1
        Local $sCTType = StringStripWS(_GUICtrlListView_GetItemText($hLV, $i, $iCTTypeCol), 3)
        If _IsMachineCalendarOptional($sCTType) Then ContinueLoop

        Local $sRaw = _GUICtrlListView_GetItemText($hLV, $i, $iCalCol)
        Local $sVal = StringStripWS($sRaw, 3)
        If $sVal = "" Then
            If $bAllowEmpty Then ContinueLoop
            ContinueLoop
        EndIf
        If StringInStr($sSet, ";" & $sVal & ";") Then ContinueLoop

        Local $sCanonical = _LookupCanonicalInSet($sSet, $sVal)
        If $sCanonical = "" Then ContinueLoop

        _GUICtrlListView_SetItemText($hLV, $i, $sCanonical, $iCalCol)
        $nFixed += 1
        _Log("AUTO-FIX: " & $sDesc & " - row " & ($i + 1) & ": '" & $sRaw & "' -> '" & $sCanonical & "'")
    Next
    Return $nFixed
EndFunc


Func _AutoFixMachineExactRef($hLV, $iCol, $sSet, $sDesc, $bAllowEmpty = False)
    Local $nFixed = 0
    Local $nRows = _GUICtrlListView_GetItemCount($hLV)
    For $i = 0 To $nRows - 1
        Local $sRaw = _GUICtrlListView_GetItemText($hLV, $i, $iCol)
        Local $sVal = StringStripWS($sRaw, 3)
        If _IsStandByMachine($sVal) Then ContinueLoop
        If $sVal = "" Then
            If $bAllowEmpty Then ContinueLoop
            ContinueLoop
        EndIf
        If StringInStr($sSet, ";" & $sVal & ";") Then ContinueLoop

        Local $sCanonical = _LookupCanonicalInSet($sSet, $sVal)
        If $sCanonical = "" Then ContinueLoop

        _GUICtrlListView_SetItemText($hLV, $i, $sCanonical, $iCol)
        $nFixed += 1
        _Log("AUTO-FIX: " & $sDesc & " - row " & ($i + 1) & ": '" & $sRaw & "' -> '" & $sCanonical & "'")
    Next
    Return $nFixed
EndFunc


Func _AutoFixItemRoutingRef($hLV, $iTypeCol, $iRoutingCol, $sSet, $sDesc, $bAllowEmpty = False)
    Local $nFixed = 0
    Local $nRows = _GUICtrlListView_GetItemCount($hLV)

    For $i = 0 To $nRows - 1
        Local $sType = StringStripWS(_GUICtrlListView_GetItemText($hLV, $i, $iTypeCol), 3)
        If _IsRawMaterialType($sType) Then ContinueLoop

        Local $sRaw = _GUICtrlListView_GetItemText($hLV, $i, $iRoutingCol)
        Local $sVal = StringStripWS($sRaw, 3)
        If $sVal = "" Then
            If $bAllowEmpty Then ContinueLoop
            ContinueLoop
        EndIf
        If StringInStr($sSet, ";" & $sVal & ";") Then ContinueLoop

        Local $sCanonical = _LookupCanonicalInSet($sSet, $sVal)
        If $sCanonical = "" Then ContinueLoop

        _GUICtrlListView_SetItemText($hLV, $i, $sCanonical, $iRoutingCol)
        $nFixed += 1
        _Log("AUTO-FIX: " & $sDesc & " - row " & ($i + 1) & ": '" & $sRaw & "' -> '" & $sCanonical & "'")
    Next
    Return $nFixed
EndFunc

Func _LookupCanonicalInSet($sSet, $sValue)
    Local $sNeedle = StringLower(StringStripWS($sValue, 3))
    If $sNeedle = "" Then Return ""

    Local $aItems = StringSplit($sSet, ";", 1)
    Local $sMatch = ""
    Local $nMatches = 0

    For $i = 1 To $aItems[0]
        Local $sItem = $aItems[$i]
        If $sItem = "" Then ContinueLoop
        If StringLower(StringStripWS($sItem, 3)) = $sNeedle Then
            $nMatches += 1
            $sMatch = $sItem
            If $nMatches > 1 Then Return ""
        EndIf
    Next
    Return $sMatch
EndFunc

Func _AutoFixMachineByWC($hLV, $iWCCol, $iMachineCol, $sDesc)
    Local $nFixed = 0
    Local $nRows = _GUICtrlListView_GetItemCount($hLV)
    Local $setMachIDs = _LVColSet($g_hLV_Mach, 7)

    For $i = 0 To $nRows - 1
        Local $sWC = StringStripWS(_GUICtrlListView_GetItemText($hLV, $i, $iWCCol), 3)
        Local $sMachRaw = _GUICtrlListView_GetItemText($hLV, $i, $iMachineCol)
        Local $sMach = StringStripWS($sMachRaw, 3)
        If $sWC = "" Then ContinueLoop
        If _IsStandByMachine($sMach) Then ContinueLoop
        If $sMach <> "" And StringInStr($setMachIDs, ";" & $sMach & ";") Then ContinueLoop

        Local $sResolved = _LookupUniqueMachineByWC($sWC)
        If $sResolved = "" Then ContinueLoop

        _GUICtrlListView_SetItemText($hLV, $i, $sResolved, $iMachineCol)
        $nFixed += 1
        _Log("AUTO-FIX: " & $sDesc & " - row " & ($i + 1) & ": '" & $sMachRaw & "' -> '" & $sResolved & "' (derived from WC '" & $sWC & "')")
    Next
    Return $nFixed
EndFunc

Func _LookupUniqueMachineByWC($sWC)
    Local $sWCNorm = StringLower(StringStripWS($sWC, 3))
    If $sWCNorm = "" Then Return ""

    Local $sFound = ""
    Local $sSeen = ";"
    Local $nUnique = 0
    Local $nRows = _GUICtrlListView_GetItemCount($g_hLV_Mach)

    For $i = 0 To $nRows - 1
        Local $sRowWC = StringLower(StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Mach, $i, 2), 3))
        If $sRowWC <> $sWCNorm Then ContinueLoop

        Local $sMachine = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Mach, $i, 7), 3)
        If $sMachine = "" Then ContinueLoop
        If StringInStr($sSeen, ";" & $sMachine & ";") Then ContinueLoop

        $sSeen &= $sMachine & ";"
        $nUnique += 1
        $sFound = $sMachine
        If $nUnique > 1 Then Return ""
    Next
    Return $sFound
EndFunc

Func _AutoFixRoutingByItemVersion($hLV, $iItemCol, $iVerCol, $iRoutingCol, $sDesc)
    Local $nFixed = 0
    Local $nRows = _GUICtrlListView_GetItemCount($hLV)
    Local $setRoutIDs = _LVColSet($g_hLV_Rout, 0)

    For $i = 0 To $nRows - 1
        Local $sItem = StringStripWS(_GUICtrlListView_GetItemText($hLV, $i, $iItemCol), 3)
        Local $sVer = StringStripWS(_GUICtrlListView_GetItemText($hLV, $i, $iVerCol), 3)
        Local $sRoutingRaw = _GUICtrlListView_GetItemText($hLV, $i, $iRoutingCol)
        Local $sRouting = StringStripWS($sRoutingRaw, 3)

        If $sItem = "" Or $sVer = "" Then ContinueLoop
        If $sRouting <> "" And StringInStr($setRoutIDs, ";" & $sRouting & ";") Then ContinueLoop

        Local $sResolved = _LookupUniqueRoutingByItemVersion($sItem, $sVer)
        If $sResolved = "" Then ContinueLoop

        _GUICtrlListView_SetItemText($hLV, $i, $sResolved, $iRoutingCol)
        $nFixed += 1
        _Log("AUTO-FIX: " & $sDesc & " - row " & ($i + 1) & ": '" & $sRoutingRaw & "' -> '" & $sResolved & "' (derived from item '" & $sItem & "' / version '" & $sVer & "')")
    Next
    Return $nFixed
EndFunc

Func _LookupUniqueRoutingByItemVersion($sItem, $sVer)
    Local $sItemNorm = StringLower(StringStripWS($sItem, 3))
    Local $sVerNorm = StringLower(StringStripWS($sVer, 3))
    If $sItemNorm = "" Or $sVerNorm = "" Then Return ""

    Local $sFound = ""
    Local $sSeen = ";"
    Local $nUnique = 0
    Local $nRows = _GUICtrlListView_GetItemCount($g_hLV_Mat)

    For $i = 0 To $nRows - 1
        Local $sRowItem = StringLower(StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Mat, $i, 0), 3))
        Local $sRowVer = StringLower(StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Mat, $i, 3), 3))
        If $sRowItem <> $sItemNorm Or $sRowVer <> $sVerNorm Then ContinueLoop

        Local $sType = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Mat, $i, 2), 3)
        If _IsRawMaterialType($sType) Then ContinueLoop

        Local $sRouting = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Mat, $i, 4), 3)
        If $sRouting = "" Then ContinueLoop
        If StringInStr($sSeen, ";" & $sRouting & ";") Then ContinueLoop

        $sSeen &= $sRouting & ";"
        $nUnique += 1
        $sFound = $sRouting
        If $nUnique > 1 Then Return ""
    Next
    Return $sFound
EndFunc

Func _AutoFixRoutingByWO($hLV, $iWOCol, $iRoutingCol, $sDesc)
    Local $nFixed = 0
    Local $nRows = _GUICtrlListView_GetItemCount($hLV)
    Local $setRoutIDs = _LVColSet($g_hLV_Rout, 0)

    For $i = 0 To $nRows - 1
        Local $sWO = StringStripWS(_GUICtrlListView_GetItemText($hLV, $i, $iWOCol), 3)
        Local $sRoutingRaw = _GUICtrlListView_GetItemText($hLV, $i, $iRoutingCol)
        Local $sRouting = StringStripWS($sRoutingRaw, 3)

        If $sWO = "" Then ContinueLoop
        If $sRouting <> "" And StringInStr($setRoutIDs, ";" & $sRouting & ";") Then ContinueLoop

        Local $sResolved = _LookupUniqueRoutingByWO($sWO)
        If $sResolved = "" Then ContinueLoop

        _GUICtrlListView_SetItemText($hLV, $i, $sResolved, $iRoutingCol)
        $nFixed += 1
        _Log("AUTO-FIX: " & $sDesc & " - row " & ($i + 1) & ": '" & $sRoutingRaw & "' -> '" & $sResolved & "' (derived from WO '" & $sWO & "')")
    Next
    Return $nFixed
EndFunc

Func _LookupUniqueRoutingByWO($sWO)
    Local $sWONorm = StringLower(StringStripWS($sWO, 3))
    If $sWONorm = "" Then Return ""

    Local $sFound = ""
    Local $sSeen = ";"
    Local $nUnique = 0
    Local $nRows = _GUICtrlListView_GetItemCount($g_hLV_WO)

    For $i = 0 To $nRows - 1
        Local $sRowWO = StringLower(StringStripWS(_GUICtrlListView_GetItemText($g_hLV_WO, $i, 0), 3))
        If $sRowWO <> $sWONorm Then ContinueLoop

        Local $sRouting = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_WO, $i, 2), 3)
        If $sRouting = "" Then ContinueLoop
        If StringInStr($sSeen, ";" & $sRouting & ";") Then ContinueLoop

        $sSeen &= $sRouting & ";"
        $nUnique += 1
        $sFound = $sRouting
        If $nUnique > 1 Then Return ""
    Next
    Return $sFound
EndFunc

Func _ValidateCalendarFormats()
    Local $nIssues = 0
    Local $nRows = _GUICtrlListView_GetItemCount($g_hLV_Cal)
    For $i = 0 To $nRows - 1
        Local $sDiaI  = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Cal, $i, 2), 3)
        Local $sHoraI = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Cal, $i, 3), 3)
        Local $sDiaF  = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Cal, $i, 4), 3)
        Local $sHoraF = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Cal, $i, 5), 3)

        If Number($sDiaI) < 1 Or Number($sDiaI) > 7 Then
            $nIssues += 1
            _IntegrityIssue("Calendars tab - row " & ($i + 1) & ": start day '" & $sDiaI & "' should be between 1 and 7.")
        EndIf
        If Number($sDiaF) < 1 Or Number($sDiaF) > 7 Then
            $nIssues += 1
            _IntegrityIssue("Calendars tab - row " & ($i + 1) & ": end day '" & $sDiaF & "' should be between 1 and 7.")
        EndIf
        If _NormalizeTimeText($sHoraI) = "" Then
            $nIssues += 1
            _IntegrityIssue("Calendars tab - row " & ($i + 1) & ": start time '" & $sHoraI & "' is not valid.")
        EndIf
        If _NormalizeTimeText($sHoraF) = "" Then
            $nIssues += 1
            _IntegrityIssue("Calendars tab - row " & ($i + 1) & ": end time '" & $sHoraF & "' is not valid.")
        EndIf
    Next
    Return $nIssues
EndFunc

Func _ValidateCapacityFormats()
    Local $nIssues = 0
    Local $nRows = _GUICtrlListView_GetItemCount($g_hLV_Cap)
    For $i = 0 To $nRows - 1
        Local $sDiaI  = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Cap, $i, 1), 3)
        Local $sHoraI = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Cap, $i, 2), 3)
        Local $sDiaF  = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Cap, $i, 3), 3)
        Local $sHoraF = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Cap, $i, 4), 3)
        Local $sRes   = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Cap, $i, 5), 3)

        If Number($sDiaI) < 1 Or Number($sDiaI) > 7 Then
            $nIssues += 1
            _IntegrityIssue("Capacity tab - row " & ($i + 1) & ": start day '" & $sDiaI & "' should be between 1 and 7.")
        EndIf
        If Number($sDiaF) < 1 Or Number($sDiaF) > 7 Then
            $nIssues += 1
            _IntegrityIssue("Capacity tab - row " & ($i + 1) & ": end day '" & $sDiaF & "' should be between 1 and 7.")
        EndIf
        If _NormalizeTimeText($sHoraI) = "" Then
            $nIssues += 1
            _IntegrityIssue("Capacity tab - row " & ($i + 1) & ": start time '" & $sHoraI & "' is not valid.")
        EndIf
        If _NormalizeTimeText($sHoraF) = "" Then
            $nIssues += 1
            _IntegrityIssue("Capacity tab - row " & ($i + 1) & ": end time '" & $sHoraF & "' is not valid.")
        EndIf
        If $sRes <> "" And Not StringRegExp($sRes, "^-?\d+([.,]\d+)?$") Then
            $nIssues += 1
            _IntegrityIssue("Capacity tab - row " & ($i + 1) & ": #Resources '" & $sRes & "' is not numeric.")
        EndIf
    Next
    Return $nIssues
EndFunc

Func _ValidateWODates()
    Local $nIssues = 0
    Local $nRows = _GUICtrlListView_GetItemCount($g_hLV_WO)
    For $i = 0 To $nRows - 1
        For $c = 5 To 6
            Local $sVal = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_WO, $i, $c), 3)
            If $sVal = "" Then ContinueLoop
            If Not StringRegExp($sVal, "^\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2}$") Then
                $nIssues += 1
                _IntegrityIssue("Work Orders tab - row " & ($i + 1) & ": date '" & $sVal & "' should be dd/mm/yyyy hh:mm.")
            EndIf
        Next
    Next
    Return $nIssues
EndFunc

Func _ValidateStockDates()
    Local $nIssues = 0
    Local $nRows = _GUICtrlListView_GetItemCount($g_hLV_Stk)
    For $i = 0 To $nRows - 1
        Local $sVal = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Stk, $i, 3), 3)
        If $sVal = "" Then ContinueLoop
        If Not StringRegExp($sVal, "^\d{2}/\d{2}/\d{4}$") Then
            $nIssues += 1
            _IntegrityIssue("Inventory Movements tab - row " & ($i + 1) & ": move date '" & $sVal & "' should be dd/mm/yyyy.")
        EndIf
    Next
    Return $nIssues
EndFunc

Func _RequiredFieldCheck($hLV, $sCols, $sDesc)
    Local $nIssues = 0
    Local $aCols = StringSplit($sCols, ",", 1)
    Local $nRows = _GUICtrlListView_GetItemCount($hLV)
    For $i = 0 To $nRows - 1
        For $j = 1 To $aCols[0]
            Local $iCol = Number($aCols[$j])
            Local $sVal = StringStripWS(_GUICtrlListView_GetItemText($hLV, $i, $iCol), 3)
            If $sVal = "" Then
                $nIssues += 1
                _IntegrityIssue($sDesc & " - row " & ($i + 1) & ": required field in column " & ($iCol + 1) & " is empty.")
            EndIf
        Next
    Next
    Return $nIssues
EndFunc

Func _ValidateRequiredFields()
    Local $nIssues = 0
    $nIssues += _RequiredFieldCheck($g_hLV_Cal,  "0,1,2,3,4,5", "Calendars tab")
    $nIssues += _RequiredFieldCheck($g_hLV_Mach, "0,2,4,7,9",     "Machines tab")
    $nIssues += _RequiredFieldCheck($g_hLV_Ops,  "0,2,3,4,5,6",   "Operations tab")
    $nIssues += _RequiredFieldCheck($g_hLV_Rout, "0,2,3",         "Routings tab")
    $nIssues += _RequiredFieldCheck($g_hLV_Mat,  "0,1,2",         "Items tab")
    $nIssues += _ValidateManufacturedItemRequiredFields()
    $nIssues += _RequiredFieldCheck($g_hLV_BOM,  "0,2,3,4,5,6",   "BOM tab")
    $nIssues += _RequiredFieldCheck($g_hLV_WO,   "0,1,2,3,4",     "Work Orders tab")
    $nIssues += _RequiredFieldCheck($g_hLV_SR,   "0,1,2,3",       "Secondary Resources tab")
    $nIssues += _RequiredFieldCheck($g_hLV_Cap,  "0,1,2,3,4,5",   "Capacity tab")
    $nIssues += _RequiredFieldCheck($g_hLV_Stk,  "0,3,4",         "Inventory Movements tab")
    $nIssues += _ValidateStockRequiredFieldsByItemType()
    Return $nIssues
EndFunc

Func _ValidateDuplicateKey($hLV, $sCols, $sDesc)
    Local $nIssues = 0
    Local $sSeen = ";"
    Local $nRows = _GUICtrlListView_GetItemCount($hLV)
    For $i = 0 To $nRows - 1
        Local $sKey = _LVComposeKey($hLV, $i, $sCols)
        If $sKey = "" Then ContinueLoop
        Local $sNorm = StringLower(StringStripWS($sKey, 3))
        If StringInStr($sSeen, ";" & $sNorm & ";") Then
            $nIssues += 1
            _IntegrityIssue($sDesc & " - row " & ($i + 1) & ": duplicate key '" & $sKey & "'.")
        Else
            $sSeen &= $sNorm & ";"
        EndIf
    Next
    Return $nIssues
EndFunc

Func _ValidateDuplicateMaterialKeys()
    Local $nIssues = 0
    Local $sSeenRaw = ";"
    Local $sSeenMfg = ";"
    Local $nRows = _GUICtrlListView_GetItemCount($g_hLV_Mat)

    For $i = 0 To $nRows - 1
        Local $sItem = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Mat, $i, 0), 3)
        Local $sType = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Mat, $i, 2), 3)
        Local $sVer  = StringStripWS(_GUICtrlListView_GetItemText($g_hLV_Mat, $i, 3), 3)
        If $sItem = "" Then ContinueLoop

        If _IsRawMaterialType($sType) Then
            Local $sRawKey = StringLower($sItem)
            If StringInStr($sSeenRaw, ";" & $sRawKey & ";") Then
                $nIssues += 1
                _IntegrityIssue("Items tab / MP Item - row " & ($i + 1) & ": duplicate raw material item '" & $sItem & "'.")
            Else
                $sSeenRaw &= $sRawKey & ";"
            EndIf
        Else
            If $sVer = "" Then ContinueLoop
            Local $sMfgKey = StringLower($sItem & "|" & $sVer)
            If StringInStr($sSeenMfg, ";" & $sMfgKey & ";") Then
                $nIssues += 1
                _IntegrityIssue("Items tab / Item + Version - row " & ($i + 1) & ": duplicate key '" & $sItem & "|" & $sVer & "'.")
            Else
                $sSeenMfg &= $sMfgKey & ";"
            EndIf
        EndIf
    Next
    Return $nIssues
EndFunc

Func _ValidateDuplicateKeys()
    Local $nIssues = 0
    $nIssues += _ValidateDuplicateKey($g_hLV_Cal,  "0,2,3",   "Calendars tab / Calendar + start day + start time")
    $nIssues += _ValidateDuplicateKey($g_hLV_Mach, "7",       "Machines tab / Machine ID")
    $nIssues += _ValidateDuplicateKey($g_hLV_Mach, "2,7",     "Machines tab / WC + Machine")
    $nIssues += _ValidateDuplicateKey($g_hLV_Ops,  "0,2,3",   "Operations tab / Operation + WC + Machine")
    $nIssues += _ValidateDuplicateKey($g_hLV_Rout, "0,2",     "Routings tab / Routing + Phase")
    $nIssues += _ValidateDuplicateMaterialKeys()
    $nIssues += _ValidateDuplicateKey($g_hLV_WO,   "0",       "Work Orders tab / WO ID")
    $nIssues += _ValidateDuplicateKey($g_hLV_WOL,  "0,1,2,3,4,5", "WO Links tab / full link")
    $nIssues += _ValidateDuplicateKey($g_hLV_Cap,  "0,1,2",   "Capacity tab / Capacity calendar + start day + start time")
    Return $nIssues
EndFunc

Func _TableListSQL($aTbl)
    Local $s = ""
    For $i = 0 To UBound($aTbl) - 1
        If $i > 0 Then $s &= ","
        $s &= "'" & $aTbl[$i] & "'"
    Next
    Return $s
EndFunc

Func _ClearDatabase()
    If Not $g_bConnected Then
        MsgBox(262144+48, "Warning", "Not connected to the database.",0,$g_hMain)
        Return
    EndIf

    Local $nRet = MsgBox(262144+4 + 16, "Confirm clear", _
        "WARNING: This action will DELETE ALL demo data from the database!" & @CRLF & @CRLF & _
        "Database: " & $g_sDatabase & @CRLF & @CRLF & "Do you want to continue?",0,$g_hMain)

    If $nRet <> 6 Then Return

    Local $oConn = ObjCreate("ADODB.Connection")
    If Not IsObj($oConn) Then
        _Log("ERROR: Could not create connection object.")
        Return
    EndIf

    $g_sLastComError = ""
    $oConn.Open($g_sConnStr)
    If $g_sLastComError <> "" Or $oConn.State <> 1 Then
        _Log("ERROR opening connection for clear: " & $g_sLastComError)
        MsgBox(262144+16, "Error", "Could not open connection: " & @CRLF & $g_sLastComError,0,$g_hMain)
        Return
    EndIf

    Local $aStmts = _GetClearSQL()   ; array: [0]=count, [1..n]=statements
    Local $nOK = 0, $nErr = 0

    For $i = 1 To $aStmts[0]
        Local $sStmt = StringStripWS($aStmts[$i], 3)
        If $sStmt = "" Then ContinueLoop

        _LogVerboseSQL("Clear database stmt " & $i, $sStmt)
        $g_sLastComError = ""
        $oConn.Execute($sStmt)
        If $g_sLastComError <> "" Then
            $nErr += 1
            _Log("ERROR clear stmt " & $i & ": " & $g_sLastComError)
        Else
            $nOK += 1
        EndIf
    Next

    $oConn.Close()

    If $nErr = 0 Then
        _Log("Database cleared successfully (" & $nOK & " statements OK).")
        MsgBox(262144+64, "OK", "Demo data removed from the database successfully.",0,$g_hMain)
    Else
        _Log("Clear completed with " & $nErr & " error(s) - check the session log file for details.")
        MsgBox(262144+48, "Completed with errors", "Clear finished with " & $nErr & " error(s)." & @CRLF & "Check the execution log and the session log file for details.",0,$g_hMain)
    EndIf
EndFunc


;=============================================================================
; IMPORT / EXPORT WORKBOOK
;=============================================================================
Func _DatasetMap()
    Local $aMap[11][4] = [ _
        [$g_hLV_Cal,  "Calendars",           "calendars",             "Work Calendars"], _
        [$g_hLV_Mach, "Machines",            "machines",              "Machines and Work Centers"], _
        [$g_hLV_Ops,  "Operations",          "operations",            "Operations"], _
        [$g_hLV_Rout, "Routings",            "routings",              "Production Routings"], _
        [$g_hLV_Mat,  "Items",               "items",                 "Items"], _
        [$g_hLV_BOM,  "BOM",                 "bom",                   "Bill of Materials"], _
        [$g_hLV_WO,   "WorkOrders",          "workorders",            "Work Orders"], _
        [$g_hLV_WOL,  "WOLinks",             "wo_links",              "WO Links"], _
        [$g_hLV_SR,   "SecondaryResources",  "secondary_resources",   "Secondary Resources"], _
        [$g_hLV_Cap,  "Capacity",            "capacity",              "Capacity Calendars"], _
        [$g_hLV_Stk,  "InventoryMovements",  "stock",                 "Inventory Movements"] _
    ]
    Return $aMap
EndFunc

Func _ExcelCreateApp()
    $g_sLastComError = ""
    Local $oExcel = ObjCreate("Excel.Application")
    If $g_sLastComError <> "" Or Not IsObj($oExcel) Then
        MsgBox(262144+16, "Excel export/import", "Microsoft Excel could not be started." & @CRLF & @CRLF & _
            "The workbook import/export requires Excel to be installed on this Windows machine." & @CRLF & @CRLF & _
            "Technical details: " & $g_sLastComError,0,$g_hMain)
        Return 0
    EndIf
    $oExcel.Visible = False
    $oExcel.DisplayAlerts = False
    Return $oExcel
EndFunc

Func _LV_GetHeaders($hLV)
    Local $nCols = _GUICtrlListView_GetColumnCount($hLV)
    Local $nDataCols = $nCols - 1
    Local $aHeaders[$nDataCols]
    For $c = 0 To $nDataCols - 1
        Local $aCol = _GUICtrlListView_GetColumn($hLV, $c)
        Local $sHdr = "Column " & ($c + 1)
        If IsArray($aCol) And UBound($aCol) > 5 Then $sHdr = $aCol[5]
        $aHeaders[$c] = $sHdr
    Next
    Return $aHeaders
EndFunc

Func _NormalizeHeader($sText)
    Local $s = StringLower(StringStripWS($sText, 3))
    $s = StringReplace($s, " ", "")
    $s = StringReplace($s, "_", "")
    $s = StringReplace($s, "-", "")
    $s = StringReplace($s, "/", "")
    Return $s
EndFunc

Func _XLWriteReadMe($oSheet, $sStatus, $nIssues)
    $oSheet.Name = "README"
    $oSheet.Cells.NumberFormat = "@"
    $oSheet.Cells(1, 1).Value = "Ortems Toolbox Data Workbook"
    $oSheet.Cells(1, 1).Font.Bold = True
    $oSheet.Cells(1, 1).Font.Size = 16
    $oSheet.Cells(3, 1).Value = "Purpose"
    $oSheet.Cells(3, 2).Value = "Round-trip data file generated by Ortems Toolbox. Edit the data sheets, keep the headers unchanged, and import the workbook back into the tool."
    $oSheet.Cells(4, 1).Value = "Generated on"
    $oSheet.Cells(4, 2).Value = @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC
    $oSheet.Cells(5, 1).Value = "Tool version"
    $oSheet.Cells(5, 2).Value = $TITLE
    $oSheet.Cells(6, 1).Value = "Integrity status at export"
    $oSheet.Cells(6, 2).Value = $sStatus
    $oSheet.Cells(7, 1).Value = "Integrity issue count"
    $oSheet.Cells(7, 2).Value = $nIssues
    $oSheet.Cells(9, 1).Value = "Rules"
    $oSheet.Cells(9, 2).Value = "Do not rename sheets or headers. All data is exported as text to avoid date/time conversion issues. The Line column is regenerated by the app during import."
    $oSheet.Columns("A:B").AutoFit()
    $oSheet.Range("A3:A9").Font.Bold = True
    $oSheet.Range("A1:B1").Interior.Color = 15189684
EndFunc

Func _LV_ExportToWorksheet($hLV, $oSheet, $sFriendlyName)
    Local $nCols = _GUICtrlListView_GetColumnCount($hLV)
    Local $nDataCols = $nCols - 1
    Local $nRows = _GUICtrlListView_GetItemCount($hLV)
    If $nDataCols < 1 Then Return 0

    $oSheet.Cells.NumberFormat = "@"
    $oSheet.Cells(1, 1).Value = $sFriendlyName
    $oSheet.Range($oSheet.Cells(1, 1), $oSheet.Cells(1, $nDataCols)).Merge
    $oSheet.Cells(1, 1).Font.Bold = True
    $oSheet.Cells(1, 1).Font.Size = 14
    $oSheet.Cells(2, 1).Value = "Keep headers unchanged. Data starts on row 4."
    $oSheet.Range($oSheet.Cells(2, 1), $oSheet.Cells(2, $nDataCols)).Merge
    $oSheet.Cells(2, 1).Font.Italic = True

    Local $aHeaders = _LV_GetHeaders($hLV)
    For $c = 0 To $nDataCols - 1
        $oSheet.Cells(3, $c + 1).Value = $aHeaders[$c]
        $oSheet.Cells(3, $c + 1).Font.Bold = True
        $oSheet.Cells(3, $c + 1).Interior.Color = 14277081
    Next

    For $r = 0 To $nRows - 1
        For $c = 0 To $nDataCols - 1
            $oSheet.Cells($r + 4, $c + 1).Value = _GUICtrlListView_GetItemText($hLV, $r, $c)
        Next
    Next

    If $nRows > 0 Then
        Local $oRange = $oSheet.Range($oSheet.Cells(3, 1), $oSheet.Cells($nRows + 3, $nDataCols))
        $oRange.Borders.LineStyle = 1
        $oRange.Borders.Color = 15132390
        $oRange.AutoFilter()
    EndIf
    $oSheet.Columns.AutoFit()
    Return $nRows
EndFunc

Func _XLSheetByName($oBook, $sSheetName)
    $g_sLastComError = ""
    Local $oSheet = $oBook.Worksheets($sSheetName)
    If $g_sLastComError <> "" Or Not IsObj($oSheet) Then Return 0
    Return $oSheet
EndFunc

Func _XLValidateHeaders($hLV, $oSheet, $sSheetName)
    Local $aHeaders = _LV_GetHeaders($hLV)
    For $c = 0 To UBound($aHeaders) - 1
        Local $sExpected = _NormalizeHeader($aHeaders[$c])
        Local $sActual = _NormalizeHeader($oSheet.Cells(3, $c + 1).Text)
        If $sExpected <> $sActual Then
            Return "Sheet '" & $sSheetName & "' header mismatch at column " & ($c + 1) & ". Expected '" & $aHeaders[$c] & "', found '" & $oSheet.Cells(3, $c + 1).Text & "'."
        EndIf
    Next
    Return ""
EndFunc

Func _XLRowIsEmpty($oSheet, $iRow, $nCols)
    For $c = 1 To $nCols
        If StringStripWS($oSheet.Cells($iRow, $c).Text, 3) <> "" Then Return False
    Next
    Return True
EndFunc

Func _XLImportWorksheetToLV($hLV, $oSheet)
    Local $nCols = _GUICtrlListView_GetColumnCount($hLV)
    Local $nDataCols = $nCols - 1
    Local $nUsedRows = $oSheet.UsedRange.Rows.Count
    Local $nImported = 0
    _GUICtrlListView_DeleteAllItems($hLV)

    For $r = 4 To $nUsedRows
        If _XLRowIsEmpty($oSheet, $r, $nDataCols) Then ContinueLoop
        Local $sItem = ""
        For $c = 1 To $nDataCols
            If $c > 1 Then $sItem &= "|"
            Local $sCell = $oSheet.Cells($r, $c).Text
            $sItem &= StringReplace($sCell, "|", "/")
        Next
        _LV_AppendDataRow($hLV, $sItem)
        $nImported += 1
    Next
    _LV_Renumber($hLV)
    _LogVerbose("Workbook sheet imported: " & $oSheet.Name & " = " & $nImported & " row(s)")
    Return $nImported
EndFunc

Func _ExportExcel()
    Local $sFile = FileSaveDialog("Export Ortems Toolbox workbook", @WorkingDir, "Excel Workbook (*.xlsx)|All (*.*)", 16, "ortems_toolbox_export.xlsx")
    If Not @error And $sFile <> "" Then _LogVerbose("Workbook export requested: " & $sFile)
    If @error Or $sFile = "" Then Return
    If StringLower(StringRight($sFile, 5)) <> ".xlsx" Then $sFile &= ".xlsx"

    Local $nIssues = _ValidateFKBeforeImport()
    Local $sStatus = ($nIssues = 0 ? "PASS" : "WARNING - " & $nIssues & " integrity issue(s)")

    Local $oExcel = _ExcelCreateApp()
    If Not IsObj($oExcel) Then Return

    $g_sLastComError = ""
    Local $oBook = $oExcel.Workbooks.Add()
    If $g_sLastComError <> "" Or Not IsObj($oBook) Then
        MsgBox(262144+16, "Export workbook", "Could not create an Excel workbook." & @CRLF & $g_sLastComError,0,$g_hMain)
        $oExcel.Quit()
        Return
    EndIf

    While $oBook.Worksheets.Count > 1
        $oBook.Worksheets($oBook.Worksheets.Count).Delete()
    WEnd

    Local $oReadMe = $oBook.Worksheets(1)
    _XLWriteReadMe($oReadMe, $sStatus, $nIssues)

    Local $aMap = _DatasetMap()
    Local $nTotalRows = 0
    For $i = 0 To UBound($aMap) - 1
        Local $oSheet = $oBook.Worksheets.Add(Default, $oBook.Worksheets($oBook.Worksheets.Count))
        $oSheet.Name = $aMap[$i][1]
        $nTotalRows += _LV_ExportToWorksheet($aMap[$i][0], $oSheet, $aMap[$i][3])
    Next

    $oReadMe.Activate()
    $g_sLastComError = ""
    $oBook.SaveAs($sFile, 51)
    Local $sErr = $g_sLastComError
    $oBook.Close(False)
    $oExcel.Quit()

    If $sErr <> "" Then
        MsgBox(262144+16, "Export workbook", "Could not save the workbook." & @CRLF & @CRLF & $sErr,0,$g_hMain)
        Return
    EndIf

    _Log("Workbook export finished: " & $nTotalRows & " data rows -> " & $sFile)
    MsgBox(262144+64, "Export workbook", "Exported " & $nTotalRows & " data row(s) to:" & @CRLF & $sFile & @CRLF & @CRLF & "Integrity at export: " & $sStatus,0,$g_hMain)
EndFunc

Func _ImportExcel()
    Local $sFile = FileOpenDialog("Import Ortems Toolbox workbook", @WorkingDir, "Excel Workbook (*.xlsx;*.xlsm;*.xls)|All (*.*)", 1)
    If Not @error And $sFile <> "" Then _LogVerbose("Workbook import requested: " & $sFile)
    If @error Or $sFile = "" Then Return

    Local $oExcel = _ExcelCreateApp()
    If Not IsObj($oExcel) Then Return

    $g_sLastComError = ""
    Local $oBook = $oExcel.Workbooks.Open($sFile, False, True)
    If $g_sLastComError <> "" Or Not IsObj($oBook) Then
        MsgBox(262144+16, "Import workbook", "Could not open the selected workbook." & @CRLF & @CRLF & $g_sLastComError,0,$g_hMain)
        $oExcel.Quit()
        Return
    EndIf

    Local $aMap = _DatasetMap()
    Local $sErrors = ""
    For $i = 0 To UBound($aMap) - 1
        Local $oSheet = _XLSheetByName($oBook, $aMap[$i][1])
        If Not IsObj($oSheet) Then
            $sErrors &= "Missing sheet: " & $aMap[$i][1] & @CRLF
            ContinueLoop
        EndIf
        Local $sHeaderError = _XLValidateHeaders($aMap[$i][0], $oSheet, $aMap[$i][1])
        If $sHeaderError <> "" Then $sErrors &= $sHeaderError & @CRLF
    Next

    If $sErrors <> "" Then
        $oBook.Close(False)
        $oExcel.Quit()
        MsgBox(262144+16, "Import workbook", "The workbook cannot be imported because its structure does not match the Toolbox export format:" & @CRLF & @CRLF & $sErrors,0,$g_hMain)
        Return
    EndIf

    Local $iAns = MsgBox(262144+4 + 32, "Import workbook", _
        "This will replace the rows currently loaded in all data tabs with the contents of the selected workbook." & @CRLF & @CRLF & _
        "File:" & @CRLF & $sFile & @CRLF & @CRLF & _
        "Continue?",0,$g_hMain)
    If $iAns <> 6 Then
        $oBook.Close(False)
        $oExcel.Quit()
        Return
    EndIf

    Local $nTotalRows = 0
    For $i = 0 To UBound($aMap) - 1
        Local $oSheet = _XLSheetByName($oBook, $aMap[$i][1])
        $nTotalRows += _XLImportWorksheetToLV($aMap[$i][0], $oSheet)
    Next

    $oBook.Close(False)
    $oExcel.Quit()

    Local $nIssues = _ValidateFKBeforeImport()
    _Log("Workbook import finished: " & $nTotalRows & " data rows <- " & $sFile)

    Local $sMsg = "Imported " & $nTotalRows & " data row(s) from:" & @CRLF & $sFile
    If $nIssues > 0 Then
        $sMsg &= @CRLF & @CRLF & "Integrity check found " & $nIssues & " issue(s). Review the execution log or click Integrity Check to fix simple references."
        MsgBox(262144+48, "Import workbook", $sMsg,0,$g_hMain)
    Else
        $sMsg &= @CRLF & @CRLF & "Integrity check passed."
        MsgBox(262144+64, "Import workbook", $sMsg,0,$g_hMain)
    EndIf
EndFunc

;=============================================================================
; GENERIC CSV IMPORT / EXPORT FOR A LISTVIEW
;=============================================================================
Func _CSVEscape($s)
    If StringInStr($s, ";") Or StringInStr($s, '"') Or StringInStr($s, @CR) Or StringInStr($s, @LF) Then
        $s = StringReplace($s, '"', '""')
        Return '"' & $s & '"'
    EndIf
    Return $s
EndFunc

Func _CSVSplit($sLine, $nCols)
    Local $aResult[$nCols]
    For $k = 0 To $nCols - 1
        $aResult[$k] = ""
    Next

    Local $iCol = 0
    Local $sCur = ""
    Local $bInQuote = False
    Local $iLen = StringLen($sLine)
    Local $i = 1
    While $i <= $iLen
        Local $ch = StringMid($sLine, $i, 1)
        If $bInQuote Then
            If $ch = '"' Then
                If $i < $iLen And StringMid($sLine, $i + 1, 1) = '"' Then
                    $sCur &= '"'
                    $i += 1
                Else
                    $bInQuote = False
                EndIf
            Else
                $sCur &= $ch
            EndIf
        Else
            If $ch = '"' Then
                $bInQuote = True
            ElseIf $ch = ';' Then
                If $iCol < $nCols Then $aResult[$iCol] = $sCur
                $iCol += 1
                $sCur = ""
            Else
                $sCur &= $ch
            EndIf
        EndIf
        $i += 1
    WEnd
    If $iCol < $nCols Then $aResult[$iCol] = $sCur
    Return $aResult
EndFunc

; Export a ListView to a CSV file. Returns row count (>=0) or -1 on error.
Func _LV_ExportCSVToFile($hLV, $sFile)
    Local $nCols = _GUICtrlListView_GetColumnCount($hLV)
    Local $nDataCols = $nCols - 1
    Local $nRows = _GUICtrlListView_GetItemCount($hLV)

    Local $sCSV = ""
    ; Header
    For $c = 0 To $nDataCols - 1
        If $c > 0 Then $sCSV &= ";"
        Local $aCol = _GUICtrlListView_GetColumn($hLV, $c)
        Local $sHdr = ""
        If IsArray($aCol) And UBound($aCol) > 5 Then $sHdr = $aCol[5]
        $sCSV &= _CSVEscape($sHdr)
    Next
    $sCSV &= @CRLF

    ; Rows
    For $i = 0 To $nRows - 1
        For $c = 0 To $nDataCols - 1
            If $c > 0 Then $sCSV &= ";"
            $sCSV &= _CSVEscape(_GUICtrlListView_GetItemText($hLV, $i, $c))
        Next
        $sCSV &= @CRLF
    Next

    Local $hFile = FileOpen($sFile, 2 + 8)  ; erase + create
    If $hFile = -1 Then
        _Log("Export CSV failed (cannot open): " & $sFile)
        Return -1
    EndIf
    FileWrite($hFile, $sCSV)
    FileClose($hFile)
    Return $nRows
EndFunc

; Import a CSV file into a ListView. If $bReplace=True, clears existing rows first.
; Returns row count imported, or -1 on error.
Func _LV_ImportCSVFromFile($hLV, $sFile, $bReplace)
    If Not FileExists($sFile) Then Return -1
    Local $hFile = FileOpen($sFile, 0)
    If $hFile = -1 Then Return -1

    Local $nCols = _GUICtrlListView_GetColumnCount($hLV)
    Local $nDataCols = $nCols - 1
    If $bReplace Then _GUICtrlListView_DeleteAllItems($hLV)

    Local $nImported = 0
    Local $bFirst = True
    While 1
        Local $sLine = FileReadLine($hFile)
        If @error Then ExitLoop
        If StringStripWS($sLine, 3) = "" Then ContinueLoop
        If $bFirst Then
            $bFirst = False
            ContinueLoop  ; skip header
        EndIf
        Local $aFields = _CSVSplit($sLine, $nDataCols)
        Local $sItem = ""
        For $c = 0 To $nDataCols - 1
            If $c > 0 Then $sItem &= "|"
            ; Replace any pipe in field to avoid breaking ListView separator
            $sItem &= StringReplace($aFields[$c], "|", "/")
        Next
        _LV_AppendDataRow($hLV, $sItem)
        $nImported += 1
    WEnd
    FileClose($hFile)
    _Log("Imported " & $nImported & " rows from " & $sFile)
    Return $nImported
EndFunc

; Interactive single-tab CSV export: asks for filename and calls export helper.
Func _LV_ExportCSVInteractive($hLV, $sDefaultName)
    Local $sFile = FileSaveDialog("Export " & $sDefaultName & " to CSV", @WorkingDir, _
        "CSV (*.csv)", 16, $sDefaultName & ".csv")
    If @error Or $sFile = "" Then Return
    If StringRight($sFile, 4) <> ".csv" Then $sFile &= ".csv"
    Local $n = _LV_ExportCSVToFile($hLV, $sFile)
    If $n < 0 Then
        MsgBox(262144+16, "Export CSV", "Failed to write file: " & $sFile,0,$g_hMain)
    Else
        MsgBox(262144+64, "Export CSV", "Exported " & $n & " rows to:" & @CRLF & $sFile,0,$g_hMain)
    EndIf
EndFunc

; Interactive single-tab CSV import: asks for file, asks replace/append, calls import helper.
Func _LV_ImportCSVInteractive($hLV, $sName)
    Local $sFile = FileOpenDialog("Import CSV for " & $sName, @WorkingDir, "CSV (*.csv)|All (*.*)", 1)
    If @error Or $sFile = "" Then Return
    Local $iAns = MsgBox(262144+4 + 32 + 3, "Import CSV", _
        "Replace existing rows?" & @CRLF & @CRLF & _
        "Yes = replace all rows in this tab" & @CRLF & _
        "No  = append to existing rows" & @CRLF & _
        "Cancel = abort",0,$g_hMain)
    If $iAns = 2 Then Return  ; cancel
    Local $bReplace = ($iAns = 6)
    Local $n = _LV_ImportCSVFromFile($hLV, $sFile, $bReplace)
    If $n < 0 Then
        MsgBox(262144+16, "Import CSV", "Failed to read file: " & $sFile,0,$g_hMain)
    Else
        MsgBox(262144+64, "Import CSV", "Imported " & $n & " rows from:" & @CRLF & $sFile,0,$g_hMain)
    EndIf
EndFunc

;=============================================================================
; PER-TAB CSV IMPORT / EXPORT WRAPPERS
;=============================================================================
Func _Cal_ImpCSV()
    _LV_ImportCSVInteractive($g_hLV_Cal,  "calendars")
EndFunc
Func _Cal_ExpCSV()
    _LV_ExportCSVInteractive($g_hLV_Cal,  "calendars")
EndFunc
Func _Mach_ImpCSV()
    _LV_ImportCSVInteractive($g_hLV_Mach, "machines")
EndFunc
Func _Mach_ExpCSV()
    _LV_ExportCSVInteractive($g_hLV_Mach, "machines")
EndFunc
Func _Ops_ImpCSV()
    _LV_ImportCSVInteractive($g_hLV_Ops,  "operations")
EndFunc
Func _Ops_ExpCSV()
    _LV_ExportCSVInteractive($g_hLV_Ops,  "operations")
EndFunc
Func _Rout_ImpCSV()
    _LV_ImportCSVInteractive($g_hLV_Rout, "routings")
EndFunc
Func _Rout_ExpCSV()
    _LV_ExportCSVInteractive($g_hLV_Rout, "routings")
EndFunc
Func _Mat_ImpCSV()
    _LV_ImportCSVInteractive($g_hLV_Mat,  "items")
EndFunc
Func _Mat_ExpCSV()
    _LV_ExportCSVInteractive($g_hLV_Mat,  "items")
EndFunc
Func _BOM_ImpCSV()
    _LV_ImportCSVInteractive($g_hLV_BOM,  "bom")
EndFunc
Func _BOM_ExpCSV()
    _LV_ExportCSVInteractive($g_hLV_BOM,  "bom")
EndFunc
Func _WO_ImpCSV()
    _LV_ImportCSVInteractive($g_hLV_WO,   "workorders")
EndFunc
Func _WO_ExpCSV()
    _LV_ExportCSVInteractive($g_hLV_WO,   "workorders")
EndFunc
Func _WOL_ImpCSV()
    _LV_ImportCSVInteractive($g_hLV_WOL,  "wo_links")
EndFunc
Func _WOL_ExpCSV()
    _LV_ExportCSVInteractive($g_hLV_WOL,  "wo_links")
EndFunc
Func _SR_ImpCSV()
    _LV_ImportCSVInteractive($g_hLV_SR,   "secondary_resources")
EndFunc
Func _SR_ExpCSV()
    _LV_ExportCSVInteractive($g_hLV_SR,   "secondary_resources")
EndFunc
Func _Cap_ImpCSV()
    _LV_ImportCSVInteractive($g_hLV_Cap,  "capacity")
EndFunc
Func _Cap_ExpCSV()
    _LV_ExportCSVInteractive($g_hLV_Cap,  "capacity")
EndFunc
Func _Stk_ImpCSV()
    _LV_ImportCSVInteractive($g_hLV_Stk,  "stock")
EndFunc
Func _Stk_ExpCSV()
    _LV_ExportCSVInteractive($g_hLV_Stk,  "stock")
EndFunc


;=============================================================================
; LOAD FROM DB - reads Ortems tables and populates all data tabs
;=============================================================================
Func _LoadFromDB()
    If Not $g_bConnected Then
        MsgBox(262144+48, "Not connected", "Connect to a database first (tab 1. Database).",0,$g_hMain)
        Return
    EndIf

    Local $nRet = MsgBox(262144+4 + 64, "Load from DB", _
        "This will CLEAR all data in the tabs and reload from the database:" & @CRLF & @CRLF & _
        "  " & $g_sDatabase & " @ " & $g_sServer & @CRLF & @CRLF & _
        "Continue?",0,$g_hMain)
    If $nRet <> 6 Then Return

    Local $oConn = ObjCreate("ADODB.Connection")
    If Not IsObj($oConn) Then
        _Log("ERROR: Could not create ADODB.Connection.")
        Return
    EndIf
    $g_sLastComError = ""
    $oConn.Open($g_sConnStr)
    If $g_sLastComError <> "" Or $oConn.State <> 1 Then
        _Log("ERROR opening connection for load: " & $g_sLastComError)
        MsgBox(262144+16, "Error", "Could not open connection: " & @CRLF & $g_sLastComError,0,$g_hMain)
        Return
    EndIf

    _Log("=== Load from DB started ===")

    Local $nTabs = 0
    $nTabs += _DB_LoadCalendars($oConn)
    $nTabs += _DB_LoadMachines($oConn)
    $nTabs += _DB_LoadOperations($oConn)
    $nTabs += _DB_LoadRoutings($oConn)
    $nTabs += _DB_LoadMaterials($oConn)
    $nTabs += _DB_LoadBOM($oConn)
    $nTabs += _DB_LoadWO($oConn)
    If $g_bModWOL Then $nTabs += _DB_LoadWOLinks($oConn)

    $oConn.Close()
    _Log("=== Load from DB complete - " & $nTabs & " tab(s) populated ===")
    MsgBox(262144+64, "Load complete", "Data loaded from database successfully." & @CRLF & @CRLF & _
        $nTabs & " tab(s) were populated." & @CRLF & "Check the execution log for row counts.",0,$g_hMain)
EndFunc

; Helper: open a recordset, return 0 on error
Func _DBQuery($oConn, $sSQL)
    _LogVerboseSQL("Load from DB query", $sSQL)
    $g_sLastComError = ""
    Local $oRS = $oConn.Execute($sSQL)
    If $g_sLastComError <> "" Then
        _Log("DB query error: " & $g_sLastComError & " | SQL: " & StringLeft($sSQL, 100))
        Return 0
    EndIf
    Return $oRS
EndFunc

; Format DB time back to "HH:MM".
; Supports text values already stored as HH:MM, numeric hour values (6 => 06:00),
; numeric HHMM values (600 => 06:00), and Excel-style decimals.
Func _IntToTime($nVal)
    Return _NormalizeTimeText($nVal)
EndFunc

;------------------------------------------------------------------------------
; CALENDARS  - B_CAL joined to B_PERI
;------------------------------------------------------------------------------
Func _DB_LoadCalendars($oConn)
    _GUICtrlListView_DeleteAllItems($g_hLV_Cal)
    Local $oRS = _DBQuery($oConn, _
        "SELECT c.NOCALHEBD, c.NOMCAL, p.NOJOUR_DEB, p.DEB_PERIO, p.NOJOUR_FIN, p.FIN_PERIO " & _
        "FROM B_CAL c " & _
        "JOIN B_PERI p ON c.NOCALHEBD = p.NOCALHEBD " & _
        "ORDER BY c.NOCALHEBD, p.NOJOUR_DEB, p.DEB_PERIO")
    If Not IsObj($oRS) Then Return 0

    Local $nRows = 0
    While Not $oRS.EOF
        Local $sID    = $oRS.Fields("NOCALHEBD").Value
        Local $sNome  = $oRS.Fields("NOMCAL").Value
        Local $sDiaI  = $oRS.Fields("NOJOUR_DEB").Value
        Local $sHoraI = _IntToTime($oRS.Fields("DEB_PERIO").Value)
        Local $sDiaF  = $oRS.Fields("NOJOUR_FIN").Value
        Local $sHoraF = _IntToTime($oRS.Fields("FIN_PERIO").Value)
        _LV_AppendDataRow($g_hLV_Cal, $sID & "|" & $sNome & "|" & $sDiaI & "|" & $sHoraI & "|" & $sDiaF & "|" & $sHoraF)
        $nRows += 1
        $oRS.MoveNext()
    WEnd
    $oRS.Close()
    _LV_Renumber($g_hLV_Cal)
    _Log("Calendars loaded: " & $nRows & " rows")
    Return 1
EndFunc

;------------------------------------------------------------------------------
; MACHINES  - B_MACH joined to B_ILOT, B_SECT, B_ZONE
;------------------------------------------------------------------------------
Func _DB_LoadMachines($oConn)
    _GUICtrlListView_DeleteAllItems($g_hLV_Mach)
    Local $oRS = _DBQuery($oConn, _
        "SELECT ISNULL(z.NOZONE,'') AS NOZONE, ISNULL(z.LIBZONE,'') AS LIBZONE, " & _
        "       ISNULL(i.ILOT,'') AS ILOT, ISNULL(i.LIBILOT,'') AS LIBILOT, " & _
        "       ISNULL(CAST(i.TYPEILOT AS VARCHAR),'') AS TYPEILOT, " & _
        "       ISNULL(s.CODESECTI,'') AS CODESECTI, ISNULL(s.DESIGSECT,'') AS DESIGSECT, " & _
        "       m.MACHINE, ISNULL(m.LIBMACH,'') AS LIBMACH, " & _
        "       ISNULL(m.MACH_MODEMACH,'') AS MACH_MODEMACH, " & _
        "       ISNULL(m.NOCALHEBD,'') AS NOCALHEBD " & _
        "FROM B_MACH m " & _
        "LEFT JOIN B_ILOT i ON m.ILOT = i.ILOT " & _
        "LEFT JOIN B_SECT s ON m.CODESECTI = s.CODESECTI " & _
        "LEFT JOIN B_ZONE z ON m.NOZONE = z.NOZONE " & _
        "ORDER BY m.MACHINE")
    If Not IsObj($oRS) Then Return 0

    Local $nRows = 0
    While Not $oRS.EOF
        _LV_AppendDataRow($g_hLV_Mach, _
            $oRS.Fields("NOZONE").Value & "|" & _
            $oRS.Fields("LIBZONE").Value & "|" & _
            $oRS.Fields("ILOT").Value & "|" & _
            $oRS.Fields("LIBILOT").Value & "|" & _
            $oRS.Fields("TYPEILOT").Value & "|" & _
            $oRS.Fields("CODESECTI").Value & "|" & _
            $oRS.Fields("DESIGSECT").Value & "|" & _
            $oRS.Fields("MACHINE").Value & "|" & _
            $oRS.Fields("LIBMACH").Value & "|" & _
            $oRS.Fields("MACH_MODEMACH").Value & "|" & _
            $oRS.Fields("NOCALHEBD").Value & "|")
        $nRows += 1
        $oRS.MoveNext()
    WEnd
    $oRS.Close()
    _LV_Renumber($g_hLV_Mach)
    _Log("Machines loaded: " & $nRows & " rows")
    Return 1
EndFunc

;------------------------------------------------------------------------------
; OPERATIONS  - B_OPE joined to B_CADE (first machine per operation)
;------------------------------------------------------------------------------
Func _DB_LoadOperations($oConn)
    _GUICtrlListView_DeleteAllItems($g_hLV_Ops)
    Local $oRS = _DBQuery($oConn, _
        "SELECT o.OPE, ISNULL(o.LIBOP,'') AS LIBOP, ISNULL(o.ILOT,'') AS ILOT, " & _
        "       ISNULL(c.MACHINE,'') AS MACHINE, " & _
        "       ISNULL(CAST(o.CODEBASET AS VARCHAR),'0') AS CODEBASET, " & _
        "       ISNULL(CAST(c.CADE_DURREAL AS VARCHAR),'0') AS CADE_DURREAL, " & _
        "       ISNULL(o.UNITE,'H') AS UNITE, " & _
        "       ISNULL(CAST(o.DURPREP AS VARCHAR),'0') AS DURPREP, " & _
        "       ISNULL(CAST(o.THM AS VARCHAR),'0') AS THM, " & _
        "       ISNULL(CAST(o.INTERUPT AS VARCHAR),'0') AS INTERUPT " & _
        "FROM B_OPE o " & _
        "LEFT JOIN (SELECT OPE, MIN(MACHINE) AS MACHINE, MIN(CADE_DURREAL) AS CADE_DURREAL FROM B_CADE GROUP BY OPE) c ON o.OPE = c.OPE " & _
        "ORDER BY o.OPE")
    If Not IsObj($oRS) Then Return 0

    Local $nRows = 0
    While Not $oRS.EOF
        _LV_AppendDataRow($g_hLV_Ops, _
            $oRS.Fields("OPE").Value & "|" & _
            $oRS.Fields("LIBOP").Value & "|" & _
            $oRS.Fields("ILOT").Value & "|" & _
            $oRS.Fields("MACHINE").Value & "|" & _
            $oRS.Fields("CODEBASET").Value & "|" & _
            $oRS.Fields("CADE_DURREAL").Value & "|" & _
            $oRS.Fields("UNITE").Value & "|" & _
            $oRS.Fields("DURPREP").Value & "|" & _
            $oRS.Fields("THM").Value & "|" & _
            $oRS.Fields("INTERUPT").Value)
        $nRows += 1
        $oRS.MoveNext()
    WEnd
    $oRS.Close()
    _LV_Renumber($g_hLV_Ops)
    _Log("Operations loaded: " & $nRows & " rows")
    Return 1
EndFunc

;------------------------------------------------------------------------------
; ROUTINGS  - B_GAMM joined to B_PHAS
;------------------------------------------------------------------------------
Func _DB_LoadRoutings($oConn)
    _GUICtrlListView_DeleteAllItems($g_hLV_Rout)
    Local $oRS = _DBQuery($oConn, _
        "SELECT g.NOMG, ISNULL(g.LIBGAM,'') AS LIBGAM, " & _
        "       p.NOPHASE, ISNULL(p.OPE,'') AS OPE, ISNULL(p.LIBPHASE,'') AS LIBPHASE " & _
        "FROM B_GAMM g " & _
        "JOIN B_PHAS p ON g.NOMG = p.NOMG " & _
        "ORDER BY g.NOMG, p.NOPHASE")
    If Not IsObj($oRS) Then Return 0

    Local $nRows = 0
    While Not $oRS.EOF
        _LV_AppendDataRow($g_hLV_Rout, _
            $oRS.Fields("NOMG").Value & "|" & _
            $oRS.Fields("LIBGAM").Value & "|" & _
            $oRS.Fields("NOPHASE").Value & "|" & _
            $oRS.Fields("OPE").Value & "|" & _
            $oRS.Fields("LIBPHASE").Value)
        $nRows += 1
        $oRS.MoveNext()
    WEnd
    $oRS.Close()
    _LV_Renumber($g_hLV_Rout)
    _Log("Routings loaded: " & $nRows & " rows")
    Return 1
EndFunc

;------------------------------------------------------------------------------
; MATERIALS  - B_ART joined to B_VER_ART (first active version per item)
;------------------------------------------------------------------------------
Func _DB_LoadMaterials($oConn)
    _GUICtrlListView_DeleteAllItems($g_hLV_Mat)
    Local $oRS = _DBQuery($oConn, _
        "SELECT a.CODEARTIC, ISNULL(a.LIBARTIC,'') AS LIBARTIC, " & _
        "       ISNULL(a.TYPEMATI,'') AS TYPEMATI, " & _
        "       ISNULL(v.VER_ART,'') AS VER_ART, " & _
        "       ISNULL(v.NOMG,'') AS NOMG, " & _
        "       ISNULL(CAST(a.QTE_STOCK AS VARCHAR),'0') AS QTE_STOCK " & _
        "FROM B_ART a " & _
        "LEFT JOIN (SELECT CODEARTIC, MIN(VER_ART) AS VER_ART, MIN(NOMG) AS NOMG FROM B_VER_ART GROUP BY CODEARTIC) v " & _
        "ON a.CODEARTIC = v.CODEARTIC " & _
        "ORDER BY a.CODEARTIC")
    If Not IsObj($oRS) Then Return 0

    Local $nRows = 0
    While Not $oRS.EOF
        Local $sType = StringStripWS(String($oRS.Fields("TYPEMATI").Value), 3)
        Local $sVer = String($oRS.Fields("VER_ART").Value)
        Local $sRout = String($oRS.Fields("NOMG").Value)
        If _IsRawMaterialType($sType) Then
            If StringStripWS($sVer, 3) <> "" Or StringStripWS($sRout, 3) <> "" Then _LogVerbose("Load from DB: ignored B_VER_ART version/routing for MP item '" & $oRS.Fields("CODEARTIC").Value & "'.")
            $sVer = ""
            $sRout = ""
        EndIf
        _LV_AppendDataRow($g_hLV_Mat, _
            $oRS.Fields("CODEARTIC").Value & "|" & _
            $oRS.Fields("LIBARTIC").Value & "|" & _
            $sType & "|" & _
            $sVer & "|" & _
            $sRout & "|" & _
            $oRS.Fields("QTE_STOCK").Value)
        $nRows += 1
        $oRS.MoveNext()
    WEnd
    $oRS.Close()
    _LV_Renumber($g_hLV_Mat)
    _Log("Materials loaded: " & $nRows & " rows")
    Return 1
EndFunc

;------------------------------------------------------------------------------
; BOM  - B_NOME (quantity columns detected at runtime)
;------------------------------------------------------------------------------
Func _DB_LoadBOM($oConn)
    _GUICtrlListView_DeleteAllItems($g_hLV_BOM)
    Local $sQRef = _DetectColumn("B_NOME", "qty_base", "QTEREF;QTEBASET;CODEBASET;QTE_BASE;QTBASE;QTE1")
    Local $sQNec = _DetectColumn("B_NOME", "qty_nec",  "QTENEC;QTEBESO;BES_CODEARTIC;QTE_BES;QTBESO;QTE2")

    Local $oRS = _DBQuery($oConn, _
        "SELECT B_V_CODEARTIC, ISNULL(VER_ART,'') AS VER_ART, " & _
        "       ISNULL(CODEARTIC,'') AS CODEARTIC, ISNULL(NOMG,'') AS NOMG, " & _
        "       ISNULL(CAST(NOPHASE AS VARCHAR),'0') AS NOPHASE, " & _
        "       ISNULL(CAST([" & $sQRef & "] AS VARCHAR),'0') AS QTYREF, " & _
        "       ISNULL(CAST([" & $sQNec & "] AS VARCHAR),'0') AS QTYNEC " & _
        "FROM B_NOME " & _
        "ORDER BY B_V_CODEARTIC, VER_ART, NOPHASE, CODEARTIC")
    If Not IsObj($oRS) Then Return 0

    Local $nRows = 0
    While Not $oRS.EOF
        _LV_AppendDataRow($g_hLV_BOM, _
            $oRS.Fields("B_V_CODEARTIC").Value & "|" & _
            $oRS.Fields("VER_ART").Value & "|" & _
            $oRS.Fields("CODEARTIC").Value & "|" & _
            $oRS.Fields("NOMG").Value & "|" & _
            $oRS.Fields("NOPHASE").Value & "|" & _
            $oRS.Fields("QTYREF").Value & "|" & _
            $oRS.Fields("QTYNEC").Value)
        $nRows += 1
        $oRS.MoveNext()
    WEnd
    $oRS.Close()
    _LV_Renumber($g_hLV_BOM)
    _Log("BOM loaded: " & $nRows & " rows")
    Return 1
EndFunc

;------------------------------------------------------------------------------
; WORK ORDERS  - B_OF (QTE column detected at runtime)
;------------------------------------------------------------------------------
Func _DB_LoadWO($oConn)
    _GUICtrlListView_DeleteAllItems($g_hLV_WO)
    Local $sQte = _DetectColumn("B_OF", "qty_order", "QTE;QTEORDER;QTEOF;QTE_ORDER;QTEOFORDER;QTECMD")

    Local $oRS = _DBQuery($oConn, _
        "SELECT NOF, ISNULL(CODEARTIC,'') AS CODEARTIC, ISNULL(NOMG,'') AS NOMG, " & _
        "       ISNULL(VER_ART,'') AS VER_ART, " & _
        "       ISNULL(CAST([" & $sQte & "] AS VARCHAR),'0') AS QTE, " & _
        "       ISNULL(CONVERT(VARCHAR,DPLUSTOT,103),'') AS DT_I, " & _
        "       ISNULL(CONVERT(VARCHAR,FPLUSTARD,103),'') AS DT_F " & _
        "FROM B_OF " & _
        "ORDER BY NOF")
    If Not IsObj($oRS) Then Return 0

    Local $nRows = 0
    While Not $oRS.EOF
        ; Append HH:MM to date if not already present
        Local $sDtI = $oRS.Fields("DT_I").Value
        Local $sDtF = $oRS.Fields("DT_F").Value
        If StringLen($sDtI) = 10 Then $sDtI &= " 00:00"
        If StringLen($sDtF) = 10 Then $sDtF &= " 23:59"
        _LV_AppendDataRow($g_hLV_WO, _
            $oRS.Fields("NOF").Value & "|" & _
            $oRS.Fields("CODEARTIC").Value & "|" & _
            $oRS.Fields("NOMG").Value & "|" & _
            $oRS.Fields("VER_ART").Value & "|" & _
            $oRS.Fields("QTE").Value & "|" & _
            $sDtI & "|" & $sDtF)
        $nRows += 1
        $oRS.MoveNext()
    WEnd
    $oRS.Close()
    _LV_Renumber($g_hLV_WO)
    _Log("Work Orders loaded: " & $nRows & " rows")
    Return 1
EndFunc

;------------------------------------------------------------------------------
; WO LINKS  - B_PROF
;------------------------------------------------------------------------------
Func _DB_LoadWOLinks($oConn)
    _GUICtrlListView_DeleteAllItems($g_hLV_WOL)
    Local $oRS = _DBQuery($oConn, _
        "SELECT ISNULL(B_O_NOF,'') AS B_O_NOF, ISNULL(B_P_NOMG,'') AS B_P_NOMG, " & _
        "       ISNULL(CAST(B_P_NOPHASE AS VARCHAR),'') AS B_P_NOPHASE, " & _
        "       ISNULL(NOF,'') AS NOF, ISNULL(NOMG,'') AS NOMG, " & _
        "       ISNULL(CAST(NOPHASE AS VARCHAR),'') AS NOPHASE, " & _
        "       ISNULL(PROF_TYPEPREC,'FS') AS PROF_TYPEPREC " & _
        "FROM B_PROF " & _
        "ORDER BY NOF, NOPHASE")
    If Not IsObj($oRS) Then Return 0

    Local $nRows = 0
    While Not $oRS.EOF
        _LV_AppendDataRow($g_hLV_WOL, _
            $oRS.Fields("B_O_NOF").Value & "|" & _
            $oRS.Fields("B_P_NOMG").Value & "|" & _
            $oRS.Fields("B_P_NOPHASE").Value & "|" & _
            $oRS.Fields("NOF").Value & "|" & _
            $oRS.Fields("NOMG").Value & "|" & _
            $oRS.Fields("NOPHASE").Value & "|" & _
            $oRS.Fields("PROF_TYPEPREC").Value)
        $nRows += 1
        $oRS.MoveNext()
    WEnd
    $oRS.Close()
    _LV_Renumber($g_hLV_WOL)
    _Log("WO Links loaded: " & $nRows & " rows")
    Return 1
EndFunc

Func _SaveSQL()
    Local $sSQL = GUICtrlRead($g_hLog)
    If $sSQL = "" Then
        MsgBox(262144+48, "Warning", "No SQL to save. Click 'GENERATE SQL' first.",0,$g_hMain)
        Return
    EndIf
    Local $sFile = FileSaveDialog("Save SQL", @WorkingDir, "SQL (*.sql)|All (*.*)", 16, "ortems_demo.sql")
    If $sFile <> "" Then
        FileWrite($sFile, $sSQL)
        _Log("SQL saved to: " & $sFile)
        MsgBox(262144+64, "Saved", "SQL saved successfully to:" & @CRLF & $sFile,0,$g_hMain)
    EndIf
EndFunc

;=============================================================================
; LOG
;=============================================================================
Func _EnsureLogFile()
    If Not FileExists($g_sLogDir) Then DirCreate($g_sLogDir)
    If $g_sLogFile = "" Then
        $g_sLogFile = $g_sLogDir & "\Toolbox_" & @YEAR & @MON & @MDAY & "_" & @HOUR & @MIN & @SEC & ".log"
    EndIf
    Return $g_sLogFile
EndFunc

Func _Log($sMsg, $bErrorOnly = False)
    Local $sTimestamp = "[" & @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC & "]"
    Local $sLine = $sTimestamp & " " & $sMsg

    ; Always write to the on-screen exec log (newest entry on top) when available.
    If Not $bErrorOnly And IsDeclared("g_hExecLog") And $g_hExecLog <> 0 Then
        Local $sCurrent = GUICtrlRead($g_hExecLog)
        GUICtrlSetData($g_hExecLog, $sLine & @CRLF & $sCurrent)
    EndIf

    ; Always append to the session log file under .\log.
    Local $sLogFile = _EnsureLogFile()
    Local $hFile = FileOpen($sLogFile, 1)   ; 1 = append
    If $hFile <> -1 Then
        FileWriteLine($hFile, $sLine)
        FileClose($hFile)
    EndIf
EndFunc

Func _LogVerbose($sMsg)
    If Not $g_bVerbose Then Return
    _Log("[VERBOSE] " & $sMsg)
EndFunc

Func _LogVerboseSQL($sContext, $sSQL)
    ; Verbose SQL tracing is written to the session log file only, not to the UI execution box.
    ; This keeps the UI responsive while still preserving the exact queries used for troubleshooting.
    If Not $g_bVerbose Then Return
    Local $sClean = StringStripWS(String($sSQL), 3)
    If $sClean = "" Then Return
    $sClean = StringReplace($sClean, @CRLF, " | ")
    $sClean = StringReplace($sClean, @LF, " | ")
    $sClean = StringReplace($sClean, @CR, " | ")
    _Log("[VERBOSE][SQL] " & $sContext & ": " & $sClean, True)
EndFunc

Func _MaskConnectionString($sConn)
    Local $sMasked = StringRegExpReplace($sConn, "(?i)PWD=[^;]*", "PWD=****")
    $sMasked = StringRegExpReplace($sMasked, "(?i)Password=[^;]*", "Password=****")
    Return $sMasked
EndFunc

; Create one log file per application opening.
Func _LogSessionStart()
    Local $sLogFile = _EnsureLogFile()
    Local $hFile = FileOpen($sLogFile, 2) ; overwrite/create for this session
    If $hFile <> -1 Then
        FileWriteLine($hFile, "========== SESSION STARTED: " & @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC & " ==========")
        FileWriteLine($hFile, "Application: " & $TITLE)
        FileWriteLine($hFile, "Version: " & FileGetVersion(@ScriptFullPath))
        FileWriteLine($hFile, "Script path: " & @ScriptFullPath)
        FileWriteLine($hFile, "Log file: " & $sLogFile)
        FileWriteLine($hFile, "Verbose Mode: " & ($g_bVerbose ? "ON" : "OFF"))
        FileWriteLine($hFile, "Computer/User: " & @ComputerName & " / " & @UserName)
        FileWriteLine($hFile, "OS: " & @OSVersion & " " & @OSArch)
        FileWriteLine($hFile, "============================================================")
        FileClose($hFile)
    EndIf
    _Log("Session log file created: " & $sLogFile)
EndFunc

;=============================================================================
; FECHAR
;=============================================================================

;=============================================================================
; UI: MODULE SELECTION -> ENABLE/DISABLE TABS (real-time)
;=============================================================================
Func _OnModuleSelectionChanged()
;~     _RefreshModuleFlags(True)
EndFunc

Func _RefreshModuleFlags($bSilent = True)
    ; Read UI state
    If IsDeclared("g_rbPS") Then
        $g_bModPS = (GUICtrlRead($g_rbPS) = $GUI_CHECKED)
        $g_bModMP = Not $g_bModPS
    EndIf

    If IsDeclared("g_chkSRP") Then $g_bModSRP   = (GUICtrlRead($g_chkSRP)   = $GUI_CHECKED)
    If IsDeclared("g_chkWOL") Then $g_bModWOL   = (GUICtrlRead($g_chkWOL)   = $GUI_CHECKED)
    If IsDeclared("g_chkSR")  Then $g_bModSR    = (GUICtrlRead($g_chkSR)    = $GUI_CHECKED)
    If IsDeclared("g_chkINV") Then $g_bModINV   = (GUICtrlRead($g_chkINV)   = $GUI_CHECKED)
    If IsDeclared("g_chkLR")  Then $g_bModLR    = (GUICtrlRead($g_chkLR)    = $GUI_CHECKED)
    If IsDeclared("g_chkPRM") Then $g_bModPRM   = (GUICtrlRead($g_chkPRM)   = $GUI_CHECKED)
    If IsDeclared("g_chkMRK") Then $g_bModMRK   = (GUICtrlRead($g_chkMRK)   = $GUI_CHECKED)
    If IsDeclared("g_chkBATCH") Then $g_bModBATCH = (GUICtrlRead($g_chkBATCH) = $GUI_CHECKED)
    If IsDeclared("g_chkCROUT") Then $g_bModCROUT = (GUICtrlRead($g_chkCROUT) = $GUI_CHECKED)

    ; Enforce compatibility: WO Links + Secondary Resources require PS
    If $g_bModMP Then
        If IsDeclared("g_chkWOL") Then
            GUICtrlSetState($g_chkWOL, $GUI_UNCHECKED)
            GUICtrlSetState($g_chkWOL, $GUI_DISABLE)
        EndIf
        If IsDeclared("g_chkSR") Then
            GUICtrlSetState($g_chkSR, $GUI_UNCHECKED)
            GUICtrlSetState($g_chkSR, $GUI_DISABLE)
        EndIf
        $g_bModWOL = False
        $g_bModSR  = False
    Else
        If IsDeclared("g_chkWOL") Then GUICtrlSetState($g_chkWOL, $GUI_ENABLE)
        If IsDeclared("g_chkSR") Then GUICtrlSetState($g_chkSR, $GUI_ENABLE)
    EndIf

    _UpdateTabAvailability()

    _SaveSettings()   ; persist module selections automatically

    If Not $bSilent Then
        Local $sMode = ($g_bModPS ? "PS" : "MP")
        Local $sMods = ""
        If $g_bModSRP Then $sMods &= " SRP"
        If $g_bModWOL Then $sMods &= " WO-LINKS"
        If $g_bModSR  Then $sMods &= " SEC-RES"
        If $g_bModINV Then $sMods &= " INVENTORY"
        If $g_bModLR  Then $sMods &= " LIM-RES"
        If $g_bModPRM Then $sMods &= " PARAMETERS"
        If $g_bModBATCH Then $sMods &= " BATCH"
        If $g_bModCROUT Then $sMods &= " COMPLEX-ROUT"
        If $sMods = "" Then $sMods = " (none)"

        MsgBox(262144+64, "Modules applied", "Mode: " & $sMode & @CRLF & "Enabled modules:" & $sMods & @CRLF & @CRLF & _
            "The data tabs have been enabled/disabled according to your selections.",0,$g_hMain)
    EndIf
EndFunc

Func _UpdateTabAvailability()
    If $g_hTabHandle = 0 Then Return

    Local $hTab = $g_hTabHandle
    Local $bPS = $g_bModPS
    Local $bMP = $g_bModMP

    ; Core tabs (always enabled)
    Local $bCal  = True
    Local $bMach = True
    Local $bMat  = True

    ; Optional / mode-dependent tabs
    Local $bOps  = $bPS
    Local $bRout = $bPS
    Local $bBOM  = ($g_bModSRP Or $bMP)
    Local $bWO   = $bPS
    Local $bWOL  = ($bPS And $g_bModWOL)
    Local $bSR   = ($bPS And $g_bModSR)
    Local $bCap  = ($g_bModLR Or $bMP)
    Local $bStk  = $g_bModINV

    _GUICtrlTab_EnableTab($hTab, $g_iTabCal,  $bCal)
    _GUICtrlTab_EnableTab($hTab, $g_iTabMach, $bMach)
    _GUICtrlTab_EnableTab($hTab, $g_iTabOps,  $bOps)
    _GUICtrlTab_EnableTab($hTab, $g_iTabRout, $bRout)
    _GUICtrlTab_EnableTab($hTab, $g_iTabMat,  $bMat)
    _GUICtrlTab_EnableTab($hTab, $g_iTabBOM,  $bBOM)
    _GUICtrlTab_EnableTab($hTab, $g_iTabWO,   $bWO)
    _GUICtrlTab_EnableTab($hTab, $g_iTabWOL,  $bWOL)
    _GUICtrlTab_EnableTab($hTab, $g_iTabSR,   $bSR)
    _GUICtrlTab_EnableTab($hTab, $g_iTabCap,  $bCap)
    _GUICtrlTab_EnableTab($hTab, $g_iTabStk,  $bStk)

    ; If the currently visible tab just became disabled, bounce back to Modules
    Local $iCur = _GUICtrlTab_GetCurSel($hTab)
    If Not _TabIsEnabled($iCur) Then
        $g_bAllowProgrammaticTabChange = True
        _GUICtrlTab_SetCurSel($hTab, $g_iTabMod)
        $g_bAllowProgrammaticTabChange = False
    EndIf

    ; Force a full repaint so AutoIt re-draws all tab child controls correctly.
    ; Without this, any TCM_SETITEM call above leaves controls visually stale
    ; (blank until mouse-over triggers a WM_PAINT).
    DllCall("user32.dll", "bool", "RedrawWindow", "hwnd", $g_hMain, "ptr", 0, "ptr", 0, "uint", 0x0185)
    ; 0x0185 = RDW_INVALIDATE(0x01) | RDW_ERASE(0x04) | RDW_ALLCHILDREN(0x80) | RDW_UPDATENOW(0x100)
EndFunc
;=============================================================================
; TAB: Enable/disable emulation
; Tab control items cannot be truly disabled in standard Win32.
; We emulate this by (1) marking the caption with ' (off)' and (2) blocking
; selection changes via WM_NOTIFY/TCN_SELCHANGING.
;=============================================================================
Func _InitTabState()
    If $g_hTabHandle = 0 Then Return
    Local $hTab = $g_hTabHandle
    Local $cnt = _GUICtrlTab_GetItemCount($hTab)
    If $cnt <= 0 Then Return
    ReDim $g_aTabEnabled[$cnt]
    ReDim $g_aTabBaseText[$cnt]
    For $i = 0 To $cnt - 1
        $g_aTabEnabled[$i] = True
        $g_aTabBaseText[$i] = _GUICtrlTab_GetItemText($hTab, $i)
    Next
EndFunc

Func _GUICtrlTab_EnableTab($hTab, $iIndex, $bEnable)
    If $hTab = 0 Then Return
    If $iIndex < 0 Then Return
    If UBound($g_aTabEnabled) = 0 Then _InitTabState()
    If $iIndex >= UBound($g_aTabEnabled) Then Return

    Local $bNewState = ($bEnable <> 0)

    ; Skip if state has not changed - avoids TCM_SETITEM redraw that wipes tab controls
    If $g_aTabEnabled[$iIndex] = $bNewState Then Return

    $g_aTabEnabled[$iIndex] = $bNewState

    Local $sBase = $g_aTabBaseText[$iIndex]
    If $sBase = "" Then $sBase = _GUICtrlTab_GetItemText($hTab, $iIndex)

    If Not $g_aTabEnabled[$iIndex] Then
        _GUICtrlTab_SetItemText($hTab, $iIndex, $sBase & " [off]")
    Else
        _GUICtrlTab_SetItemText($hTab, $iIndex, $sBase)
    EndIf
EndFunc

Func _TabIsEnabled($iIndex)
    If UBound($g_aTabEnabled) = 0 Then Return True
    If $iIndex < 0 Or $iIndex >= UBound($g_aTabEnabled) Then Return True
    Return $g_aTabEnabled[$iIndex]
EndFunc

Func _WM_NOTIFY($hWnd, $iMsg, $wParam, $lParam)
    If $g_hTabHandle = 0 Then Return $GUI_RUNDEFMSG

    Local $tNMHDR = DllStructCreate($tagNMHDR, $lParam)
    Local $hFrom  = HWnd(DllStructGetData($tNMHDR, "hWndFrom"))
    Local $iCode = DllStructGetData($tNMHDR, "Code")

    ; While a command is running, keep the window visible but block tab/list interactions.
    If $g_iBusyDepth > 0 Then
        If $hFrom = $g_hTabHandle And $iCode = $TCN_SELCHANGING Then Return 1
        Return 0
    EndIf

    ; Generic sortable tables: click a ListView header once for ascending, click again for descending.
    If $iCode = $ORTEMS_LVN_COLUMNCLICK Then
        Local $hLVFromNotify = _LV_GetControlFromNotifyHandle($hFrom)
        If $hLVFromNotify <> 0 Then
            Local $tNMLV = DllStructCreate($tagNMHDR & ";int Item;int SubItem;uint NewState;uint OldState;uint Changed;long ActionX;long ActionY;long Param", $lParam)
            _LV_SortByColumn($hLVFromNotify, DllStructGetData($tNMLV, "SubItem"))
            Return 0
        EndIf
    EndIf

    If $hFrom <> $g_hTabHandle Then Return $GUI_RUNDEFMSG

    Switch $iCode
        Case $TCN_SELCHANGE
            ; Skip if this notification was raised by our own programmatic switch
            If $g_bAllowProgrammaticTabChange Then Return $GUI_RUNDEFMSG

            Local $iNew = _GUICtrlTab_GetCurSel($g_hTabHandle)
            If Not _TabIsEnabled($iNew) Then
                ; User clicked a disabled tab - bounce back to Modules immediately
                $g_bAllowProgrammaticTabChange = True
                _GUICtrlTab_SetCurSel($g_hTabHandle, $g_iTabMod)
                $g_bAllowProgrammaticTabChange = False

                ; Force repaint so the Modules tab content reappears cleanly
                DllCall("user32.dll", "bool", "RedrawWindow", "hwnd", $g_hMain, "ptr", 0, "ptr", 0, "uint", 0x0185)

                If $g_lblFooter <> 0 Then GUICtrlSetData($g_lblFooter, "Tab [off] - enable the module in tab 2. Modules.")
            Else
                If $g_lblFooter <> 0 Then GUICtrlSetData($g_lblFooter, $DevBy)
            EndIf
    EndSwitch

    Return $GUI_RUNDEFMSG
EndFunc


;=============================================================================
; UI: RESIZE HANDLING (avoid controls overlapping)
;=============================================================================
Func _OnResize()
    If $g_hMain = 0 Then Return
    Local $a = WinGetClientSize($g_hMain)
    If @error Then Return
    _LayoutMain($a[0], $a[1])
EndFunc

Func _LayoutMain($w, $h)
    Local $bottomBarTop = $h - $BOTTOM_BAR_H
    If $bottomBarTop < ($TAB_TOP_Y + 200) Then $bottomBarTop = $TAB_TOP_Y + 200

    ; Resize tab control
    If $g_hTab <> 0 Then GUICtrlSetPos($g_hTab, 5, $TAB_TOP_Y, $w - 10, $bottomBarTop - $TAB_TOP_Y - 6)

    ; Bottom separator line
    If $g_idBottomLine <> 0 Then GUICtrlSetPos($g_idBottomLine, 0, $bottomBarTop + 5, $w, 2)

    If IsDeclared("g_btnHelp") Then GUICtrlSetPos($g_btnHelp, $w - 42, 8, 40, 40)

    ; Bottom toolbar buttons
    Local $yBtn = $bottomBarTop + 12
    If IsDeclared("g_btnImportXLS") Then GUICtrlSetPos($g_btnImportXLS, 10,  $yBtn, 132, 32)
    If IsDeclared("g_btnExportXLS") Then GUICtrlSetPos($g_btnExportXLS, 150, $yBtn, 132, 32)
    If IsDeclared("g_btnLoadDB")    Then GUICtrlSetPos($g_btnLoadDB,    290, $yBtn, 122, 32)
    If IsDeclared("g_btnClearDB")   Then GUICtrlSetPos($g_btnClearDB,   420, $yBtn, 92, 32)
    If IsDeclared("g_btnGenerate")  Then GUICtrlSetPos($g_btnGenerate,  520, $yBtn, 118, 32)
    If IsDeclared("g_btnExecute")   Then GUICtrlSetPos($g_btnExecute,   646, $yBtn, 115, 32)
    If IsDeclared("g_btnSaveSQL")   Then GUICtrlSetPos($g_btnSaveSQL,   769, $yBtn, 105, 32)
    If IsDeclared("g_btnIntegrity") Then GUICtrlSetPos($g_btnIntegrity, 882, $yBtn, 125, 32)
    If $g_lblFooter <> 0 Then GUICtrlSetPos($g_lblFooter, $w - 445, $yBtn + 35, 395, 18)

    ; Usable content area bottom edge (inside the tab, minus a small margin)
    Local $tabInnerBottom = $bottomBarTop - $TAB_TOP_Y - 30
    Local $bottomContent  = $TAB_TOP_Y + $tabInnerBottom - 8

    ; Resize all ListViews to fill available height
    _ResizeToBottom($g_hLV_Cal,  $w, $bottomContent)
    _ResizeToBottom($g_hLV_Mach, $w, $bottomContent)
    _ResizeToBottom($g_hLV_Ops,  $w, $bottomContent)
    _ResizeToBottom($g_hLV_Rout, $w, $bottomContent)
    _ResizeToBottom($g_hLV_Mat,  $w, $bottomContent)
    _ResizeToBottom($g_hLV_BOM,  $w, $bottomContent)
    _ResizeToBottom($g_hLV_WO,   $w, $bottomContent)
    _ResizeToBottom($g_hLV_WOL,  $w, $bottomContent)
    _ResizeToBottom($g_hLV_SR,   $w, $bottomContent)
    _ResizeToBottom($g_hLV_Cap,  $w, $bottomContent)
    _ResizeToBottom($g_hLV_Stk,  $w, $bottomContent)

    ; SQL tab: the exec log is anchored 95px above the bottom content edge.
    ; The SQL edit box fills the space above it.
    If IsDeclared("g_hExecLog") Then
        Local $yExec = $bottomContent - 75
        If $yExec < 400 Then $yExec = 400
        GUICtrlSetPos($g_hExecLog, 20, $yExec, $w - 40, 55)
    EndIf
    If $g_hLog <> 0 Then
        Local $pLog = ControlGetPos($g_hMain, "", $g_hLog)
        If Not @error Then
            Local $logBottom = (IsDeclared("g_hExecLog") ? ($bottomContent - 95) : $bottomContent)
            Local $newH = $logBottom - $pLog[1]
            If $newH < 120 Then $newH = 120
            GUICtrlSetPos($g_hLog, $pLog[0], $pLog[1], $w - 40, $newH)
        EndIf
    EndIf
EndFunc

Func _ResizeToBottom($ctrlId, $w, $bottomY)
    If $ctrlId = 0 Then Return
    Local $p = ControlGetPos($g_hMain, "", $ctrlId)
    If @error Then Return
    Local $newH = $bottomY - $p[1]
    If $newH < 80 Then $newH = 80
    GUICtrlSetPos($ctrlId, $p[0], $p[1], $w - 40, $newH)
EndFunc

; Enforce minimum window size
Func _WM_GETMINMAXINFO($hWnd, $iMsg, $wParam, $lParam)
    Local $tMMI = DllStructCreate("long;long;long;long;long;long;long;long;long;long", $lParam)
    DllStructSetData($tMMI, 7, $MIN_W)
    DllStructSetData($tMMI, 8, $MIN_H)
    Return 0
EndFunc

Func _OnClose()
    _SaveSettings()
    Local $nRet = MsgBox(262144+4 + 32, "Exit", "Do you want to exit Ortems Toolbox?" & @CRLF & @CRLF & "Unsaved data will be lost.",0,$g_hMain)
    If $nRet = 6 Then Exit
EndFunc

;=============================================================================
; SETTINGS: LOAD / SAVE  (settings.ini next to the .au3 / .exe)
;=============================================================================
Func _SaveSettings()
    ; [Connection]
    IniWrite($g_sIniFile, "Connection", "Server",   GUICtrlRead($g_edtServer))
    IniWrite($g_sIniFile, "Connection", "Database", GUICtrlRead($g_edtDatabase))
    IniWrite($g_sIniFile, "Connection", "Auth",     GUICtrlRead($g_cmbAuth))
    IniWrite($g_sIniFile, "Connection", "User",     GUICtrlRead($g_edtUser))
    ; NOTE: password is NOT saved for security reasons

    ; [Modules]
    IniWrite($g_sIniFile, "Modules", "Mode",   ($g_bModPS ? "PS" : "MP"))
    IniWrite($g_sIniFile, "Modules", "SRP",    ($g_bModSRP   ? "1" : "0"))
    IniWrite($g_sIniFile, "Modules", "WOL",    ($g_bModWOL   ? "1" : "0"))
    IniWrite($g_sIniFile, "Modules", "SR",     ($g_bModSR    ? "1" : "0"))
    IniWrite($g_sIniFile, "Modules", "INV",    ($g_bModINV   ? "1" : "0"))
    IniWrite($g_sIniFile, "Modules", "MRK",    ($g_bModMRK   ? "1" : "0"))
    IniWrite($g_sIniFile, "Modules", "LR",     ($g_bModLR    ? "1" : "0"))
    IniWrite($g_sIniFile, "Modules", "PRM",    ($g_bModPRM   ? "1" : "0"))
    IniWrite($g_sIniFile, "Modules", "BATCH",  ($g_bModBATCH ? "1" : "0"))
    IniWrite($g_sIniFile, "Modules", "CROUT",  ($g_bModCROUT ? "1" : "0"))

    ; [SQL]
    If IsDeclared("g_chkClearFirst") Then
        IniWrite($g_sIniFile, "SQL", "ClearFirst",  (GUICtrlRead($g_chkClearFirst)  = $GUI_CHECKED ? "1" : "0"))
        IniWrite($g_sIniFile, "SQL", "Transaction", (GUICtrlRead($g_chkTransaction) = $GUI_CHECKED ? "1" : "0"))
        If IsDeclared("g_chkVerbose") Then IniWrite($g_sIniFile, "SQL", "Verbose", (GUICtrlRead($g_chkVerbose) = $GUI_CHECKED ? "1" : "0"))
    EndIf

    ; [Window]
    Local $aPos = WinGetPos($g_hMain)
    If Not @error Then
        IniWrite($g_sIniFile, "Window", "X", $aPos[0])
        IniWrite($g_sIniFile, "Window", "Y", $aPos[1])
        IniWrite($g_sIniFile, "Window", "W", $aPos[2])
        IniWrite($g_sIniFile, "Window", "H", $aPos[3])
    EndIf
EndFunc

Func _LoadSettings()
    If Not FileExists($g_sIniFile) Then Return   ; first run, keep defaults

    ; [Connection]
    Local $sServer = IniRead($g_sIniFile, "Connection", "Server",   "")
    Local $sDB     = IniRead($g_sIniFile, "Connection", "Database", "")
    Local $sAuth   = IniRead($g_sIniFile, "Connection", "Auth",     "Windows Authentication")
    Local $sUser   = IniRead($g_sIniFile, "Connection", "User",     "")

    If $sServer <> "" Then GUICtrlSetData($g_edtServer,   $sServer)
    If $sDB     <> "" Then GUICtrlSetData($g_edtDatabase, $sDB)
    GUICtrlSetData($g_cmbAuth, $sAuth)
    _OnAuthChange()   ; enable/disable user+pass fields based on auth type
    If $sUser <> "" Then GUICtrlSetData($g_edtUser, $sUser)

    ; [Modules]
    Local $sMode = IniRead($g_sIniFile, "Modules", "Mode", "PS")
    If $sMode = "MP" Then
        GUICtrlSetState($g_rbMP, $GUI_CHECKED)
        GUICtrlSetState($g_rbPS, $GUI_UNCHECKED)
    Else
        GUICtrlSetState($g_rbPS, $GUI_CHECKED)
        GUICtrlSetState($g_rbMP, $GUI_UNCHECKED)
    EndIf

    Local $aChks[][2] = [ _
        [$g_chkSRP,   IniRead($g_sIniFile, "Modules", "SRP",   "0")], _
        [$g_chkWOL,   IniRead($g_sIniFile, "Modules", "WOL",   "0")], _
        [$g_chkSR,    IniRead($g_sIniFile, "Modules", "SR",    "0")], _
        [$g_chkINV,   IniRead($g_sIniFile, "Modules", "INV",   "0")], _
        [$g_chkMRK,   IniRead($g_sIniFile, "Modules", "MRK",   "0")], _
        [$g_chkLR,    IniRead($g_sIniFile, "Modules", "LR",    "0")], _
        [$g_chkPRM,   IniRead($g_sIniFile, "Modules", "PRM",   "0")], _
        [$g_chkBATCH, IniRead($g_sIniFile, "Modules", "BATCH", "0")], _
        [$g_chkCROUT, IniRead($g_sIniFile, "Modules", "CROUT", "0")] _
    ]
    For $i = 0 To UBound($aChks) - 1
        If $aChks[$i][1] = "1" Then
            GUICtrlSetState($aChks[$i][0], $GUI_CHECKED)
        Else
            GUICtrlSetState($aChks[$i][0], $GUI_UNCHECKED)
        EndIf
    Next

    ; [SQL]
    If IniRead($g_sIniFile, "SQL", "ClearFirst", "1") = "1" Then
        GUICtrlSetState($g_chkClearFirst, $GUI_CHECKED)
    Else
        GUICtrlSetState($g_chkClearFirst, $GUI_UNCHECKED)
    EndIf
    If IniRead($g_sIniFile, "SQL", "Transaction", "1") = "1" Then
        GUICtrlSetState($g_chkTransaction, $GUI_CHECKED)
    Else
        GUICtrlSetState($g_chkTransaction, $GUI_UNCHECKED)
    EndIf
    $g_bVerbose = (IniRead($g_sIniFile, "SQL", "Verbose", "0") = "1")
    If IsDeclared("g_chkVerbose") Then
        If $g_bVerbose Then
            GUICtrlSetState($g_chkVerbose, $GUI_CHECKED)
        Else
            GUICtrlSetState($g_chkVerbose, $GUI_UNCHECKED)
        EndIf
    EndIf

    ; [Window] - restore size/position
    Local $iX = Int(IniRead($g_sIniFile, "Window", "X", "-1"))
    Local $iY = Int(IniRead($g_sIniFile, "Window", "Y", "-1"))
    Local $iW = Int(IniRead($g_sIniFile, "Window", "W", $APP_WIDTH))
    Local $iH = Int(IniRead($g_sIniFile, "Window", "H", $APP_HEIGHT))
    If $iW < $MIN_W Then $iW = $MIN_W
    If $iH < $MIN_H Then $iH = $MIN_H
    ; Clamp to screen bounds
    Local $iSW = @DesktopWidth, $iSH = @DesktopHeight
    If $iX < 0 Or $iX > $iSW - 100 Then $iX = -1
    If $iY < 0 Or $iY > $iSH - 100 Then $iY = -1
    If $iX <> -1 Then WinMove($g_hMain, "", $iX, $iY, $iW, $iH)
EndFunc