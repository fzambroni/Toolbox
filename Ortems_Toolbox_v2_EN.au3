#NoTrayIcon
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <TabConstants.au3>
#include <GuiTab.au3>
#include <ListViewConstants.au3>
#include <GuiListView.au3>
#include <GuiComboBox.au3>
#include <Array.au3>
#include <File.au3>
#include <String.au3>
#include <EditConstants.au3>


;=============================================================================
; ORTEMS TOOLBOX - AutoIt Interface
; Substitui o arquivo Excel Toolbox_v2.0.1.xlsm
; Base de dados: SQL Server
;=============================================================================

Opt("GUIOnEventMode", 1)
;~ Opt("MustDeclareVars", 0)

; === Window sizing ===
;~ Global $WM_GETMINMAXINFO = 0x24
Global $MIN_W = 1050
Global $MIN_H = 700
Global $BOTTOM_BAR_H = 70
Global $TAB_TOP_Y = 58

; === CONSTANTES GLOBAIS ===
Global $TITLE       = "Ortems Toolbox v2.0"
Global $APP_WIDTH   = 1100
Global $APP_HEIGHT  = 720
Global $COL_W       = 140

; === VARIAVEIS GLOBAIS DE ESTADO ===
Global $g_sServer   = ""
Global $g_sDatabase = ""
Global $g_sConnStr  = ""
Global $g_bConnected = False

; Modulos selecionados
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

; Dados em memoria (arrays de arrays)
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

; Handles dos controles principais
Global $g_hMain, $g_hTab
Global $g_hTabHandle = 0
Global $g_idBottomLine = 0
Global $g_lblFooter = 0
Global $g_hLV_Cal, $g_hLV_Mach, $g_hLV_Ops, $g_hLV_Rout
Global $g_hLV_Mat, $g_hLV_BOM, $g_hLV_WO, $g_hLV_WOL
Global $g_hLV_SR, $g_hLV_Cap, $g_hLV_Stk
Global $g_hLog

; Indices das tabs
Global $g_iTabDB = 0, $g_iTabMod = 1
Global $g_iTabCal = 2, $g_iTabMach = 3, $g_iTabOps = 4
Global $g_iTabRout = 5, $g_iTabMat = 6, $g_iTabBOM = 7
Global $g_iTabWO = 8, $g_iTabWOL = 9, $g_iTabSR = 10
Global $g_iTabCap = 11, $g_iTabStk = 12, $g_iTabSQL = 13

;=============================================================================
; PONTO DE ENTRADA
;=============================================================================
Main()

Func Main()
    CreateMainWindow()
    GUISetState(@SW_SHOW, $g_hMain)
    _RefreshModuleFlags(True)
    _OnResize()

    While 1
        Sleep(100)
    WEnd
EndFunc

;=============================================================================
; CRIACAO DA JANELA PRINCIPAL
;=============================================================================
Func CreateMainWindow()
    $g_hMain = GUICreate($TITLE, $APP_WIDTH, $APP_HEIGHT, -1, -1, _
        BitOR($GUI_SS_DEFAULT_GUI, $WS_SIZEBOX, $WS_MAXIMIZEBOX))
    GUISetOnEvent($GUI_EVENT_CLOSE, "_OnClose")
    GUISetOnEvent($GUI_EVENT_RESIZED, "_OnResize")
    GUIRegisterMsg($WM_GETMINMAXINFO, "_WM_GETMINMAXINFO")

    ; Barra de titulo/logo
    GUICtrlCreateLabel("ORTEMS TOOLBOX", 10, 8, 300, 26)
    GUICtrlSetFont(-1, 14, 800, 0, "Segoe UI")
    GUICtrlSetColor(-1, 0x003366)

    GUICtrlCreateLabel("Custom demo builder for Ortems - SQL Database", 10, 34, 500, 18)
    GUICtrlSetFont(-1, 9, 400, 2, "Segoe UI")
    GUICtrlSetColor(-1, 0x666666)

    ; Status de conexao
    GUICtrlCreateLabel("Status:", 650, 10, 50, 18)
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    Global $g_lblStatus = GUICtrlCreateLabel("Disconnected", 705, 10, 200, 18)
    GUICtrlSetFont(-1, 9, 700, 0, "Segoe UI")
    GUICtrlSetColor($g_lblStatus, 0xCC0000)

    ; === ABAS PRINCIPAIS ===
    $g_hTab = GUICtrlCreateTab(5, $TAB_TOP_Y, $APP_WIDTH - 10, $APP_HEIGHT - 130)
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    $g_hTabHandle = GUICtrlGetHandle($g_hTab)

    ; ---- TAB 1: BANCO DE DADOS ----
    GUICtrlCreateTabItem("1. Database")
    _CreateTabDatabase()

    ; ---- ABA 2: MODULOS ----
    GUICtrlCreateTabItem("2. Modules")
    _CreateTabModules()

    ; ---- ABA 3: CALENDARIOS ----
    GUICtrlCreateTabItem("3. Calendars")
    _CreateTabCalendars()

    ; ---- ABA 4: MAQUINAS ----
    GUICtrlCreateTabItem("4. Machines")
    _CreateTabMachines()

    ; ---- ABA 5: OPERACOES ----
    GUICtrlCreateTabItem("5. Operations")
    _CreateTabOperations()

    ; ---- ABA 6: ROTEIROS ----
    GUICtrlCreateTabItem("6. Routings")
    _CreateTabRoutings()

    ; ---- ABA 7: MATERIAIS ----
    GUICtrlCreateTabItem("7. Items")
    _CreateTabMaterials()

    ; ---- ABA 8: BOM ----
    GUICtrlCreateTabItem("8. BOM")
    _CreateTabBOM()

    ; ---- ABA 9: ORDENS DE PRODUCAO ----
    GUICtrlCreateTabItem("9. Work Orders (WO)")
    _CreateTabWO()

    ; ---- ABA 10: LINKS WO ----
    GUICtrlCreateTabItem("10. WO Links")
    _CreateTabWOLinks()

    ; ---- ABA 11: RECURSOS SECUNDARIOS ----
    GUICtrlCreateTabItem("11. Secondary Resources")
    _CreateTabSecResources()

    ; ---- ABA 12: CAPACIDADE ----
    GUICtrlCreateTabItem("12. Capacity")
    _CreateTabCapacity()

    ; ---- ABA 13: ESTOQUES ----
    GUICtrlCreateTabItem("13. Inventory Movements")
    _CreateTabStockMov()

    ; ---- ABA 14: GERAR SQL ----
    GUICtrlCreateTabItem("14. Generate & Run SQL")
    _CreateTabSQL()

    GUICtrlCreateTabItem("")

    ; === BARRA INFERIOR ===
    $g_idBottomLine = GUICtrlCreateLabel("", 0, $APP_HEIGHT - 65, $APP_WIDTH, 2)
    GUICtrlSetBkColor(-1, 0xCCCCCC)

    Global $g_btnImportXLS = GUICtrlCreateButton("Import Excel...", 10, $APP_HEIGHT - 58, 140, 32)
    GUICtrlSetOnEvent($g_btnImportXLS, "_ImportExcel")
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")

    Global $g_btnExportXLS = GUICtrlCreateButton("Export Excel...", 160, $APP_HEIGHT - 58, 140, 32)
    GUICtrlSetOnEvent($g_btnExportXLS, "_ExportExcel")
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")

    Global $g_btnClearDB = GUICtrlCreateButton("Clear DB", 310, $APP_HEIGHT - 58, 130, 32)
    GUICtrlSetOnEvent($g_btnClearDB, "_ClearDatabase")
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    GUICtrlSetBkColor($g_btnClearDB, 0xFFDDDD)

    Global $g_btnGenerate = GUICtrlCreateButton("GENERATE SQL", 460, $APP_HEIGHT - 58, 130, 32)
    GUICtrlSetOnEvent($g_btnGenerate, "_GenerateSQL")
    GUICtrlSetFont(-1, 9, 700, 0, "Segoe UI")

    Global $g_btnExecute = GUICtrlCreateButton("RUN ON DB", 600, $APP_HEIGHT - 58, 170, 32)
    GUICtrlSetOnEvent($g_btnExecute, "_ExecuteSQL")
    GUICtrlSetFont(-1, 9, 700, 0, "Segoe UI")
    GUICtrlSetBkColor($g_btnExecute, 0xDDFFDD)

    Global $g_btnSaveSQL = GUICtrlCreateButton("Save SQL...", 780, $APP_HEIGHT - 58, 130, 32)
    GUICtrlSetOnEvent($g_btnSaveSQL, "_SaveSQL")
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")

    $g_lblFooter = GUICtrlCreateLabel($TITLE & " | Replaces Toolbox_v2.0.1.xlsm | " & @YEAR, 920, $APP_HEIGHT - 45, 170, 20)
    GUICtrlSetFont(-1, 7, 400, 0, "Segoe UI")
    GUICtrlSetColor(-1, 0x999999)
EndFunc

;=============================================================================
; ABA: BANCO DE DADOS
;=============================================================================
Func _CreateTabDatabase()
    Local $y = 85, $xL = 20, $xV = 200

    GUICtrlCreateGroup("SQL Server Connection", $xL, $y, 600, 200)
    $y += 25

    GUICtrlCreateLabel("Server / Instance:", $xL + 10, $y, 170, 20)
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    Global $g_edtServer = GUICtrlCreateInput("localhost\SQLEXPRESS", $xV, $y, 200, 22)
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    $y += 30

    GUICtrlCreateLabel("Database name:", $xL + 10, $y, 170, 20)
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    Global $g_edtDatabase = GUICtrlCreateInput("ORTEMS_DEMO", $xV, $y, 200, 22)
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    $y += 30

    GUICtrlCreateLabel("Authentication:", $xL + 10, $y, 170, 20)
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    Global $g_cmbAuth = GUICtrlCreateCombo("Windows Authentication", $xV, $y, 200, 22)
    GUICtrlSetData($g_cmbAuth, "SQL Server Authentication")
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    GUICtrlSetOnEvent($g_cmbAuth, "_OnAuthChange")
    $y += 30

    GUICtrlCreateLabel("User:", $xL + 10, $y, 170, 20)
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    Global $g_edtUser = GUICtrlCreateInput("", $xV, $y, 200, 22)
    GUICtrlSetState($g_edtUser, $GUI_DISABLE)
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    $y += 30

    GUICtrlCreateLabel("Password:", $xL + 10, $y, 170, 20)
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    Global $g_edtPass = GUICtrlCreateInput("", $xV, $y, 200, 22, $ES_PASSWORD)
    GUICtrlSetState($g_edtPass, $GUI_DISABLE)
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    $y += 40

    Global $g_btnConnect = GUICtrlCreateButton("Test Connection", $xV, $y, 150, 30)
    GUICtrlSetOnEvent($g_btnConnect, "_TestConnection")
    GUICtrlSetFont(-1, 9, 700, 0, "Segoe UI")

    ; Info do caminho do banco (compatibilidade com Excel)
    $y = 320
    GUICtrlCreateGroup("Information", $xL, $y, 600, 120)
    $y += 25

    GUICtrlCreateLabel("Connection string:", $xL + 10, $y, 170, 20)
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    Global $g_edtConnStr = GUICtrlCreateEdit("", $xL + 10, $y + 22, 570, 60, $ES_READONLY + $WS_VSCROLL)
    GUICtrlSetFont(-1, 8, 400, 0, "Courier New")
    GUICtrlSetBkColor($g_edtConnStr, 0xF5F5F5)

    ; Instrucoes
    $y = 470
    GUICtrlCreateGroup("How to use the Toolbox", $xL, $y, 700, 150)
    $y += 22
    Local $sInfo = "WORKFLOW:" & @CRLF & _
        "  1. Configure the connection to the Ortems database (above)" & @CRLF & _
        "  2. Go to the '2. Modules' tab and select the required Ortems modules for the demo" & @CRLF & _
        "  3. Fill in the data in the tabs (Calendars, Machines, Operations, Routings, Items, etc.)" & @CRLF & _
        "  4. Click 'GENERATE SQL' to build the SQL statements" & @CRLF & _
        "  5. Click 'RUN ON DB' to insert the data into the Ortems database" & @CRLF & _
        "  Tip: You can import data from an existing Excel file using 'Import Excel...'"
    GUICtrlCreateEdit($sInfo, $xL + 10, $y, 680, 110, $ES_READONLY + $WS_VSCROLL)
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    GUICtrlSetBkColor(-1, 0xFFFFF0)
EndFunc

;=============================================================================
; ABA: SELECAO DE MODULOS
;=============================================================================
Func _CreateTabModules()
    GUICtrlCreateLabel("Ortems Module Selection", 20, 80, 500, 22)
    GUICtrlSetFont(-1, 12, 700, 0, "Segoe UI")
    GUICtrlSetColor(-1, 0x003366)

    GUICtrlCreateLabel("Answer the questions below to configure the demo modules:", 20, 105, 700, 18)
    GUICtrlSetFont(-1, 9, 400, 2, "Segoe UI")

    Local $y = 135, $xL = 20
    GUICtrlCreateGroup("Module Configuration", $xL, $y, 900, 480)
    $y += 20

    ; Q1: PS vs MP
    _CreateModuleQuestion($y, "Q1:", "Does the customer need continuous operation scheduling (PS) or", "bucket planning with load leveling (MP)?")
    Global $g_rbPS = GUICtrlCreateRadio("PS - Production Scheduling (continuous scheduling)", $xL + 30, $y + 40, 420, 20)
    GUICtrlSetOnEvent($g_rbPS, "_OnModuleSelectionChanged")
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    GUICtrlSetState($g_rbPS, $GUI_CHECKED)
    Global $g_rbMP = GUICtrlCreateRadio("MP - Master Planning (bucket planning)", $xL + 30, $y + 62, 420, 20)
    GUICtrlSetOnEvent($g_rbMP, "_OnModuleSelectionChanged")
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    $y += 92

    ; Q2: SRP
    Global $g_chkSRP = _CreateModuleCheckbox($y, $xL, "Q2:", "Automatic synchronization across BOM levels? (SRP module)")
    $y += 45

    ; Q3: WO Links
    Global $g_chkWOL = _CreateModuleCheckbox($y, $xL, "Q3:", "Pre-defined precedence constraints between WOs/operations? (WO Links)")
    $y += 45

    ; Q4: Complex Routings
    Global $g_chkCROUT = _CreateModuleCheckbox($y, $xL, "Q4:", "Complex, non-linear routings, or start-start synchronization? (Complex routings)")
    $y += 45

    ; Q5: Secondary Resources
    Global $g_chkSR = _CreateModuleCheckbox($y, $xL, "Q5:", "Labor/tool constraints? Multi-resource optimization? (Secondary resources)")
    $y += 45

    ; Q6: Batch Machines
    Global $g_chkBATCH = _CreateModuleCheckbox($y, $xL, "Q6:", "Batch machines where operations can be grouped (e.g., oven)? (Batch machines)")
    $y += 45

    ; Q7: Inventory
    Global $g_chkINV = _CreateModuleCheckbox($y, $xL, "Q7:", "Show detailed inventory features (e.g., raw material replenishment)? (Inventory)")
    $y += 45

    ; Q8: Markers
    Global $g_chkMRK = _CreateModuleCheckbox($y, $xL, "Q8:", "Visual markers in planning for special events / collaboration? (Markers)")
    $y += 45

    ; Q9: Limited Resources
    Global $g_chkLR = _CreateModuleCheckbox($y, $xL, "Q9:", "Limited shared resources across multiple operations? (Limited resources)")
    $y += 45

    ; Q10: Parameters
    Global $g_chkPRM = _CreateModuleCheckbox($y, $xL, "Q10:", "Grouping by categories (color, temperature), changeovers between tools? (Parameters)")
    $y += 45

    ; Botao aplicar
    $y = $APP_HEIGHT - 145
    Global $g_btnApplyMod = GUICtrlCreateButton("Apply module selection", 20, $y, 250, 35)
    GUICtrlSetOnEvent($g_btnApplyMod, "_ApplyModules")
    GUICtrlSetFont(-1, 10, 700, 0, "Segoe UI")
    GUICtrlSetBkColor($g_btnApplyMod, 0xDDEEFF)

    Global $g_lblModHint = GUICtrlCreateLabel("Selected modules determine which data tabs are enabled.", 290, $y + 8, 600, 20)
    GUICtrlSetFont(-1, 9, 400, 2, "Segoe UI")
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
    Local $hChk = GUICtrlCreateCheckbox($sText, $xL + 50, $y, 800, 20)
    GUICtrlSetOnEvent($hChk, "_OnModuleSelectionChanged")
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    Return $hChk
EndFunc

;=============================================================================
; ABA: CALENDARIOS (SV_CALENDARS)
;=============================================================================
Func _CreateTabCalendars()
    ; Descricao
    GUICtrlCreateLabel("Work Calendars (SV_CALENDARS)", 20, 80, 600, 20)
    GUICtrlSetFont(-1, 11, 700, 0, "Segoe UI")
    GUICtrlCreateLabel("A calendar defines working hours for machines or labor. Each row represents a shift for a given calendar.", 20, 103, 900, 18)
    GUICtrlSetFont(-1, 8, 400, 2, "Segoe UI")

    ; Botoes CRUD
    Local $y = 125
    _CreateCRUDButtons($y, "_Cal_Add", "_Cal_Edit", "_Cal_Del", "_Cal_DelAll")

    ; ListView
    $g_hLV_Cal = GUICtrlCreateListView("Calendar ID|Calendar name|Start day|Start time|End day|End time|", _
        20, $y + 35, $APP_WIDTH - 40, 390, _
        BitOR($LVS_REPORT, $LVS_SHOWSELALWAYS, $WS_BORDER))
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    _GUICtrlListView_SetColumnWidth($g_hLV_Cal, 0, 130)
    _GUICtrlListView_SetColumnWidth($g_hLV_Cal, 1, 150)
    _GUICtrlListView_SetColumnWidth($g_hLV_Cal, 2, 90)
    _GUICtrlListView_SetColumnWidth($g_hLV_Cal, 3, 90)
    _GUICtrlListView_SetColumnWidth($g_hLV_Cal, 4, 90)
    _GUICtrlListView_SetColumnWidth($g_hLV_Cal, 5, 90)

    ; Exemplo
    GUICtrlCreateLabel("Example: Cal_1x8 | Cal 1x8 | 1 (Mon) | 08:00 | 1 (Mon) | 17:00   |   Days: 1=Mon, 2=Tue, ... 7=Sun", 20, $y + 440, 900, 18)
    GUICtrlSetFont(-1, 8, 400, 2, "Segoe UI")
    GUICtrlSetColor(-1, 0x555555)

    _LoadExampleCalendars()
EndFunc

Func _LoadExampleCalendars()
    ; Adiciona exemplos padrão
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
    GUICtrlCreateListViewItem($sID & "|" & $sNome & "|" & $iDiaI & "|" & $sHoraI & "|" & $iDiaF & "|" & $sHoraF, $g_hLV_Cal)
EndFunc

;=============================================================================
; ABA: MAQUINAS (SV_MACHINE)
;=============================================================================
Func _CreateTabMachines()
    GUICtrlCreateLabel("Machines & Work Centers (SV_MACHINE)", 20, 80, 600, 20)
    GUICtrlSetFont(-1, 11, 700, 0, "Segoe UI")
    GUICtrlCreateLabel("A machine is a resource that executes operations. Each machine belongs to a single work center.", 20, 103, 900, 18)
    GUICtrlSetFont(-1, 8, 400, 2, "Segoe UI")

    Local $y = 125
    _CreateCRUDButtons($y, "_Mach_Add", "_Mach_Edit", "_Mach_Del", "_Mach_DelAll")

    $g_hLV_Mach = GUICtrlCreateListView("Site ID|Site Nome|CT ID|CT Nome|Tipo CT|Secao ID|Secao Nome|Maq ID|Maq Nome|Tipo Maq|Cal ID|Cal Cap ID|", _
        20, $y + 35, $APP_WIDTH - 40, 390, _
        BitOR($LVS_REPORT, $LVS_SHOWSELALWAYS, $WS_BORDER))
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    For $i = 0 To 11
        _GUICtrlListView_SetColumnWidth($g_hLV_Mach, $i, $COL_W - 30)
    Next

    GUICtrlCreateLabel("WC type: 2=Finite capacity | 4=Infinite capacity   |   Machine type: NR=Standard, BA=Batch, RN=Run, CU=Tank", 20, $y + 440, 900, 18)
    GUICtrlSetFont(-1, 8, 400, 2, "Segoe UI")
    GUICtrlSetColor(-1, 0x555555)

    _LoadExampleMachines()
EndFunc

Func _LoadExampleMachines()
    _AddMachine("LYON", "LYON", "Milling", "Fresagem", 2, "Sec_A", "Secao A", "Milling_1", "Fresagem 1", "NR", "Cal_1x8", "Cal_1x8")
    _AddMachine("LYON", "LYON", "Milling", "Fresagem", 2, "Sec_A", "Secao A", "Milling_2", "Fresagem 2", "NR", "Cal_1x8", "Cal_1x8")
    _AddMachine("LYON", "LYON", "Drilling", "Furadeira", 2, "Sec_A", "Secao A", "Drilling_1", "Furadeira 1", "NR", "Cal_1x8", "Cal_1x8")
    _AddMachine("LYON", "LYON", "Assembly_Robot", "Robotica", 2, "Sec_B", "Secao B", "Assy_Robot_1", "Robo Montagem 1", "NR", "Cal_2x8", "Cal_2x8")
EndFunc

Func _AddMachine($sSiteID, $sSiteNm, $sCTID, $sCTNm, $iCtTp, $sSecID, $sSecNm, $sMaqID, $sMaqNm, $sMaqTp, $sCalID, $sCalCap)
    Local $n = UBound($g_aMach, 1)
    ReDim $g_aMach[$n + 1][13]
    $g_aMach[$n][0] = $sSiteID ; ...store all fields
    GUICtrlCreateListViewItem($sSiteID & "|" & $sSiteNm & "|" & $sCTID & "|" & $sCTNm & "|" & $iCtTp & "|" & $sSecID & "|" & $sSecNm & "|" & $sMaqID & "|" & $sMaqNm & "|" & $sMaqTp & "|" & $sCalID & "|" & $sCalCap, $g_hLV_Mach)
EndFunc

;=============================================================================
; ABA: OPERACOES (SV_OPERATIONS)
;=============================================================================
Func _CreateTabOperations()
    GUICtrlCreateLabel("Operations (SV_OPERATIONS)", 20, 80, 600, 20)
    GUICtrlSetFont(-1, 11, 700, 0, "Segoe UI")
    GUICtrlCreateLabel("An operation describes a process executed by a machine (e.g., Milling, Drilling, Assembly).", 20, 103, 900, 18)
    GUICtrlSetFont(-1, 8, 400, 2, "Segoe UI")

    Local $y = 125
    _CreateCRUDButtons($y, "_Ops_Add", "_Ops_Edit", "_Ops_Del", "_Ops_DelAll")

    $g_hLV_Ops = GUICtrlCreateListView("Operation ID|Operation name|WC ID|Machine ID|Ref qty|Ref duration|Unit|Setup|Break|Interrupt|", _
        20, $y + 35, $APP_WIDTH - 40, 390, _
        BitOR($LVS_REPORT, $LVS_SHOWSELALWAYS, $WS_BORDER))
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    For $i = 0 To 9
        _GUICtrlListView_SetColumnWidth($g_hLV_Ops, $i, $COL_W - 20)
    Next

    GUICtrlCreateLabel("Unit: D=Days | H=Hours | C=Hundredths of an hour   |   Interruptible: 1=Yes, 0=No", 20, $y + 440, 900, 18)
    GUICtrlSetFont(-1, 8, 400, 2, "Segoe UI")
    GUICtrlSetColor(-1, 0x555555)

    _LoadExampleOperations()
EndFunc

Func _LoadExampleOperations()
    _AddOperation("MILL", "Fresagem Eixos", "Milling", "Milling_1", 26, 1, "H", 0, 0, 1)
    _AddOperation("MILL", "Fresagem Eixos", "Milling", "Milling_2", 24, 1, "H", 0, 0, 1)
    _AddOperation("DRIL", "Furadeira Carter", "Drilling", "Drilling_1", 49, 1, "H", 0, 0, 1)
    _AddOperation("ASSY", "Montagem Caixa", "Assembly_Robot", "Assy_Robot_1", 19, 1, "H", 0, 0, 1)
EndFunc

Func _AddOperation($sID, $sNome, $sCT, $sMaq, $nQtdRef, $nDurRef, $sUni, $nPrep, $nTout, $bInterr)
    Local $n = UBound($g_aOps, 1)
    ReDim $g_aOps[$n + 1][11]
    GUICtrlCreateListViewItem($sID & "|" & $sNome & "|" & $sCT & "|" & $sMaq & "|" & $nQtdRef & "|" & $nDurRef & "|" & $sUni & "|" & $nPrep & "|" & $nTout & "|" & $bInterr, $g_hLV_Ops)
EndFunc

;=============================================================================
; ABA: ROTEIROS (SV_ROUTINGS)
;=============================================================================
Func _CreateTabRoutings()
    GUICtrlCreateLabel("Production Routings (SV_ROUTINGS)", 20, 80, 600, 20)
    GUICtrlSetFont(-1, 11, 700, 0, "Segoe UI")
    GUICtrlCreateLabel("A routing describes the steps of a manufacturing process (sequence of phases/operations).", 20, 103, 900, 18)
    GUICtrlSetFont(-1, 8, 400, 2, "Segoe UI")

    Local $y = 125
    _CreateCRUDButtons($y, "_Rout_Add", "_Rout_Edit", "_Rout_Del", "_Rout_DelAll")

    $g_hLV_Rout = GUICtrlCreateListView("Routing ID|Routing name|Phase code|Operation ID|Phase name|", _
        20, $y + 35, $APP_WIDTH - 40, 390, _
        BitOR($LVS_REPORT, $LVS_SHOWSELALWAYS, $WS_BORDER))
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    _GUICtrlListView_SetColumnWidth($g_hLV_Rout, 0, 160)
    _GUICtrlListView_SetColumnWidth($g_hLV_Rout, 1, 200)
    _GUICtrlListView_SetColumnWidth($g_hLV_Rout, 2, 110)
    _GUICtrlListView_SetColumnWidth($g_hLV_Rout, 3, 160)
    _GUICtrlListView_SetColumnWidth($g_hLV_Rout, 4, 250)

    GUICtrlCreateLabel("Phase code: number (10, 20, 30...) that defines the operation sequence in the routing.", 20, $y + 440, 900, 18)
    GUICtrlSetFont(-1, 8, 400, 2, "Segoe UI")
    GUICtrlSetColor(-1, 0x555555)

    _LoadExampleRoutings()
EndFunc

Func _LoadExampleRoutings()
    _AddRouting("STD_AX", "Eixos Standard", 10, "MILL", "Fresagem Eixos Standard")
    _AddRouting("SPT_AX", "Eixos Sport", 10, "MILL", "Fresagem Eixos Sport")
    _AddRouting("SPT_AX", "Eixos Sport", 20, "MILL", "Acabamento Eixos Sport")
    _AddRouting("CT", "Carter", 10, "DRIL", "Furadeira Carter")
    _AddRouting("STD_GB", "Caixa Standard", 10, "ASSY", "Montagem Caixa Standard")
    _AddRouting("SPT_GB", "Caixa Sport", 10, "ASSY", "Montagem Caixa Sport")
EndFunc

Func _AddRouting($sID, $sNome, $nFase, $sOp, $sNomFase)
    Local $n = UBound($g_aRout, 1)
    ReDim $g_aRout[$n + 1][5]
    GUICtrlCreateListViewItem($sID & "|" & $sNome & "|" & $nFase & "|" & $sOp & "|" & $sNomFase, $g_hLV_Rout)
EndFunc

;=============================================================================
; ABA: MATERIAIS (SV_MATERIALS)
;=============================================================================
Func _CreateTabMaterials()
    GUICtrlCreateLabel("Items (SV_MATERIALS)", 20, 80, 600, 20)
    GUICtrlSetFont(-1, 11, 700, 0, "Segoe UI")
    GUICtrlCreateLabel("An item is a product that is consumed or produced during the production process.", 20, 103, 900, 18)
    GUICtrlSetFont(-1, 8, 400, 2, "Segoe UI")

    Local $y = 125
    _CreateCRUDButtons($y, "_Mat_Add", "_Mat_Edit", "_Mat_Del", "_Mat_DelAll")

    $g_hLV_Mat = GUICtrlCreateListView("Item ID|Item name|Type|Version|Routing ID|On-hand qty|", _
        20, $y + 35, $APP_WIDTH - 40, 390, _
        BitOR($LVS_REPORT, $LVS_SHOWSELALWAYS, $WS_BORDER))
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    _GUICtrlListView_SetColumnWidth($g_hLV_Mat, 0, 150)
    _GUICtrlListView_SetColumnWidth($g_hLV_Mat, 1, 200)
    _GUICtrlListView_SetColumnWidth($g_hLV_Mat, 2, 80)
    _GUICtrlListView_SetColumnWidth($g_hLV_Mat, 3, 80)
    _GUICtrlListView_SetColumnWidth($g_hLV_Mat, 4, 150)
    _GUICtrlListView_SetColumnWidth($g_hLV_Mat, 5, 110)

    GUICtrlCreateLabel("Type: RM=Raw material (in) | SF=Semi-finished (in/out) | FG=Finished good (out)", 20, $y + 440, 900, 18)
    GUICtrlSetFont(-1, 8, 400, 2, "Segoe UI")
    GUICtrlSetColor(-1, 0x555555)

    _LoadExampleMaterials()
EndFunc

Func _LoadExampleMaterials()
    _AddMaterial("R_Eixos", "Eixos Brutos", "MP", "00", "", 300)
    _AddMaterial("R_Carter", "Carter Bruto", "MP", "00", "", 72)
    _AddMaterial("STD_Eixos", "Eixos Standard", "SF", "STD", "STD_AX", 23)
    _AddMaterial("SPT_Eixos", "Eixos Sport", "SF", "SPT", "SPT_AX", 33)
    _AddMaterial("Carter", "Carter", "SF", "00", "CT", 36)
    _AddMaterial("STD_Caixa", "Caixa Standard", "PF", "STD", "STD_GB", 0)
    _AddMaterial("SPT_Caixa", "Caixa Sport", "PF", "SPT", "SPT_GB", 0)
EndFunc

Func _AddMaterial($sID, $sNome, $sTipo, $sVer, $sRot, $nStk)
    Local $n = UBound($g_aMat, 1)
    ReDim $g_aMat[$n + 1][6]
    GUICtrlCreateListViewItem($sID & "|" & $sNome & "|" & $sTipo & "|" & $sVer & "|" & $sRot & "|" & $nStk, $g_hLV_Mat)
EndFunc

;=============================================================================
; ABA: BOM (SV_BOM)
;=============================================================================
Func _CreateTabBOM()
    GUICtrlCreateLabel("Bill of Materials - BOM (SV_BOM)", 20, 80, 600, 20)
    GUICtrlSetFont(-1, 11, 700, 0, "Segoe UI")
    GUICtrlCreateLabel("The BOM defines the relationship between the parent item and the component items required to produce it.", 20, 103, 900, 18)
    GUICtrlSetFont(-1, 8, 400, 2, "Segoe UI")

    Local $y = 125
    _CreateCRUDButtons($y, "_BOM_Add", "_BOM_Edit", "_BOM_Del", "_BOM_DelAll")

    $g_hLV_BOM = GUICtrlCreateListView("Parent item ID|Parent version|Component item ID|Routing ID|Phase|Ref qty|Required qty|", _
        20, $y + 35, $APP_WIDTH - 40, 390, _
        BitOR($LVS_REPORT, $LVS_SHOWSELALWAYS, $WS_BORDER))
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    For $i = 0 To 6
        _GUICtrlListView_SetColumnWidth($g_hLV_BOM, $i, $COL_W)
    Next

    GUICtrlCreateLabel("Example: STD_Box | STD | STD_Axes | STD_GB | 10 | 1 | 1   (1 axis to make 1 box in phase 10)", 20, $y + 440, 900, 18)
    GUICtrlSetFont(-1, 8, 400, 2, "Segoe UI")
    GUICtrlSetColor(-1, 0x555555)

    _LoadExampleBOM()
EndFunc

Func _LoadExampleBOM()
    _AddBOM("STD_Caixa", "STD", "R_Eixos",  "STD_GB", 10, 1, 2)
    _AddBOM("STD_Caixa", "STD", "STD_Eixos","STD_GB", 10, 1, 1)
    _AddBOM("STD_Caixa", "STD", "Carter",   "STD_GB", 10, 1, 1)
    _AddBOM("SPT_Caixa", "SPT", "R_Eixos",  "SPT_GB", 10, 1, 2)
    _AddBOM("SPT_Caixa", "SPT", "SPT_Eixos","SPT_GB", 10, 1, 1)
EndFunc

Func _AddBOM($sPaiID, $sVPai, $sCompID, $sRotID, $nFase, $nQRef, $nQNec)
    Local $n = UBound($g_aBOM, 1)
    ReDim $g_aBOM[$n + 1][7]
    GUICtrlCreateListViewItem($sPaiID & "|" & $sVPai & "|" & $sCompID & "|" & $sRotID & "|" & $nFase & "|" & $nQRef & "|" & $nQNec, $g_hLV_BOM)
EndFunc

;=============================================================================
; ABA: ORDENS DE PRODUCAO (SV_WO)
;=============================================================================
Func _CreateTabWO()
    GUICtrlCreateLabel("Work Orders (SV_WO)", 20, 80, 600, 20)
    GUICtrlSetFont(-1, 11, 700, 0, "Segoe UI")
    GUICtrlCreateLabel("A work order (WO) is a production order with item, quantity, and planned due date.", 20, 103, 900, 18)
    GUICtrlSetFont(-1, 8, 400, 2, "Segoe UI")

    Local $y = 125
    _CreateCRUDButtons($y, "_WO_Add", "_WO_Edit", "_WO_Del", "_WO_DelAll")

    $g_hLV_WO = GUICtrlCreateListView("WO ID|Item ID|Routing|Version|Qty|Start date|End date|", _
        20, $y + 35, $APP_WIDTH - 40, 390, _
        BitOR($LVS_REPORT, $LVS_SHOWSELALWAYS, $WS_BORDER))
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    _GUICtrlListView_SetColumnWidth($g_hLV_WO, 0, 120)
    _GUICtrlListView_SetColumnWidth($g_hLV_WO, 1, 140)
    _GUICtrlListView_SetColumnWidth($g_hLV_WO, 2, 140)
    _GUICtrlListView_SetColumnWidth($g_hLV_WO, 3, 80)
    _GUICtrlListView_SetColumnWidth($g_hLV_WO, 4, 80)
    _GUICtrlListView_SetColumnWidth($g_hLV_WO, 5, 140)
    _GUICtrlListView_SetColumnWidth($g_hLV_WO, 6, 140)

    GUICtrlCreateLabel("Date format: dd/mm/yyyy hh:mm   |   WO number must be unique", 20, $y + 440, 900, 18)
    GUICtrlSetFont(-1, 8, 400, 2, "Segoe UI")
    GUICtrlSetColor(-1, 0x555555)

    _LoadExampleWOs()
EndFunc

Func _LoadExampleWOs()
    _AddWO("WO001", "STD_Caixa", "STD_GB", "STD", 10, "01/02/2025 00:00", "28/02/2025 23:59")
    _AddWO("WO002", "STD_Caixa", "STD_GB", "STD", 15, "01/02/2025 00:00", "28/02/2025 23:59")
    _AddWO("WO003", "SPT_Caixa", "SPT_GB", "SPT", 8,  "01/03/2025 00:00", "31/03/2025 23:59")
EndFunc

Func _AddWO($sNum, $sMatID, $sRotID, $sVer, $nQtd, $sDtI, $sDtF)
    Local $n = UBound($g_aWO, 1)
    ReDim $g_aWO[$n + 1][7]
    GUICtrlCreateListViewItem($sNum & "|" & $sMatID & "|" & $sRotID & "|" & $sVer & "|" & $nQtd & "|" & $sDtI & "|" & $sDtF, $g_hLV_WO)
EndFunc

;=============================================================================
; ABA: WO LINKS
;=============================================================================
Func _CreateTabWOLinks()
    GUICtrlCreateLabel("WO / Operation links (Precedence constraints) (SV_WO_LINKS)", 20, 80, 600, 20)
    GUICtrlSetFont(-1, 11, 700, 0, "Segoe UI")
    GUICtrlCreateLabel("A link is a precedence constraint between two WOs/operations. Example: WO2 can only start after phase 10 of WO1 ends.", 20, 103, 900, 18)
    GUICtrlSetFont(-1, 8, 400, 2, "Segoe UI")

    Local $y = 125
    _CreateCRUDButtons($y, "_WOL_Add", "_WOL_Edit", "_WOL_Del", "_WOL_DelAll")

    $g_hLV_WOL = GUICtrlCreateListView("Predecessor WO|Pred routing|Pred phase|Successor WO|Succ routing|Succ phase|Link type|", _
        20, $y + 35, $APP_WIDTH - 40, 390, _
        BitOR($LVS_REPORT, $LVS_SHOWSELALWAYS, $WS_BORDER))
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    For $i = 0 To 6
        _GUICtrlListView_SetColumnWidth($g_hLV_WOL, $i, $COL_W)
    Next

    GUICtrlCreateLabel("Link types: FS=Finish-Start | SS=Start-Start | FF=Finish-Finish | SF=Start-Finish", 20, $y + 440, 900, 18)
    GUICtrlSetFont(-1, 8, 400, 2, "Segoe UI")
    GUICtrlSetColor(-1, 0x555555)
EndFunc

;=============================================================================
; ABA: RECURSOS SECUNDARIOS
;=============================================================================
Func _CreateTabSecResources()
    GUICtrlCreateLabel("Secondary Resources (SV_SEC_RESOURCES)", 20, 80, 600, 20)
    GUICtrlSetFont(-1, 11, 700, 0, "Segoe UI")
    GUICtrlCreateLabel("A secondary resource is required during an operation in addition to the machine (e.g., labor).", 20, 103, 900, 18)
    GUICtrlSetFont(-1, 8, 400, 2, "Segoe UI")

    Local $y = 125
    _CreateCRUDButtons($y, "_SR_Add", "_SR_Edit", "_SR_Del", "_SR_DelAll")

    $g_hLV_SR = GUICtrlCreateListView("Operation ID|WC ID|Machine ID|Qualification ID|Capacity calendar ID|", _
        20, $y + 35, $APP_WIDTH - 40, 390, _
        BitOR($LVS_REPORT, $LVS_SHOWSELALWAYS, $WS_BORDER))
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    For $i = 0 To 4
        _GUICtrlListView_SetColumnWidth($g_hLV_SR, $i, $COL_W + 20)
    Next
EndFunc

;=============================================================================
; ABA: CAPACIDADE
;=============================================================================
Func _CreateTabCapacity()
    GUICtrlCreateLabel("Capacity Calendars (SV_CAPACITY)", 20, 80, 600, 20)
    GUICtrlSetFont(-1, 11, 700, 0, "Segoe UI")
    GUICtrlCreateLabel("A capacity calendar defines shifts and the number of available resources for load calculation.", 20, 103, 900, 18)
    GUICtrlSetFont(-1, 8, 400, 2, "Segoe UI")

    Local $y = 125
    _CreateCRUDButtons($y, "_Cap_Add", "_Cap_Edit", "_Cap_Del", "_Cap_DelAll")

    $g_hLV_Cap = GUICtrlCreateListView("Cap cal ID|Start day|Start time|End day|End time|#Resources|", _
        20, $y + 35, $APP_WIDTH - 40, 390, _
        BitOR($LVS_REPORT, $LVS_SHOWSELALWAYS, $WS_BORDER))
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    For $i = 0 To 5
        _GUICtrlListView_SetColumnWidth($g_hLV_Cap, $i, $COL_W + 10)
    Next
EndFunc

;=============================================================================
; ABA: MOVIMENTACOES DE ESTOQUE
;=============================================================================
Func _CreateTabStockMov()
    GUICtrlCreateLabel("Inventory Movements (SV_STOCK_MOVEMENTS)", 20, 80, 600, 20)
    GUICtrlSetFont(-1, 11, 700, 0, "Segoe UI")
    GUICtrlCreateLabel("An inventory movement is an in/out transaction of an item on a given date.", 20, 103, 900, 18)
    GUICtrlSetFont(-1, 8, 400, 2, "Segoe UI")

    Local $y = 125
    _CreateCRUDButtons($y, "_Stk_Add", "_Stk_Edit", "_Stk_Del", "_Stk_DelAll")

    $g_hLV_Stk = GUICtrlCreateListView("Item ID|Routing ID|Version|Move date|Quantity|", _
        20, $y + 35, $APP_WIDTH - 40, 390, _
        BitOR($LVS_REPORT, $LVS_SHOWSELALWAYS, $WS_BORDER))
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    For $i = 0 To 4
        _GUICtrlListView_SetColumnWidth($g_hLV_Stk, $i, $COL_W + 30)
    Next
EndFunc

;=============================================================================
; ABA: GERAR SQL
;=============================================================================
Func _CreateTabSQL()
    GUICtrlCreateLabel("SQL Generation & Execution", 20, 80, 500, 20)
    GUICtrlSetFont(-1, 11, 700, 0, "Segoe UI")
    GUICtrlCreateLabel("Click 'GENERATE SQL' in the bottom bar to build the script. Then click 'RUN ON DB' to insert the data.", 20, 103, 900, 18)
    GUICtrlSetFont(-1, 8, 400, 2, "Segoe UI")

    ; Opcoes de geracao
    Local $y = 128
    GUICtrlCreateGroup("Opcoes", 20, $y, 400, 80)
    Global $g_chkClearFirst = GUICtrlCreateCheckbox("Clear existing data before insert (DELETE FROM ...)", 35, $y + 22, 380, 20)
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    GUICtrlSetState($g_chkClearFirst, $GUI_CHECKED)
    Global $g_chkTransaction = GUICtrlCreateCheckbox("Use transaction (rollback on error)", 35, $y + 44, 380, 20)
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
    GUICtrlSetState($g_chkTransaction, $GUI_CHECKED)

    ; Area do SQL gerado
    $y = 218
    GUICtrlCreateLabel("Generated SQL:", 20, $y, 200, 18)
    GUICtrlSetFont(-1, 9, 700, 0, "Segoe UI")

    $g_hLog = GUICtrlCreateEdit("-- Click 'GENERATE SQL' to view the SQL statements..." & @CRLF & _
        "-- The SQL is based on the data entered in the previous tabs.", _
        20, $y + 22, $APP_WIDTH - 40, $APP_HEIGHT - 360, _
        BitOR($ES_MULTILINE, $WS_VSCROLL, $WS_HSCROLL, $ES_READONLY, $ES_AUTOVSCROLL))
    GUICtrlSetFont(-1, 9, 400, 0, "Courier New")
    GUICtrlSetBkColor($g_hLog, 0x1E1E1E)
    GUICtrlSetColor($g_hLog, 0x00FF00)

    ; Log de execucao
    $y = $APP_HEIGHT - 130
    GUICtrlCreateLabel("Execution log:", 20, $y, 200, 18)
    GUICtrlSetFont(-1, 9, 700, 0, "Segoe UI")
    Global $g_hExecLog = GUICtrlCreateEdit("", 20, $y + 20, $APP_WIDTH - 40, 60, _
        BitOR($ES_MULTILINE, $WS_VSCROLL, $ES_READONLY))
    GUICtrlSetFont(-1, 8, 400, 0, "Segoe UI")
    GUICtrlSetBkColor($g_hExecLog, 0xFFFFF0)
EndFunc

;=============================================================================
; HELPER: BOTOES CRUD PADRAO
;=============================================================================
Func _CreateCRUDButtons($y, $sAdd, $sEdit, $sDel, $sDelAll)
    Local $btnAdd = GUICtrlCreateButton("+ Add", 20, $y - 2, 110, 28)
    GUICtrlSetOnEvent($btnAdd, $sAdd)
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")

    Local $btnEdit = GUICtrlCreateButton("Edit", 140, $y - 2, 90, 28)
    GUICtrlSetOnEvent($btnEdit, $sEdit)
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")

    Local $btnDel = GUICtrlCreateButton("- Remove", 240, $y - 2, 100, 28)
    GUICtrlSetOnEvent($btnDel, $sDel)
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")

    Local $btnDelAll = GUICtrlCreateButton("Clear All", 350, $y - 2, 105, 28)
    GUICtrlSetOnEvent($btnDelAll, $sDelAll)
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")

    Local $btnImp = GUICtrlCreateButton("Import CSV...", 465, $y - 2, 120, 28)
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")

    Local $btnExp = GUICtrlCreateButton("Export CSV...", 595, $y - 2, 120, 28)
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
EndFunc

;=============================================================================
; DIALOGO GENERICO PARA ADICIONAR/EDITAR LINHA
;=============================================================================
Func _ShowRowDialog($sTitle, $aFields, $aValues)
    ; $aFields = array de nomes de campos
    ; $aValues = array de valores atuais (para edicao)
    Local $nF = UBound($aFields)
    Local $hDlg = GUICreate($sTitle, 480, 80 + $nF * 36 + 50, -1, -1, _
        BitOR($WS_DLGFRAME, $WS_POPUP, $WS_CAPTION))

    Local $aInputs[$nF]
    Local $y = 20
    For $i = 0 To $nF - 1
        GUICtrlCreateLabel($aFields[$i] & ":", 20, $y + 3, 170, 20)
        GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
        $aInputs[$i] = GUICtrlCreateInput(IsArray($aValues) ? $aValues[$i] : "", 200, $y, 250, 24)
        GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
        $y += 36
    Next

    Local $btnOK = GUICtrlCreateButton("OK", 200, $y + 10, 100, 30)
    GUICtrlSetFont(-1, 9, 700, 0, "Segoe UI")
    Local $btnCancel = GUICtrlCreateButton("Cancel", 310, $y + 10, 100, 30)
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")

    GUISetState(@SW_SHOW, $hDlg)

    Local $aResult[$nF]
    Local $bOK = False

    While 1
        Local $nMsg = GUIGetMsg()
        If $nMsg = $GUI_EVENT_CLOSE Or $nMsg = $btnCancel Then
            ExitLoop
        ElseIf $nMsg = $btnOK Then
            For $i = 0 To $nF - 1
                $aResult[$i] = GUICtrlRead($aInputs[$i])
            Next
            $bOK = True
            ExitLoop
        EndIf
    WEnd

    GUIDelete($hDlg)
    If $bOK Then Return $aResult
    Return 0
EndFunc

;=============================================================================
; HANDLERS DE EVENTOS - CRUD CALENDARIOS
;=============================================================================
Func _Cal_Add()
    Local $aFields[] = ["Calendar ID (no spaces)", "Calendar name", "Start day (1=Mon..7=Sun)", "Start time (HH:MM)", "End day (1=Mon..7=Sun)", "End time (HH:MM)"]
    Local $aVals[] = ["Cal_1x8", "Cal 1x8", "1", "08:00", "1", "17:00"]
    Local $aResult = _ShowRowDialog("Add calendar", $aFields, $aVals)
    If IsArray($aResult) Then
        _AddCalendar($aResult[0], $aResult[1], $aResult[2], $aResult[3], $aResult[4], $aResult[5])
    EndIf
EndFunc
;~ Func _Cal_Edit() EndFunc
Func _Cal_Del()
    Local $iSel = _GUICtrlListView_GetSelectedIndices($g_hLV_Cal)
    If $iSel <> "" Then GUICtrlDelete(_GUICtrlListView_GetItemText($g_hLV_Cal, $iSel))
    _GUICtrlListView_DeleteItem($g_hLV_Cal, $iSel)
EndFunc
Func _Cal_DelAll()
	_GUICtrlListView_DeleteAllItems($g_hLV_Cal)
EndFunc

;=============================================================================
; HANDLERS - MAQUINAS
;=============================================================================
Func _Mach_Add()
    Local $aFields[] = ["Site ID", "Site Nome", "CT ID (sem espacos)", "CT Nome", "Tipo CT (2=Finito/4=Infinito)", "Secao ID", "Secao Nome", "Machine ID", "Machine name", "Machine type (NR/BA/RN/CU)", "Cal. Abertura ID", "Capacity calendar ID"]
    Local $aVals[] = ["LYON","LYON","Milling","Fresagem","2","Sec_A","Secao A","Milling_1","Fresagem 1","NR","Cal_1x8","Cal_1x8"]
    Local $aResult = _ShowRowDialog("Add machine", $aFields, $aVals)
    If IsArray($aResult) Then
        _AddMachine($aResult[0],$aResult[1],$aResult[2],$aResult[3],$aResult[4],$aResult[5],$aResult[6],$aResult[7],$aResult[8],$aResult[9],$aResult[10],$aResult[11])
    EndIf
EndFunc
Func _Mach_Edit()
EndFunc
Func _Mach_Del()
	_GUICtrlListView_DeleteItem($g_hLV_Mach, _GUICtrlListView_GetSelectedIndices($g_hLV_Mach))
EndFunc
Func _Mach_DelAll()
	_GUICtrlListView_DeleteAllItems($g_hLV_Mach)
EndFunc

;=============================================================================
; HANDLERS - OPERACOES
;=============================================================================
Func _Ops_Add()
    Local $aFields[] = ["Operation ID","Operation name","CT ID","Machine ID","Qtd Referencia","Dur Referencia","Unidade (J/H/C)","Prep (HH)","Time Out (HH)","Interruptible (1/0)"]
    Local $aResult = _ShowRowDialog("Add operation", $aFields, 0)
    If IsArray($aResult) Then _AddOperation($aResult[0],$aResult[1],$aResult[2],$aResult[3],$aResult[4],$aResult[5],$aResult[6],$aResult[7],$aResult[8],$aResult[9])
EndFunc
Func _Ops_Edit()
EndFunc
Func _Ops_Del()
	_GUICtrlListView_DeleteItem($g_hLV_Ops, _GUICtrlListView_GetSelectedIndices($g_hLV_Ops))
EndFunc
Func _Ops_DelAll()
	_GUICtrlListView_DeleteAllItems($g_hLV_Ops)
EndFunc

;=============================================================================
; HANDLERS - ROTEIROS
;=============================================================================
Func _Rout_Add()
    Local $aFields[] = ["Routing ID","Routing name","Codigo Fase (10,20,...)","Operation ID","Nome Fase"]
    Local $aResult = _ShowRowDialog("Add routing", $aFields, 0)
    If IsArray($aResult) Then _AddRouting($aResult[0],$aResult[1],$aResult[2],$aResult[3],$aResult[4])
EndFunc
Func _Rout_Edit()
EndFunc
Func _Rout_Del()
	_GUICtrlListView_DeleteItem($g_hLV_Rout, _GUICtrlListView_GetSelectedIndices($g_hLV_Rout))
EndFunc
Func _Rout_DelAll()
	_GUICtrlListView_DeleteAllItems($g_hLV_Rout)
EndFunc

;=============================================================================
; HANDLERS - MATERIAIS
;=============================================================================
Func _Mat_Add()
    Local $aFields[] = ["Item ID","Item name","Tipo (MP/SF/PF)","Versao","Routing ID","On-hand qty"]
    Local $aResult = _ShowRowDialog("Add item", $aFields, 0)
    If IsArray($aResult) Then _AddMaterial($aResult[0],$aResult[1],$aResult[2],$aResult[3],$aResult[4],$aResult[5])
EndFunc
Func _Mat_Edit()
EndFunc
Func _Mat_Del()
	_GUICtrlListView_DeleteItem($g_hLV_Mat, _GUICtrlListView_GetSelectedIndices($g_hLV_Mat))
EndFunc
Func _Mat_DelAll()
	_GUICtrlListView_DeleteAllItems($g_hLV_Mat)
EndFunc

;=============================================================================
; HANDLERS - BOM
;=============================================================================
Func _BOM_Add()
    Local $aFields[] = ["Parent item ID","Versao Pai","Component item ID","Routing ID","Codigo Fase","Qtd Referencia","Qtd Necessaria"]
    Local $aResult = _ShowRowDialog("Add BOM row", $aFields, 0)
    If IsArray($aResult) Then _AddBOM($aResult[0],$aResult[1],$aResult[2],$aResult[3],$aResult[4],$aResult[5],$aResult[6])
EndFunc
Func _BOM_Edit()
EndFunc
Func _BOM_Del()
	_GUICtrlListView_DeleteItem($g_hLV_BOM, _GUICtrlListView_GetSelectedIndices($g_hLV_BOM))
EndFunc
Func _BOM_DelAll()
	_GUICtrlListView_DeleteAllItems($g_hLV_BOM)
EndFunc

;=============================================================================
; HANDLERS - WO
;=============================================================================
Func _WO_Add()
    Local $aFields[] = ["Numero OP","Item ID","Routing ID","Versao","Quantidade","Data Inicio (dd/mm/aaaa hh:mm)","Data Fim (dd/mm/aaaa hh:mm)"]
    Local $aResult = _ShowRowDialog("Add work order", $aFields, 0)
    If IsArray($aResult) Then _AddWO($aResult[0],$aResult[1],$aResult[2],$aResult[3],$aResult[4],$aResult[5],$aResult[6])
EndFunc
Func _WO_Edit()
EndFunc
Func _WO_Del()
	_GUICtrlListView_DeleteItem($g_hLV_WO, _GUICtrlListView_GetSelectedIndices($g_hLV_WO))
EndFunc
Func _WO_DelAll()
	_GUICtrlListView_DeleteAllItems($g_hLV_WO)
EndFunc

;=============================================================================
; HANDLERS - WO LINKS
;=============================================================================
Func _WOL_Add()
    Local $aFields[] = ["OP Predecessor","Pred routing","Fase Pred","OP Sucessor","Succ routing","Fase Suc","Link type (FS/SS/FF)"]
    Local $aResult = _ShowRowDialog("Add WO link", $aFields, 0)
    If IsArray($aResult) Then GUICtrlCreateListViewItem($aResult[0] & "|" & $aResult[1] & "|" & $aResult[2] & "|" & $aResult[3] & "|" & $aResult[4] & "|" & $aResult[5] & "|" & $aResult[6], $g_hLV_WOL)
EndFunc
Func _WOL_Edit()
EndFunc
Func _WOL_Del()
	_GUICtrlListView_DeleteItem($g_hLV_WOL, _GUICtrlListView_GetSelectedIndices($g_hLV_WOL))
EndFunc
Func _WOL_DelAll()
	_GUICtrlListView_DeleteAllItems($g_hLV_WOL)
EndFunc

;=============================================================================
; HANDLERS - RECURSOS SECUNDARIOS
;=============================================================================
Func _SR_Add()
    Local $aFields[] = ["Operation ID","CT ID","Machine ID","Qualification ID","Capacity calendar ID"]
    Local $aResult = _ShowRowDialog("Add secondary resource", $aFields, 0)
    If IsArray($aResult) Then GUICtrlCreateListViewItem($aResult[0] & "|" & $aResult[1] & "|" & $aResult[2] & "|" & $aResult[3] & "|" & $aResult[4], $g_hLV_SR)
EndFunc
Func _SR_Edit()
EndFunc
Func _SR_Del()
	_GUICtrlListView_DeleteItem($g_hLV_SR, _GUICtrlListView_GetSelectedIndices($g_hLV_SR))
EndFunc
Func _SR_DelAll()
	_GUICtrlListView_DeleteAllItems($g_hLV_SR)
EndFunc

;=============================================================================
; HANDLERS - CAPACIDADE
;=============================================================================
Func _Cap_Add()
    Local $aFields[] = ["Capacity calendar ID","Dia Inicio (1-7)","Start time (HH:MM)","Dia Fim (1-7)","End time (HH:MM)","Resource count"]
    Local $aResult = _ShowRowDialog("Add capacity calendar", $aFields, 0)
    If IsArray($aResult) Then GUICtrlCreateListViewItem($aResult[0] & "|" & $aResult[1] & "|" & $aResult[2] & "|" & $aResult[3] & "|" & $aResult[4] & "|" & $aResult[5], $g_hLV_Cap)
EndFunc
Func _Cap_Edit()
EndFunc
Func _Cap_Del()
	_GUICtrlListView_DeleteItem($g_hLV_Cap, _GUICtrlListView_GetSelectedIndices($g_hLV_Cap))
EndFunc
Func _Cap_DelAll()
	_GUICtrlListView_DeleteAllItems($g_hLV_Cap)
EndFunc

;=============================================================================
; HANDLERS - ESTOQUE
;=============================================================================
Func _Stk_Add()
    Local $aFields[] = ["Item ID","Routing ID","Versao","Movement date (dd/mm/yyyy)","Quantidade"]
    Local $aResult = _ShowRowDialog("Add inventory movement", $aFields, 0)
    If IsArray($aResult) Then GUICtrlCreateListViewItem($aResult[0] & "|" & $aResult[1] & "|" & $aResult[2] & "|" & $aResult[3] & "|" & $aResult[4], $g_hLV_Stk)
EndFunc
Func _Stk_Edit()
EndFunc
Func _Stk_Del()
	_GUICtrlListView_DeleteItem($g_hLV_Stk, _GUICtrlListView_GetSelectedIndices($g_hLV_Stk))
EndFunc
Func _Stk_DelAll()
	_GUICtrlListView_DeleteAllItems($g_hLV_Stk)
EndFunc

;=============================================================================
; CONEXAO COM O BANCO
;=============================================================================
Func _TestConnection()
    $g_sServer   = GUICtrlRead($g_edtServer)
    $g_sDatabase = GUICtrlRead($g_edtDatabase)
    Local $sAuth  = GUICtrlRead($g_cmbAuth)

    If $g_sServer = "" Or $g_sDatabase = "" Then
        MsgBox(48, "Error", "Server and database are required.")
        Return
    EndIf

    If $sAuth = "Windows Authentication" Then
        $g_sConnStr = "Driver={SQL Server};Server=" & $g_sServer & ";Database=" & $g_sDatabase & ";Trusted_Connection=yes;"
    Else
        $g_sConnStr = "Driver={SQL Server};Server=" & $g_sServer & ";Database=" & $g_sDatabase & ";UID=" & GUICtrlRead($g_edtUser) & ";PWD=" & GUICtrlRead($g_edtPass) & ";"
    EndIf

    GUICtrlSetData($g_edtConnStr, $g_sConnStr)

    ; Testa conexao via ADO
    Local $oConn = ObjCreate("ADODB.Connection")
    If Not IsObj($oConn) Then
        GUICtrlSetData($g_lblStatus, "Error: ADO not available")
        GUICtrlSetColor($g_lblStatus, 0xCC0000)
        MsgBox(16, "Error", "Could not create ADODB.Connection." & @CRLF & "Check that the SQL Server ODBC / OLE DB provider is installed.")
        Return
    EndIf

    $oConn.ConnectionString = $g_sConnStr
    $oConn.ConnectionTimeout = 5

    Local $bOK = False
    Local $sErr = ""
    ; Try to open
    ; Usar On Error seria ideal, aqui simulamos com objecto de erro
    $oConn.Open($g_sConnStr)
    If $oConn.State = 1 Then ; adStateOpen = 1
        $bOK = True
        $oConn.Close()
    Else
        $sErr = "Could not open the connection."
    EndIf

    If $bOK Then
        $g_bConnected = True
        GUICtrlSetData($g_lblStatus, "Conectado: " & $g_sDatabase & "@" & $g_sServer)
        GUICtrlSetColor($g_lblStatus, 0x007700)
        MsgBox(64, "Connection", "Successfully connected to the database!" & @CRLF & @CRLF & "Database: " & $g_sDatabase & @CRLF & "Servidor: " & $g_sServer)
    Else
        $g_bConnected = False
        GUICtrlSetData($g_lblStatus, "Connection error")
        GUICtrlSetColor($g_lblStatus, 0xCC0000)
        MsgBox(16, "Connection error", "Could not connect to the database." & @CRLF & @CRLF & $sErr & @CRLF & @CRLF & "Connection string used:" & @CRLF & $g_sConnStr)
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

;=============================================================================
; APLICAR MODULOS
;=============================================================================
Func _ApplyModules()
    _RefreshModuleFlags(False)
EndFunc

;=============================================================================
; GERACAO DO SQL
;=============================================================================
Func _GenerateSQL()
    Local $sSQL = ""
    Local $bClear = (GUICtrlRead($g_chkClearFirst) = $GUI_CHECKED)
    Local $bTrans = (GUICtrlRead($g_chkTransaction) = $GUI_CHECKED)

    $sSQL &= "-- =============================================" & @CRLF
    $sSQL &= "-- ORTEMS TOOLBOX - Import SQL" & @CRLF
    $sSQL &= "-- Gerado em: " & @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & @CRLF
    $sSQL &= "-- Database: " & GUICtrlRead($g_edtDatabase) & @CRLF
    $sSQL &= "-- =============================================" & @CRLF & @CRLF

    If $bTrans Then $sSQL &= "BEGIN TRANSACTION;" & @CRLF & @CRLF

    ; Limpar tabelas (ordem inversa de FK)
    If $bClear Then
        $sSQL &= "-- ===== LIMPEZA DAS TABELAS =====" & @CRLF
        $sSQL &= _GetClearSQL() & @CRLF
    EndIf

    ; Calendarios
    $sSQL &= "-- ===== CALENDARIOS (B_CAL / B_PERI) =====" & @CRLF
    $sSQL &= _GenerateCalSQL() & @CRLF

    ; Maquinas
    $sSQL &= "-- ===== MAQUINAS E CENTROS DE TRABALHO =====" & @CRLF
    $sSQL &= _GenerateMachSQL() & @CRLF

    ; Operacoes
    $sSQL &= "-- ===== OPERACOES (B_OPE / B_CADE) =====" & @CRLF
    $sSQL &= _GenerateOpsSQL() & @CRLF

    ; Roteiros
    $sSQL &= "-- ===== ROTEIROS (B_GAMM / B_PHAS) =====" & @CRLF
    $sSQL &= _GenerateRoutSQL() & @CRLF

    ; Materiais
    $sSQL &= "-- ===== MATERIAIS (B_ART / B_VER_ART) =====" & @CRLF
    $sSQL &= _GenerateMatSQL() & @CRLF

    ; BOM
    $sSQL &= "-- ===== BOM (B_NOME) =====" & @CRLF
    $sSQL &= _GenerateBOMSQL() & @CRLF

    ; Ordens de Producao
    $sSQL &= "-- ===== ORDENS DE PRODUCAO (B_OF) =====" & @CRLF
    $sSQL &= _GenerateWOSQL() & @CRLF

    ; WO Links (se modulo ativo)
    If $g_bModWOL Then
        $sSQL &= "-- ===== LINKS WO (B_PROF) =====" & @CRLF
        $sSQL &= _GenerateWOLSQL() & @CRLF
    EndIf

    If $bTrans Then
        $sSQL &= @CRLF & "COMMIT TRANSACTION;" & @CRLF
        $sSQL &= "-- Se ocorrer erro, execute: ROLLBACK TRANSACTION;" & @CRLF
    EndIf

    $sSQL &= @CRLF & "-- ===== FIM DO SCRIPT =====" & @CRLF

    GUICtrlSetData($g_hLog, $sSQL)
    _Log("SQL generated successfully. " & StringLen($sSQL) & " caracteres.")

    ; Mudar para a aba SQL
    GUICtrlSetState($g_hTab, $g_iTabSQL)
EndFunc

Func _GetClearSQL()
    Local $s = ""
    $s &= "DELETE FROM B_PROF WHERE 1=1;" & @CRLF
    $s &= "DELETE FROM B_NOME WHERE 1=1;" & @CRLF
    $s &= "DELETE FROM E_OF2 WHERE 1=1;" & @CRLF
    $s &= "DELETE FROM B_OF WHERE 1=1;" & @CRLF
    $s &= "DELETE FROM B_SER WHERE 1=1;" & @CRLF
    $s &= "DELETE FROM B_VER_ART WHERE 1=1;" & @CRLF
    $s &= "DELETE FROM B_ART WHERE 1=1;" & @CRLF
    $s &= "DELETE FROM B_PHAS WHERE 1=1;" & @CRLF
    $s &= "DELETE FROM B_GAMM WHERE 1=1;" & @CRLF
    $s &= "DELETE FROM B_CADE WHERE 1=1;" & @CRLF
    $s &= "DELETE FROM B_OPE WHERE 1=1;" & @CRLF
    $s &= "DELETE FROM B_MACH WHERE 1=1;" & @CRLF
    $s &= "DELETE FROM B_SECT WHERE 1=1;" & @CRLF
    $s &= "DELETE FROM B_ILOT WHERE 1=1;" & @CRLF
    $s &= "DELETE FROM B_ZONE WHERE 1=1;" & @CRLF
    $s &= "DELETE FROM B_PERI WHERE 1=1;" & @CRLF
    $s &= "DELETE FROM B_CAL WHERE 1=1;" & @CRLF
    Return $s
EndFunc

Func _GenerateCalSQL()
    Local $s = ""
    Local $nRows = _GUICtrlListView_GetItemCount($g_hLV_Cal)
    For $i = 0 To $nRows - 1
        Local $sID    = _GUICtrlListView_GetItemText($g_hLV_Cal, $i, 0)
        Local $sNome  = _GUICtrlListView_GetItemText($g_hLV_Cal, $i, 1)
        Local $sDiaI  = _GUICtrlListView_GetItemText($g_hLV_Cal, $i, 2)
        Local $sHoraI = _GUICtrlListView_GetItemText($g_hLV_Cal, $i, 3)
        Local $sDiaF  = _GUICtrlListView_GetItemText($g_hLV_Cal, $i, 4)
        Local $sHoraF = _GUICtrlListView_GetItemText($g_hLV_Cal, $i, 5)

        ; Insercao em B_CAL (um registro por ID unico)
        $s &= "IF NOT EXISTS (SELECT 1 FROM B_CAL WHERE NOCALHEBD='" & $sID & "')" & @CRLF
        $s &= "    INSERT INTO B_CAL (NOCALHEBD, NOMCAL) VALUES ('" & $sID & "', '" & $sNome & "');" & @CRLF

        ; Insercao dos periodos em B_PERI
        Local $sHI = StringReplace($sHoraI, ":", "")
        Local $sHF = StringReplace($sHoraF, ":", "")
        $s &= "INSERT INTO B_PERI (NOCALHEBD, NOJOUR_DEB, DEB_PERIO, NOJOUR_FIN, FIN_PERIO) " & @CRLF
        $s &= "    VALUES ('" & $sID & "', " & $sDiaI & ", " & $sHI & ", " & $sDiaF & ", " & $sHF & ");" & @CRLF
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
        $s &= "    INSERT INTO B_MACH (MACHINE, LIBMACH, ILOT, CODESECTI, NOZONE, MACH_MODEMACH, NOCALHEBD, CODE_CYCLE) " & @CRLF
        $s &= "    VALUES ('" & $sMaqID & "', '" & $sMaqNm & "', '" & $sCTID & "', '" & $sSecID & "', '" & $sSiteID & "', '" & $sMaqTp & "', '" & $sCalID & "', '" & $sCalCap & "');" & @CRLF
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
        Local $sID   = _GUICtrlListView_GetItemText($g_hLV_Rout, $i, 0)
        Local $sNome = _GUICtrlListView_GetItemText($g_hLV_Rout, $i, 1)
        Local $nFase = _GUICtrlListView_GetItemText($g_hLV_Rout, $i, 2)
        Local $sOp   = _GUICtrlListView_GetItemText($g_hLV_Rout, $i, 3)
        Local $sNomF = _GUICtrlListView_GetItemText($g_hLV_Rout, $i, 4)

        $s &= "IF NOT EXISTS (SELECT 1 FROM B_GAMM WHERE NOMG='" & $sID & "')" & @CRLF
        $s &= "    INSERT INTO B_GAMM (NOMG, LIBGAM) VALUES ('" & $sID & "', '" & $sNome & "');" & @CRLF
        $s &= "INSERT INTO B_PHAS (NOMG, NOPHASE, OPE, LIBPHASE) VALUES ('" & $sID & "', " & $nFase & ", '" & $sOp & "', '" & $sNomF & "');" & @CRLF
    Next
    Return $s
EndFunc

Func _GenerateMatSQL()
    Local $s = ""
    Local $nRows = _GUICtrlListView_GetItemCount($g_hLV_Mat)
    For $i = 0 To $nRows - 1
        Local $sID   = _GUICtrlListView_GetItemText($g_hLV_Mat, $i, 0)
        Local $sNome = _GUICtrlListView_GetItemText($g_hLV_Mat, $i, 1)
        Local $sTipo = _GUICtrlListView_GetItemText($g_hLV_Mat, $i, 2)
        Local $sVer  = _GUICtrlListView_GetItemText($g_hLV_Mat, $i, 3)
        Local $sRot  = _GUICtrlListView_GetItemText($g_hLV_Mat, $i, 4)
        Local $nStk  = _GUICtrlListView_GetItemText($g_hLV_Mat, $i, 5)

        $s &= "IF NOT EXISTS (SELECT 1 FROM B_ART WHERE CODEARTIC='" & $sID & "')" & @CRLF
        $s &= "    INSERT INTO B_ART (CODEARTIC, LIBARTIC, TYPEMATI, QTE_STOCK) VALUES ('" & $sID & "', '" & $sNome & "', '" & $sTipo & "', " & $nStk & ");" & @CRLF
        $s &= "INSERT INTO B_VER_ART (CODEARTIC, VER_ART, VER_DESC, NOMG, VER_EFFET_DEBUT, VER_EFFET_FIN)" & @CRLF
        $s &= "    VALUES ('" & $sID & "', '" & $sVer & "', '" & $sNome & "', '" & $sRot & "', '01/01/1995 00:00', '01/01/2050 00:00');" & @CRLF
    Next
    Return $s
EndFunc

Func _GenerateBOMSQL()
    Local $s = ""
    Local $nRows = _GUICtrlListView_GetItemCount($g_hLV_BOM)
    For $i = 0 To $nRows - 1
        Local $sPaiID  = _GUICtrlListView_GetItemText($g_hLV_BOM, $i, 0)
        Local $sVPai   = _GUICtrlListView_GetItemText($g_hLV_BOM, $i, 1)
        Local $sCompID = _GUICtrlListView_GetItemText($g_hLV_BOM, $i, 2)
        Local $sRotID  = _GUICtrlListView_GetItemText($g_hLV_BOM, $i, 3)
        Local $nFase   = _GUICtrlListView_GetItemText($g_hLV_BOM, $i, 4)
        Local $nQRef   = _GUICtrlListView_GetItemText($g_hLV_BOM, $i, 5)
        Local $nQNec   = _GUICtrlListView_GetItemText($g_hLV_BOM, $i, 6)

        $s &= "INSERT INTO B_NOME (B_V_CODEARTIC, VER_ART, CODEARTIC, NOMG, NOPHASE, CODEBASET, BES_CODEARTIC)" & @CRLF
        $s &= "    VALUES ('" & $sPaiID & "', '" & $sVPai & "', '" & $sCompID & "', '" & $sRotID & "', " & $nFase & ", " & $nQRef & ", " & $nQNec & ");" & @CRLF
    Next
    Return $s
EndFunc

Func _GenerateWOSQL()
    Local $s = ""
    Local $nRows = _GUICtrlListView_GetItemCount($g_hLV_WO)
    For $i = 0 To $nRows - 1
        Local $sNum  = _GUICtrlListView_GetItemText($g_hLV_WO, $i, 0)
        Local $sMat  = _GUICtrlListView_GetItemText($g_hLV_WO, $i, 1)
        Local $sRot  = _GUICtrlListView_GetItemText($g_hLV_WO, $i, 2)
        Local $sVer  = _GUICtrlListView_GetItemText($g_hLV_WO, $i, 3)
        Local $nQtd  = _GUICtrlListView_GetItemText($g_hLV_WO, $i, 4)
        Local $sDtI  = _GUICtrlListView_GetItemText($g_hLV_WO, $i, 5)
        Local $sDtF  = _GUICtrlListView_GetItemText($g_hLV_WO, $i, 6)

        $s &= "INSERT INTO B_OF (NOF, CODEARTIC, NOMG, VER_ART, QTEORDER, DPLUSTOT, FPLUSTARD," & @CRLF
        $s &= "    CODEGEST, ETATOF, VER_EFFET_DEBUT, VER_EFFET_FIN, MODE_UTIL)" & @CRLF
        $s &= "    VALUES ('" & $sNum & "', '" & $sMat & "', '" & $sRot & "', '" & $sVer & "', " & $nQtd & "," & @CRLF
        $s &= "    CONVERT(datetime, '" & $sDtI & "', 103), CONVERT(datetime, '" & $sDtF & "', 103)," & @CRLF
        $s &= "    'F', 'S', '01/01/1995 00:00', '01/01/2050 00:00', 'C');" & @CRLF
    Next
    Return $s
EndFunc

Func _GenerateWOLSQL()
    Local $s = ""
    Local $nRows = _GUICtrlListView_GetItemCount($g_hLV_WOL)
    For $i = 0 To $nRows - 1
        Local $sNOF_P = _GUICtrlListView_GetItemText($g_hLV_WOL, $i, 0)
        Local $sNOMG_P= _GUICtrlListView_GetItemText($g_hLV_WOL, $i, 1)
        Local $sFas_P = _GUICtrlListView_GetItemText($g_hLV_WOL, $i, 2)
        Local $sNOF_S = _GUICtrlListView_GetItemText($g_hLV_WOL, $i, 3)
        Local $sNOMG_S= _GUICtrlListView_GetItemText($g_hLV_WOL, $i, 4)
        Local $sFas_S = _GUICtrlListView_GetItemText($g_hLV_WOL, $i, 5)
        Local $sTipo  = _GUICtrlListView_GetItemText($g_hLV_WOL, $i, 6)

        $s &= "INSERT INTO B_PROF (NOF, NOMG, NOPHASE, B_O_NOF, B_P_NOMG, B_P_NOPHASE, PROF_TYPEPREC)" & @CRLF
        $s &= "    VALUES ('" & $sNOF_S & "', '" & $sNOMG_S & "', " & $sFas_S & ", '" & $sNOF_P & "', '" & $sNOMG_P & "', " & $sFas_P & ", '" & $sTipo & "');" & @CRLF
    Next
    Return $s
EndFunc

;=============================================================================
; EXECUTAR SQL NO BANCO
;=============================================================================
Func _ExecuteSQL()
    If Not $g_bConnected Then
        MsgBox(48, "Warning", "You are not connected to the database." & @CRLF & "Go to the '1. Database' tab and test the connection first.")
        Return
    EndIf

    Local $sSQL = GUICtrlRead($g_hLog)
    If $sSQL = "" Or StringInStr($sSQL, "INSERT") = 0 Then
        MsgBox(48, "Warning", "No SQL has been generated yet. Click 'GENERATE SQL' first.")
        Return
    EndIf

    Local $nRet = MsgBox(4 + 48, "Confirm execution", _
        "You are about to run SQL on the database:" & @CRLF & @CRLF & _
        "Database: " & $g_sDatabase & @CRLF & "Servidor: " & $g_sServer & @CRLF & @CRLF & _
        "This action can DELETE and RECREATE existing data!" & @CRLF & @CRLF & _
        "Do you want to continue?")

    If $nRet <> 6 Then Return  ; 6 = Yes

    Local $oConn = ObjCreate("ADODB.Connection")
    If Not IsObj($oConn) Then
        _Log("ERROR: Could not create the connection object.")
        Return
    EndIf

    $oConn.Open($g_sConnStr)
    If $oConn.State <> 1 Then
        _Log("ERROR: Could not open the connection.")
        Return
    EndIf

    ; Executa cada statement separado por ";"
    Local $aStmts = StringSplit($sSQL, ";", 1)
    Local $nOK = 0, $nErr = 0

    For $i = 1 To $aStmts[0]
        Local $sStmt = StringStripWS($aStmts[$i], 3)
        If $sStmt = "" Or StringLeft($sStmt, 2) = "--" Then ContinueLoop

        $oConn.Execute($sStmt)
        If @error Then
            $nErr += 1
            _Log("ERROR in statement " & $i & ": " & $sStmt)
        Else
            $nOK += 1
        EndIf
    Next

    $oConn.Close()
    _Log("Execucao concluida: " & $nOK & " statements OK, " & $nErr & " errors.")
    MsgBox(64, "Resultado", "SQL executed successfully on the database!" & @CRLF & @CRLF & $nOK & " statements executed successfully." & @CRLF & ($nErr > 0 ? $nErr & " errors found (check the log)." : "No errors!"))
EndFunc

;=============================================================================
; LIMPAR BANCO
;=============================================================================
Func _ClearDatabase()
    If Not $g_bConnected Then
        MsgBox(48, "Warning", "Not connected to the database.")
        Return
    EndIf

    Local $nRet = MsgBox(4 + 16, "Confirm clear", _
        "WARNING: This action will DELETE ALL demo data from the database!" & @CRLF & @CRLF & _
        "Database: " & $g_sDatabase & @CRLF & @CRLF & "Do you want to continue?")

    If $nRet = 6 Then
        Local $oConn = ObjCreate("ADODB.Connection")
        $oConn.Open($g_sConnStr)
        $oConn.Execute(_GetClearSQL())
        $oConn.Close()
        _Log("Database cleared successfully.")
        MsgBox(64, "OK", "Demo data removed from the database successfully.")
    EndIf
EndFunc

;=============================================================================
; IMPORTAR / EXPORTAR EXCEL
;=============================================================================
Func _ImportExcel()
    Local $sFile = FileOpenDialog("Select Excel/CSV file", @WorkingDir, "Excel/CSV (*.xlsx;*.xlsm;*.csv)|Todos (*.*)")
    If $sFile = "" Then Return

    If StringRight($sFile, 4) = ".csv" Then
        _Log("Importando CSV: " & $sFile)
        MsgBox(64, "Import", "CSV import not implemented yet. File: " & $sFile & @CRLF & "Select the corresponding target tab.")
    Else
        MsgBox(64, "Import Excel", "Para importar do Excel Toolbox original:" & @CRLF & @CRLF & _
            "1. Open the file: " & $sFile & @CRLF & _
            "2. Export each SV_ tab as CSV" & @CRLF & _
            "3. Use 'Import CSV' in each data tab if needed" & @CRLF & @CRLF & _
            "(Suporte nativo a .xlsm requer Microsoft Excel instalado)")
    EndIf
EndFunc

Func _ExportExcel()
    Local $sSaveFile = FileSaveDialog("Export data", @WorkingDir, "CSV (*.csv)|Todos (*.*)", 16, "ortems_demo_export.csv")
    If $sSaveFile = "" Then Return

    Local $sCSV = ""
    ; Exporta calendarios
    $sCSV &= "TIPO;ID;NOME;DIA_INICIO;HORA_INICIO;DIA_FIM;HORA_FIM" & @CRLF
    Local $n = _GUICtrlListView_GetItemCount($g_hLV_Cal)
    For $i = 0 To $n - 1
        $sCSV &= "CAL;" & _GUICtrlListView_GetItemText($g_hLV_Cal, $i, 0) & ";" & _GUICtrlListView_GetItemText($g_hLV_Cal, $i, 1) & ";" & _GUICtrlListView_GetItemText($g_hLV_Cal, $i, 2) & ";" & _GUICtrlListView_GetItemText($g_hLV_Cal, $i, 3) & ";" & _GUICtrlListView_GetItemText($g_hLV_Cal, $i, 4) & ";" & _GUICtrlListView_GetItemText($g_hLV_Cal, $i, 5) & @CRLF
    Next

    FileWrite($sSaveFile, $sCSV)
    _Log("Data exported to: " & $sSaveFile)
    MsgBox(64, "Export", "Data exported successfully to:" & @CRLF & $sSaveFile)
EndFunc

Func _SaveSQL()
    Local $sSQL = GUICtrlRead($g_hLog)
    If $sSQL = "" Then
        MsgBox(48, "Warning", "No SQL to save. Click 'GENERATE SQL' first.")
        Return
    EndIf
    Local $sFile = FileSaveDialog("Save SQL", @WorkingDir, "SQL (*.sql)|Todos (*.*)", 16, "ortems_demo.sql")
    If $sFile <> "" Then
        FileWrite($sFile, $sSQL)
        _Log("SQL saved to: " & $sFile)
        MsgBox(64, "Salvo", "SQL saved successfully to:" & @CRLF & $sFile)
    EndIf
EndFunc

;=============================================================================
; LOG
;=============================================================================
Func _Log($sMsg)
    Local $sCurrent = GUICtrlRead($g_hExecLog)
    GUICtrlSetData($g_hExecLog, "[" & @HOUR & ":" & @MIN & ":" & @SEC & "] " & $sMsg & @CRLF & $sCurrent)
EndFunc

;=============================================================================
; FECHAR
;=============================================================================

;=============================================================================
; UI: MODULE SELECTION -> ENABLE/DISABLE TABS (real-time)
;=============================================================================
Func _OnModuleSelectionChanged()
    _RefreshModuleFlags(True)
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
            GUICtrlSetState($g_chkWOL, $GUI_DISABLE)
            GUICtrlSetState($g_chkWOL, $GUI_UNCHECKED)
        EndIf
        If IsDeclared("g_chkSR") Then
            GUICtrlSetState($g_chkSR, $GUI_DISABLE)
            GUICtrlSetState($g_chkSR, $GUI_UNCHECKED)
        EndIf
        $g_bModWOL = False
        $g_bModSR  = False
    Else
        If IsDeclared("g_chkWOL") Then GUICtrlSetState($g_chkWOL, $GUI_ENABLE)
        If IsDeclared("g_chkSR") Then GUICtrlSetState($g_chkSR, $GUI_ENABLE)
    EndIf

    _UpdateTabAvailability()

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

        MsgBox(64, "Modules applied", "Mode: " & $sMode & @CRLF & "Enabled modules:" & $sMods & @CRLF & @CRLF & _
            "The data tabs have been enabled/disabled according to your selections.")
    EndIf
EndFunc

Func _UpdateTabAvailability()
    If $g_hTabHandle = 0 Then Return

    Local $hTab = $g_hTabHandle
    Local $bPS = $g_bModPS
    Local $bMP = $g_bModMP

    ; Core tabs
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
    Local $bCap  = ($g_bModLR Or $bMP) ; Capacity shown when needed for Limited Resources or in MP mode
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

    ; If current tab became disabled, move back to Modules
    Local $iCur = _GUICtrlTab_GetCurSel($hTab)
    If ($iCur = $g_iTabOps And Not $bOps) Or _
       ($iCur = $g_iTabRout And Not $bRout) Or _
       ($iCur = $g_iTabBOM And Not $bBOM) Or _
       ($iCur = $g_iTabWO And Not $bWO) Or _
       ($iCur = $g_iTabWOL And Not $bWOL) Or _
       ($iCur = $g_iTabSR And Not $bSR) Or _
       ($iCur = $g_iTabCap And Not $bCap) Or _
       ($iCur = $g_iTabStk And Not $bStk) Then
        _GUICtrlTab_SetCurSel($hTab, $g_iTabMod)
    EndIf
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

    ; Bottom buttons
    Local $yBtn = $bottomBarTop + 12
    If IsDeclared("g_btnImportXLS") Then GUICtrlSetPos($g_btnImportXLS, 10, $yBtn, 140, 32)
    If IsDeclared("g_btnExportXLS") Then GUICtrlSetPos($g_btnExportXLS, 160, $yBtn, 140, 32)
    If IsDeclared("g_btnClearDB")   Then GUICtrlSetPos($g_btnClearDB,   310, $yBtn, 130, 32)
    If IsDeclared("g_btnGenerate")  Then GUICtrlSetPos($g_btnGenerate,  460, $yBtn, 130, 32)
    If IsDeclared("g_btnExecute")   Then GUICtrlSetPos($g_btnExecute,   600, $yBtn, 170, 32)
    If IsDeclared("g_btnSaveSQL")   Then GUICtrlSetPos($g_btnSaveSQL,   780, $yBtn, 130, 32)

    ; Footer
    If $g_lblFooter <> 0 Then GUICtrlSetPos($g_lblFooter, $w - 220, $yBtn + 13, 210, 18)

    ; Resize ListViews/Edit areas to the available height (best-effort)
    Local $bottomContent = $bottomBarTop - 15
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

    ; SQL tab: resize main SQL box and exec log if present
    If $g_hLog <> 0 Then
        Local $p = ControlGetPos($g_hMain, "", $g_hLog)
        If Not @error Then
            Local $newW = $w - 40
            Local $newH = ($bottomContent - 105) - $p[1]
            If $newH < 120 Then $newH = 120
            GUICtrlSetPos($g_hLog, $p[0], $p[1], $newW, $newH)
        EndIf
    EndIf
    If IsDeclared("g_hExecLog") Then
        Local $yExec = $bottomContent - 80
        If $yExec < 260 Then $yExec = 260
        GUICtrlSetPos($g_hExecLog, 20, $yExec, $w - 40, 60)
    EndIf

    ; Modules tab: keep hint aligned (optional)
    If IsDeclared("g_btnApplyMod") Then
        Local $yApply = $bottomContent - 55
        If $yApply < 520 Then $yApply = 520
        GUICtrlSetPos($g_btnApplyMod, 20, $yApply, 250, 35)
        If IsDeclared("g_lblModHint") Then GUICtrlSetPos($g_lblModHint, 290, $yApply + 8, $w - 320, 20)
    EndIf
EndFunc

Func _ResizeToBottom($ctrlId, $w, $bottomY)
    If $ctrlId = 0 Then Return
    Local $p = ControlGetPos($g_hMain, "", $ctrlId)
    If @error Then Return
    Local $newW = $w - 40
    Local $newH = $bottomY - $p[1]
    If $newH < 120 Then $newH = 120
    GUICtrlSetPos($ctrlId, $p[0], $p[1], $newW, $newH)
EndFunc

; Enforce a minimum window size (prevents overlap when resizing smaller than the UI layout)
Func _WM_GETMINMAXINFO($hWnd, $iMsg, $wParam, $lParam)
    Local $tMMI = DllStructCreate("long;long;long;long;long;long;long;long;long;long", $lParam)
    DllStructSetData($tMMI, 7, $MIN_W) ; ptMinTrackSize.x
    DllStructSetData($tMMI, 8, $MIN_H) ; ptMinTrackSize.y
    Return 0
EndFunc


Func _OnClose()
    Local $nRet = MsgBox(4 + 32, "Sair", "Deseja sair do Ortems Toolbox?" & @CRLF & @CRLF & "Unsaved data will be lost.")
    If $nRet = 6 Then Exit
EndFunc
