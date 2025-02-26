#NoTrayIcon
#RequireAdmin
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=..\..\Pictures\Iconshock-Real-Vista-General-Wizard.ico
#AutoIt3Wrapper_UseUpx=y
#AutoIt3Wrapper_Res_Comment=WXB helper, made by Leniter
#AutoIt3Wrapper_Res_Description=WXB Helper Tool
#AutoIt3Wrapper_Res_Fileversion=1.9.0.2
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_ProductName=WXBHelper
#AutoIt3Wrapper_Res_ProductVersion=1.9
#AutoIt3Wrapper_Res_CompanyName=Leniter
#AutoIt3Wrapper_Res_LegalCopyright=Leniter
#AutoIt3Wrapper_Res_LegalTradeMarks=Leniter
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_requestedExecutionLevel=requireAdministrator
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
; wxb simple automation, made by Leniter
DllCall("User32.dll","bool","SetProcessDPIAware")
#include <SQLite.au3>
#include <Misc.au3>
#include <ButtonConstants.au3>
#include <ComboConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <GuiListView.au3>
#include <ListViewConstants.au3>
#include <ScreenCapture.au3>
#include <GDIPlus.au3>
#include <WinAPIGdi.au3>
#include <ColorConstants.au3>
#include <Date.au3>
;#include <Clipboard.au3>
#include "Notify.au3"
#include "UWPOCR.au3"
#pragma compile(inputboxres, true)"

Opt("GUICloseOnESC", 0)
;Opt("GUIEventOptions", 1)
;Opt("TrayAutoPause", 0)
;Opt("TrayMenuMode", 3)
Opt("WinTitleMatchMode", 2)

$title = "WXBhelper"

If _Singleton($title, 1) = 0 Then
    Msgbox(64, $title, "The program is already running.")
	WinActivate($title)
    Exit
EndIf

$logging = True
$logfile = "wxbhelper.log"
$version = "1.9 built on " & FileGetTime(@ScriptName, 0, 5)
$logtxt = "---- " & $title & " v" & $version & " started @ " & _Now() & "." & @MSEC & " ----"
$hLog = FileOpen(".\" & $logfile, 2)
If $hLog = -1 Then
	MsgBox(64, $title, "Unable to write log file. Logging disabled.")
	$logging = False
Else
	FileWriteLine($hLog, $logtxt)
EndIf

$clp = ClipPut("")
If $clp = 0 Then
	MsgBox(262192, $title, "Unable to control clipboard! Program terminating. Check permissions?")
	writelog("Control clipboard failed! Exit.")
	Exit
EndIf


$dbfile = ".\info.dbf"
$sSQliteDll = _SQLite_Startup("", 1)
If @error Then
	_SQLite_Shutdown()
	writelog("sqlite3 library not found. Installing.")
	$sSQLiteinstall = FileInstall(".\sqlite3.dll", @WindowsDir&"\sqlite3.dll")
	writelog("sqlite3.dll install code " & $sSQLiteinstall)
	$sSQLiteinstall = FileInstall(".\sqlite3_x64.dll", @WindowsDir&"\sqlite3_x64.dll")
	writelog("sqlite3_x64.dll install code " & $sSQLiteinstall)
	If $sSQLiteinstall=0 Then
		MsgBox(262192, "SQLite Error", "SQLite3 not found and install attempt failed!" & @CRLF & "You have to install SQLite3 for this application to work.")
		writelog("sqlite3 installation failed. Exiting.")
		Exit
	EndIf
	$sSQliteDll = _SQLite_Startup("", 1)
	If @error Then
		MsgBox(262192, $title, "Sqlite3 startup failed. Program now terminating.")
		writelog("sqlite3 startup failed. Exiting.")
		Exit
	EndIf
EndIf
;MsgBox(0, "SQLite3.dll Loaded", $sSQliteDll & " (" & _SQLite_LibVersion() & ")")
writelog("Using SQLite version: " & $sSQliteDll & ", " & _SQLite_LibVersion())
OnAutoItExitRegister("closeres")

_SQLite_Open($dbfile) ; Open a permanent disk database
If @error Then
	MsgBox(262192, "DB Error", "Open/Create database failed!")
	writelog("Failed to open db. Exiting.")
	Exit -1
EndIf


Global $ODW = @DesktopWidth, $ODH = @DesktopHeight, $isResChanged = False
writelog("Current desktop resolution: " & $ODW & "x" & $ODH)

If @DesktopWidth<>1280 And @DesktopHeight<>800 Then
	If Not IsDeclared("iMsgBoxAnswer") Then Local $iMsgBoxAnswer
	$iMsgBoxAnswer = MsgBox(262195,$title,"Current desktop resolution is not 1280x800. " & @CRLF & "Program may not work as expected. " & @CRLF & "Press Yes, program will attempt to change current resolution and restore when exit." & @CRLF & "Press No, ignore and continue." & @CRLF & "Press Cancel, exit the program.")
	Select
		Case $iMsgBoxAnswer = 6 ;Yes
			$isResChanged = True
			$result = _ChangeScreenRes(1280, 800, @DesktopDepth, @DesktopRefresh)
			writelog("User choose change resolution, " & $result)
			MsgBox(262208,$title,"You should re-open 微小宝.",20)
		Case $iMsgBoxAnswer = 7 ;No
			writelog("User choose ignore.")
		Case $iMsgBoxAnswer = 2 ;Cancel
			Exit
	EndSelect
EndIf



_SQLite_Exec(-1, "PRAGMA analysis_limit=100")
$sexec = _SQLite_Exec(-1, "CREATE TABLE IF NOT EXISTS accounts (gzhid UNIQUE, gzhname, onlylink, rm1, rm2, rm3, rm4, rm5, rm6, rm7, rm8, fans, prevfans, topiccount, sortnum, targettime, nextday);")
writelog("Init table, " & $sexec)
$sexec = _SQLite_Exec(-1, "ANALYZE accounts")
writelog("Analyze table, " & $sexec)


Local $btnSave = Null, $btnCancel = Null, $frmAdd = Null, $hQuery = "", $aRow, $ckbOnlyfilllink
Local $txtSetTargetHour = Null, $txtSetTargetMin = Null, $ckbSetNextDay = Null, $sOCRTextResult = 0
Local $frmFansQtyTool = Null, $btnCollectFans = Null, $btnPasteFans = Null, $btnCloseFansTool = Null, $lstFans = ""
Local $btnSearchWord = Null, $btnStopSearchWord = Null, $frmWordSearch = Null, $btnCloseWordSearch = Null, $varWordSearchProgress = 0
Local $txtRM1, $txtRM2, $txtRM3, $txtRM4, $txtRM5, $txtRM6, $txtRM7, $txtRM8, $cmbTopicCount, $txtSort, $stbFQT
Local $txtSStr[11]
Local $editmethod, $txtGzhid, $txtGzhname, $onlyfilllink = 0, $topiccount = 6
Local $txtRM[9], $txtSstrings[11]
Local $curGzhid = "", $curGzhname = "", $curViewport = "1280x800", $isAutovp = True
Local $curSort = "", $curTargetHour = 0, $curTargetMin = 0, $curNextday = $GUI_UNCHECKED, $varTargetDate = "", $varTargetTime = ""



#Region ### START Koda GUI section ### Form=C:\Users\Jin\Documents\wxbhelper\frmMain.kxf
$frmMain = GUICreate($title & " Workbench", 325, 350, 422, 388, -1, $WS_EX_TOPMOST)
GUISetFont(10, 400, 0, "Verdana")
GUICtrlSetDefColor(0x000080)
$menuViewport = GUICtrlCreateMenu("&Viewport")
GUICtrlCreateMenuItem("NOT implemented!", $menuViewport)
$menuV1280x800 = GUICtrlCreateMenuItem("1280 x 800", $menuViewport, -1 , 1)
GUICtrlSetState($menuV1280x800, $GUI_CHECKED)
$menuV1696x808 = GUICtrlCreateMenuItem("1696 x 808 (s22)", $menuViewport, -1 , 1)
$menuV1740x808 = GUICtrlCreateMenuItem("1740 x 808 (s21)", $menuViewport, -1 , 1)
GUICtrlCreateMenuItem("", $menuViewport)
$menuAutovp = GUICtrlCreateMenuItem("Auto switch", $menuViewport)
GUICtrlSetState($menuAutovp, $GUI_CHECKED)
$menuAccount = GUICtrlCreateMenu("&Account")
$menuEdit = GUICtrlCreateMenuItem("Edit current account", $menuAccount)
$menuRemove = GUICtrlCreateMenuItem("Remove current account", $menuAccount)
$menuAdd = GUICtrlCreateMenuItem("Add new account", $menuAccount)
GUICtrlCreateMenuItem("", $menuAccount)
$menuDbOp = GUICtrlCreateMenu("DB operator", $menuAccount)
$menuDbAddCol = GUICtrlCreateMenuItem("Add Column", $menuDbOp)
$menuDbDropCol = GUICtrlCreateMenuItem("Drop Column", $menuDbOp)
$menuTools = GUICtrlCreateMenu("&Tools")
$menuFansQtyTool = GUICtrlCreateMenuItem("Fans Qty Tool", $menuTools)
$menuWordSearchTool = GUICtrlCreateMenuItem("Word Search Helper", $menuTools)
$menuHelp = GUICtrlCreateMenu("&Help")
$menuAbout = GUICtrlCreateMenuItem("About", $menuHelp)
$stbMain = GUICtrlCreateLabel("Loading", 0, 307, 324, 20, BitOR($SS_SIMPLE,$SS_SUNKEN))
GUICtrlSetColor(-1, 0x008000)
$cmbAccount = GUICtrlCreateCombo("- Select Account -", 16, 10, 147, 25, BitOR($CBS_DROPDOWNLIST,$CBS_AUTOHSCROLL))
$lblCurId = GUICtrlCreateLabel("currentGzhid", 168, 12, 71, 20)
$ckbAutoPublish = GUICtrlCreateCheckbox("Auto publish", 16, 48, 101, 17)
GUICtrlSetState(-1, $GUI_CHECKED)
GUICtrlSetColor(-1, 0x000080)
GUICtrlSetBkColor(-1, 0xA6CAF0)
$ckbIsNextDay = GUICtrlCreateCheckbox("Next day", 120, 48, 81, 17)
GUICtrlSetColor(-1, 0x800000)
$txtTargetHour = GUICtrlCreateInput("0", 204, 44, 45, 24, BitOR($ES_CENTER,$ES_NUMBER))
GUICtrlSetLimit(-1, 2)
GUICtrlSetColor(-1, 0x000080)
$updTargetHour = GUICtrlCreateUpdown($txtTargetHour)
GUICtrlSetLimit(-1, 22, 6)
$lblColumn = GUICtrlCreateLabel(":", 252, 48, 10, 20)
GUICtrlSetColor(-1, 0x000080)
$txtTargetMin = GUICtrlCreateInput("0", 264, 44, 41, 24, BitOR($ES_CENTER,$ES_NUMBER))
GUICtrlSetColor(-1, 0x000080)
$updTargetMin = GUICtrlCreateUpdown($txtTargetMin)
GUICtrlSetLimit(-1, 59, 0)
$btnGo = GUICtrlCreateButton("Go", 236, 10, 69, 25)
GUICtrlSetColor(-1, 0x800000)
GUICtrlSetBkColor(-1, "0xF0FFF0")
GUICtrlSetState(-1, $GUI_DISABLE)
$ckbIsFillLinks = GUICtrlCreateCheckbox("Dont touch template, fill links only", 16, 76, 289, 17)
GUICtrlSetColor(-1, 0x000080)
$ckbAutoCloseEdit = GUICtrlCreateCheckbox("Auto close after save, invoke pub timer", 16, 104, 289, 17)
GUICtrlSetState(-1, $GUI_CHECKED)
GUICtrlSetColor(-1, 0x000080)
$grpAttr = GUICtrlCreateGroup("No account loaded", 16, 130, 291, 171)
$curlink1 = GUICtrlCreateLabel("link1", 24, 154, 280, 19)
GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
$curlink2 = GUICtrlCreateLabel("link2", 24, 178, 280, 19)
GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
$curlink3 = GUICtrlCreateLabel("link3", 24, 202, 280, 19)
GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
$curlink4 = GUICtrlCreateLabel("link4", 24, 226, 280, 19)
GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
$curlink5 = GUICtrlCreateLabel("link5", 24, 250, 280, 19)
GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
$curlink6 = GUICtrlCreateLabel("link6", 24, 274, 280, 19)
GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
;$curlink7 = GUICtrlCreateLabel("link7", 24, 244, 280, 19)
;GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
;$curlink8 = GUICtrlCreateLabel("link8", 24, 268, 280, 19)
;GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
GUICtrlCreateGroup("", -99, -99, 1, 1)
GUISetState(@SW_SHOW, $frmMain)
refreshCombo()
GUICtrlSetData($stbMain, "Ready")
#EndRegion ### END Koda GUI section ###


_Notify_RegMsg()

While 1
	$nMsg = GUIGetMsg()
	If $frmFansQtyTool<>Null Then
		Switch $nMsg
			Case $btnCollectFans
				collectFans()
			Case $btnPasteFans
				pasteFans()
			;Case $btnEditFans
			;	modFans()
			Case $btnCloseFansTool
				destroyFansToolFrm()
		EndSwitch
	ElseIf $frmWordSearch<>Null Then
		; word search
		Switch $nMsg
			Case $btnSearchWord
				startWordSearch()
			Case $btnStopSearchWord
				stopWordSearch()
			Case $btnCloseFansTool
				destroyFansToolFrm()
			Case $btnCloseWordSearch
				destroyWordSearchFrm()
			Case $menuDbAddCol
				dbAddColumn()
			Case $menuDbDropCol
				dbDropColumn()
		EndSwitch
	ElseIf $frmAdd<>Null Then
		; editing
		Switch $nMsg
			Case $btnSave
				modAccount()
			Case $btnCancel
				destroyEditFrm()
		EndSwitch
	Else
		; main
		Switch $nMsg
			Case $GUI_EVENT_CLOSE
				Exit
			Case $cmbAccount
				refreshDisplayedAttr()
			Case $menuEdit
				showEditFrm("modify")
			Case $menuAdd
				showEditFrm("add")
			Case $btnGo
				runRoutine()
			Case $menuRemove
				removeAccount()
			Case $ckbAutoCloseEdit
				If GUICtrlRead($ckbAutoCloseEdit)=$GUI_CHECKED Then
					GUICtrlSetState($ckbAutoPublish, $GUI_ENABLE)
					GUICtrlSetState($ckbAutoPublish, $GUI_CHECKED)
					GUICtrlSetBkColor($ckbAutoPublish, 0xA6CAF0)
				Else
					GUICtrlSetState($ckbAutoPublish, $GUI_DISABLE)
					GUICtrlSetState($ckbAutoPublish, $GUI_UNCHECKED)
					GUICtrlSetBkColor($ckbAutoPublish, $GUI_BKCOLOR_TRANSPARENT)
				EndIf
			Case $ckbIsFillLinks
				If GUICtrlRead($ckbIsFillLinks)=$GUI_CHECKED Then
					$onlyfilllink = $GUI_CHECKED
				Else
					$onlyfilllink = $GUI_UNCHECKED
				EndIf
			Case $curlink1
				copyLink("1")
			Case $curlink2
				copyLink("2")
			Case $curlink3
				copyLink("3")
			Case $curlink4
				copyLink("4")
			Case $curlink5
				copyLink("5")
			Case $curlink6
				copyLink("6")
			;Case $curlink7
				;copyLink("7")
			;Case $curlink8
				;copyLink("8")
			Case $menuAbout
				MsgBox(262208, $title, "Automatic procedure tool for wxb." & @CRLF & "Ver." & $version & @CRLF & "Made by Leniter" & @CRLF & @CRLF & "Using libraries:" & @CRLF & "SQLite " & _SQLite_LibVersion())
			Case $menuV1280x800
				swV1280x800()
			Case $menuV1696x808
				swV1696x808()
			Case $menuV1740x808
				swV1740x808()
			Case $menuAutovp
				swAutovp()
			Case $menuFansQtyTool
				showFansQtyFrm()
			Case $menuWordSearchTool
				showWordSearchFrm()
			Case $menuDbAddCol
				dbAddColumn()
			Case $menuDbDropCol
				dbDropColumn()
		EndSwitch
	EndIf

	If Mod(@SEC, 5)=0 Then detectViewport()
WEnd



Func writelog($logtxt = "")
	If $logging = True Then
		FileWriteLine($hLog, "[" & _NowCalc() & "." & @MSEC & "] " & $logtxt)
	EndIf
EndFunc

Func _ChangeScreenRes($i_Width = @DesktopWidth, $i_Height = @DesktopHeight, $i_BitsPP = @DesktopDepth, $i_RefreshRate = @DesktopRefresh)
    Local Const $DM_PELSWIDTH = 0x00080000
    Local Const $DM_PELSHEIGHT = 0x00100000
    Local Const $DM_BITSPERPEL = 0x00040000
    Local Const $DM_DISPLAYFREQUENCY = 0x00400000
    Local Const $CDS_TEST = 0x00000002
    Local Const $CDS_UPDATEREGISTRY = 0x00000001
    Local Const $DISP_CHANGE_RESTART = 1
    Local Const $DISP_CHANGE_SUCCESSFUL = 0
    Local Const $HWND_BROADCAST = 0xffff
    Local Const $WM_DISPLAYCHANGE = 0x007E
    If $i_Width = "" Or $i_Width = -1 Then $i_Width = @DesktopWidth ; default to current setting
    If $i_Height = "" Or $i_Height = -1 Then $i_Height = @DesktopHeight ; default to current setting
    If $i_BitsPP = "" Or $i_BitsPP = -1 Then $i_BitsPP = @DesktopDepth ; default to current setting
    If $i_RefreshRate = "" Or $i_RefreshRate = -1 Then $i_RefreshRate = @DesktopRefresh ; default to current setting
    Local $DEVMODE = DllStructCreate("byte[32];int[10];byte[32];int[6]")
    Local $B = DllCall("user32.dll", "int", "EnumDisplaySettings", "ptr", 0, "long", 0, "ptr", DllStructGetPtr($DEVMODE))
    If @error Then
        $B = 0
        SetError(1)
        Return $B
    Else
        $B = $B[0]
    EndIf
    If $B <> 0 Then
        DllStructSetData($DEVMODE, 2, BitOR($DM_PELSWIDTH, $DM_PELSHEIGHT, $DM_BITSPERPEL, $DM_DISPLAYFREQUENCY), 5)
        DllStructSetData($DEVMODE, 4, $i_Width, 2)
        DllStructSetData($DEVMODE, 4, $i_Height, 3)
        DllStructSetData($DEVMODE, 4, $i_BitsPP, 1)
        DllStructSetData($DEVMODE, 4, $i_RefreshRate, 5)
        $B = DllCall("user32.dll", "int", "ChangeDisplaySettings", "ptr", DllStructGetPtr($DEVMODE), "int", $CDS_TEST)
        If @error Then
            $B = -1
        Else
            $B = $B[0]
        EndIf
        Select
            Case $B = $DISP_CHANGE_RESTART
                $DEVMODE = ""
                Return 2
            Case $B = $DISP_CHANGE_SUCCESSFUL
                DllCall("user32.dll", "int", "ChangeDisplaySettings", "ptr", DllStructGetPtr($DEVMODE), "int", $CDS_UPDATEREGISTRY)
                DllCall("user32.dll", "int", "SendMessage", "hwnd", $HWND_BROADCAST, "int", $WM_DISPLAYCHANGE, _
                        "int", $i_BitsPP, "int", $i_Height * 2 ^ 16 + $i_Width)
                $DEVMODE = ""
                Return 1
            Case Else
                $DEVMODE = ""
                SetError(1)
                Return $B
        EndSelect
    EndIf
EndFunc ;==>_ChangeScreenRes

Func showEditFrm($editmethod)
	GUISetState(@SW_HIDE, $frmMain)
	$frmAdd = GUICreate("Account Management - " & $editmethod, 406, 450, 502, 200, 0, -1, $frmMain)
	GUISetFont(10, 400, 0, "Verdana")
	$lblGzhid = GUICtrlCreateLabel("Account ID", 12, 8, 78, 20)
	GUICtrlSetColor(-1, 0x000080)
	$txtGzhid = GUICtrlCreateInput("", 96, 6, 295, 24)
	GUICtrlSetColor(-1, 0x000080)
	$lblGzhname = GUICtrlCreateLabel("Name", 12, 38, 40, 20)
	GUICtrlSetColor(-1, 0x000080)
	$txtGzhname = GUICtrlCreateInput("", 96, 36, 295, 24)
	GUICtrlSetColor(-1, 0x000080)
	$lblProcess = GUICtrlCreateLabel("Process", 12, 72, 78, 20)
	GUICtrlSetColor(-1, 0x000080)
	$ckbOnlyfilllink = GUICtrlCreateCheckbox("Link only", 96, 72, 88, 17)
	GUICtrlSetColor(-1, 0x000080)
	$cmbTopicCount = GUICtrlCreateCombo("6", 210, 68, 50, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
	GUICtrlSetLimit(-1, 1)
	GUICtrlSetData(-1, "8|4")
	GUICtrlSetColor(-1, 0x000080)
	GUICtrlSetTip(-1, "How many topics per routine", "Topic count")
	$txtSort = GUICtrlCreateInput("", 282, 68, 50, 24)
	GUICtrlSetLimit(-1, 3)
	GUICtrlSetColor(-1, 0x000080)
	GUICtrlSetTip(-1, "Listing with ascend number order", "Sorting")
	GUICtrlCreateUpdown(-1)
	GUICtrlCreateLabel("link1", 12, 108, 33, 20)
	GUICtrlSetColor(-1, 0x000080)
	$txtRM1 = GUICtrlCreateInput("", 96, 102, 295, 24)
	GUICtrlSetColor(-1, 0x000080)
	GUICtrlCreateLabel("link2", 12, 138, 33, 20)
	GUICtrlSetColor(-1, 0x000080)
	$txtRM2 = GUICtrlCreateInput("", 96, 132, 295, 24)
	GUICtrlSetColor(-1, 0x000080)
	GUICtrlCreateLabel("link3", 12, 168, 33, 20)
	GUICtrlSetColor(-1, 0x000080)
	$txtRM3 = GUICtrlCreateInput("", 96, 162, 295, 24)
	GUICtrlSetColor(-1, 0x000080)
	GUICtrlCreateLabel("link4", 12, 198, 33, 20)
	GUICtrlSetColor(-1, 0x000080)
	$txtRM4 = GUICtrlCreateInput("", 96, 192, 295, 24)
	GUICtrlSetColor(-1, 0x000080)
	GUICtrlCreateLabel("link5", 12, 228, 33, 20)
	GUICtrlSetColor(-1, 0x000080)
	$txtRM5 = GUICtrlCreateInput("", 96, 222, 295, 24)
	GUICtrlSetColor(-1, 0x000080)
	GUICtrlCreateLabel("link6", 12, 258, 33, 20)
	GUICtrlSetColor(-1, 0x000080)
	$txtRM6 = GUICtrlCreateInput("", 96, 252, 295, 24)
	GUICtrlSetColor(-1, 0x000080)
	GUICtrlCreateLabel("link7", 12, 288, 33, 20)
	GUICtrlSetColor(-1, 0x000080)
	$txtRM7 = GUICtrlCreateInput("", 96, 282, 295, 24)
	GUICtrlSetColor(-1, 0x000080)
	GUICtrlCreateLabel("link8", 12, 318, 33, 20)
	GUICtrlSetColor(-1, 0x000080)
	$txtRM8 = GUICtrlCreateInput("", 96, 312, 295, 24)
	GUICtrlSetColor(-1, 0x000080)
	$lblTargetTime = GUICtrlCreateLabel("Publish time", 12, 352, 84, 20)
	GUICtrlSetColor(-1, 0x000080)
	$txtSetTargetHour = GUICtrlCreateInput("0", 96, 348, 49, 24, BitOR($ES_CENTER,$ES_NUMBER))
	GUICtrlSetColor(-1, 0x000080)
	$updSetTargetHour = GUICtrlCreateUpdown($txtSetTargetHour)
	GUICtrlSetLimit(-1, 22, 6)
	$lblSetTargetCol = GUICtrlCreateLabel(":", 148, 352, 10, 20)
	GUICtrlSetColor(-1, 0x000080)
	$txtSetTargetMin = GUICtrlCreateInput("0", 160, 348, 49, 24, BitOR($ES_CENTER,$ES_NUMBER))
	GUICtrlSetColor(-1, 0x000080)
	$updSetTargetMin = GUICtrlCreateUpdown($txtSetTargetMin)
	GUICtrlSetLimit(-1, 59, 0)
	$ckbSetNextDay = GUICtrlCreateCheckbox("next day", 220, 352, 97, 17)
	GUICtrlSetColor(-1, 0x000080)
	$btnSave = GUICtrlCreateButton($editmethod, 96, 386, 75, 25)
	GUICtrlSetColor(-1, 0xFF0000)
	$btnCancel = GUICtrlCreateButton("cancel", 192, 386, 75, 25)
	GUICtrlSetColor(-1, 0x000080)
	If $editmethod = "modify" Then
		GUICtrlSetData($txtGzhid, $curGzhid)
		GUICtrlSetState($txtGzhid, $GUI_DISABLE)
		GUICtrlSetData($txtGzhname, $curGzhname)
		GUICtrlSetState($ckbOnlyfilllink, $onlyfilllink)
		GUICtrlSetData($txtRM1, $txtRM[1])
		GUICtrlSetData($txtRM2, $txtRM[2])
		GUICtrlSetData($txtRM3, $txtRM[3])
		GUICtrlSetData($txtRM4, $txtRM[4])
		GUICtrlSetData($txtRM5, $txtRM[5])
		GUICtrlSetData($txtRM6, $txtRM[6])
		GUICtrlSetData($txtRM7, $txtRM[7])
		GUICtrlSetData($txtRM8, $txtRM[8])
		GUICtrlSetData($cmbTopicCount, $topiccount)
		GUICtrlSetData($txtSort, $curSort)
		GUICtrlSetData($txtSetTargetHour, $curTargetHour)
		GUICtrlSetData($txtSetTargetMin, $curTargetMin)
		GUICtrlSetState($ckbSetNextDay, $curNextday)
	EndIf
	GUISetState(@SW_SHOW, $frmAdd)
EndFunc

Func showFansQtyFrm()
	GUISetState(@SW_HIDE, $frmMain)
	$frmFansQtyTool = GUICreate("Fans Quantity Tool", 615, 490, 650, 100, 0, $WS_EX_TOPMOST, $frmMain)
	GUISetFont(10, 400, 0, "Verdana")
	$lstFans = GUICtrlCreateListView("gzhId|gzhName|current fans|previous fans|diff", 6, 42, 600, 384)
	GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, 0, 120)
	GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, 1, 150)
	GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, 2, 120)
	GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, 3, 120)
	GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, 4, 60)
	$btnCollectFans = GUICtrlCreateButton("Collect", 6, 6, 123, 25)
	GUICtrlSetBkColor(-1, 0xD7E4F2)
	$btnPasteFans = GUICtrlCreateButton("Paste", 138, 6, 141, 25)
	GUICtrlSetBkColor(-1, 0xD7E4F2)
	;$btnEditFans = GUICtrlCreateButton("Edit", 300, 6, 75, 25)
	;GUICtrlSetBkColor(-1, 0xCCCCCC)
	$btnCloseFansTool = GUICtrlCreateButton("Close", 528, 6, 75, 25)
	GUICtrlSetBkColor(-1, 0xC0C0C0)
	$stbFQT = GUICtrlCreateLabel("Ready", 0, 440, 614, 20, BitOR($SS_SIMPLE,$SS_SUNKEN))
	GUICtrlSetColor(-1, 0x008000)

	_SQLite_Query(-1, "SELECT gzhid,gzhname,fans,prevfans FROM accounts", $hQuery)
	Local $aFansList[]
	$i = 0
	While _SQLite_FetchData($hQuery, $aRow) = $SQLITE_OK
		$i = $i + 1
        $aFansList[$i] = GUICtrlCreateListViewItem($aRow[0]&"|"&$aRow[1]&"|"&$aRow[2]&"|"&$aRow[3], $lstFans)
	WEnd
	_SQLite_QueryFinalize($hQuery)
	GUISetState(@SW_SHOW, $frmFansQtyTool)
EndFunc

Func showWordSearchFrm()
	GUISetState(@SW_HIDE, $frmMain)
	$frmWordSearch = GUICreate("Word Search Helper Tool", 574, 180, 192, 124, 0, $WS_EX_TOPMOST, $frmMain)
	GUISetFont(9, 400, 0, "Verdana")
	$txtSStr[1] = GUICtrlCreateInput($txtSstrings[1], 8, 8, 230, 22)
	GUICtrlSetTip(-1, "Enter string you wanna search")
	$txtSStr[2] = GUICtrlCreateInput($txtSstrings[2], 248, 8, 230, 22)
	GUICtrlSetTip(-1, "Enter string you wanna search")
	$txtSStr[3] = GUICtrlCreateInput($txtSstrings[3], 8, 36, 230, 22)
	GUICtrlSetTip(-1, "Enter string you wanna search")
	$txtSStr[4] = GUICtrlCreateInput($txtSstrings[4], 248, 36, 230, 22)
	GUICtrlSetTip(-1, "Enter string you wanna search")
	$txtSStr[5] = GUICtrlCreateInput($txtSstrings[5], 8, 64, 230, 22)
	GUICtrlSetTip(-1, "Enter string you wanna search")
	$txtSStr[6] = GUICtrlCreateInput($txtSstrings[6], 248, 64, 230, 22)
	GUICtrlSetTip(-1, "Enter string you wanna search")
	$txtSStr[7] = GUICtrlCreateInput($txtSstrings[7], 8, 92, 230, 22)
	GUICtrlSetTip(-1, "Enter string you wanna search")
	$txtSStr[8] = GUICtrlCreateInput($txtSstrings[8], 248, 92, 230, 22)
	GUICtrlSetTip(-1, "Enter string you wanna search")
	$txtSStr[9] = GUICtrlCreateInput($txtSstrings[9], 8, 120, 230, 22)
	GUICtrlSetTip(-1, "Enter string you wanna search")
	$txtSStr[10] = GUICtrlCreateInput($txtSstrings[10], 248, 120, 230, 22)
	GUICtrlSetTip(-1, "Enter string you wanna search")
	$btnSearchWord = GUICtrlCreateButton("Start", 488, 8, 75, 49)
	$btnStopSearchWord = GUICtrlCreateButton("Stop", 488, 64, 75, 49)
	$btnCloseWordSearch = GUICtrlCreateButton("Close", 488, 120, 75, 22)
	GUISetState(@SW_SHOW, $frmWordSearch)
EndFunc

Func destroyEditFrm()
	GUIDelete($frmAdd)
	$frmAdd = Null
	GUISetState(@SW_SHOW, $frmMain)
	refreshCombo()
	WinActivate($frmMain)
	GUICtrlSetData($stbMain, "Ready")
EndFunc

Func destroyFansToolFrm()
	GUIDelete($frmFansQtyTool)
	$frmFansQtyTool = Null
	GUISetState(@SW_SHOW, $frmMain)
	WinActivate($frmMain)
	GUICtrlSetData($stbMain, "Ready")
EndFunc

Func destroyWordSearchFrm()
	GUIDelete($frmWordSearch)
	$frmWordSearch = Null
	GUISetState(@SW_SHOW, $frmMain)
	WinActivate($frmMain)
	GUICtrlSetData($stbMain, "Ready")
EndFunc

Func modAccount()
	$curGzhid = StringReplace(GUICtrlRead($txtGzhid), " ", "")
	$curGzhname = StringReplace(GUICtrlRead($txtGzhname), " ", "")
	If $curGzhid = "" Or $curGzhname = "" Then
		MsgBox(262192,"","GZHid or GZHname is empty!")
	Else
		writelog("User submit account edit - " & $editmethod & " " & $curGzhid & " " & $curGzhname)
		$editmethod = GUICtrlRead($btnSave)
		$onlyfilllink = GUICtrlRead($ckbOnlyfilllink)
		$txtRM[1] = GUICtrlRead($txtRM1)
		$txtRM[2] = GUICtrlRead($txtRM2)
		$txtRM[3] = GUICtrlRead($txtRM3)
		$txtRM[4] = GUICtrlRead($txtRM4)
		$txtRM[5] = GUICtrlRead($txtRM5)
		$txtRM[6] = GUICtrlRead($txtRM6)
		$txtRM[7] = GUICtrlRead($txtRM7)
		$txtRM[8] = GUICtrlRead($txtRM8)
		$topiccount = GUICtrlRead($cmbTopicCount)
		If $topiccount = "" Then $topiccount = 6
		If Not IsNumber($topiccount) Then $topiccount = 6
		$sortnum = GUICtrlRead($txtSort)
		$varDesHour = GUICtrlRead($txtSetTargetHour)
		If $varDesHour<1 Or $varDesHour>23 Then $varDesHour = 0
		$varDesMin = GUICtrlRead($txtSetTargetMin)
		If $varDesMin<1 Or $varDesMin>59 Then $varDesMin = 0
		$varDesignatedTime = $varDesHour & ":" & $varDesMin
		$varNextday = GUICtrlRead($ckbSetNextDay)
		If $editmethod = "add" Then
			GUICtrlSetData($stbMain, "Adding new account.", 0)
			$sexec = _SQLite_Exec(-1, "INSERT INTO accounts VALUES ('" & $curGzhid & "', '" & $curGzhname & "', '" & $onlyfilllink & "', '" & $txtRM[1] & "', '" & $txtRM[2] & "', '" & $txtRM[3] & "', '" & $txtRM[4] & "', '" & $txtRM[5] & "', '" & $txtRM[6] & "', '" & $txtRM[7] & "', '" & $txtRM[8] & "', 0, 0, '" & $topiccount & "', '" & $sortnum & "', '" & $varDesignatedTime &  "', '" & $varNextday & "');")
			writelog("sqlite insert new line, " & $sexec)
		Else
			GUICtrlSetData($stbMain, "Saving account changes.", 0)
			$sexec = _SQLite_Exec(-1, "UPDATE accounts SET gzhname='" & $curGzhname & "', onlylink='" & $onlyfilllink & "', rm1='" & $txtRM[1] & "', rm2='" & $txtRM[2] & "', rm3='" & $txtRM[3] & "', rm4='" & $txtRM[4] & "', rm5='" & $txtRM[5] & "', rm6='" & $txtRM[6] & "', rm7='" & $txtRM[7] & "', rm8='" & $txtRM[8] & "', topiccount='" & $topiccount & "', sortnum='" & $sortnum & "', targettime='" & $varDesignatedTime & "', nextday='" & $varNextday & "' WHERE gzhid='" & $curGzhid & "';")
			writelog("sqlite update line, " & $sexec)
		EndIf
		destroyEditFrm()
	EndIf
EndFunc

Func removeAccount()
	writelog("User select remove account: " & $curGzhid & " (" & $curGzhname & ")")
	$confirmRemove = MsgBox(262436, $title, "Are you quite sure to DELETE this account profile?" & @CRLF & $curGzhid & " (" & $curGzhname & ")" & @CRLF & @CRLF & "This action CANNOT BE UNDONE!")
	Select
		Case $confirmRemove = 6 ;Yes
			writelog(" - Confirmed removal.")
			$sexec = _SQLite_Exec(-1, "DELETE FROM accounts WHERE gzhid='" & $curGzhid & "'")
			writelog("sqlite delete line, " & $sexec)
			refreshCombo()
		Case $confirmRemove = 7 ;No
			writelog(" - Cancelled.")
	EndSelect
EndFunc

Func refreshCombo()
	GUISetState(@SW_DISABLE, $frmMain)
	GUICtrlSetData($stbMain, "Loading data...")
	GUICtrlSetData($cmbAccount, "")
	_SQLite_Query(-1, "SELECT gzhname FROM accounts ORDER BY sortnum ASC, rowid ASC", $hQuery)
	$aNamelist = ""
	While _SQLite_FetchData($hQuery, $aRow) = $SQLITE_OK
		If $aNamelist = "" Then
			$aNamelist = $aRow[0]
		Else
			$aNamelist = $aNamelist & "|" & $aRow[0]
		EndIf
	WEnd
	_SQLite_QueryFinalize($hQuery)
	GUICtrlSetData($cmbAccount, $aNamelist)
	GUICtrlSetData($grpAttr, "Select an account to reload")
	GUICtrlSetData($lblCurId, "")
	GUICtrlSetState($btnGo, $GUI_DISABLE)
	GUICtrlSetState($menuEdit, $GUI_DISABLE)
	GUICtrlSetState($menuRemove, $GUI_DISABLE)
	GUICtrlSetData($curlink1, "")
	GUICtrlSetData($curlink2, "")
	GUICtrlSetData($curlink3, "")
	GUICtrlSetData($curlink4, "")
	GUICtrlSetData($curlink5, "")
	GUICtrlSetData($curlink6, "")
	;GUICtrlSetData($curlink7, "")
	;GUICtrlSetData($curlink8, "")
	GUICtrlSetData($stbMain, "Ready")
	GUISetState(@SW_ENABLE, $frmMain)
EndFunc

Func refreshDisplayedAttr()
	$curGzhname = GUICtrlRead($cmbAccount)
	_SQLite_QuerySingleRow(-1, "SELECT * FROM accounts WHERE gzhname='" & $curGzhname & "' LIMIT 1", $aRow)
	;_SQLite_FetchData($hQuery, $aRow)
	$curGzhid = $aRow[0]
	;$curGzhname = $aRow[1]
	$onlyfilllink = $aRow[2]
	$txtRM[1] = $aRow[3]
	$txtRM[2] = $aRow[4]
	$txtRM[3] = $aRow[5]
	$txtRM[4] = $aRow[6]
	$txtRM[5] = $aRow[7]
	$txtRM[6] = $aRow[8]
	;$txtRM[7] = $aRow[9]
	;$txtRM[8] = $aRow[10]
	;$topiccount = $aRow[13]
	$curSort = $aRow[14]
	$varDesignatedTime = $aRow[15]
	$curTargetHour = 0
	$curTargetMin = 0
	;$isAutoPub = $GUI_UNCHECKED
	;$varColorAutopub = $GUI_BKCOLOR_TRANSPARENT
	If $varDesignatedTime<>"" And StringLen($varDesignatedTime)>3 And StringInStr($varDesignatedTime, ":")>1 Then
		$aDesTime = StringSplit($varDesignatedTime, ":")
		;$isAutoPub = $GUI_CHECKED
		;$varColorAutopub = "0xA6CAF0"
		$curTargetHour = $aDesTime[1]
		$curTargetMin = $aDesTime[2]
	EndIf
	$curNextday = $aRow[16]
	$varColorNextday = $GUI_BKCOLOR_TRANSPARENT
	If $curNextday="" Or $curNextday<>$GUI_CHECKED Then
		$curNextday = $GUI_UNCHECKED
	Else
		If @HOUR < 8 Then
			; nextday set but maybe it's 'nextday' now, changing state
			$curNextday = $GUI_UNCHECKED
			$varColorNextday = $COLOR_YELLOW
		Else
			$varColorNextday = $COLOR_AQUA
		EndIf
	EndIf
	;_SQLite_QueryFinalize($hQuery)

	GUICtrlSetData($lblCurId, $curGzhid)
	;GUICtrlSetState($ckbAutoPublish, $isAutoPub)
	;GUICtrlSetBkColor($ckbAutoPublish, $varColorAutopub)
	GUICtrlSetState($ckbIsNextDay, $curNextday)
	GUICtrlSetBkColor($ckbIsNextDay, $varColorNextday)
	GUICtrlSetData($txtTargetHour, $curTargetHour)
	GUICtrlSetData($txtTargetMin, $curTargetMin)
	GUICtrlSetData($grpAttr, "Read-orig Links | Count: " & $topiccount)
	If $onlyfilllink = $GUI_CHECKED Then
		;GUICtrlSetData($grpAttr, "Process: Fill links only | Count: " & $topiccount)
		GUICtrlSetState($ckbIsFillLinks, $GUI_CHECKED)
	Else
		;GUICtrlSetData($grpAttr, "Process: Full | Count: " & $topiccount)
		GUICtrlSetState($ckbIsFillLinks, $GUI_UNCHECKED)
	EndIf
	GUICtrlSetData($curlink1, $txtRM[1])
	GUICtrlSetData($curlink2, $txtRM[2])
	GUICtrlSetData($curlink3, $txtRM[3])
	GUICtrlSetData($curlink4, $txtRM[4])
	GUICtrlSetData($curlink5, $txtRM[5])
	GUICtrlSetData($curlink6, $txtRM[6])
	;GUICtrlSetData($curlink7, $txtRM[7])
	;GUICtrlSetData($curlink8, $txtRM[8])
	GUICtrlSetData($stbMain, "Account profile loaded.")
	GUICtrlSetState($menuEdit, $GUI_ENABLE)
	GUICtrlSetState($menuRemove, $GUI_ENABLE)
	GUICtrlSetState($btnGo, $GUI_ENABLE + $GUI_FOCUS)
EndFunc

Func runRoutine()
	GUICtrlSetData($stbMain, "Starting routine.")
	BlockInput(1)
	GUICtrlSetData($btnGo, "Running")
	writelog("Edit routine started for " & $curGzhid & ", screen: " & $curViewport)
	_Notify_Size(100, 200, 400)
	_Notify_Set(1, Default, 0xCCCCCC, "Tahoma", Default, 250)
	$hNotify = _Notify_Show(0, "Edit routine for " & $curGzhname, "Initializing...", 200)
	If WinExists("微小宝") Then
		; get auto publish settings
		$isAutoCloseEdit = GUICtrlRead($ckbAutoCloseEdit)
		$isAutoPub = $GUI_UNCHECKED
		If $isAutoCloseEdit=$GUI_CHECKED Then $isAutoPub = GUICtrlRead($ckbAutoPublish)
		If $isAutoPub=$GUI_CHECKED Then
			$curNextday = GUICtrlRead($ckbIsNextDay)
			$curTargetHour = GUICtrlRead($txtTargetHour)
			$curTargetMin = GUICtrlRead($txtTargetMin)
			If $curTargetHour<6 Or $curTargetHour>22 Then
				$isAutoPub = $GUI_UNCHECKED
			ElseIf $curTargetMin>59 Then
				$curTargetMin = 30
			Else
				If $curNextday=$GUI_UNCHECKED Then
					If @HOUR > $curTargetHour Or (@HOUR = $curTargetHour And @MIN > $curTargetMin-5) Then
						$curTargetHour=@HOUR
						$curTargetMin=@MIN+7
					EndIf
				EndIf
				If StringLen($curTargetHour)=1 Then $curTargetHour = "0" & $curTargetHour
				If StringLen($curTargetMin)=1 Then $curTargetMin = "0" & $curTargetMin
			EndIf
		EndIf
		GUISetState(@SW_MINIMIZE, $frmMain)
		_Notify_Modify($hNotify, 0x0000FF, 0xFFFFCC, Default, "Processing...")
		;GUICtrlSetData($stbMain, "Target located. Running...", 0)
		WinActivate("微小宝")
		WinWaitActive("微小宝", "", 3)
		$i = 1
		; full 8 topics action
		;$y = 130
		;MouseMove(66, $y, 0)
		;Sleep(10)
		;MouseWheel("down", 3)
		; 6 topics action
		$y = 322
		MouseMove(76, $y, 0)
		Sleep(10)
		MouseWheel("up", 5)
		Sleep(10)
		MouseDown("primary")
		Sleep(100)
		MouseUp("primary")
		Sleep(1200)
		MouseClick("")
		Sleep(1400)
		$isAborted = False
		Do
			If detectViewport(True) = False Then
				_Notify_Hide($hNotify)
				_Notify_Set(1, 0xFF0000, 0xFFFF00, "Tahoma", True, 250, -1000, True)
				_Notify_Show(16, "Routine for " & $curGzhname, "Aborted...", 40)
				GUICtrlSetData($stbMain, "Viewport changed, routine aborted!")
				writelog("Edit routine aborted due to viewport changed.")
				BlockInput(0)
				GUISetState(@SW_RESTORE, $frmMain)
				ExitLoop
				Return False
			EndIf
			If WinExists("[Class:TaskManagerWindow]") Then
				_Notify_Hide($hNotify)
				_Notify_Set(1, 0xFF0000, 0xFFFF00, "Tahoma", True, 250, -1000, True)
				_Notify_Show(16, "Routine for " & $curGzhname, "Aborted...", 60)
				GUICtrlSetData($stbMain, "Taskmgr detected, routine aborted!")
				writelog("Edit routine aborted due to taskmgr opened.")
				BlockInput(0)
				GUISetState(@SW_RESTORE, $frmMain)
				ExitLoop
				Return False
			EndIf
			$isWxbActive = WinActivate("微小宝")
			Sleep(100)
			If $isWxbActive=0 Then
				_Notify_Hide($hNotify)
				_Notify_Set(1, 0xFF0000, 0xFFFF00, "Tahoma", True, 250, -1000, True)
				_Notify_Show(16, "Routine for " & $curGzhname, "Aborted...", 120)
				GUICtrlSetData($stbMain, "Target lost unexpectfully, routine aborted!")
				writelog("Edit routine aborted due to target lost!")
				BlockInput(0)
				$isAborted = True
				GUISetState(@SW_RESTORE, $frmMain)
				ExitLoop
				Return False
			EndIf
			WinWaitActive("微小宝", "", 1)
			_Notify_Modify($hNotify, 0x0000FF, 0xFFFFCC, Default, "Processing topic " & $i)
			Sleep(10)
			If $i > 1 Then
				; Moving cursor to the topic and click.")
				;Sleep(5)
				$y = $y + 70
				MouseMove(76, $y, 0)
				;$aPos = MouseGetPos()
				;writelog("Mouse pos: " & $aPos[0] & "," & $aPos[1])
				Sleep(10)
				MouseDown("primary")
				Sleep(100)
				MouseUp("primary")
				Sleep(1000)
				MouseClick("")
				Sleep(2000)
			EndIf
			;If $i < 3 Then
				;GUICtrlSetData($stbMain, $i & " - Moving cursor to the right and scroll down.")
				;Sleep(5)
				MouseMove(1074, 644, 0)
				Sleep(10)
				MouseWheel("down", 2)
				Sleep(10)
				MouseWheel("down", 2)
				Sleep(10)
			;EndIf
			If $onlyfilllink = $GUI_UNCHECKED Then
				; full process
				; header replacement
				;nav to topic top
				;MouseMove(296, 643, 0)
				;Sleep(10)
				;MouseDown("primary")
				;Sleep(100)
				;MouseUp("primary")
				;Sleep(5)
				;Send("{CTRLDOWN}")
				;Send("{HOME}")
				;Send("{CTRLUP}")
				;Sleep(5)
				;floating toolbar may be upper
				MouseMove(666, 520, 0)
				Sleep(10)
				MouseDown("primary")
				Sleep(100)
				MouseUp("primary")
				Sleep(10)
				MouseMove(514, 230, 0)
				Sleep(10)
				MouseDown("primary")
				Sleep(100)
				MouseUp("primary")
				Sleep(10)
				;or may be lower
				MouseMove(666, 520, 0)
				Sleep(10)
				MouseDown("primary")
				Sleep(100)
				MouseUp("primary")
				Sleep(10)
				MouseMove(515, 632, 0)
				Sleep(10)
				MouseDown("primary")
				Sleep(100)
				MouseUp("primary")
				Sleep(10)
				;or may be even lower
				MouseMove(382, 574, 0)
				Sleep(10)
				MouseDown("primary")
				Sleep(100)
				MouseUp("primary")
				Sleep(10)
				MouseMove(516, 656, 0)
				Sleep(10)
				MouseDown("primary")
				Sleep(100)
				MouseUp("primary")
				Sleep(10)
				; 20230505 add: footer replacement
				;nav to topic bottom
				MouseMove(296, 643, 0)
				Sleep(10)
				MouseDown("primary")
				Sleep(100)
				MouseUp("primary")
				Sleep(5)
				Send("{CTRLDOWN}")
				Sleep(5)
				Send("{END}")
				Sleep(5)
				Send("{END}")
				Sleep(5)
				Send("{CTRLUP}")
				Sleep(20)
				;click footer and delete, possible position 1
				MouseMove(574, 644, 0)
				Sleep(10)
				MouseDown("primary")
				Sleep(100)
				MouseUp("primary")
				Sleep(10)
				MouseMove(512, 474, 0)
				Sleep(10)
				MouseDown("primary")
				Sleep(100)
				MouseUp("primary")
				Sleep(5)
				;click footer and delete, possible position 2
				MouseMove(702, 604, 0)
				Sleep(10)
				MouseDown("primary")
				Sleep(100)
				MouseUp("primary")
				Sleep(10)
				MouseMove(510, 426, 0)
				Sleep(10)
				MouseDown("primary")
				Sleep(100)
				MouseUp("primary")
				Sleep(5)
				;click footer and delete, possible position 3
				MouseMove(722, 630, 0)
				Sleep(10)
				MouseDown("primary")
				Sleep(100)
				MouseUp("primary")
				Sleep(10)
				MouseMove(513, 453, 0)
				Sleep(10)
				MouseDown("primary")
				Sleep(100)
				MouseUp("primary")
				Sleep(5)
				; Moving cursor to guanzhuliuyan and click.
				;Sleep(5)
				MouseMove(1076, 642, 0)
				Sleep(10)
				MouseDown("primary")
				Sleep(100)
				MouseUp("primary")
				Sleep(10)
				; Moving cursor to insert template and click.
				;Sleep(5)
				MouseMove(1020, 678, 0)
				Sleep(10)
				MouseDown("primary")
				Sleep(100)
				MouseUp("primary")
				Sleep(1500)
				; uncheck yueduyindao now not needed
				;Sleep(5)
				;MouseMove(1034, 568, 0)
				;Sleep(10)
				;MouseDown("primary")
				;Sleep(100)
				;MouseUp("primary")
				;Sleep(5)
				; Moving cursor to insert button and click.
				;Sleep(5)
				MouseMove(1094, 626, 0)
				;$aPos = MouseGetPos()
				;writelog("Mouse pos: " & $aPos[0] & "," & $aPos[1])
				Sleep(10)
				MouseDown("primary")
				Sleep(100)
				MouseUp("primary")
				Sleep(1500)
			EndIf
			; Paste the corresponding readmore link.
			;Sleep(5)
			MouseMove(1100, 572, 0)
			;$aPos = MouseGetPos()
			;writelog("Mouse pos: " & $aPos[0] & "," & $aPos[1])
			Sleep(10)
			MouseDown("primary")
			Sleep(100)
			MouseUp("primary")
			Sleep(10)
			Send("^a")
			Sleep(5)
			;writelog("Send link - " & $txtRM[$i])
			;Sleep(5)
			;Send("https://mp.weixin.qq.com/s/QxlTJn7i64hMLJFNsU7LkQ", 1)
			ClipPut($txtRM[$i])
			Sleep(5)
			Send("^v")
			Sleep(5)
			$i = $i + 1
			;$y = $y + 50
			;If $i > 2 Then $y = $y + 30
		Until $i > $topiccount
		If $isAborted = False Then
			; click save
			MouseMove(746, 748, 0)
			Sleep(10)
			MouseClick("primary")
			Sleep(10)
			GUICtrlSetData($stbMain, "Routine done for " & $curGzhname & " at " & @HOUR & ":" & @MIN & ":" & @SEC)
			_Notify_Modify($hNotify, Default, 0xCCFFCC, Default, "Completed. Saving.")
			writelog("Edit routine done for " & $curGzhname & ", now saving.")
			Sleep(19000)
			If GUICtrlRead($ckbAutoCloseEdit)=$GUI_CHECKED Then
				; click close
				MouseMove(919, 745, 0)
				Sleep(10)
				MouseClick("primary")
				Sleep(8000)
				_Notify_Modify($hNotify, Default, 0xCCFFCC, Default, "Invoking publish.")
				Sleep(5)
				; move cursor to somewhere and scroll down
				MouseMove(198, 396, 0)
				Sleep(10)
				MouseWheel("down", 5)
				Sleep(20)
				; move cursor to first material's post button
				MouseMove(336, 713, 0)
				Sleep(10)
				; click timed post
				MouseClickDrag("primary", 336, 714, 336, 676)
				Sleep(20)
				MouseMove(336, 676, 0)
				Sleep(10)
				MouseClick("primary")
				Sleep(500)
				;If $isAutoPub=$GUI_CHECKED Then
					$varTargetDate = "Today"
					; whether to enter pub date
					If $curNextday=$GUI_CHECKED And @HOUR <= 23 And @MIN < 55 Then
						MouseMove(562, 222, 0)
						Sleep(10)
						MouseClick("primary")
						Sleep(500)
						MouseClick("primary")
						Sleep(10)
						Send("^a")
						Sleep(10)
						$varTargetDate = _DateAdd('d', 1, _NowCalcDate())
						$varTargetDate = StringReplace($varTargetDate, "/", "-")
						Sleep(10)
						ClipPut($varTargetDate)
						Sleep(5)
						Send("^v")
						Sleep(10)
						MouseMove(638, 178, 0)
						Sleep(10)
						MouseClick("primary")
						Sleep(5)
					EndIf
					; enter pub time
					MouseMove(710, 222, 0)
					MouseClick("primary")
					Sleep(500)
					MouseClick("primary")
					Sleep(10)
					Send("^a")
					Sleep(10)
					$varTargetTime = $curTargetHour & ":" & $curTargetMin
					Sleep(5)
					ClipPut($varTargetTime)
					Sleep(5)
					Send("^v")
					Sleep(10)
					MouseMove(638, 178, 0)
					Sleep(10)
					MouseClick("primary")
					Sleep(10)
					; click pub!
					If $isAutoPub=$GUI_CHECKED Then
						writelog(" - Auto pub selected, procceeding with timer (" & $varTargetDate & " " & $varTargetTime & ")")
						_Notify_Modify($hNotify, Default, 0xCCFFCC, Default, "Auto publishing...")
						Sleep(6000)
						MouseMove(848, 388, 0)
						Sleep(15)
						MouseClick("primary")
						Sleep(12000)
						MouseMove(751, 247, 0)
						Sleep(15)
						MouseClick("primary")
						Sleep(5)
						_Notify_Modify($hNotify, Default, 0xE0FFFF, Default, "Pub set: " & $varTargetDate & " " & $varTargetTime)
					Else
						_Notify_Modify($hNotify, Default, 0x87CEFA, Default, "Done. ")
					EndIf
				;EndIf
			EndIf ;autoclose and pub
		EndIf
	Else
		GUICtrlSetData($stbMain, "Target not found!!!")
		_Notify_Hide($hNotify)
		_Notify_Set(1, 0xFF0000, 0xFFFF00, "Tahoma", True, 250, -1000, True)
		_Notify_Show(16, "Routine for " & $curGzhname, "Error...", 60)
	EndIf
	;GUICtrlSetState($btnGo, @SW_ENABLE)
	GUICtrlSetData($btnGo, "Go")
	BlockInput(0)
	;Sleep(20000)
	;_Notify_Hide($hNotify)
	;MouseMove(920, 744, 0)
EndFunc

Func copyLink($i)
	$clp = ClipPut($txtRM[$i])
	If $clp = 1 Then
		GUICtrlSetData($stbMain, "Link " & $i & " copied to clipboard.")
	Else
		GUICtrlSetData($stbMain, "Access clipboard failed! Please restart the program.")
	EndIf
EndFunc

Func swV1280x800()
	$curViewport = "1280x800"
	GUICtrlSetState($menuV1280x800, $GUI_CHECKED)
	GUICtrlSetState($menuV1696x808, $GUI_UNCHECKED)
	GUICtrlSetState($menuV1740x808, $GUI_UNCHECKED)
	;GUICtrlSetData($menuViewport, "Viewport: 1280x800")
	writelog("Viewport resolution changed to 1280x800.")
	GUICtrlSetData($stbMain, "Viewport changed to 1280x800.")
EndFunc

Func swV1696x808()
	$curViewport = "1696x808"
	GUICtrlSetState($menuV1280x800, $GUI_UNCHECKED)
	GUICtrlSetState($menuV1696x808, $GUI_CHECKED)
	GUICtrlSetState($menuV1740x808, $GUI_UNCHECKED)
	;GUICtrlSetData($menuViewport, "Viewport: 1696x808")
	writelog("Viewport resolution changed to 1696x808.")
	GUICtrlSetData($stbMain, "Viewport changed to 1696x808.")
EndFunc

Func swV1740x808()
	$curViewport = "1740x808"
	GUICtrlSetState($menuV1280x800, $GUI_UNCHECKED)
	GUICtrlSetState($menuV1696x808, $GUI_UNCHECKED)
	GUICtrlSetState($menuV1740x808, $GUI_CHECKED)
	;GUICtrlSetData($menuViewport, "Viewport: 1740x808")
	writelog("Viewport resolution changed to 1740x808.")
	GUICtrlSetData($stbMain, "Viewport changed to 1740x808.")
EndFunc

Func swAutovp()
	If BitAND(GUICtrlRead($menuAutovp), $GUI_CHECKED) = $GUI_CHECKED Then
		$isAutovp = False
		GUICtrlSetState($menuAutovp, $GUI_UNCHECKED)
	Else
		$isAutovp = True
		GUICtrlSetState($menuAutovp, $GUI_CHECKED)
	EndIf
EndFunc

Func detectViewport($isInroutine = False)
	$strViewport = @DesktopWidth & "x" & @DesktopHeight
	If $strViewport <> $curViewport Then
		GUICtrlSetData($stbMain, "Caution! Detected viewport: " & $strViewport & "!")
		If $isInroutine = True Then
			MsgBox(262436, $title, "Screen resolution changed during the process." & @CRLF & "Routine halted." & @CRLF & "You should check your materials and start over.")
			Return False
		Else
			If $isAutovp = True Then
				If $strViewport = "1280x800" Then
					swV1280x800()
				ElseIf $strViewport = "1696x808" Then
					swV1696x808()
				ElseIf $strViewport = "1740x808" Then
					swV1740x808()
				EndIf
			EndIf
		EndIf
	Else
		Return True
	EndIf
EndFunc

Func collectFans()
	; ocr fans
	$tmpPicfile = @ScriptDir&"\fanscptr.tif"
	FileDelete($tmpPicfile)
	; mp fans area
	$iX1 = 432
	$iY1 = 370
	$iX2 = 540
	$iY2 = 502
	Local $giTIFColorDepth = 24
	Local $giTIFCompression = $GDIP_EVTCOMPRESSIONNONE
	Local $Ext=""
	writelog("Fans capture routine started, screen: " & $curViewport)
	_Notify_Size(100, 180, 360)
	_Notify_Set(1, Default, 0xCCCCCC, "Tahoma", Default, 250)
	$hNotify = _Notify_Show(0, "Fans capture routine", "Initializing...", 20)
	BlockInput(1)
	GUICtrlSetData($btnCollectFans, "Collecting")
	GUISetState($frmFansQtyTool, @SW_DISABLE)
	If WinExists("微小宝") Then
		_Notify_Modify($hNotify, 0x0000FF, 0xFFFFCC, Default, "Processing...")
		WinActivate("微小宝")
		$hWnd = WinWaitActive("微小宝")
		;WinSetState($hWnd, "", @SW_MAXIMIZE)
		;;; to-do
		;click account
		; fetch current account name
		MouseMove(428, 372, 0)
		MouseWheel("up", 10)
		;MouseClick("primary")
		Sleep(600)
		MouseClickDrag("primary", 432, 274, 580, 274)
		Sleep(20)
		Send("^c")
		Sleep(5)
		Send("^c")
		Sleep(5)
		;$tmpGzhname = StringRegExpReplace(ClipGet(), "[^\u4e00-\u9fa5]", "")
		$tmpGzhname = ClipGet()
		writelog("Got gzhName: " & $tmpGzhname)
		_Notify_Modify($hNotify, 0x0000FF, 0xFFFFCC, Default, "Fetched gzhName: " & $tmpGzhname)
		; do screenshot of fans area
		$hBitmap = _ScreenCapture_Capture("", $iX1, $iY1, $iX2, $iY2, False)
		_GDIPlus_Startup()
		$hImage = _GDIPlus_BitmapCreateFromHBITMAP($hBitmap)
		$hWnd = _WinAPI_GetDesktopWindow()
		$hDC = _WinAPI_GetDC($hWnd)
		$hBMP = _WinAPI_CreateCompatibleBitmap($hDC, $iX2*3-$iX1*3, $iY2*3-$iY1*3)
		_WinAPI_ReleaseDC($hWnd, $hDC)
		$hImage1 = _GDIPlus_BitmapCreateFromHBITMAP($hBMP)
		$hGraphic = _GDIPlus_ImageGetGraphicsContext($hImage1)
		_GDIPLus_GraphicsDrawImageRect($hGraphic, $hImage, 0, 0, ($iX2-$iX1)*3, ($iY2-$iY1)*3)
		$CLSID = _GDIPlus_EncodersGetCLSID($Ext)
		$tParams = _GDIPlus_ParamInit(2)
		$tData = DllStructCreate("int ColorDepth;int Compression")
		DllStructSetData($tData, "ColorDepth", $giTIFColorDepth)
		DllStructSetData($tData, "Compression", $giTIFCompression)
		_GDIPlus_ParamAdd($tParams, $GDIP_EPGCOLORDEPTH, 1, $GDIP_EPTLONG, DllStructGetPtr($tData, "ColorDepth"))
		_GDIPlus_ParamAdd($tParams, $GDIP_EPGCOMPRESSION, 1, $GDIP_EPTLONG, DllStructGetPtr($tData, "Compression"))
		If IsDllStruct($tParams) Then $pParams = DllStructGetPtr($tParams)
		_GDIPlus_ImageSaveToFileEx($hImage1, $tmpPicfile, $CLSID, $pParams)
		_GDIPlus_ImageDispose($hImage1)
		_GDIPlus_ImageDispose($hImage)
		_GDIPlus_GraphicsDispose ($hGraphic)
		_WinAPI_DeleteObject($hBMP)
		_GDIPlus_Shutdown()
		; do ocr and filter non-digit
		$sOCRTextResult = _UWPOCR_GetText($tmpPicfile)
		Sleep(5)
		writelog("OCRed raw result: " & $sOCRTextResult)
		$sOCRTextResult = StringRegExpReplace($sOCRTextResult, "[*\D]", "")
		_Notify_Modify($hNotify, 0x0000FF, 0xFFFFCC, Default, $tmpGzhname & " : " & $sOCRTextResult)
		ClipPut($sOCRTextResult)
		GUICtrlSetData($stbFQT, "Fetched fans for " & $tmpGzhname & ": " & $sOCRTextResult)
		; read previous fans
		$squery = _SQLite_QuerySingleRow(-1, "SELECT fans FROM accounts WHERE gzhname = '" & $tmpGzhname & "'", $aRow)
		$tmpFans = $aRow[0]
		; update db
		$squery = "UPDATE accounts SET fans='" & $sOCRTextResult & "', prevfans = '" & $tmpFans & "' WHERE gzhname = '" & $tmpGzhname & "'"
		$sexec = _SQLite_Exec(-1, $squery)
		writelog("sql update line, " & $sexec & ", sql: " & $squery)
		; refresh list
		;GUICtrlSetData($lstFans, "")
		_GUICtrlListView_DeleteAllItems($lstFans)
		_SQLite_Query(-1, "SELECT gzhid,gzhname,fans,prevfans FROM accounts", $hQuery)
		Local $aFansList[]
		$i = 0
		While _SQLite_FetchData($hQuery, $aRow) = $SQLITE_OK
			$i = $i + 1
			$tmpFansDiff = $aRow[2] - $aRow[3]
			$aFansList[$i] = GUICtrlCreateListViewItem($aRow[0]&"|"&$aRow[1]&"|"&$aRow[2]&"|"&$aRow[3]&"|"&$tmpFansDiff, $lstFans)
		WEnd
		_SQLite_QueryFinalize($hQuery)
	EndIf
	BlockInput(0)
	GUICtrlSetData($btnCollectFans, "Collect")
	GUISetState(@SW_ENABLE, $frmFansQtyTool)
	;_Notify_Hide($hNotify)
	;_Notify_Set(1, Default, 0xCCFFCC, "Tahoma", True, 250, -1000, True)
	;_Notify_Show(0, "Fans collect routine", "Completed.", 10)
	_Notify_Modify($hNotify, Default, 0xCCFFCC, Default, "Completed.")
EndFunc

Func pasteFans()
	; fill in the fans and click color
	WinActivate("总号表")
	Sleep(5)
	MouseMove(423, 737, 0)
	Sleep(10)
	MouseClick("primary")
	Sleep(5)
	Send("^v")
	;Send(ClipGet())
	Sleep(20)
	;Send($sOCRTextResult)
	;Sleep(20)
	MouseMove(446, 188, 0)
	Sleep(10)
	MouseClick("primary")
	Sleep(100)
	MouseClick("primary")
EndFunc

Func modFans()
	; edit one's fans manually
EndFunc

Func startWordSearch()
	; start strings search helper sequence
	If WinExists("微小宝") Then
		WinActivate("微小宝")
		If $varWordSearchProgress>0 Then
			GUICtrlSetBkColor($txtSStr[$varWordSearchProgress], 0x32CD32)
		EndIf
		$varWordSearchProgress += 1
		$txtSstrings[$varWordSearchProgress] = GUICtrlRead($txtSStr[$varWordSearchProgress])
		If $txtSstrings[$varWordSearchProgress] <> "" Then
			BlockInput(1)
			GUICtrlSetData($btnSearchWord, "Next")
			GUICtrlSetBkColor($txtSStr[$varWordSearchProgress], 0xADD8E6)
			; click somewhere and scroll to top
			MouseMove(335, 421, 0)
			Sleep(10)
			;MouseClick("primary")
			Sleep(5)
			;Send("{HOME}")
			MouseWheel("up", 10)
			Sleep(300)
			; click "mass-send record"
			If $varWordSearchProgress = 1 Then
				MouseMove(182, 480, 0)
				Sleep(10)
				MouseClick("primary")
				Sleep(4500)
			EndIf
			; move cursor to searchbox and send string
			MouseMove(1130, 320, 0)
			Sleep(10)
			MouseClick("primary")
			Sleep(10)
			Send("^a")
			Sleep(5)
			;Send($txtSstrings[$varWordSearchProgress], 1)
			ClipPut($txtSstrings[$varWordSearchProgress])
			Sleep(5)
			Send("^v")
			Sleep(5)
			Send("{ENTER}")
			BlockInput(0)
		Else
			stopWordSearch()
		EndIf

		If $varWordSearchProgress = 10 Then
			$varWordSearchProgress = 0
			GUICtrlSetBkColor($txtSStr[$varWordSearchProgress], 0x32CD32)
			GUICtrlSetData($btnSearchWord, "Start")
		EndIf
	Else
		MsgBox(262208, $title, "Target not running!")
	EndIf
EndFunc

Func stopWordSearch()
	; stop strings search sequence
	$varWordSearchProgress = 0
	$i = 1
	Do
		GUICtrlSetBkColor($txtSStr[$i], 0xFFFFFF)
		$i += 1
	Until $i = 11
	GUICtrlSetData($btnSearchWord, "Start")
EndFunc

Func dbAddColumn()
	writelog("User select DbOp-AddCol.")
	#Region --- CodeWizard generated code Start ---
	;InputBox features: Title=Yes, Prompt=Yes, Default Text=No
	If Not IsDeclared("sInputBoxAnswer") Then Local $sInputBoxAnswer
	$sInputBoxAnswer = InputBox("CAUTION! Direct DB operation can't be undone", "Add a column to the account db with name:",""," M","400","140",Default,Default,120, $frmMain)
	Select
		Case @Error = 0 ;OK - The string returned is valid
			If $sInputBoxAnswer="" Then
				writelog(" - blank input.")
			Else
				writelog(" - input column name: " & $sInputBoxAnswer)
				$squery = "ALTER TABLE accounts ADD COLUMN " & $sInputBoxAnswer
				$sexec = _SQLite_Exec(-1, $squery)
				writelog(" - sql update line, " & $sexec & ", sql: " & $squery)
			EndIf
		Case @Error = 1 ;The Cancel button was pushed
			writelog(" - cancel")
		Case @Error = 3 ;The InputBox failed to open
			writelog("Unable to init dialog. Exiting.")
			Exit
	EndSelect
	#EndRegion --- CodeWizard generated code End ---
	writelog(" - operation end.")
EndFunc

Func dbDropColumn()
	writelog("User select DbOp-DropCol.")
	#Region --- CodeWizard generated code Start ---
	;InputBox features: Title=Yes, Prompt=Yes, Default Text=No
	If Not IsDeclared("sInputBoxAnswer") Then Local $sInputBoxAnswer
	$sInputBoxAnswer = InputBox("CAUTION! Direct DB operation can't be undone", "DROP a column from the account db with name:",""," M","400","140",Default,Default,120, $frmMain)
	Select
		Case @Error = 0 ;OK - The string returned is valid
			If $sInputBoxAnswer="" Then
				writelog(" - blank input.")
			Else
				writelog(" - input column name: " & $sInputBoxAnswer)
				$squery = "ALTER TABLE accounts DROP COLUMN " & $sInputBoxAnswer
				$sexec = _SQLite_Exec(-1, $squery)
				writelog(" - sql update line, " & $sexec & ", sql: " & $squery)
			EndIf
		Case @Error = 1 ;The Cancel button was pushed
			writelog(" - cancel")
		Case @Error = 3 ;The InputBox failed to open
			writelog("Unable to init dialog. Exiting.")
			Exit
	EndSelect
	#EndRegion --- CodeWizard generated code End ---
	writelog(" - operation end.")
EndFunc


Func closeres()
	BlockInput(0)
	_SQLite_Exec(-1, "PRAGMA optimize")
	_SQLite_Close()
	_SQLite_Shutdown()
	If $isResChanged = True Then
		$result = _ChangeScreenRes($ODW, $ODH, @DesktopDepth, @DesktopRefresh)
		writelog("Restore desktop resolution: " & $ODW & "x" & $ODH)
	EndIf
	If $hLog Then FileClose($hLog)
EndFunc