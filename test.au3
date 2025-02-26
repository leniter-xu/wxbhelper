#NoTrayIcon
DllCall("User32.dll","bool","SetProcessDPIAware")

#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <ScreenCapture.au3>
#include <GDIPlus.au3>
#include <WinAPIGdi.au3>
#include "UWPOCR.au3"

#Region ### START Koda GUI section ### Form=
$Form1 = GUICreate("Form1", 507, 309, 600, 100)
$bFile = GUICtrlCreateButton("get text from file", 12, 12, 150, 25)
$bScr = GUICtrlCreateButton("get text from screen", 166, 12, 150, 25)
$bLang = GUICtrlCreateButton("languages", 320, 12)
$strTxt = GUICtrlCreateLabel("", 12, 44, 330, 200)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

_UWPOCR_Log(__UWPOCR_Log)

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $bFile
			$sOCRTextResult = _UWPOCR_GetText(FileOpenDialog("Select Image", @ScriptDir & "\", "Images (*.jpg;*.bmp;*.png;*.tif;*.gif)"))
			MsgBox(0,"",$sOCRTextResult)
		Case $bScr
			;_GDIPlus_Startup()
			$tmpPicfile = @ScriptDir&"\test.tif"
			$iX1 = 424
			$iY1 = 462
			$iX2 = 532
			$iY2 = 596
			;_ScreenCapture_Capture("", 424, 462, 524, 596, False)

			Local $giTIFColorDepth = 24
			Local $giTIFCompression = $GDIP_EVTCOMPRESSIONNONE
			Local $Ext=""
			$hBitmap = _ScreenCapture_Capture("", $iX1, $iY1, $iX2, $iY2, False)
			_GDIPlus_Startup()
			; Convert the image to a bitmap
			$hImage = _GDIPlus_BitmapCreateFromHBITMAP($hBitmap)
			$hWnd = _WinAPI_GetDesktopWindow()
			$hDC = _WinAPI_GetDC($hWnd)
			$hBMP = _WinAPI_CreateCompatibleBitmap($hDC, $iX2*3-$iX1*3, $iY2*3-$iY1*3)
			_WinAPI_ReleaseDC($hWnd, $hDC)
			$hImage1 = _GDIPlus_BitmapCreateFromHBITMAP($hBMP)
			$hGraphic = _GDIPlus_ImageGetGraphicsContext($hImage1)
			_GDIPLus_GraphicsDrawImageRect($hGraphic, $hImage, 0, 0, ($iX2-$iX1)*3, ($iY2-$iY1)*3)
			$CLSID = _GDIPlus_EncodersGetCLSID($Ext)
			; Set TIFF parameters
			$tParams = _GDIPlus_ParamInit(2)
			$tData = DllStructCreate("int ColorDepth;int Compression")
			DllStructSetData($tData, "ColorDepth", $giTIFColorDepth)
			DllStructSetData($tData, "Compression", $giTIFCompression)
			_GDIPlus_ParamAdd($tParams, $GDIP_EPGCOLORDEPTH, 1, $GDIP_EPTLONG, DllStructGetPtr($tData, "ColorDepth"))
			_GDIPlus_ParamAdd($tParams, $GDIP_EPGCOMPRESSION, 1, $GDIP_EPTLONG, DllStructGetPtr($tData, "Compression"))
			If IsDllStruct($tParams) Then $pParams = DllStructGetPtr($tParams)
			; Save TIFF and cleanup
			_GDIPlus_ImageSaveToFileEx($hImage1, $tmpPicfile, $CLSID, $pParams)
			_GDIPlus_ImageDispose($hImage1)
			_GDIPlus_ImageDispose($hImage)
			_GDIPlus_GraphicsDispose ($hGraphic)
			_WinAPI_DeleteObject($hBMP)
			_GDIPlus_Shutdown()

			;$scrcap = _ScreenCapture_Capture("", 430, 470, 518, 506, False)
			;$hBitmap = _GDIPlus_BitmapCreateFromHBITMAP($scrcap)
			;Local $sOCRTextResult = _UWPOCR_GetText($hBitmap, Default, True)
			$sOCRTextResult = _UWPOCR_GetText($tmpPicfile)
			GUICtrlSetData($strTxt, $sOCRTextResult)
			$sOCRTextResult = StringRegExpReplace($sOCRTextResult, "[*\D]", "")
			MsgBox(64,"",Int($sOCRTextResult))
			;_WinAPI_DeleteObject($hBitmap)
			;_GDIPlus_Shutdown()
		Case $bLang
			MsgBox(64, "", _UWPOCR_GetSupportedLanguages())
		Case $GUI_EVENT_CLOSE
			Exit

	EndSwitch
WEnd
