#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#Region ### START Koda GUI section ### Form=
$Form1 = GUICreate("Form1", 440, 284, 192, 124)
$btnEnd = GUICtrlCreateButton("move to end", 20, 8, 135, 25)
$btnHome = GUICtrlCreateButton("move to home", 180, 8, 147, 25)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Exit
		Case $btnEnd
			movetoend()
	EndSwitch
WEnd


Func movetoend()
	MouseMove(296, 643, 0)
	MouseClick("primary")
	Send("{CTRLDOWN}")
	Send("{END}")
	Send("{CTRLUP}")
EndFunc