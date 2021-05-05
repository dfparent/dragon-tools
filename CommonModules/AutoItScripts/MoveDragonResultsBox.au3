#include <Array.au3>
#include <MsgBoxConstants.au3>
#include <AutoItConstants.au3>
#include <StringConstants.au3>
#include <WinAPIGdi.au3> ;for _WinOnMonitor()
#include <WindowsConstants.au3>
#include <WinAPI.au3>
#include <WinAPIConv.au3>
#include <SendMessage.au3>

RunScript()

; THIS DOES NOT WORK

; Command line options
;	1: Left|Right (move window to the left or right screen)

Func RunScript()

	if $CmdLine[0] < 1 Then
		MsgBox($MB_SYSTEMMODAL + $MB_ICONWARNING, "Move Results Box", "You must supply ""Left"" or ""Right"" in the command line.")
		return
	endif

	local $direction = ""
	$direction = $CmdLine[1]
	$direction = StringLower($direction)
	if $direction <> "left" and $direction <> "right" Then
		MsgBox($MB_SYSTEMMODAL + $MB_ICONWARNING, "Move Results Box", "You must supply ""Left"" or ""Right"" in the command line.")
		return
	EndIf

	; Get dragon results window handle
	local $dgnHwnd = WinGetHandle("[CLASS:DgnResultsBoxWindow]")
	If Not WinActivate($dgnHwnd) Then
		Beep(500, 100)
	EndIf

	; Move window to Left or right screen
	_WinOnMonitor($dgnHwnd, $direction)
	_SendMessage($dgnHwnd, $WM_LBUTTONDOWN, 1, _WinAPI_MakeLong(10, 10))
	Sleep(100)
	_SendMessage($dgnHwnd, $WM_LBUTTONUP, 0, _WinAPI_MakeLong(10, 10))

EndFunc   ;==>Example


; Returns found monitor handle:

Func _WinOnMonitor($winHandle, $moveDirection)
    Local $aMonitors = _WinAPI_EnumDisplayMonitors()
	; This method returns the following array:
	; [0][0] - Number of rows in array (n)
    ; [0][1] - Unused
    ; [n][0] - A handle to the display monitor.
    ; [n][1] - $tagRECT structure defining a display monitor rectangle or the clipping area.

    If IsArray($aMonitors) Then
		; Add space for dimensions of each monitor
        ReDim $aMonitors[$aMonitors[0][0] + 1][5]

		; For each monitor, replace the position rect with position values
		; [0][0] - Number of rows in array (n)
		; [0][1] - Unused
		; [n][0] - A handle to the display monitor.
		; [n][1] - X Pos
		; [n][2] - Y Pos
		; [n][3] - Width
		; [n][4] - Height

        For $ix = 1 To $aMonitors[0][0]
            $aPos = _WinAPI_GetPosFromRect($aMonitors[$ix][1])
            For $j = 0 To 3
                $aMonitors[$ix][$j + 1] = $aPos[$j]
            Next
        Next

		$winPos = WinGetPos($winHandle)
		; $aArray[0] = X position
		; $aArray[1] = Y position
		; $aArray[2] = Width
		; $aArray[3] = Height

		$iXPos = $winPos[0]
		$iYPos = $winPos[1]

		; Find monitor for given position
		For $ixMonitor = 1 to $aMonitors[0][0] ; Step through array of monitors
			If $iXPos > $aMonitors[$ixMonitor][1] And $iXPos <  $aMonitors[$ixMonitor][1] + $aMonitors[$ixMonitor][3] Then
				If $iYPos > $aMonitors[$ixMonitor][2] And $iYPos < $aMonitors[$ixMonitor][2] + $aMonitors[$ixMonitor][4] Then
					; Found the monitor.  Now get new window coords
					$newX = $iXPos + $aMonitors[$ixMonitor][3]
					WinMove($winHandle, "", $newX, $iYPos)
					Return
				EndIf
			EndIf
		Next
    EndIf

    Return
EndFunc ;==>  _WinOnMonitor