;#include <MsgBoxConstants.au3>
#include <Misc.au3>
#include <AutoItConstants.au3>

; Only runs once instance of the script
_Singleton( "GlobalHotKeys" )
;MsgBox($MB_SYSTEMMODAL, "", "Run")
Opt("WinTitleMatchMode", 2)

; Shortcut combination to run the "type command" script.
HotKeySet("^+!t", "TypeCommand")

; Shortcut combination to run mute / unmute mic for Zoom
HotKeySet("^+!z", "MuteUnmuteZoom")
HotKeySet("^+!m", "MuteUnmuteTeams")
HotKeySet("^+!y", "MuteUnmutePhone")

While 1
	Sleep(100)
WEnd

Func TypeCommand()
	ShellExecute(@ScriptDir & "\..\TypeCommand\TypeCommand.vbs")
Endfunc

Func MuteUnmuteZoom()
	; Switch to Zoom
	If WinActivate("Zoom Meeting") Then
		;MsgBox($MB_SYSTEMMODAL + $MB_ICONINFORMATION, "Debug", "Zoom activated")
		if WinActive("Zoom") Then
			;MsgBox($MB_SYSTEMMODAL + $MB_ICONINFORMATION, "Debug", "Zoom active")
			; Mute/unmute
			Send("!a")
		EndIf
	Else
		MsgBox($MB_SYSTEMMODAL + $MB_ICONINFORMATION, "Mute / Unmute Zoom", "No Zoom meeting is in progress.")
	EndIf
EndFunc

Func MuteUnmuteTeams()
	; Switch to Teams
	Opt("WinTitleMatchMode", 2)

	If WinActivate("Microsoft Teams") Then
		;MsgBox($MB_SYSTEMMODAL + $MB_ICONINFORMATION, "Debug", "Teams activated")
		if WinActive("Microsoft Teams") Then
			;MsgBox($MB_SYSTEMMODAL + $MB_ICONINFORMATION, "Debug", "Teams active")
			; Mute/unmute
			Send("^+m")
		EndIf
	Else
		MsgBox($MB_SYSTEMMODAL + $MB_ICONINFORMATION, "Mute / Unmute Teams", "Teams is not running.")
	EndIf
EndFunc

Func MuteUnmutePhone()
	; Using WinActivate when Avaya is minimized to tray causes it to display BLACK and it needs to be killed
	;If WinActivate("Avaya one-X® Communicator") Then

	; Switch to Avaya using its built in global shortcut.  You have to define the shortcut that brings up the app
	; as ctrl+alt+y

	Send("^!y")
	Sleep(100)
	local $hwnd = WinActive("Avaya")
	if not $hwnd Then
		; If Avaya was already up when it does !y, then it goes away.  Need to !y again.
		Send("^!y")
		Sleep(100)
	endIf

	$hwnd = WinActive("Avaya")
	if $hwnd Then
		;MsgBox($MB_SYSTEMMODAL + $MB_ICONINFORMATION, "Debug", "Avaya active")
		; Mute/unmute
		Opt("MouseCoordMode", 2)
		MouseClick($MOUSE_CLICK_LEFT, 364,93)
	Else
		MsgBox($MB_SYSTEMMODAL + $MB_ICONINFORMATION, "Mute / Unmute Avaya", "Avaya is not running.")
	EndIf
EndFunc