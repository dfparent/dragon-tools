;#include <MsgBoxConstants.au3>
#include <Misc.au3>

; Only runs once instance of the script
_Singleton( "GlobalHotKeys" )
;MsgBox($MB_SYSTEMMODAL, "", "Run")

; Shortcut combination to run the "type command" script.
HotKeySet("^+!t", "TypeCommand")

While 1
	Sleep(100)
WEnd

Func TypeCommand()
	ShellExecute(@ScriptDir & "\..\TypeCommand\TypeCommand.vbs")
Endfunc
