#include <Array.au3>
#include <MsgBoxConstants.au3>
#include <AutoItConstants.au3>
#include <StringConstants.au3>

; Command line options
;	1: Title string to search
RunScript()

Func RunScript()
   ; Match any substring, case insensitive
   Opt("WinTitleMatchMode", -2)

   if $CmdLine[0] < 1 Then
	  MsgBox($MB_SYSTEMMODAL + $MB_ICONWARNING, "Activate Window", "You must supply a string in the command line.")
	  return
   endif

   ; Get dragon results window handle
   local $dgnHwnd = WinGetHandle("[CLASS:DgnResultsBoxWindow]")

   ; Get active window handle
   local $activeHwnd = WinGetHandle("[ACTIVE]")

   ; Get windows that match the text
   local $titleText = ""
   local $titleTextNoSpaces = ""
   $titleText = $CmdLine[1]
   $titleTextNoSpaces = StringStripWS($CmdLine[1], $STR_STRIPALL)

	; Get matches both with spaces and without.  Combine and remove dupes.
   ;MsgBox($MB_SYSTEMMODAL + $MB_ICONINFORMATION, "Debug", $titleText)
   Local $aList = WinList($titleText)
   local $aSubList = WinList($titleTextNoSpaces)
   local $count = $aList[0][0] + $aSubList[0][0]
   _ArrayConcatenate ($aList, $aSubList, 1)
   $aList[0][0] = $count
   ;_ArrayDisplay($aList)

	; Remove dupes
	;local $uniqueList = []	; map
	;local $handle = ""
	;for $i = 1 to $aList[0][0]
	;	$uniqueList[($aList[$i][0])] = $aList[$i][1]
	;Next
	;_ArrayDisplay($uniqueList)

   ; Loop through the array.  Do not activate the dragon resuls box which will always have the desired text
   local $countMax = 1
    For $i = 1 To $aList[0][0]
		$count = 0
        If $aList[$i][1] <> $dgnHwnd and $aList[$i][1] <> $activeHwnd Then
			; Try to activate a few times
		   Do
				local $handle = WinActivate($aList[$i][1])
				local $state = WinGetState($handle)
				;MsgBox($MB_SYSTEMMODAL + $MB_ICONWARNING, "Done", $state)
				If BitAND($state, $WIN_STATE_ACTIVE) and BitAND($state, $WIN_STATE_VISIBLE) Then
				    ;MsgBox($MB_SYSTEMMODAL + $MB_ICONWARNING, "Done", "Hi")
				    ;If BitAND($state, $WIN_STATE_MINIMIZED) Then
					;   WinSetState($handle, "", @SW_RESTORE)
					;endIf
				    SplashTextOn("", "Found window: " & @CRLF & $aList[$i][0], 500, 100, @DesktopWidth - 500, @DesktopHeight - 100, $DLG_NOTITLE)
				    Sleep(3000)
				   ;MsgBox($MB_SYSTEMMODAL + $MB_ICONWARNING, "Done", $title)
				   Exit
				Else
					;MsgBox($MB_SYSTEMMODAL + $MB_ICONWARNING, "Debug", $state)
					SplashTextOn("", "Trying: " & @CRLF & $aList[$i][0], 500, 100, @DesktopWidth - 500, @DesktopHeight - 100, $DLG_NOTITLE)
					$count = $count + 1
					;Sleep(10)
				endif
			Until $count >= $countMax
			;SplashOff()
        EndIf
	 Next

	 ; Not found
	 ;MsgBox($MB_SYSTEMMODAL + $MB_ICONINFORMATION, "Debug", $titleText)
	 Beep(500, 100)
EndFunc   ;==>Example
