#include <AutoItConstants.au3>
#include <WinAPIFiles.au3>
#include <array.au3>

Local $aArray = DriveGetDrive($DT_REMOVABLE)
;_arraydisplay($aArray)

If @error Then
    ; An error occurred when retrieving the drives.;
    MsgBox(0, "", "An error occurred while obtaining the list of removable drives.")
	Exit
EndIf

$out = ""
For $i = 1 To $aArray[0]
	$out = $out & $aArray[$i] & " " & DriveGetLabel($aArray[$i]) & @CRLF
Next

MsgBox (0, "List Removable Drives", $out)
