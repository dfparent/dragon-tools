#include <AutoItConstants.au3>
#include <WinAPIFiles.au3>
#include <array.au3>
#include <Process.au3>

; Command line options
;	1: Drive to eject. If omitted, eject all.

if $CmdLine[0] < 1 Then
  MsgBox($MB_SYSTEMMODAL + $MB_ICONWARNING, "Eject Removable Drive", "Please provide a removable drive letter to eject.")
  Exit
endif

$ejectDrive = StringUpper($CmdLine[1])
if StringRight($ejectDrive, 1) <> ":" Then
	$ejectDrive = $ejectDrive & ":"
endif

Local $aDrives = DriveGetDrive($DT_REMOVABLE)
;_arraydisplay($aDrives)

If @error Then
    ; An error occurred when retrieving the drives.;
    MsgBox($MB_SYSTEMMODAL + $MB_ICONWARNING, "", "There are no removable drives to eject.")
	Exit
EndIf

local $DriveLabel = ""
For $i = 1 To $aDrives[0]
	if StringUpper($aDrives[$i]) == $ejectDrive and DriveStatus($aDrives[$i]) == $DS_READY Then
		$DriveLabel = DriveGetLabel($aDrives[$i])
		_RunDos("sync.exe -r "&$aDrives[$i]) ; Flush drive cache
		Sleep(2000)
        EjectDrive($aDrives[$i] & "\")
		Exit
	endif
Next

MsgBox ($MB_SYSTEMMODAL + $MB_ICONWARNING, "Eject drive", "Could not find drive " & $ejectDrive & " to eject.")

Exit

Func EjectDrive($dLetter) ; From VB script, Authors XoLoX and ChrisL
    Local CONST $SSF_DRIVES = 17
    Local $oShell, $oNameSpace, $oDrive, $strDrive
    $strDrive = $dLetter
    $strName=DriveGetLabel($strDrive) & "(" & $strDrive & ")"
    $oShell = ObjCreate("Shell.Application")
    $oNamespace = $oShell.NameSpace($SSF_DRIVES)
    $oDrive = $oNamespace.ParseName($strDrive)
    ;$oDrive.InvokeVerb ("E&ject")
	$oDrive.InvokeVerb ("Eject")
    If DriveGetLabel($strDrive) <> $DriveLabel Or DriveGetLabel($strDrive) = "Removeable Disk" Then
		TrayTip("USB Drive "&$strName&" ejected", "You can now remove the device safely.", 5, 1)
		MsgBox($MB_SYSTEMMODAL + $MB_ICONWARNING,"", $strName & " drive ejected. You can now remove the device safely.",5)
	Else
		MsgBox($MB_SYSTEMMODAL + $MB_ICONWARNING,"", "Not sure if drive " & $strDrive & " was ejected.  Drive label = " & DriveGetLabel($strDrive))
		; Some sort of error.  Sleep to allow the error message to show
		Sleep(5000)
	EndIf


Endfunc
