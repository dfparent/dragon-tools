#include <MsgBoxConstants.au3>
#include <AutoItConstants.au3>

RunScript()

Func Get_Systray_Index($sToolTipTitle)

    ; Find systray handle
    $hSysTray_Handle = ControlGetHandle("[Class:Shell_TrayWnd]", "", "[Class:ToolbarWindow32;Instance:1]")
    If @error Then
        MsgBox(16, "Error", "System tray not found")
        Exit
    EndIf

    ; Get systray item count
    Local $iSystray_ButCount = _GUICtrlToolbar_ButtonCount($hSysTray_Handle)
    If $iSystray_ButCount = 0 Then
        MsgBox(16, "Error", "No items found in system tray")
        Exit
    EndIf

    ; Look for wanted tooltip
    ;For $iSystray_ButtonNumber = 0 To $iSystray_ButCount - 1 =========== "OLD" instruction
    For $iSystray_ButtonNumber = 1 To $iSystray_ButCount  ; =========== MODIFIED instruction
         If StringInStr(_GUICtrlToolbar_GetButtonText($hSysTray_Handle, $iSystray_ButtonNumber), $sToolTipTitle) = 1 Then ExitLoop
    Next

    ;If $iSystray_ButtonNumber = $iSystray_ButCount Then ;=========== "OLD" instruction
    If $iSystray_ButtonNumber = $iSystray_ButCount + 1 Then ; =========== MODIFIED instruction
        Return 0 ; Not found
    Else
        Return $iSystray_ButtonNumber ; Found
    EndIf

EndFunc