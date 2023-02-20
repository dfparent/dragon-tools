'#Uses "utilities.bas"
'Option Explicit On

private const registryAppName = "..\Douglas Parent\Click By Numbers"
private const registrySection = "GlobalSettings"
private const registryKey = "BringToForegroundHotkey"
private const registryDefaultValue = "F1"

Public Sub PressClickByNumbersHotKey()
    ' Get  hot key
    dim hotkey as string
    hotkey = GetSetting(registryAppName, registrySection, registryKey, registryDefaultValue)
    SendKeys("%^+{" & hotkey & "}")
    Wait 0.1
End Sub

Public function IsClickByNumbersRunning() as Boolean
    IsClickByNumbersRunning = IsProcessRunning("ClickByNumbers")
End Function

' This works with the Window Labeler application
Public Sub DoClickByNumbersAction(numbers As String, Optional action As String = "click")
    Dim lcaseAction As String
    lcaseAction = LCase(action)

    PressClickByNumbersHotKey()
    SlowTypeString(numbers)

    Select Case lcaseAction
        Case "click"
            SendKeys("{Enter}")

        Case "shift click"
            SendKeys("+{Enter}")

        Case "control click"
            SendKeys("^{Enter}")

        Case "double-click"
            SendKeys("%{Enter}")

        Case "right-click"
            SendKeys("^+{Enter}")

        Case "pick up", "mouse down"
            SendKeys("{Down}")

        Case "drag"
            SendKeys("g")

        Case "put down", "mouse up"
            SendKeys("{Up}")

        Case "drop"
            SendKeys("p")

        Case "control drop"
            SendKeys("^p")

        Case "shift drop"
            SendKeys("+p")

        Case "move to", "move"
            SendKeys("{Tab}")

        Case Else
            Msgbox("Unrecognized Action: " & action)
    End Select

    Wait 0.1

End Sub