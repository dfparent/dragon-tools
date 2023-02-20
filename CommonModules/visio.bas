'#Uses "keyboard.bas"
'#Language "WWB.NET"
Option Explicit On

Private Const MACRO_MODULE_NAME = "VisioUtilities.vss"
Private Const MACRO_NAME_PREFIX = "VisioUtilities.Module1."

Public Sub RunMacro(macroName As String)
    SendKeys("%{F8}")
    Wait(0.1)
    SendKeys("%a")
    SendKeys(MACRO_MODULE_NAME)
    SendKeys("%m")
    SendKeys(MACRO_NAME_PREFIX)
    SendKeys(macroName)
    SendKeys("%r")
End Sub

Public Sub IncreaseDecreaseHeightWidth(heightWidth As String, theOperator As String, increment As Integer)

    If heightWidth = "height" Then
        If theOperator = "+" Then
            RunMacro("IncreaseHeight")
        ElseIf theOperator = "-" Then
            RunMacro("DecreaseHeight")
        Else
            Exit Sub
        End If
    ElseIf heightWidth = "width" Then
        If theOperator = "+" Then
            RunMacro("IncreaseWidth")
        ElseIf theOperator = "-" Then
            RunMacro("DecreaseWidth")
        Else
            Exit Sub
        End If
    Else
        Exit Sub
    End If

    Wait(0.1)

    ' Given numbers are arbitrary and are integers.The width is measured in inches (most likely).
    ' Divide given numbers by 10. A tenth of an inch is more precise.
    SendKeys(CStr(increment / 10))
    SendKeys("{enter}")
End Sub