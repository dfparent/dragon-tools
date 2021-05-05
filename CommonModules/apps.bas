'#Uses "clickbynumbers.bas"
'#Uses "keyboard.bas"
'#Uses "cache.bas"
'#Language "WWB.NET"

' These numbers correspond with the positions of the pinned app icons on the Windows taskbar.  
' 1 is the left most position, 0 is the right most position.
Public Function getAppNumber(appName As String) As String
    On Error GoTo ErrorHandler

    Dim apps As Object
    apps = GetApps()

    If Not apps.ContainsKey(appName) Then
        'Beep
        getAppNumber = ""
        Exit Function
    End If

    getAppNumber = apps(appName)(0)

    Exit Function

ErrorHandler:
    Msgbox("Error in getAppNumber: " & err.description)

End Function

' Returns true if successful, false if not
Public Function goToApp(appName As String, Optional count As String = "1", Optional openNew As Boolean = False) As Boolean
    On Error GoTo ErrorHandler

    Dim number As String
    number = getAppNumber(appName)
    If number = "" Then
        goToApp = False
        Exit Function
    End If

    PressWindowsKey(number, count, openNew)

    ' When running ClickByNumbers, we don't get new flags
    ' Do a flag refresh
    If IsClickByNumbersRunning() Then
        Wait(0.1)
        ' Refresh flags
        SendKeys("%^+{F5}")
    End If

    goToApp = True
    Exit Function

ErrorHandler:
    MsgBox("Error in goToApp: " & err.description)
End Function

