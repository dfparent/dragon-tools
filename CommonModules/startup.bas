'#Uses "paths.bas"

Public Sub DoStartupCommands()

    ' Disable Word Add in
    Dim result As String
    Dim command As String

    If MsgBox("Disable Word Add-In?", vbYesNo, "Startup Commands") = vbYes Then
        command = "regedit.exe /s """ & getFile("disable word add in") & """"
        'Msgbox command
        Shell(command)
    End If

    If MsgBox("Start Window Highlight?", vbYesNo, "Startup Commands") = vbYes Then
        AppBringUp(getFile("show active window"))
        Wait(1.0)
    End If

    If MsgBox("Start Show Flags utility?", vbYesNo, "Startup Commands") = vbYes Then
        AppBringUp(getFile("flags"))
        Wait(1.0)
    End If

    If MsgBox("Start Time Tracker?", vbYesNo, "Startup Commands") = vbYes Then
        AppBringUp(getFile("time tracker"))
    End If

End Sub

