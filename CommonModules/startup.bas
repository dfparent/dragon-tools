'#Language "WWB.NET"
'#Uses "autoIt.bas"
'#Uses "paths.bas"

Public Sub DoStartupCommands()

    ' Disable Word Add in
    Dim result As String
    Dim command As String
	dim objShell as Object
	
	' Need to use this on some systems for some reason
	objShell = CreateObject("WScript.Shell")
		
    If MsgBox("Disable Word Add-In?", vbYesNo, "Startup Commands") = vbYes Then
        command = "regedit.exe /s """ & getFile("disable word add in") & """"
        'Msgbox command
		objShell.Run command, 0, true
        'Shell(command)
    End If

    If MsgBox("Disable Excel Add-In?", vbYesNo, "Startup Commands") = vbYes Then
        command = "regedit.exe /s """ & getFile("disable excel add in") & """"
        'Msgbox command
        'Shell(command)
		objShell.Run command, 0, true
        
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

    ' Setup global commands
    ' The script will only allow 1 instance to run, so we can safely just run it.
    If MsgBox("Run Global Hot Keys?", vbYesNo, "Startup Commands") = vbYes Then
        runAutoIt("SetGlobalHotKeys.au3")
    End If


End Sub

