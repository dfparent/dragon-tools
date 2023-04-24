'#Uses "keyboard.bas"
'#Uses "window.bas"
'#Language "WWB.NET"
Option Explicit On

Public Function GetOutlookApp() As Object
    Return CreateObject("Outlook.Application")
End Function

Public Sub ShowCategories()
    msgbox("Hi")
    Dim obj As Object
    obj = GetOutlookApp()

End Sub

Public Sub SelectFolder(ListVar1 As String, Optional subFolderToken As String = ",")
    'SendKeys "^+i"
    'Wait 0.1

    SendKeys("^y")
    Wait(0.3)

    If Not WaitForWindow("Go to Folder", 0.5) Then
        ' Some folders do not Work with ^y for some stupid reason.
        ' For these cases, we need to First go to the inbox and then use ^y
        SendKeys("^+i")
        Wait(0.1)
        SendKeys("^y")
        Wait(0.3)
    End If


    Dim folders() As String
    folders = Split(ListVar1, subFolderToken)
    SendKeys("{Home}{Down}{Left}{Right}")
    Wait(0.1)
    Dim aFolder As String
    For Each aFolder In folders
        SendKeys(aFolder)
        Wait(0.1)
        SendKeys("{Right}")
        Wait(0.2)
    Next

    SendKeys("{Enter}")
End Sub

Public Sub MoveToFolder(Listvar1 As String, Optional subFolderToken As String = ",")
    'SendKeys("%hmvo")
    SendKeys("^+v")
    If Not WaitForWindow("Move Items", 0.5) Then
        Beep
        Exit Sub
    End If

    Dim folders() As String
    folders = Split(Listvar1, subFolderToken)
    SendKeys("{Home}{Left}{Right}")
    Wait(0.1)
    Dim aFolder As String
    For Each aFolder In folders
        SendKeys(aFolder)
        Wait(0.1)
        SendKeys("{Right}")
        Wait(0.2)
    Next


End Sub

Public Sub GoToField(fieldName As String, Optional name As String = "")
    'ShowKeyTips()
    Select Case LCase(fieldName)
        Case "to"
            SendKeys("%u+{tab 3}")
        Case "cc"
            SendKeys("%u+{tab 2}")
        Case "bcc"
            SendKeys("%u+{tab}")
		case else
			Msgbox("Unsupported field: " & fieldName)
    End Select

    'SendKeys("{end}")

    If name <> "" Then
        SendKeys(name)
    End If

End Sub

Public Sub SnoozeReminder(ListVar1 As String)
    If Not CheckWindowText("Reminder(s)") Then
        Exit Sub
    End If

    SendKeys("%c")
    Dim snooze As String
    If ListVar1 = "Half Hour" Then
        snooze = "30 Minutes"
    Else
        snooze = ListVar1
    End If
    SendKeys(snooze)
    'SendKeys "{Down}%s"
    SendKeys("%s")
    Wait(0.3)

    ' Sometimes the reminder window is now closed
    If CheckWindowText("Reminder(s)") Then
        SendKeys("{Tab 3}")
    End If


End Sub