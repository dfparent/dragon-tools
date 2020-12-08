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

Public Sub SelectFolder(ListVar1 As String)
    'SendKeys "^+i"
    'Wait 0.1

    SendKeys("^y")
    Wait(0.3)

    Dim folders() As String
    folders = Split(ListVar1, ",")
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

Public Sub MoveToFolder(Listvar1 As String)
    SendKeys("%hmvo")
    Wait(0.5)

    Dim folders() As String
    folders = Split(Listvar1, ",")
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
