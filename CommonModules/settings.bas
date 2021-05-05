'#Uses "cache.bas"
'#Uses "utilities.bas"
'#Language "WWB.NET"
Option Explicit On

Public Function getSetting(setting As String) As String
    On Error GoTo ErrorHandler
    Dim settings As Object

    settings = GetSettings()
    setting = LCase(setting)
    If Not settings.ContainsKey(setting) Then
        'MsgBox "There is no defined '" & path & "' path."
        Return ""
    End If

    Return settings(setting)(0)

ErrorHandler:
    Msgbox("Error: " & err.description)
    Return ""
End Function

