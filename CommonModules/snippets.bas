'#Uses "cache.bas"
'#Uses "utilities.bas"
'#Language "WWB.NET"
Option Explicit On

Public Function getSnippet(snippet As String) As String
    On Error GoTo ErrorHandler
    Dim snippets As Object

    snippets = GetSnippets()
    snippet = LCase(snippet)
    If Not snippets.ContainsKey(snippet) Then
        'MsgBox "There is no defined '" & path & "' path."
        Return ""
    End If

    Return snippets(snippet)(0)

ErrorHandler:
    Msgbox("Error: " & err.description)
    Return ""
End Function

Public Sub PrintSnippet(dictation As String)
    ' Need to use clipboard to avoid problems with new lines activating default buttons
    Dim text As String
    text = getSnippet(dictation)
    If text = "" Then
        Beep
        Exit Sub
    End If

    PutClipboard(text)
    SendKeys("^v")

End Sub