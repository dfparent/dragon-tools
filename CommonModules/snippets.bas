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
        Beep
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

Public Sub ManageSnippets()
    On Error GoTo ErrorHandler

    Dim snippets As Object
    snippets = GetSnippets()

    Dim newSnippets() As String
    Dim snippetsList As New System.Collections.Generic.List(Of String)
    Dim snippetsArray() As String
    Dim pair As Object
    Dim aName As String

    For Each pair In snippets
        ' The value property is an array
        For Each aName In pair.Value
            snippetsList.Add(aName)
        Next
    Next

    snippetsArray = snippetsList.ToArray()
    '    If Not ShowManageSnippetsDialog(snippetsArray, newSnippets) Then
    '    ' User Cancelled
    '    Exit Sub
    '    End If

    'msgbox(ArrayToString(newSnippets, 0))

    ' Update snippets list and save
    snippets.Clear()

    Dim initials As String
    Dim parts() As String
    Dim part As String
    Dim valueArray() As String
    For Each aName In newSnippets
        If aName = "" Then
            Continue For
        End If

        parts = Split(aName, " ")
        initials = ""
        For Each part In parts
            initials = initials & part.Substring(0, 1).ToLower()
        Next
        If Not snippets.ContainsKey(initials) Then
            snippets.Add(initials, {aName})
        Else
            '  Add to existing array
            valueArray = snippets(initials)
            ReDim Preserve valueArray(UBound(valueArray) + 1)
            valueArray(UBound(valueArray)) = aName
            'msgbox(ArrayToString(valueArray, 0, -1, ","))
            snippets(initials) = valueArray
        End If
    Next

    ' Savesnippets


    Exit Sub

ErrorHandler:
    Msgbox("Error: " & err.description)
    Return
End Sub

