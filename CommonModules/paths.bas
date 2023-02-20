'#Uses "cache.bas"
'#Uses "PathSelect-dlg.bas"
'#Language "WWB.NET"
Option Explicit On

Public Function getPath(path As String) As String
    On Error GoTo ErrorHandler
    Dim paths As Object

    paths = GetPaths()
    path = LCase(path)
    If Not paths.ContainsKey(path) Then
        ' Returns empty string if No selection
        Return SelectPath(path)
    End If

    Return paths(path)(0)

ErrorHandler:
    Msgbox "Error: " & err.description
    Return ""
End Function

Public Sub ShowPaths()
    On Error GoTo ErrorHandler
    Dim paths As Object

    paths = GetPaths()
    SelectPath("")

    Exit Sub

ErrorHandler:
    Msgbox "Error: " & err.description

End Sub

Public Function getFile(file As String) As String
    On Error GoTo ErrorHandler
    Dim files As Object
    files = GetFiles()
    file = LCase(file)
    If Not files.ContainsKey(file) Then
        ' MsgBox "There is no defined '" & file & "' file."
        Return ""
    End If

    Return files(file)(0)

ErrorHandler:
    Msgbox "Error: " & err.description
    Return ""
End Function

Public Function getUrl(url As String) As String
    On Error GoTo ErrorHandler
    Dim urls As Object
    urls = GetURLs()
    url = LCase(url)
    If Not urls.ContainsKey(url) Then
        'MsgBox "There is no defined '" & url & "' url."
        ' Returns empty string if No selection
        Return SelectUrl(url)
    End If

    Return urls(url)(0)

ErrorHandler:
    Msgbox "Error: " & err.description
    Return ""
End Function

Public Function SelectPath(name As String) As String
    On Error GoTo ErrorHandler
    Dim paths As Object

    paths = GetPaths()

    Dim selectedPath As String
    If Not SelectDataDialog(name, paths, "path", selectedPath) Then
        ' User Cancelled
        Exit Function
    End If

    Return selectedPath


ErrorHandler:
    Msgbox("Error: " & err.description)
    Return ""

End Function

Public Function SelectUrl(url As String) As String
    On Error GoTo ErrorHandler
    Dim urls As Object

    urls = GetURLs()

    Dim selectedUrl As String
    If Not SelectDataDialog(url, urls, "URL", selectedUrl) Then
        ' User Cancelled
        Exit Function
    End If

    Return selectedUrl


ErrorHandler:
    Msgbox("Error: " & err.description)
    Return ""

End Function