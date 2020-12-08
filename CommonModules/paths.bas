'#Uses "cache.bas"
'#Language "WWB.NET"
Option Explicit On

Public Function getPath(path As String) As String
    On Error GoTo ErrorHandler
    Dim paths As Object

    paths = GetPaths()
    path = LCase(path)
    If Not paths.ContainsKey(path) Then
        'MsgBox "There is no defined '" & path & "' path."
        Return ""
    End If

    Return paths(path)(0)

ErrorHandler:
    Msgbox "Error: " & err.description
    Return ""
End Function

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

Public function getUrl(url as String)
    on error goto ErrorHandler
    Dim urls As Object
    urls = GetURLs()
    url = LCase(url)
    if not urls.ContainsKey(url) then
        ' MsgBox "There is no defined '" & file & "' file."
        return ""
    end If

    Return urls(url)(0)

ErrorHandler:
	Msgbox "Error: " & err.description
	return ""
end function
