'#Language "WWB.NET"
'#Uses "utilities.bas"
'#Uses "window.bas"
'#Uses "cache.bas"

'option explicit

Private Const ADOBE_BOOKMARK_VALUE = "adobe-bookmarks"
Private Const ZOOM_INCREMENT = 10

Public Sub Zoom(inOut As String, Optional count As Integer = 1)
    SendKeys("^y")
    WaitForWindow("Zoom To")
    SendKeys("^c")
    Dim zoomValue As String
    zoomValue = GetClipboard()


    If zoomValue = "Actual Size" Then
        zoomValue = "100"
    ElseIf Instr(1, zoomValue, "Fit") <> 0 Then
        ' Guess
        zoomValue = "50"
    Else
        ' Remove % sign
        zoomValue = Left(zoomValue, Len(zoomValue) - 1)
    End If


    Dim zoomNum As Integer
    zoomNum = CInt(zoomValue)

    If inOut = "In" Then
        zoomNum = zoomNum + ZOOM_INCREMENT * count
    ElseIf inOut = "Out" Then
        zoomNum = zoomNum - ZOOM_INCREMENT * count
    End If

    SendKeys(CStr(zoomNum))
    SendKeys("~")
End Sub

Public Sub SetBookmark()
    SendKeys("%vng")
    WaitForWindow("Go To Page")
    SendKeys("^c")

    Dim pageNum As String
    pageNum = GetClipboard()

    SendKeys("{Escape}")

    Dim bookmarks() As String
    If CacheValueExists(ADOBE_BOOKMARK_VALUE) Then
        bookmarks = GetCacheValue(ADOBE_BOOKMARK_VALUE)
        ReDim Preserve bookmarks(UBound(bookmarks) + 1)
    Else
        ReDim bookmarks(0)
    End If

    bookmarks(Ubound(bookmarks)) = pageNum

    UpdateCacheValue(ADOBE_BOOKMARK_VALUE, bookmarks)

End Sub

Public Sub GoToPreviousBookmark()
    Dim bookmarks() As String

    If Not CacheValueExists(ADOBE_BOOKMARK_VALUE) Then
        MsgBox("There are no bookmarks to go back to.  Set a bookmark first.")
        Exit Sub
    End If

    bookmarks = GetCacheValue(ADOBE_BOOKMARK_VALUE)

    SendKeys("%vng")
    WaitForWindow("Go To Page")
    SendKeys(bookmarks(UBound(bookmarks)))
    SendKeys("~")

    ' Pop array
    If UBound(bookmarks) = 0 Then
        RemoveCacheValue(ADOBE_BOOKMARK_VALUE)
    Else
        ReDim Preserve bookmarks(UBound(bookmarks) - 1)
        UpdateCacheValue(ADOBE_BOOKMARK_VALUE, bookmarks)
    End If

End Sub
