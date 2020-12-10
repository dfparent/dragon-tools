'#Language "WWB.NET"
'#Uses "utilities.bas"
'#Uses "cache.bas"

'option explicit

Private Const ADOBE_BOOKMARK_VALUE = "adobe-bookmarks"

Public Sub SetBookmark()
    SendKeys("%vng")
    Wait 0.1
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
    Wait(0.1)
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
