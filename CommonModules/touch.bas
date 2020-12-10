'#Uses "utilities.bas"
'#Uses "cache.bas"
'#Uses "mouse.bas"
'#Uses "imports.bas"
'#Language "WWB.NET"

Option Explicit On

Public Function GetTouchFilePath() As String
    Dim processName As String
    processName = GetForegroundProcessName()
    GetTouchFilePath = GetTouchesFileName(processName)
End Function

Public Sub ClickTouchLocation(touchLocationName As String)
    DoTouchLocation(touchLocationName, MouseButton.Left)
End Sub

Public Sub RightClickTouchLocation(touchLocationName As String)
    DoTouchLocation(touchLocationName, MouseButton.Right)
End Sub

Public Sub MoveTouchLocation(touchLocationName As String)
    DoTouchLocation(touchLocationName)
End Sub

' Clicks the mouse at the given location
Private Sub DoTouchLocation(touchLocationName As String, Optional clickAction = Nothing)
    Dim touches As Object
    touches = GetTouches(GetForegroundProcessName())


    If touches.ContainsKey(touchLocationName) Then
        Dim values() As String
        values = touches(touchLocationName)
        PositionMouseRelativeToActiveWindow(CLng(values(2)), CLng(values(0)), CLng(values(1)))
        If clickAction IsNot Nothing Then
            ClickMouse(KeyModifier.None, clickAction)
        End If
    Else
        MsgBox("No click location named """ & touchLocationName & """")
    End If
End Sub

Public Sub ListTouchLocations()
    Dim processName As String
    processName = GetForegroundProcessName()

    Dim touches As Object
    touches = GetTouches(processName)

    If touches.Count() = 0 Then
        MsgBox("There are no touch locations for process """ & processName & """.")
        Exit Sub
    End If

    Dim touchLocations() As String
    touchLocations = touches.Keys()
    MsgBox(ArrayToString(touchLocations, 0, -1, vbcrlf))

End Sub

Public Sub AddTouchLocation(touchLocationName As String, Optional relativeTo As WindowCorner = WindowCorner.NW)

    Dim processName As String
    processName = GetForegroundProcessName()

    Dim touches As Object
    touches = GetTouches(processName)

    If touches.ContainsKey(touchLocationName) Then
        If MsgBox("Touch location """ & touchLocationName & """ already exists for process """ & processName & """.  Replace?", vbYesNo) = vbNo Then
            Exit Sub
        End If
        touches.Remove(touchLocationName)
    End If

    Dim relativePoint As POINTAPI
    Dim relativeX As Long
    Dim relativeY As Long

    ' By ref values are not being returned here for some stupid reason.
    relativePoint = GetMousePositionRelativeToActiveWindow(relativeTo, relativeX, relativeY)

    Dim values(2) As String
    values(0) = CStr(relativePoint.X)
    values(1) = CStr(relativePoint.Y)
    values(2) = CStr(relativeTo)
    touches.Add(touchLocationName, values)

    ' msgbox(values(0) & " " & values(1) & " " & values(2))
    SaveTouchLocations(processName)

    TTSPlayString("Added " & touchLocationName)

End Sub

Public Sub RemoveTouchLocation(touchlocationName As String)
    Dim processName As String
    processName = GetForegroundProcessName()
    Dim touches As Object
    touches = GetTouches(processName)

    If touches.ContainsKey(touchlocationName) Then
        If MsgBox("Remove touch location """ & touchlocationName & """ for process """ & processName & """?", vbYesNo) = vbNo Then
            Exit Sub
        End If
        touches.Remove(touchlocationName)
        SaveTouchLocations(processName)
    End If

End Sub