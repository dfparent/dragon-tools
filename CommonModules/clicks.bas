'#Uses "utilities.bas"
'#Uses "cache.bas"
'#Uses "mouse.bas"
'#Language "WWB.NET"
Option Explicit On

' Clicks the mouse at the given location
Public Sub ClickMouseLocation(clickLocationName As String)
    Dim clicks As Object
    clicks = GetClicks(GetForegroundProcessName())


    If clicks.ContainsKey(clickLocationName) Then
        Dim values() As String
        values = clicks(clickLocationName)
        PositionMouseRelativeToActiveWindow(WindowCorner.NW, CLng(values(0)), CLng(values(1)))
        ClickMouse()
    Else
        MsgBox("No click location named """ & clickLocationName & """")
    End If
End Sub

Public Sub AddMouseLocation(clickLocationName As String)
    Dim relativeX As Long, relativeY As Long
    GetMousePositionRelativeToActiveWindow(WindowCorner.NW, relativeX, relativeY)

    AddClick(GetForegroundProcessName(), clickLocationName, relativeX, relativeY)
End Sub

