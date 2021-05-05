'#Uses "utilities.bas"
'#Uses "keyboard.bas"
'#Uses "imports.bas"
'#Uses "cache.bas"
'Option Explicit On


' For Mouse wheel
Declare Function mouse_event Lib "user32.dll" (ByVal dwFlags As Long, ByVal dx As _
                                                Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long) As Long
Const WHEEL_DELTA = 120

const MOUSEEVENTF_ABSOLUTE = &H8000
const MOUSEEVENTF_LEFTDOWN = &H2
const MOUSEEVENTF_LEFTUP = &H4
const MOUSEEVENTF_MIDDLEDOWN = &H20
const MOUSEEVENTF_MIDDLEUP = &H40
const MOUSEEVENTF_MOVE = &H1
const MOUSEEVENTF_RIGHTDOWN = &H8
const MOUSEEVENTF_RIGHTUP = &H10
const MOUSEEVENTF_XDOWN = &H80
const MOUSEEVENTF_XUP = &H100
const MOUSEEVENTF_WHEEL = &H800
const MOUSEEVENTF_HWHEEL = &H1000
'Courtesy of Mike Jerry

Public Const WHEEL_CTS_ENABLED_KEY = "WheelClickToScrollEnabled"
Public Const WHEEL_CTS_HOME_LOCATION_X_KEY = "WheelClickToScrollHomeLocationX"
Public Const WHEEL_CTS_HOME_LOCATION_Y_KEY = "WheelClickToScrollHomeLocationY"
Public Const WHEEL_CTS_SCROLL_DELTA = 10

Public Enum WheelScrollDirection
    N
    NE
    E
    SE
    S
    SW
    W
    NW
End Enum

Public enum MouseButton
	Left
	Right
	Middle
end enum

public enum MouseButtonPosition
	Down
	Up
end enum

public enum KeyModifier
	none
	shift
	control
	alt
	win
end enum

Public Enum WindowCorner
    NW = 0
    NE = 1
    SW = 2
    SE = 3
End Enum

public enum WindowRegion
	Center
	LeftSide
	FarLeftSide
	RightSide
	FarRightSide
	TopHalf
	BottomHalf
end enum

Public Sub CenterMouseOnActiveWindow()
	PositionMouseOnActiveWindow(WindowRegion.Center)
End Sub

Public Sub PositionMouseOnActiveWindow(region as WindowRegion)
    Dim r As RECT
    Dim handle As Long
    Dim bRet As Boolean

    handle = GetForegroundWindow()
    bRet = GetWindowRect(handle, r)
    If Not bRet Then
        Beep
		'msgbox("Can't get window rectangle.")
        Exit Sub
    End If

	select case region
	case Center
		SetMousePosition(1, (r.Right - r.Left) / 2, (r.Bottom - r.Top) / 2)
		
	case LeftSide
		SetMousePosition(1, (r.Right - r.Left) / 4, (r.Bottom - r.Top) / 2)
		
	case FarLeftSide
		SetMousePosition(1, (r.Right - r.Left) / 8, (r.Bottom - r.Top) / 2)

	case RightSide
		SetMousePosition(1, (r.Right - r.Left) * 3 / 4, (r.Bottom - r.Top) / 2)
		
	case FarRightSide
		SetMousePosition(1, (r.Right - r.Left) * 7 / 8, (r.Bottom - r.Top) / 2)

	case TopHalf
		SetMousePosition(1, (r.Right - r.Left) / 2, (r.Bottom - r.Top) / 4)
		
	case BottomHalf
		SetMousePosition(1, (r.Right - r.Left) / 2, (r.Bottom - r.Top) * 3 / 4)
		
	end select

End Sub

Public Sub PositionMouseRelativeToActiveWindow(relativeTo As WindowCorner, relativeX As Integer, relativeY As Integer)
    Dim handle As Long
    Dim bRet As Boolean
    Dim rectangle As RECT

    ' Get window rectangle in system coordinates
    handle = GetForegroundWindow()
    bRet = GetWindowRect(handle, rectangle)
    If Not bRet Then
        Beep
		'msgbox("Can't get window rectangle.")
        Exit Sub
    End If

    ' Need to get coordinates relative to client
    Select Case relativeTo
        Case WindowCorner.NW
            SetMousePosition(1, relativeX, relativeY)

        Case WindowCorner.NE
            SetMousePosition(1, rectangle.Right - rectangle.Left - relativeX, relativeY)

        Case WindowCorner.SW
            SetMousePosition(1, relativeX, rectangle.Bottom - rectangle.Top - relativeY)

        Case WindowCorner.SE
            SetMousePosition(1, rectangle.Right - rectangle.Left - relativeX, rectangle.Bottom - rectangle.Top - relativeY)
    End Select

End Sub

Public Function GetMousePositionRelativeToActiveWindow(relativeTo As WindowCorner, ByRef relativeX As Long, ByRef relativeY As Long) As POINTAPI
    Dim handle As Long
    Dim bRet As Boolean
    Dim rectangle As RECT

    handle = GetForegroundWindow()
    bRet = GetWindowRect(handle, rectangle)
    If Not bRet Then
        Beep
        'msgbox("Can't get window rectangle.")
        Exit Function
    End If

    ' Get mouse location in system coordinates
    Dim thePoint As POINTAPI
    GetCursorPos(thePoint)

    Select Case relativeTo
        Case WindowCorner.NW
            relativeX = thePoint.x - rectangle.Left
            relativeY = thePoint.y - rectangle.Top

        Case WindowCorner.NE
            relativeX = rectangle.Right - thePoint.x
            relativeY = thePoint.y - rectangle.Top

        Case WindowCorner.SW
            relativeX = thePoint.x - rectangle.Left
            relativeY = rectangle.Bottom - thePoint.y

        Case WindowCorner.SE
            relativeX = rectangle.Right - thePoint.x
            relativeY = rectangle.Bottom - thePoint.y

    End Select

    'msgbox("RelativeX: " & CStr(relativeX) & vbcrlf & "RelativeY: " & CStr(relativeY))
    Dim ret As POINTAPI
    ret.X = relativeX
    ret.Y = relativeY

    GetMousePositionRelativeToActiveWindow = ret

End Function

Public Function GetMousePositionRelativeToScreen(relativeTo As WindowCorner, byref relativeX as long, byref relativeY as long)
    Dim handle As Long
    Dim bRet As Boolean
    Dim rectangle As RECT

    ' Get mouse location in system coordinates
    Dim thePoint As POINTAPI
    GetCursorPos(thePoint)

	' Get Screen dimensions
	'TBD
	dim screenWidth as long
	dim screenHeight as long
	
    Select Case relativeTo
        Case WindowCorner.NW
            relativeX = thePoint.x
            relativeY = thePoint.y

        Case WindowCorner.NE
            relativeX = screenWidth - thePoint.x
            relativeY = thePoint.y

        Case WindowCorner.SW
            relativeX = thePoint.x
            relativeY = screenHeight - thePoint.y

        Case WindowCorner.SE
            relativeX = screenWidth - thePoint.x
            relativeY = screenHeight - thePoint.y

    End Select
    
End Function

' Use slowScroll in applications that don't respond well to fast scrolling
Public Sub WheelMouse(Direction As String, Optional numMoves As Integer = 5, Optional slowScoll As Boolean = False)

    Dim loopCount As Integer
    Dim scrollSize As Integer

    If slowScoll Then
        ' Each mouse event is a smaller scroll, and loop through and do multiple mouse events
        ' For stupid apps like Adobe
        loopCount = numMoves
        scrollSize = WHEEL_DELTA
    Else
        ' Do it all in a single mouse event
        ' For  most modern apps
        loopCount = 1
        scrollSize = WHEEL_DELTA * numMoves
    End If

    Dim i As Integer
    For i = 1 To loopCount
        Select Case Direction
            Case "Up"
                mouse_event(MOUSEEVENTF_WHEEL, 0, 0, scrollSize, 0)
            Case "Down"
                mouse_event(MOUSEEVENTF_WHEEL, 0, 0, -scrollSize, 0)
            Case "Right"
                mouse_event(MOUSEEVENTF_HWHEEL, 0, 0, scrollSize, 0)
            Case "Left"
                mouse_event(MOUSEEVENTF_HWHEEL, 0, 0, -scrollSize, 0)
        End Select
        Wait 0.1
    Next


End Sub

' Enables the wheel click / move the mouse to scroll feature
Public Sub WheelClickToScroll(enable As Boolean, Optional scrollDirection As WheelScrollDirection = WheelScrollDirection.S, Optional scrollSpeedFactor As Integer = 1)
    Dim x As Long
    Dim y As Long

    If enable Then
        If Not CacheValueExists(WHEEL_CTS_ENABLED_KEY) Then
            ' Turn it on
            UpdateCacheValueSingle(WHEEL_CTS_ENABLED_KEY, "true")
            GetMousePositionRelativeToScreen(WindowCorner.NE, x, y)
            UpdateCacheValueSingle(WHEEL_CTS_HOME_LOCATION_X_KEY, CStr(x))
            UpdateCacheValueSingle(WHEEL_CTS_HOME_LOCATION_Y_KEY, CStr(y))

            ClickMouse(KeyModifier.none, MouseButton.Middle)
        Else
            x = GetCacheValueSingle(WHEEL_CTS_HOME_LOCATION_X_KEY)
            y = GetCacheValueSingle(WHEEL_CTS_HOME_LOCATION_Y_KEY)
        End If
    Else
        ' Disable
        If CacheValueExists(WHEEL_CTS_ENABLED_KEY) Then
            RemoveCacheValue(WHEEL_CTS_ENABLED_KEY)
            RemoveCacheValue(WHEEL_CTS_HOME_LOCATION_X_KEY)
            RemoveCacheValue(WHEEL_CTS_HOME_LOCATION_Y_KEY)
        End If
        Exit Sub
    End If

    ' Calculate mouse position
    Dim newX As Integer
    Dim newY As Integer
    Dim scrollDelta As Integer
    scrollDelta = WHEEL_CTS_SCROLL_DELTA * scrollSpeedFactor

    Select Case scrollDirection
        Case WheelScrollDirection.N
            newY = y - scrollDelta
        Case WheelScrollDirection.NE

        Case WheelScrollDirection.E
            newX = x + scrollDelta
        Case WheelScrollDirection.SE
        Case WheelScrollDirection.S
            newY = y + scrollDelta
        Case WheelScrollDirection.SW
        Case WheelScrollDirection.W
            newX = x - scrollDelta
        Case WheelScrollDirection.NW
    End Select

    SetMousePosition(0, newX, newY)

End Sub


Public Sub DoMouseEvent(button As MouseButton, position As MouseButtonPosition)

    If button = MouseButton.Left And position = MouseButtonPosition.Down Then
        mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    ElseIf button = MouseButton.Left And position = MouseButtonPosition.Up Then
        mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
    ElseIf button = MouseButton.Middle And position = MouseButtonPosition.Down Then
        mouse_event(MOUSEEVENTF_MIDDLEDOWN, 0, 0, 0, 0)
    ElseIf button = MouseButton.Middle And position = MouseButtonPosition.Up Then
        mouse_event(MOUSEEVENTF_MIDDLEUP, 0, 0, 0, 0)
    ElseIf button = MouseButton.Right And position = MouseButtonPosition.Down Then
        mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0)
    ElseIf button = MouseButton.Right And position = MouseButtonPosition.Up Then
        mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0)
    End If

End Sub

Public sub PressMouse(position as MouseButtonPosition, optional button as MouseButton = MouseButton.Left)
	
	select case position
		case MouseButtonPosition.Up
			mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
		case MouseButtonPosition.Down
			mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
	end select
end Sub

Public Sub ClickMouse(Optional key As KeyModifier = KeyModifier.none, Optional button As MouseButton = MouseButton.Left, Optional clickCount As Integer = 1)
    Dim modifier As String
    modifier = ""

    Select Case key
        Case KeyModifier.shift
            modifier = "+"

        Case KeyModifier.control
            modifier = "^"

        Case KeyModifier.alt
            modifier = "%"

    End Select

    Dim buttonStr As String
    Select Case button
        Case MouseButton.Left
            buttonStr = "ClickLeft"

        Case MouseButton.Middle
            buttonStr = "ClickMiddle"

        Case MouseButton.Right
            buttonStr = "ClickRight"
    End Select

    ' Use current mouse position
    Dim thePoint As POINTAPI
    GetCursorPos(thePoint)

    Dim command As String
    command = modifier & "{" & buttonStr & " " & CStr(thePoint.x) & "," & CStr(thePoint.y) & "}"

    SendKeys(command)

End Sub

'Public Sub ClickMouse(Optional key As KeyModifier = KeyModifier.none, Optional button As MouseButton = MouseButton.Left, Optional clickCount As Integer = 1)
'    Select Case key
'        Case KeyModifier.shift
'            KeyDown(VK_SHIFT)

'        Case KeyModifier.control
'            KeyDown(VK_CONTROL)

'        Case KeyModifier.alt
'            KeyDown(VK_MENU)

'    End Select

'    Wait(0.1)

'    DoMouseEvent(button, MouseButtonPosition.Down)
'    DoMouseEvent(button, MouseButtonPosition.Up)

'    Wait(0.1)

'    Select Case key
'        Case KeyModifier.shift
'            KeyUp(VK_SHIFT)

'        Case KeyModifier.control
'            KeyUp(VK_CONTROL)

'        Case KeyModifier.alt
'            KeyUp(VK_MENU)

'    End Select

'End Sub

' Converts mouse coordinates from client coordinates relative to active window to screen coordinates
Public Sub ConvertMousePositionToScreenCoordinates(relativeX as long, relativeY as long, byref screenX as long, byref screenY as long)
	dim handle as Long
    handle = GetForegroundWindow()
    Dim screenPoint As POINTAPI
    screenPoint.x = relativeX
    screenPoint.y = relativeY
    ClientToScreen(handle, screenPoint)
    screenX = screenPoint.x
    screenY = screenPoint.y

End Sub


