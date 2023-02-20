'#Language "WWB.NET"
'#Uses "imports.bas"
'#Uses "utilities.bas"
'#Uses "keyboard.bas"
'#Uses "KeyboardConstants.bas"
'#Uses "cache.bas"
Option Explicit On
Option Strict On


' For Mouse wheel
Declare Sub mouse_event Lib "user32.dll" (dwFlags As Integer, dx As Integer, dy As Integer, dwData As Integer, ByRef dwExtraInfo As Long)
Declare Function GetAsyncKeyState Lib "user32" (ByVal vkey As Integer) As Short

Const WHEEL_DELTA = 120

Const MOUSEEVENTF_ABSOLUTE = &H8000
Const MOUSEEVENTF_LEFTDOWN = &H2
Const MOUSEEVENTF_LEFTUP = &H4
Const MOUSEEVENTF_MIDDLEDOWN = &H20
Const MOUSEEVENTF_MIDDLEUP = &H40
Const MOUSEEVENTF_MOVE = &H1
Const MOUSEEVENTF_RIGHTDOWN = &H8
Const MOUSEEVENTF_RIGHTUP = &H10
Const MOUSEEVENTF_XDOWN = &H80
Const MOUSEEVENTF_XUP = &H100
Const MOUSEEVENTF_WHEEL = &H800
Const MOUSEEVENTF_HWHEEL = &H1000
'Courtesy of Mike Jerry

Public Const WHEEL_CTS_DIRECTION_KEY = "WheelClickToScrollDirection"
Public Const WHEEL_CTS_HOME_LOCATION_X_KEY = "WheelClickToScrollHomeLocationX"
Public Const WHEEL_CTS_HOME_LOCATION_Y_KEY = "WheelClickToScrollHomeLocationY"
Public Const WHEEL_CTS_SCROLL_DELTA = 40

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

Public Enum MouseButton
    Left
    Right
    Middle
End Enum

Public enum MouseButtonPosition
	Down
	Up
end Enum

Public Enum WindowCorner
    NW = 0
    NE = 1
    SW = 2
    SE = 3
End Enum

public Enum WindowRegion
	Center
	LeftSide
	FarLeftSide
	RightSide
	FarRightSide
	TopHalf
	BottomHalf
end Enum

Public Sub CenterMouseOnActiveWindow()
    PositionMouseOnActiveWindow(WindowRegion.Center)
End Sub

Public Sub PositionMouseOnActiveWindow(region As WindowRegion)
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

    Select Case region
        Case WindowRegion.Center
            SetMousePosition(1, (r.Right - r.Left) / 2, (r.Bottom - r.Top) / 2)

        Case WindowRegion.LeftSide
            SetMousePosition(1, (r.Right - r.Left) / 4, (r.Bottom - r.Top) / 2)

        Case WindowRegion.FarLeftSide
            SetMousePosition(1, (r.Right - r.Left) / 8, (r.Bottom - r.Top) / 2)

        Case WindowRegion.RightSide
            SetMousePosition(1, (r.Right - r.Left) * 3 / 4, (r.Bottom - r.Top) / 2)

        Case WindowRegion.FarRightSide
            SetMousePosition(1, (r.Right - r.Left) * 7 / 8, (r.Bottom - r.Top) / 2)

        Case WindowRegion.TopHalf
            SetMousePosition(1, (r.Right - r.Left) / 2, (r.Bottom - r.Top) / 4)

        Case WindowRegion.BottomHalf
            SetMousePosition(1, (r.Right - r.Left) / 2, (r.Bottom - r.Top) * 3 / 4)

    End Select

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

Public Sub PositionMouseRelativeToScreen(relativeTo As WindowCorner, relativeX As Integer, relativeY As Integer)
    Dim handle As Long
    Dim bRet As Boolean
    Dim rectangle As RECT

    ' Get Screen dimensions
    'TBD

    Select Case relativeTo
        Case WindowCorner.NW
            SetMousePosition(0, relativeX, relativeY)

        Case WindowCorner.NE
            SetMousePosition(0, rectangle.Right - rectangle.Left - relativeX, relativeY)

        Case WindowCorner.SW
            SetMousePosition(0, relativeX, rectangle.Bottom - rectangle.Top - relativeY)

        Case WindowCorner.SE
            SetMousePosition(0, rectangle.Right - rectangle.Left - relativeX, rectangle.Bottom - rectangle.Top - relativeY)
    End Select

End Sub

Public Function GetMousePositionRelativeToActiveWindow(relativeTo As WindowCorner, ByRef relativeX As Integer, ByRef relativeY As Integer) As POINTAPI
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

Public Function GetMousePositionRelativeToScreen(relativeTo As WindowCorner, ByRef relativeX As Integer, ByRef relativeY As Integer)
    Dim handle As Long
    Dim bRet As Boolean
    Dim rectangle As RECT

    ' Get mouse location in system coordinates
    Dim thePoint As POINTAPI
    GetCursorPos(thePoint)

    ' Get Screen dimensions
    'TBD
    Dim screenWidth As Integer
    Dim screenHeight As Integer

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
    'MsgBox(relativeX.ToString() & ", " & relativeY.ToString())
End Function

Public Function GetScreenMousePositionCode(Optional relativeTo As WindowCorner = WindowCorner.NW) As String
    Dim x As Integer
    Dim y As Integer
    GetMousePositionRelativeToScreen(relativeTo, x, y)

    Dim clip As String
    clip = "SetMousePosition 0," & CStr(x) & "," & CStr(y) & Chr(13) & Chr(10)
    clip = clip & "Wait(0.5)" & Chr(13) & Chr(10)
    clip = clip & "ButtonClick 1,1" & Chr(13) & Chr(10)

    GetScreenMousePositionCode = clip

End Function

Public Function GetClientMousePositionCode(Optional relativeTo As WindowCorner = WindowCorner.NW) As String
    Dim x As Integer
    Dim y As Integer
    GetMousePositionRelativeToActiveWindow(relativeTo, x, y)

    Dim clip As String
    clip = "SetMousePosition 1," & CStr(x) & "," & CStr(y) & Chr(13) & Chr(10)
    clip = clip & "Wait(0.5)" & Chr(13) & Chr(10)
    clip = clip & "ButtonClick 1,1" & Chr(13) & Chr(10)

    GetClientMousePositionCode = clip

End Function

' Use slowScroll in applications that don't respond well to fast scrolling
Public Sub WheelMouse(Direction As String, Optional numMoves As Integer = 5, Optional slowScoll As Boolean = False)

    Dim loopCount As Integer
    Dim scrollSize As UInteger

    If slowScoll Then
        ' Each mouse event is a smaller scroll, and loop through and do multiple mouse events
        ' For stupid apps like Adobe
        loopCount = numMoves
        scrollSize = WHEEL_DELTA
    Else
        ' Do it all in a single mouse event
        ' For  most modern apps
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
                mouse_event(MOUSEEVENTF_HWHEEL, 0, 0, 0 - scrollSize, 0)
        End Select
        Wait(0.1)
    Next


End Sub

Public Enum ScrollMode
    Enable
    Disable
    ChangeSpeed
End Enum

' Enables the wheel click / move the mouse to scroll feature
Public Sub WheelClickToScroll(mode As ScrollMode, Optional scrollDirection As WheelScrollDirection = WheelScrollDirection.S, Optional scrollSpeedDelta As Integer = 0)

    Dim scrollDelta As Integer

    If mode = ScrollMode.Enable Then
        'If Not scrolling Then
        ' Store scroll direction
        'UpdateCacheValueSingle(WHEEL_CTS_HOME_LOCATION_X_KEY, CStr(x))
        'UpdateCacheValueSingle(WHEEL_CTS_HOME_LOCATION_Y_KEY, CStr(y))
        UpdateCacheValueSingle(WHEEL_CTS_DIRECTION_KEY, CStr(scrollDirection))
        'MsgBox("Starting scrolling")
        'DoMouseEvent(MouseButtonPosition.Down, MouseButton.Middle)
        SendKeys("{esc}")
        ButtonClick(4, 1)
        Wait(0.1)
        scrollDelta = WHEEL_CTS_SCROLL_DELTA + scrollSpeedDelta
        'Else
        '   GetMousePositionRelativeToActiveWindow(WindowCorner.NW, x, y)
    ElseIf ScrollMode.ChangeSpeed Then
        If (Not CacheValueExists(WHEEL_CTS_DIRECTION_KEY)) Then
            Beep
            Exit Sub
        End If
        scrollDirection = GetCacheValueSingle(WHEEL_CTS_DIRECTION_KEY)
        If scrollSpeedDelta < 0 Then
            ' Slowing down, so move mouse in opposte direction
            If scrollDirection = WheelScrollDirection.N Then
                scrollDirection = WheelScrollDirection.S
            ElseIf scrollDirection = WheelScrollDirection.S Then
                scrollDirection = WheelScrollDirection.N
            ElseIf scrollDirection = WheelScrollDirection.E Then
                scrollDirection = WheelScrollDirection.W
            ElseIf scrollDirection = WheelScrollDirection.W Then
                scrollDirection = WheelScrollDirection.E
            End If
        End If

        scrollDelta = System.Math.Abs(scrollSpeedDelta)

        'msgbox CStr(x) & "," & CStr(y)

    ElseIf mode = ScrollMode.Disable Then
        ' Disable
        'If scrolling Then
        'RemoveCacheValue(WHEEL_CTS_HOME_LOCATION_X_KEY)
        'RemoveCacheValue(WHEEL_CTS_HOME_LOCATION_Y_KEY)
        'RemoveCacheValue(WHEEL_CTS_ENABLED_KEY)
        'MsgBox("Stopping scrolling")
        'DoMouseEvent(MouseButtonPosition.Up, MouseButton.Middle)
        SendKeys("{esc}")
        Exit Sub
        'End If
    End If

    ' Calculate mouse position
    Dim x As Integer
    Dim y As Integer
    Dim newX As Integer
    Dim newY As Integer

    GetMousePositionRelativeToActiveWindow(WindowCorner.NW, x, y)

    newX = x
    newY = y


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

    'msgbox CStr(newX) & "," & CStr(newY)
    SetMousePosition(1, newX, newY)
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

    Wait(0.1)

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

'   For Wait(0.1)

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
Public Sub ConvertMousePositionToScreenCoordinates(relativeX As Integer, relativeY As Integer, ByRef screenX As Integer, ByRef screenY As Integer)
    Dim handle As Long
    handle = GetForegroundWindow()
    Dim screenPoint As POINTAPI
    screenPoint.x = relativeX
    screenPoint.y = relativeY
    ClientToScreen(handle, screenPoint)
    screenX = screenPoint.x
    screenY = screenPoint.y

End Sub



'Private Structure MOUSEINPUT
'Public dx As Integer
'Public dy As Integer
'Public mouseData As Integer
'Public dwFlags As Integer
'Public time As Integer
'Public dwExtraInfo As System.IntPtr
'End Structure

'Private Structure KEYBDINPUT
'Public wVk As Short
'Public wScan As Short
'Public dwFlags As Integer
'Public time As Integer
'Public dwExtraInfo As System.IntPtr
'End Structure

'Private Structure HARDWAREINPUT
'Public uMsg As Integer
'Public wParamL As Short
'Public wParamH As Short
'End Structure

' For use with the INPUT struct, see SendInput for an example
'Public Enum Win32Constants
'INPUT_MOUSE = 0
'INPUT_KEYBOARD = 1
'INPUT_HARDWARE = 2
'End Enum

'Private Structure INPUT
'Public type As Integer
'Public xi(0 To 23) As Byte
'End Structure

'Declare Function SendInput Lib "user32.dll" (nInputs As Integer, pInputs() As INPUT, cbSize As Integer)
'Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal pDst As System.IntPtr, ByVal pSrc As System.IntPtr, ByVal ByteLen As UInteger)
'<SuppressMessage("Microsoft.Security", "CA2118:RevThis workiewSuppressUnmanagedCodeSecurityUsage")>
'<DllImport("Ntdll.dll", SetLastError:=True, ExactSpelling:=True, EntryPoint:="RtlMoveMemory", CharSet:=System.Runtime.InteropServices.CharSet.Auto)>
'Public Shared Sub CopyMemory(ByVal destData As HandleRef, ByVal srcData As HandleRef, ByVal size As Integer)

'End Sub

' This doesn't work
'Public Sub DoMouseInput(button As MouseButton, position As MouseButtonPosition)
'Dim theInput(1) As INPUT
'Dim mi As MOUSEINPUT

'mi.dwFlags = MOUSEEVENTF_LEFTDOWN Or MOUSEEVENTF_ABSOLUTE

'theInput(0).type = Win32Constants.INPUT_MOUSE
'CopyMemory(theInput(0).xi(0), mi, Len(mi))

'SendInput(1, theInput, System.Runtime.InteropServices.Marshal.SizeOf(GetType(INPUT)))
'End Sub

