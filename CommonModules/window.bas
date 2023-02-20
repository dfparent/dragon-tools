'#Language "WWB.NET"
'#Uses "imports.bas"
'#Uses "utilities.bas"
'#Uses "mouse.bas"
Option Explicit On

Public Enum SHOW_WINDOW_COMMAND As Integer
    Hide = 0
    Normal = 1
    ShowMinimized = 2
    Maximize = 3
    ShowMaximized = 3
    ShowNoActivate = 4
    Show = 5
    Minimize = 6
    ShowMinNoActive = 7
    ShowNA = 8
    Restore = 9
    ShowDefault = 10
    ForceMinimize = 11
End Enum

Public Enum WindowDock
    Center
    LeftSide
    RightSide
    TopHalf
    BottomHalf
End Enum

Private Sub PrintRect(theRect As RECT)
    Msgbox("Top: " & CStr(theRect.Top) & vbcrlf &
            "Left: " & CStr(theRect.Left) & vbcrlf &
            "Right: " & CStr(theRect.Right) & vbcrlf &
            "Bottom: " & CStr(theRect.Bottom) & vbcrlf)
End Sub

Public Sub CenterActiveWindow()
    Dim screenRect As RECT
    GetActiveScreenRect(screenRect)

    'PrintRect(screenRect)

    Dim windowRect As RECT
    GetActiveWindowRect(windowRect)

    'PrintRect(windowRect)

    Dim screenWidth As Integer
    Dim screenHeight As Integer
    Dim windowWidth As Integer
    Dim windowHeight As Integer

    screenWidth = screenRect.right - screenRect.left
    screenHeight = screenRect.bottom - screenRect.top
    windowWidth = windowRect.right - windowRect.left
    windowHeight = windowRect.bottom - windowRect.top

    MoveActiveWindow(screenRect.left + (screenWidth / 2) - (windowWidth / 2), screenRect.top + (screenHeight / 2) - (windowHeight / 2), windowWidth, windowHeight)

End Sub

Public Sub DockActiveWindow(docPosition As WindowDock)
    Dim windowRect As RECT
    Dim handle As Long
    Dim bRet As Boolean

    ' Get screen size

    handle = GetForegroundWindow()

    Dim hMonitor As System.IntPtr
    hMonitor = MonitorFromWindow(handle, MONITOR_DEFAULTTONEAREST)

    Dim info As MONITORINFO
    'info.size = LenB(info)
    info.size = Len(info)

    GetMonitorInfo(hMonitor, info)

    'PrintRect(info.monitorRect)

    Dim screenWidth As Integer
    Dim screenHeight As Integer
    screenWidth = info.monitorRect.right - info.monitorRect.left
    screenHeight = info.monitorRect.bottom - info.monitorRect.top

    'msgbox "Screen size: " & CStr(screenWidth) & "x" & Cstr(screenHeight)
    'msgbox "Screen location: " & CStr(info.monitorRect.left) & "x" & Cstr(info.monitorRect.top)

    ' Get window size
    bRet = GetWindowRect(handle, windowRect)
    If Not bRet Then
        Beep
        'msgbox("Can't get window rectangle.")
        Exit Sub
    End If

    Dim width As Integer
    Dim height As Integer
    width = windowRect.right - windowRect.left
    height = windowRect.Bottom - windowRect.Top

    'msgbox CStr(width) & "x" & Cstr(height)

    Select Case docPosition
        Case WindowDock.Center
            ' Need to add the monitor location in case it is not the primary monitor
            'Resize window
            width = screenWidth / 2
            height = screenHeight / 2
            MoveWindow(handle, info.monitorRect.left + (screenWidth / 2) - (width / 2), info.monitorRect.top + (screenHeight / 2) - (height / 2), width, height, True)

        Case WindowDock.LeftSide
            SendKeys("{WindowsHold}{Left}")

        Case WindowDock.RightSide
            SendKeys("{WindowsHold}{Right}")

        Case WindowDock.TopHalf
            MoveWindow(handle, info.monitorRect.left, info.monitorRect.top, screenWidth, screenHeight / 2, True)

        Case WindowDock.BottomHalf
            MoveWindow(handle, info.monitorRect.left, info.monitorRect.top + (screenHeight / 2), screenWidth, screenHeight / 2, True)
    End Select
End Sub

Public Sub ResizeActiveWindow(makeLarger As Boolean, widthresizeby As Integer, heightresizeby As Integer)
    Dim windowRect As RECT
    Dim handle As Long
    Dim bRet As Boolean

    ' Get window size
    handle = GetForegroundWindow()
    bRet = GetWindowRect(handle, windowRect)

    Dim x, y, w, h As Integer
    x = windowRect.Left
    y = windowRect.Top
    w = windowRect.Right - windowRect.Left
    h = windowRect.Bottom - windowRect.Top

    ' Change size
    If makeLarger Then
        MoveWindow(handle, x, y, w + widthresizeby, h + heightresizeby, True)
    Else
        MoveWindow(handle, x, y, w - widthresizeby, h - heightresizeby, True)
    End If

    'SendKeys "%{Space}"	
    'Wait 0.5
    'SendKeys "s"
    'Wait 0.5

    'if makeLarger then
    '	SendKeys "{Right}"
    '	RepeatKeyStrokes("{Right}", resizeby)
    '	RepeatKeyStrokes("{Down}", resizeby)
    'else
    '	SendKeys "{Right}"
    '	RepeatKeyStrokes("{Left}", resizeby)
    '	SendKeys "{Down}"
    '	RepeatKeyStrokes("{Up}", resizeby)
    'end if
    'Wait 0.3
    'SendKeys "{Enter}"
End Sub

Public Function GetActiveWindowRect(ByRef windowRect As RECT)
    Dim handle As System.IntPtr
    Dim bRet As Boolean

    ' Get window size
    handle = GetForegroundWindow()
    bRet = GetWindowRect(handle, windowRect)


End Function

' Gets the rectangle for the screen containing the active window
Public Function GetActiveScreenRect(ByRef screenRect As RECT)
    Dim handle As Long
    Dim bRet As Boolean

    handle = GetForegroundWindow()

    Dim hMonitor As Long
    hMonitor = MonitorFromWindow(handle, MONITOR_DEFAULTTONEAREST)

    Dim info As MONITORINFO
    'info.size = LenB(info)
    info.size = Len(info)
    GetMonitorInfo(hMonitor, info)

    'PrintRect(info.monitorRect)

    Dim screenWidth As Integer
    Dim screenHeight As Integer
    screenRect.Top = info.monitorRect.top
    screenRect.Left = info.monitorRect.left
    screenRect.Bottom = info.monitorRect.bottom
    screenRect.Right = info.monitorRect.right

End Function

Public Sub MoveActiveWindow(x As Integer, y As Integer, w As Integer, h As Integer)
    Dim handle As Long

    ' Get window size
    handle = GetForegroundWindow()

    ' Change size
    MoveWindow(handle, x, y, w, h, True)

End Sub

Public Sub MoveActiveWindowToMouse()
    ' Get mouse position
    Dim x As Integer
    Dim y As Integer

    ' X, y are out parameters
    GetMousePositionRelativeToScreen(WindowCorner.NW, x, y)
    'MsgBox(CStr(x) & ", " & CStr(y))

    Dim windowRect As RECT
    GetActiveWindowRect(windowRect)

    MoveActiveWindow(x, y, windowRect.Right - windowRect.Left, windowRect.Bottom - windowRect.Top)

End Sub

Public Sub SetWindowState(command As SHOW_WINDOW_COMMAND, Optional handle As Long = -1)
    Dim handlePtr As System.IntPtr
    If handle = -1 Then
        handlePtr = GetForegroundWindow()
    Else
        handlePtr = handle
    End If
    ShowWindow(handlePtr, command)
End Sub

Public Sub ActivateWindow(handle As Long)
    SetForegroundWindow(handle)
End Sub

Public Function GetWindowTitleText(Optional handle As Long = 0) As String
    Dim length As Integer
    Dim handlePtr As System.IntPtr
    If handle = 0 Then
        handlePtr = GetForegroundWindow()
    Else
        handlePtr = handle
    End If
    length = GetWindowTextLength(handlePtr)
    If length = 0 Then
        GetWindowTitleText = ""
        Exit Function
    End If

    ' Make a string of max length.  Tricky to do in this environment.
    Dim outTextBuilder As New System.Text.StringBuilder()
    outTextBuilder.Append(" "c, 255)
    Dim outText As String
    outText = outTextBuilder.ToString()

    If length > 254 Then
        length = 254
    End If

    GetWindowText(handlePtr, outText, length + 1)
    'msgbox(Trim(outText))
    GetWindowTitleText = Trim(outText)
End Function

' Returns true if the given window title contains the given text.
' If a window handle is not provided, the current foreground window is used
Public Function CheckWindowText(checkText, Optional handle As Long = -1) As Boolean
    Dim checkArray() As String
    If Not IsArray(checkText) Then
        Dim saveText As String
        saveText = CStr(checkText)
        ReDim checkArray(0)
        checkArray(0) = saveText
    Else
        checkArray = checkText
    End If

    Dim handlePtr As System.IntPtr
    If handle = -1 Then
        handlePtr = GetForegroundWindow()
    Else
        handlePtr = handle
    End If

    Dim text As String
    text = LCase(GetWindowTitleText(handlePtr))

    Dim i As Integer
    For i = 0 To UBound(checkArray)
        If InStr(text, LCase(checkArray(i))) > 0 Then
            CheckWindowText = True
            Exit Function
        End If
    Next i

    CheckWindowText = False
End Function

Public Function WaitForWindow(windowTitleText As String, Optional maxWaitSeconds As Long = 5, Optional handle As Long = -1) As Boolean
    Dim startTime As Date
    startTime = Now

    Wait(0.1)

    Do Until DateDiff("s", startTime, Now) > maxWaitSeconds
        If CheckWindowText(windowTitleText, handle) Then
            WaitForWindow = True
            Exit Function
        End If
    Loop

    WaitForWindow = False
    Exit Function
End Function

Public Function WaitForWindowTitleChange(Optional maxWaitSeconds As Long = 5, Optional handle As Long = -1) As Boolean
    Dim windowText As String
    If handle = -1 Then
        windowText = GetWindowTitleText()
    Else
        windowText = GetWindowTitleText(handle)
    End If

    Dim startTime As Date
    startTime = Now

    Do Until DateDiff("s", startTime, Now) > maxWaitSeconds
        ' Wait until the title CHANGES
        If Not CheckWindowText(windowText, handle) Then
            WaitForWindowTitleChange = True
            Exit Function
        End If
    Loop

    WaitForWindowTitleChange = False

End Function