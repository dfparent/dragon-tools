'#Uses "utilities.bas"
'#Uses "mouse.bas"
'#Uses "imports.bas"
'Option Explicit On

Public enum SHOW_WINDOW_COMMAND
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


Public Sub CenterActiveWindow()
    Dim windowRect As RECT
    Dim handle As Long
	Dim bRet As Boolean
	
	' Get screen size

    handle = GetForegroundWindow()
	
	dim hMonitor as long
	hMonitor = MonitorFromWindow(handle, MONITOR_DEFAULTTONEAREST)
	
	dim info as MONITORINFO
	info.size = LenB(info)
	GetMonitorInfo(hMonitor, info)

	dim screenWidth as integer
	dim screenHeight as integer
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

	dim width as integer
	dim height as integer
	width = windowRect.right - windowRect.left
	height = windowRect.Bottom - windowRect.Top
	
	'msgbox CStr(width) & "x" & Cstr(height)
	
	' Need to add the monitor location in case it is not the primary monitor
	MoveWindow(handle, info.monitorRect.left + (screenWidth / 2) - (width / 2), info.monitorRect.top + (screenHeight / 2) - (height / 2), width, height, true)
	
	

End Sub

public sub ResizeActiveWindow(makeLarger as boolean, widthresizeby as integer, heightresizeby as integer)
    Dim windowRect As RECT
    Dim handle As Long
	
	' Get window size
    handle = GetForegroundWindow()
	bRet = GetWindowRect(handle, windowRect)

	dim x, y, w, h as integer
	x = windowRect.Left
	y = windowRect.Top
	w = windowRect.Right - windowRect.Left
	h = windowRect.Bottom - windowRect.Top
	
	' Change size
	if makeLarger then
		MoveWindow(handle, x, y, w + widthresizeby, h + heightresizeby, true)
	else
		MoveWindow(handle, x, y, w - widthresizeby, h - heightresizeby, true)	
	end if
	
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
end Sub

Public Function GetActiveWindowRect(ByRef windowRect As RECT)
    Dim handle As Long

    ' Get window size
    handle = GetForegroundWindow()
    bRet = GetWindowRect(handle, windowRect)


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
    Dim x As Long
    Dim y As Long

    ' X, y are out parameters
    GetMousePositionRelativeToScreen(WindowCorner.NW, x, y)

    Dim windowRect As RECT
    GetActiveWindowRect(windowRect)

    MoveActiveWindow(x, y, windowRect.Right - windowRect.Left, windowRect.Bottom - windowRect.Top)

End Sub

Public Sub SetWindowState(command As SHOW_WINDOW_COMMAND, Optional handle As Long = -1)
    If handle = -1 Then
        handle = GetForegroundWindow()
    End If
    ShowWindow(handle, command)
End Sub

Public Sub ActivateWindow(handle As Long)
    SetForegroundWindow(handle)
End Sub

Public Function GetWindowTitleText(handle As Long) As String
    Dim length As Integer
    If handle = 0 Then
        GetWindowTitleText = ""
        Exit Function
    End If
    length = GetWindowTextLength(handle)
    If length = 0 Then
        GetWindowTitleText = ""
        Exit Function
    End If

    Dim outText As String * 255
    If length > 254 Then
        length = 254
    End If

    GetWindowText(handle, outText, length + 1)
    GetWindowTitleText = outText
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

    If handle = -1 Then
        handle = GetForegroundWindow()
    End If

    Dim text As String
    text = LCase(GetWindowTitleText(handle))

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

    Do Until DateDiff("s", startTime, Now) > maxWaitSeconds
        If CheckWindowText(windowTitleText, handle) Then
            WaitForWindow = True
            Exit Function
        End If
    Loop

    WaitForWindow = False
    Exit Function
End Function