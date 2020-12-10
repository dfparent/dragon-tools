'#Uses "paths.bas"
'#Uses "mouse.bas"
'#Uses "keyboard.bas"
'#Uses "keyboardconstants.bas"
'#Uses "window.bas"
'#Uses "cache.bas"
'#Uses "imports.bas"
'#Language "WWB.NET"
'Option Explicit On


Public Enum ClientSection
    Top
    Left
    Right
    Bottom
    Center
End Enum

Private ID_VALUE_KEY = "grid.bas-shell-id"

Private Sub doShell(commandLine)
    Dim id
    If CacheValueExists(ID_VALUE_KEY) Then
        id = GetCacheValue(ID_VALUE_KEY)(0)
        On Error Resume Next
        Err.clear
        'msgbox(id)
        AppActivate(CLng(id))
        Wait(0.1)
        On Error GoTo 0
        ' Close currently open grid
        If CheckWindowText("Mouse Grid") Then
            ' Close
            SendKeys("x")
        End If
        RemoveCacheValue(ID_VALUE_KEY)
    End If

    id = Shell(commandLine)
    'Clipboard(commandLine)
    'msgbox commandLine
    Wait(0.1)

    AddCacheValue(ID_VALUE_KEY, {CStr(id)})
End Sub

Public Sub ShowGridByString(commandLineSwitches As String)
    doShell(getFile("mouse grid") & " " & commandLineSwitches)
End Sub

Public Sub ShowGridByValues(x As Integer, y As Integer, w As Integer, h As Integer, rowh As Integer, colw As Integer)
    doShell(getFile("mouse grid") & " --location-x " & CStr(x) & " --location-y " & CStr(y) & " --width " & CStr(w) & " --height " & CStr(h) & " --row-height " & CStr(rowh) & " --column-width " & CStr(colw))
End Sub

' Input values are relative to client window
Public Sub ShowClientGridByValues(clientx As Integer, clienty As Integer, w As Integer, h As Integer, rowh As Integer, colw As Integer)
    Dim handle As Long
    handle = GetForegroundWindow()

    doShell(getFile("mouse grid") & " --size-to-client " & CStr(handle) & " --location-x " & CStr(clientx) & " --location-y " & CStr(clienty) & " --width " & CStr(w) & " --height " & CStr(h) & " --row-height " & CStr(rowh) & " --column-width " & CStr(colw))
End Sub

' Input values are relative to client window
Public Sub ShowClientGridByString(commandLineSwitches as string)
    Dim handle As Long
    handle = GetForegroundWindow()

    doShell(getFile("mouse grid") & " --size-to-client " & CStr(handle) & " " & commandLineSwitches)
End Sub

Public Sub ShowClientGrid(optional extraArgs as string = "")
    Dim handle As Long
    handle = GetForegroundWindow()
	dim command as String
    command = getFile("mouse grid") & " --size-to-client " & CStr(handle)
    If extraArgs <> "" then
		command = command & " " & extraArgs
	end if
	'msgbox command
    'Clipboard command
	doShell(command)
End Sub

Public Sub ShowPartialClientGrid(section As ClientSection)

    Dim command As String
    command = getFile("mouse grid") & " --dock-to-client " & section.ToString().ToLower()

    'msgbox command
    'Clipboard command
    doShell(command)

End Sub

Public Sub SetGrid()
	dim command as String
    command = getFile("mouse grid")
    'msgbox command
    'Clipboard command
    doShell(command)

    ' For some reason the window in this mode is shown minimized
    'HeardWord("restore", "window")
End Sub

Public Sub ShowGridHelp()
    doShell(getFile("mouse grid"))
End Sub

Public Sub ToggleSticky()
	SendKeys "y"
end sub

public enum GridAction
	Click
	DoubleClick
	RightClick
	Move
	DragFrom
	DropTo
end enum

public Sub ClickGrid(action as GridAction, rowNum as string, colNum as string, optional pressShift as boolean = false, optional pressControl as boolean = false)
	if not IsNumeric(rowNum) then	
		msgbox "Row number is not a number: " & rowNum
		exit sub
	end if
	
	SendKeys "r" & rowNum

    SendKeys "{Enter}"
    Wait 0.1

    If Not IsNumeric(colNum) Then
        msgbox "Column number is not a number: " & colNum
            Exit Sub
    End If

    SendKeys "c" & colNum

    SendKeys "{Enter}"
    Wait 0.1

    ' Press modifiers?
    If pressShift then
		KeyDown(VK_SHIFT)
	end if
	
	if pressControl then
		KeyDown(VK_CONTROL)
	end If

    Select case action
	case GridAction.Click
            SendKeys "s"

    Case GridAction.DoubleClick
		SendKeys "d"
		
	case GridAction.RightClick
		SendKeys "t"
		
	case GridAction.Move
            SendKeys "m"

    Case GridAction.DragFrom	
		SendKeys "f"
		
	case GridAction.DropTo
		SendKeys "g"
		
	end select

	Wait 0.1
	
	' Release modifiers?
	if pressShift then
		KeyUp(VK_SHIFT)
	end if
	
	if pressControl then
		KeyUp(VK_CONTROL)
	end if

	Wait 0.1
end Sub

