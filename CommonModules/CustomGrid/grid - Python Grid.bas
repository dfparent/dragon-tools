'#Uses "paths.bas"
'#Uses "mouse.bas"
'#Uses "keyboard.bas"
'#Uses "keyboardconstants.bas"
'#Language "WWB.NET"
'Option Explicit On

'Public structure RECT
'    public Left As Long
'    public Top As Long
'    public Right As Long
'    public Bottom As Long
'End structure

'Public Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, ByRef lpRect As RECT) As Boolean

Private Sub doShell(commandLine)
    Dim id
    id = Shell(commandLine)
	'Clipboard(commandLine)
	'msgbox commandLine
    Wait 0.1
    AppActivate id
End Sub

Public Sub ShowGridByString(commandLineSwitches As String)
    doShell(getFile("custom grid") & " " & commandLineSwitches)
End Sub

Public Sub ShowGridByValues(x As Integer, y As Integer, w As Integer, h As Integer, rowh As Integer, colw As Integer)
    doShell(getFile("custom grid") & " --locationx " & CStr(x) & " --locationy " & CStr(y) & " --width " & CStr(w) & " --height " & CStr(h) & " --rowheight " & CStr(rowh) & " --columnwidth " & CStr(colw))
End Sub

' Input values are relative to client window
Public Sub ShowClientGridByValues(clientx As Integer, clienty As Integer, w As Integer, h As Integer, rowh As Integer, colw As Integer)
    Dim handle As Long
    handle = GetForegroundWindow()

    doShell(getFile("custom grid") & " --sizeToClient " & CStr(handle) & " --locationx " & CStr(clientx) & " --locationy " & CStr(clienty) & " --width " & CStr(w) & " --height " & CStr(h) & " --rowheight " & CStr(rowh) & " --columnwidth " & CStr(colw))
End Sub

' Input values are relative to client window
Public Sub ShowClientGridByString(commandLineSwitches as string)
    Dim handle As Long
    handle = GetForegroundWindow()

    doShell(getFile("custom grid") & " --sizeToClient " & CStr(handle) & " " & commandLineSwitches)
End Sub

Public Sub ShowClientGrid(optional extraArgs as string = "")
    Dim handle As Long
    handle = GetForegroundWindow()
	dim command as string
	command = getFile("custom grid") & " --sizeToClient " & CStr(handle)
	if extraArgs <> "" then
		command = command & " " & extraArgs
	end if
	'msgbox command
    'Clipboard command
	doShell(command)
End Sub

Public Sub SetGrid()
	dim command as string
	command = getFile("custom grid") & " -p --columnwidth 20 --numcolumns 20 --rowheight 20 --numrows 20"
	'msgbox command
	'Clipboard command
    doShell(command)

    ' For some reason the window in this mode is shown minimized
    HeardWord("restore", "window")
End Sub

Public Sub ShowGridHelp()
    doShell(getFile("custom grid") & " --help")
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
		
	if not IsNumeric(colNum) then	
			msgbox "Column number is not a number: " & colNum
			exit sub
	end if
	
	SendKeys "c" & colNum
		
	' Press modifiers?
	if pressShift then
		KeyDown(VK_SHIFT)
	end if
	
	if pressControl then
		KeyDown(VK_CONTROL)
	end if

	SendKeys "{Enter}"
	Wait 0.1
	
	select case action
	case GridAction.Click
		SendKeys "s"
		
	case GridAction.DoubleClick
		SendKeys "d"
		
	case GridAction.RightClick
		SendKeys "t"
		
	case GridAction.Move
		SendKeys "x"
		
	case GridAction.DragFrom	
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
end sub

'public Sub ClickGridOff(action as GridAction, rowNum1 as string, optional rowNum2 as string = "", optional colNum1 as string = "", optional colNum2 as string = "", optional pressShift as boolean = false, optional pressControl as boolean = false)
'	if not IsNumeric(rowNum1) then	
'		msgbox "Row number 1 is not a number: " & rowNum1
'		exit sub
'	end if
'	
'	SendKeys "r" & rowNum1
'	
'	if not rowNum2 = "" then
'		if not IsNumeric(rowNum2) then	
'			msgbox "Row number 2 is not a number: " & rowNum2
'			exit sub
'		end if
'		SendKeys rowNum2
'	end if
'
'	if colNum1 = "" then	
'		msgbox "You must supply a column number."
'		exit sub
'	end if
'	
'	if not IsNumeric(colNum1) then	
'			msgbox "Column number 1 is not a number: " & colNum1
'			exit sub
'	end if
'	
'	SendKeys "c" & colNum1
'	
'	if not colNum2 = "" then
'		if not IsNumeric(colNum2) then	
'			msgbox "Column number 2 is not a number: " & colNum2
'			exit sub
'		end if
'		SendKeys colNum2
'	end if
'	
'	' Press modifiers?
'	if pressShift then
'		KeyDown(VK_SHIFT)
'	end if
'	
'	if pressControl then
'		KeyDown(VK_CONTROL)
'	end if
'
'	SendKeys "{Enter}"
'	Wait 0.1
'	
'	select case action
'	case GridAction.Click
'		SendKeys "s"
'		
'	case GridAction.DoubleClick
'		SendKeys "d"
'		
'	case GridAction.RightClick
'		SendKeys "t"
'		
'	case GridAction.Move
'		SendKeys "x"
'		
'	case GridAction.DragFrom	
'		SendKeys "f"
'		
'	case GridAction.DropTo
'		SendKeys "g"
'		
'	end select
'
'	Wait 0.1
'	
'	' Release modifiers?
'	if pressShift then
'		KeyUp(VK_SHIFT)
'	end if
'	
'	if pressControl then
'		KeyUp(VK_CONTROL)
'	end if
'
'	Wait 0.1
'end sub
