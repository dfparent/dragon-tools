'#Uses "utilities.bas"
'#Uses "keyboardconstants.bas"

''''''''''''''''''''''''''''''''''''''
' Use these when SendKeys doesn't work
''''''''''''''''''''''''''''''''''''''

Private Declare Function keybd_event Lib "user32.dll" (ByVal vKey As Long, bScan As Long, ByVal Flag As Long, ByVal exInfo As Long) As Long
Private Const KEYEVENTF_EXTENDEDKEY As Integer = &H1
Private Const KEYEVENTF_KEYUP As Integer = &H2

Public Sub KeyDown(keyCode As Integer)
    keybd_event(keyCode, 0, KEYEVENTF_EXTENDEDKEY, 0)
	Wait(0.1)
End Sub

Public Sub KeyUp(keyCode As Integer)
    keybd_event(keyCode, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
	Wait(0.1)
End Sub

Public Sub KeyPress(keyCode As Integer, Optional count As Integer = 1)
    Dim i As Integer
    For i = 1 To count
        KeyDown(keyCode)
        KeyUp(keyCode)
    Next i
End Sub

Sub PressWindowsKey(appKey As String, count As String, Optional shift As Boolean = False)
    ' Windows key down
    KeyDown(VK_LWIN)

    ' Shift key down?
    If shift Then
        KeyDown(VK_SHIFT)
    End If

    'SendKeys "{" & appKey & " " & count & "}", True
    'Wait(0.5)
    RepeatKeystrokes(appKey, CInt(count), 0.2)

    ' Windows key up
    KeyUp(VK_LWIN)

    If shift Then
        KeyUp(VK_SHIFT)
    End If
End Sub

Sub PressModifiedKey(keyCode As Integer, Optional shift As Boolean = False, Optional control As Boolean = False, Optional alt As Boolean = False)
    ' Modifier keys down
    If shift Then
        KeyDown(VK_SHIFT)
    End If

    If control Then
        KeyDown(VK_CONTROL)
    End If

    If alt Then
        KeyDown(VK_ALT)
    End If

	
    ' Press Requested key down and up
    KeyDown(keyCode)
    KeyUp(keyCode)
	
    ' Release modifier keys
    If shift Then
        KeyUp(VK_SHIFT)
    End If

    If control Then
        KeyUp(VK_CONTROL)
    End If

    If alt Then
        KeyUp(VK_ALT)
    End If

End Sub

Public Sub ShowKeyTips()
    KeyPress(VK_ALT)
    Wait 0.1
End Sub

Public Sub ShowContextMenu()
    KeyPress(VK_CONTEXT_MENU)
    Wait 0.1
End Sub

public Sub ReleaseModifierKeys()
	KeyUp(VK_CONTEXT_MENU)
	KeyUp(VK_CONTROL)
	KeyUp(VK_LCONTROL)
	KeyUp(VK_RCONTROL)
	KeyUp(VK_SHIFT)
	KeyUp(VK_LSHIFT)
	KeyUp(VK_RSHIFT)
	KeyUp(VK_ALT)
	KeyUp(VK_LMENU)
	KeyUp(VK_RMENU)
	KeyUp(VK_LWIN)
End sub