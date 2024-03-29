
Public Const VK_ADD = &H6B
Public Const VK_ATTN = &HF6
Public Const VK_BACK = &H8 'back space
Public Const VK_CANCEL = &H3
Public Const VK_CAPITAL = &H14
Public Const VK_CLEAR = &HC
Public Const VK_CONTEXT_MENU = &H5D ' Right click menu; same as shfit+F10 in most apps
Public Const VK_CONTROL = &H11 'control
Public Const VK_CRSEL = &HF7
Public Const VK_DECIMAL = &H6E
Public Const VK_DELETE = &H2E 'delete
Public Const VK_DIVIDE = &H6F ' /
Public Const VK_DOWN = &H28 'arrow
Public Const VK_END = &H23 'end
Public Const VK_EREOF = &HF9
Public Const VK_ESCAPE = &H1B 'esc
Public Const VK_EXECUTE = &H2B
Public Const VK_EXSEL = &HF8
Public Const VK_F1 = &H70 'function keys
Public Const VK_F10 = &H79 'function keys
Public Const VK_F11 = &H7A 'function keys
Public Const VK_F12 = &H7B 'function keys
Public Const VK_F13 = &H7C 'function keys
Public Const VK_F14 = &H7D 'function keys
Public Const VK_F15 = &H7E 'function keys
Public Const VK_F16 = &H7F 'function keys
Public Const VK_F17 = &H80 'function keys
Public Const VK_F18 = &H81 'function keys
Public Const VK_F19 = &H82 'function keys
Public Const VK_F2 = &H71 'function keys
Public Const VK_F20 = &H83 'function keys
Public Const VK_F21 = &H84 'function keys
Public Const VK_F22 = &H85 'function keys
Public Const VK_F23 = &H86 'function keys
Public Const VK_F24 = &H87 'function keys
Public Const VK_F3 = &H72 'function keys
Public Const VK_F4 = &H73 'function keys
Public Const VK_F5 = &H74 'function keys
Public Const VK_F6 = &H75 'function keys
Public Const VK_F7 = &H76 'function keys
Public Const VK_F8 = &H77 'function keys
Public Const VK_F9 = &H78 'function keys
Public Const VK_HELP = &H2F
Public Const VK_HOME = &H24 'home
Public Const VK_INSERT = &H2D 'insert
Public Const VK_LBUTTON = &H1 'left mouse button
Public Const VK_LCONTROL = &HA2 'left control
Public Const VK_LEFT = &H25 'arrow
Public Const VK_LMENU = &HA4
Public Const VK_LSHIFT = &HA0 'left shift
Public Const VK_LWIN = &h5B
Public Const VK_MBUTTON = &H4 ' NOT contiguous with L RBUTTON
Public Const VK_MENU = &H12 ' Alternate key
Public Const VK_ALT = &H12
Public Const VK_MULTIPLY = &H6A ' *
Public Const VK_NEXT = &H22 ' Page Down
Public Const VK_NONAME = &HFC
Public Const VK_NUMLOCK = &H90 'numlock(use toggle)
Public Const VK_NUMPAD0 = &H60
Public Const VK_NUMPAD1 = &H61
Public Const VK_NUMPAD2 = &H62
Public Const VK_NUMPAD3 = &H63
Public Const VK_NUMPAD4 = &H64
Public Const VK_NUMPAD5 = &H65
Public Const VK_NUMPAD6 = &H66
Public Const VK_NUMPAD7 = &H67
Public Const VK_NUMPAD8 = &H68
Public Const VK_NUMPAD9 = &H69
Public Const VK_OEM_CLEAR = &HFE
Public Const VK_PA1 = &HFD
Public Const VK_PAUSE = &H13 'break/pause
Public Const VK_PLAY = &HFA
Public Const VK_PRINT = &H2A 'print screen
Public Const VK_PRIOR = &H21 ' Page Up
Public Const VK_PROCESSKEY = &HE5
Public Const VK_RBUTTON = &H2 'right mouse button
Public Const VK_RCONTROL = &HA3 'right control
Public Const VK_RETURN = &HD 'return/enter
Public Const VK_RIGHT = &H27 'arrow
Public Const VK_RMENU = &HA5
Public Const VK_RSHIFT = &HA1 'right shift
Public Const VK_SCROLL = &H91
Public Const VK_SELECT = &H29
Public Const VK_SEPARATOR = &H6C
Public Const VK_SHIFT = &H10 'shift
Public Const VK_SNAPSHOT = &H2C
Public Const VK_SPACE = &H20 'space bar
Public Const VK_SUBTRACT = &H6D ' -
Public Const VK_TAB = &H9 'Tab button
Public Const VK_UP = &H26 'arrow
Public Const VK_ZOOM = &HFB

Public Const VK_0 = &H30
Public Const VK_1 = &H31
Public Const VK_2 = &H32
Public Const VK_3 = &H33
Public Const VK_4 = &H34
Public Const VK_5 = &H35
Public Const VK_6 = &H36
Public Const VK_7 = &H37
Public Const VK_8 = &H38
Public Const VK_9 = &H39
Public Const VK_A = &H41
Public Const VK_B = &H42
Public Const VK_C = &H43
Public Const VK_D = &H44
Public Const VK_E = &H45
Public Const VK_F = &H46
Public Const VK_G = &H47
Public Const VK_H = &H48
Public Const VK_I = &H49
Public Const VK_J = &H4A
Public Const VK_K = &H4B
Public Const VK_L = &H4C
Public Const VK_M = &H4D
Public Const VK_N = &H4E
Public Const VK_O = &H4F
Public Const VK_P = &H50
Public Const VK_Q = &H51
Public Const VK_R = &H52
Public Const VK_S = &H53
Public Const VK_T = &H54
Public Const VK_U = &H55
Public Const VK_V = &H56
Public Const VK_W = &H57
Public Const VK_X = &H58
Public Const VK_Y = &H59
Public Const VK_Z = &H5A

Private Declare Function ToAscii Lib "user32" (ByVal uVirtKey As Integer, ByVal uScanCode As Integer, ByVal lpKeyState() As Byte, ByVal lpChar() As Byte, ByVal uFlags As Integer) As Integer

Public Function CharToKeyCode(theChar As String) As Integer

    ' Only single character strings allowed
    If Len(theChar) > 1 Then
        CharToKeyCode = 0
        Exit Function
    End If

    If IsNumeric(theChar) Then
        Dim theNum As Integer
        theNum = CInt(theChar)
        CharToKeyCode = theNum + &H30
        Exit Function
    End If

    ' Is this a letter?
    If UCase$(theChar) <> LCase$(theChar) Then
        Dim index As Integer
        index = Instr("abcdefghijklmnopqrstuvwxyz", LCase(theChar))
        CharToKeyCode = index + &H40
        Exit Function
    End If

    ' Punctuation or other
    Select Case theChar
        Case "+"
            CharToKeyCode = &H6B
        Case "."
            CharToKeyCode = &H6E
        Case "/"
            CharToKeyCode = &H6F ' /
        Case "*"
            CharToKeyCode = &H6A ' *
        Case "0"
            CharToKeyCode = &H60
        Case "1"
            CharToKeyCode = &H61
        Case "2"
            CharToKeyCode = &H62
        Case "3"
            CharToKeyCode = &H63
        Case "4"
            CharToKeyCode = &H64
        Case "5"
            CharToKeyCode = &H65
        Case "6"
            CharToKeyCode = &H66
        Case "7"
            CharToKeyCode = &H67
        Case "8"
            CharToKeyCode = &H68
        Case "9"
            CharToKeyCode = &H69
        Case " "
            CharToKeyCode = &H20 'space bar
        Case "-"
            CharToKeyCode = &H6D ' -
    End Select
End Function