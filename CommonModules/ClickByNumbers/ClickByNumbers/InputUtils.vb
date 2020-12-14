Imports System.Runtime.InteropServices
Imports System.Windows.Input
Imports System.Threading

Module InputUtils

    Private Const MOUSEEVENTF_ABSOLUTE = &H8000
    Private Const MOUSEEVENTF_LEFTDOWN = &H2
    Private Const MOUSEEVENTF_LEFTUP = &H4
    Private Const MOUSEEVENTF_MIDDLEDOWN = &H20
    Private Const MOUSEEVENTF_MIDDLEUP = &H40
    Private Const MOUSEEVENTF_MOVE = &H1
    Private Const MOUSEEVENTF_RIGHTDOWN = &H8
    Private Const MOUSEEVENTF_RIGHTUP = &H10
    Private Const MOUSEEVENTF_XDOWN = &H80
    Private Const MOUSEEVENTF_XUP = &H100
    Private Const MOUSEEVENTF_WHEEL = &H800
    Private Const MOUSEEVENTF_HWHEEL = &H1000
    Private Const MOUSEEVENTF_VIRTUALDESK = &H4000

    Private Const KEYEVENTF_KEYDOWN As Integer = &H0
    Private Const KEYEVENTF_KEYUP As Integer = &H2

    Private Const INPUT_MOUSE As Integer = 0
    Private Const INPUT_KEYBOARD As Integer = 1
    Private Const INPUT_HARDWARE As Integer = 2

    Private Structure MOUSEINPUT
        Public dx As Integer
        Public dy As Integer
        Public mouseData As Integer
        Public dwFlags As Integer
        Public time As Integer
        Public dwExtraInfo As IntPtr
    End Structure

    Private Structure KEYBDINPUT
        Public wVk As Short
        Public wScan As Short
        Public dwFlags As Integer
        Public time As Integer
        Public dwExtraInfo As IntPtr
    End Structure

    Private Structure HARDWAREINPUT
        Public uMsg As Integer
        Public wParamL As Short
        Public wParamH As Short
    End Structure

    <StructLayout(LayoutKind.Explicit)>
    Private Structure INPUT
        <FieldOffset(0)>
        Public type As Integer
        <FieldOffset(4)>
        Public mi As MOUSEINPUT
        <FieldOffset(4)>
        Public ki As KEYBDINPUT
        <FieldOffset(4)>
        Public hi As HARDWAREINPUT
    End Structure

    <DllImport("user32.dll", SetLastError:=True)>
    Private Function SendInput(<[In]()> ByVal nInput As UInt32,
                               <[In](), MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.Struct, SizeParamIndex:=0)> ByVal pInputs() As INPUT,
                               <[In]()> ByVal cbInput As Int32) As UInt32
    End Function

    '    <DllImport("user32.dll")>
    '    Private Function keybd_event(bVk As Byte, bScan As Byte, dwFlags As UInteger, dwExtraInfo As UIntPtr) As Integer
    '   End Function

    ' declare virtual key constant for the left Windows key,
    ' control key, shift key and alt key
    Const VK_LWIN As Long = &H5B
    Const VK_CTRL As Long = &H11
    Const VK_SHIFT As Long = &H10
    Const VK_ALT As Long = &H12

    Public Sub MouseLeftClick(clickPoint As Drawing.Point)

        Cursor.Position = clickPoint

        Dim theInput(0) As INPUT

        theInput(0).type = INPUT_MOUSE
        theInput(0).mi.dwFlags = MOUSEEVENTF_LEFTDOWN + MOUSEEVENTF_LEFTUP

        SendInput(1, theInput, Marshal.SizeOf(GetType(INPUT)))

    End Sub

    Public Sub MouseRightClick(clickPoint As Drawing.Point)

        Cursor.Position = clickPoint

        Dim theInput(0) As INPUT
        theInput(0).type = INPUT_MOUSE
        theInput(0).mi.dwFlags = MOUSEEVENTF_RIGHTDOWN + MOUSEEVENTF_RIGHTUP
        'theInput(0).mi.dwFlags = MOUSEEVENTF_RIGHTDOWN + MOUSEEVENTF_ABSOLUTE
        'theInput(0).mi.dx = clickPoint.X
        'theInput(0).mi.dy = clickPoint.Y

        'theInput(1).type = INPUT_MOUSE
        'theInput(1).mi.dwFlags = MOUSEEVENTF_RIGHTUP
        'theInput(1).mi.dx = clickPoint.X
        'theInput(1).mi.dy = clickPoint.Y

        SendInput(1, theInput, Marshal.SizeOf(GetType(INPUT)))

    End Sub

    Public Sub MouseDoubleClick(clickPoint As Drawing.Point)
        MouseLeftClick(clickPoint)
        MouseLeftClick(clickPoint)
    End Sub

    Public Sub MouseLeftDown(clickPoint As Drawing.Point)

        Cursor.Position = clickPoint

        Dim theInput(0) As INPUT
        theInput(0).type = INPUT_MOUSE
        'theInput(0).mi.dwFlags = MOUSEEVENTF_LEFTDOWN + MOUSEEVENTF_ABSOLUTE
        theInput(0).mi.dwFlags = MOUSEEVENTF_LEFTDOWN
        'theInput(0).mi.dx = clickPoint.X
        'theInput(0).mi.dy = clickPoint.Y

        SendInput(1, theInput, Marshal.SizeOf(GetType(INPUT)))

    End Sub

    Public Sub MouseLeftUp(clickPoint As Drawing.Point)

        Cursor.Position = clickPoint

        Dim theInput(0) As INPUT
        theInput(0).type = INPUT_MOUSE
        'theInput(0).mi.dwFlags = MOUSEEVENTF_LEFTUP + MOUSEEVENTF_ABSOLUTE
        theInput(0).mi.dwFlags = MOUSEEVENTF_LEFTUP
        'theInput(0).mi.dx = clickPoint.X
        'theInput(0).mi.dy = clickPoint.Y

        SendInput(1, theInput, Marshal.SizeOf(GetType(INPUT)))

    End Sub

    Public Sub PressKey(ByVal key As Key)
        Dim theInput(1) As INPUT

        ' press the key
        theInput(0).type = INPUT_KEYBOARD
        theInput(0).ki.wVk = KeyInterop.VirtualKeyFromKey(key)
        theInput(0).ki.dwFlags = KEYEVENTF_KEYDOWN

        ' release the key
        theInput(1).type = INPUT_KEYBOARD
        theInput(1).ki.wVk = KeyInterop.VirtualKeyFromKey(key)
        theInput(1).ki.dwFlags = KEYEVENTF_KEYUP

        SendInput(2, theInput, Marshal.SizeOf(GetType(INPUT)))

    End Sub

    Public Sub KeyDown(ByVal key As Key)
        Dim theInput(0) As INPUT

        ' press the key
        theInput(0).type = INPUT_KEYBOARD
        theInput(0).ki.wVk = KeyInterop.VirtualKeyFromKey(key)
        theInput(0).ki.dwFlags = KEYEVENTF_KEYDOWN

        SendInput(1, theInput, Marshal.SizeOf(GetType(INPUT)))

        Thread.Sleep(100)

    End Sub

    Public Sub KeyUp(ByVal key As Key)
        Dim theInput(0) As INPUT

        ' press the key
        theInput(0).type = INPUT_KEYBOARD
        theInput(0).ki.wVk = KeyInterop.VirtualKeyFromKey(key)
        theInput(0).ki.dwFlags = KEYEVENTF_KEYUP

        SendInput(1, theInput, Marshal.SizeOf(GetType(INPUT)))

        Thread.Sleep(100)
    End Sub

End Module
