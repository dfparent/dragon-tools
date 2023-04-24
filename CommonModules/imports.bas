'#Language "WWB.NET"
'option explicit

Public Enum KeyModifier
    none
    shift
    control
    alt
    win
End Enum

Public Structure RECT
    Public Left As Integer
    Public Top As Integer
    Public Right As Integer
    Public Bottom As Integer
End Structure


Public Structure POINTAPI
    Public x As Integer
    Public y As Integer
End Structure

Public Structure MONITORINFO
    Public size As Integer
    Public monitorRect As RECT
    Public workRect As RECT
    Public flags As Integer
End Structure


Public Const MONITOR_DEFAULTTONULL = 0
Public Const MONITOR_DEFAULTTOPRIMARY = 1
Public Const MONITOR_DEFAULTTONEAREST = 2

Public Const HORZRES = 8
Public Const VERTRES = 10

Public Declare Function MonitorFromWindow Lib "user32" (ByVal hwnd As System.IntPtr, ByVal dwFlags As Integer) As System.IntPtr
'<DllImport("user32.dll")>
'Public Shared Function MonitorFromWindow(ByVal hwnd As Long, ByVal dwFlags As Long) As Long
'End Function

Public Declare Function GetMonitorInfo Lib "user32" (ByVal hMonitor As System.IntPtr, ByRef lpmi As MONITORINFO) As Boolean
'<DllImport("user32.dll")>
'Public Shared Function GetMonitorInfo(ByVal hMonitor As Long, ByRef lpmi As MONITORINFO) As Boolean
'End Function

Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As System.IntPtr, ByRef lpRect As RECT) As Boolean
'<DllImport("user32.dll")>
'Public Shared Function GetWindowRect(ByVal hWnd As IntPtr, ByRef lpRect As RECT) As Boolean
'End Function

Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As System.IntPtr, ByRef lpRECT As RECT) As Integer
'<DllImport("user32.dll")>
'Public Shared Function GetClientRect(ByVal hWnd As Long, ByRef lpRECT As RECT) As Integer
'End Function

Public Declare Function GetForegroundWindow Lib "user32" () As System.IntPtr
'<DllImport("user32.dll")>
'Public Shared Function GetForegroundWindow() As IntPtr
'End Function

Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As System.IntPtr) As Boolean
'<DllImport("user32.dll")>
'Public Shared Function SetForegroundWindow(ByVal hwnd As Long) As Long
'End Function

Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As System.IntPtr, ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal bRepaint As Boolean) As Boolean
'<DllImport("user32.dll")>
'Public Shared Function MoveWindow(ByVal hWnd As Long, ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal bRepaint As Boolean) As Boolean
'End Function

Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As System.IntPtr, ByVal nCmdShow As Integer) As Boolean
'<DllImport("user32.dll")>
'Public Shared Function ShowWindow(ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
'End Function

Public Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hWnd As System.IntPtr, ByVal prmstrString As String, ByVal nMaxCount As Integer) As Integer
'<DllImport("user32.dll")>
'Public Shared Function GetWindowText(ByVal hWnd As Integer, ByVal prmstrString As System.Text.StringBuilder, ByVal nMaxCount As Integer) As Integer
'End Function

Public Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hWnd As System.IntPtr) As Integer
'<DllImport("user32.dll")>
'Public Shared Function GetWindowTextLength(ByVal hWnd As Long) As Long
'End Function

Public Declare Function GetCursorPos Lib "user32" Alias "GetCursorPos" (ByRef lpPoint As POINTAPI) As Integer
'<DllImport("user32.dll")>
'Public Shared Function GetCursorPos(ByRef lpPoint As POINTAPI) As Long
'End Function

Public Declare Function ClientToScreen Lib "user32.dll" (ByVal hWnd As System.IntPtr, ByRef lpPoint As POINTAPI) As Boolean
'<DllImport("user32.dll")>
'Public Shared Function ClientToScreen(ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Boolean
'End Function

'Public Delegate Function EnumWindowsDelegateCallBack(ByVal hwnd As Integer, ByVal lParam As Integer) As Boolean
'Public Declare Function EnumWindows Lib "user32" (ByVal x As EnumWindowsDelegateCallBack, ByVal y As Integer) As Integer
'Public Declare Function GetClassName Lib "user32" (ByVal hWnd As System.IntPtr, ByVal lpClassName As System.Text.StringBuilderQuestion, ByVal nMaxCount As Integer) As Integer

