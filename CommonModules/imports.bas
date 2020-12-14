'option explicit

public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public type POINTAPI
    x As Long
    y As Long
End type

Public Type MONITORINFO
  size as long
  monitorRect as RECT
  workRect as RECT
  flags as long
End type


Public Const MONITOR_DEFAULTTONULL = 0
Public Const MONITOR_DEFAULTTOPRIMARY = 1
Public Const MONITOR_DEFAULTTONEAREST = 2

Public Const HORZRES = 8
Public Const VERTRES = 10

Public Declare Function MonitorFromWindow Lib "user32" (ByVal hwnd As Long, ByVal dwFlags As Long) As Long
Public Declare Function GetMonitorInfo Lib "user32" (ByVal hMonitor As Long, ByRef lpmi As MONITORINFO) As Boolean
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, ByRef lpRect As RECT) As Boolean
Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, ByRef lpRECT As RECT) As Integer
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal bRepaint As Boolean) As Boolean
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal prmstrString As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long


Declare Function GetCursorPos Lib "user32" Alias "GetCursorPos" (lpPoint As POINTAPI) As Long
Declare Function ClientToScreen Lib "user32.dll" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Boolean



